from __future__ import annotations

import pickle
import re
import sys
from dataclasses import dataclass, replace
from pathlib import Path
from typing import Callable

import numpy as np
import pandas as pd
import joblib
from openpyxl.chart import Reference, ScatterChart, Series
from openpyxl.styles import PatternFill


FIRST_ISOTOPE_SPACING_DA = 1.00335483507
MASS_TOLERANCE_DA = 0.010
HIGHER_ISOTOPE_WINDOW_DA = 0.020
RT_TOLERANCE_MIN = 0.020
TC_TOLERANCE = 0.060
ISOTOPE_ORDERS = (1, 2, 3, 4)
NITROGEN_MASS_DA = 28.0134
PROTON_MASS_DA = 1.007276466621
LINEAR_CANDIDATE_PROBABILITY = 0.50
LINEAR_HIGH_CONFIDENCE_PROBABILITY = 0.80
LINEAR_MEDIUM_CONFIDENCE_PROBABILITY = 0.65
DEFAULT_DATABASE_FILENAME = "Chemical List PFAS.xlsx"

PLOT_CONFIDENCE_LABELS = np.array(
    [
        "Non-PFAS / rejected",
        "Possible PFAS",
        "Likely PFAS",
        "Most likely PFAS",
    ],
    dtype=object,
)
PLOT_CONFIDENCE_COLORS = np.array(
    ["A6A6A6", "2CA02C", "FF8C00", "D62728"],
    dtype=object,
)

# Nominal isotope shifts and natural abundances relative to the light
# monoisotopic isotope. Elements absent from this map are reported as
# unsupported rather than assigned a misleading theoretical distribution.
ISOTOPE_ABUNDANCES = {
    "H": {0: 0.999885, 1: 0.000115},
    "C": {0: 0.9893, 1: 0.0107},
    "N": {0: 0.99636, 1: 0.00364},
    "O": {0: 0.99757, 1: 0.00038, 2: 0.00205},
    "Si": {0: 0.92223, 1: 0.04685, 2: 0.03092},
    "S": {0: 0.9499, 1: 0.0075, 2: 0.0425, 4: 0.0001},
    "Cl": {0: 0.7576, 2: 0.2424},
    "Br": {0: 0.5069, 2: 0.4931},
    "F": {0: 1.0},
    "Na": {0: 1.0},
    "Al": {0: 1.0},
    "P": {0: 1.0},
    "I": {0: 1.0},
    "Au": {0: 1.0},
}

SUPPORTED_ION_MODES = {"POS", "NEG"}
SUPPORTED_SCREENING_STRATEGIES = {
    "model_only",
    "linear_only",
    "intersection",
    "union",
}
LEGACY_RF_MC_MODEL = "legacy_rf_mc"

MODEL_DESCRIPTIONS = {
    LEGACY_RF_MC_MODEL: (
        "RF-MC riMp1 +/-20% CCS +/-3%: uses m/z, CCS, riMp1, and derived features only. "
        "It does not use riMp2-riMp4 for prediction."
    ),
}

LEGACY_MODEL_THRESHOLDS = {
    "POS": {"candidate": 0.58375, "high_confidence": 0.955},
    "NEG": {"candidate": 0.5725, "high_confidence": 0.970},
}

FEATURE_COLS = [
    "mz_percentile",
    "mz",
    "CCS",
    "riMp1",
    "CCS_over_m",
    "CCS_over_Mp1",
    "M_over_Mp1",
    "CCS_over_sqrtm",
    "CCS_over_m23",
]

COLUMN_ALIASES = {
    "mz": [
        "mz",
        "m/z",
        "PrecursorMz",
        "3D_m_z",
        "3D_m_z_uncal",
        "mz_M+H",
        "mz_M-H",
    ],
    "observed_mz": [
        "observed_mz",
        "observed m/z",
        "measured_mz",
        "measured m/z",
        "mz_obs",
        "m/z observed",
        "m/z",
        "mz",
    ],
    "exact_mz": [
        "exact_mz",
        "exact m/z",
        "reference_mz",
        "reference m/z",
        "known_mz",
        "known m/z",
        "theoretical_mz",
        "theoretical m/z",
    ],
    "ccs": [
        "CCS",
        "CCS(Progensis)",
        "Reference CCS",
        "Standard CCS",
        "PrecursorCCS",
        "CCS_correct_linear",
        "CCS_M+H_ALLCCS2",
        "CCS_M-H_ALLCCS2",
        "CCS_M+H",
        "CCS_M-H",
    ],
    "tc": [
        "tc",
        "t_c",
        "corrected drift time",
        "dt",
        "drift time",
        "3D_omegaD",
        "3D_omegaD_corr_3",
    ],
    "charge": ["charge", "z", "PrecursorCharge", "3D_z"],
    "rt": ["rt", "RT", "retention time", "Retention time (min)", "PrecursorRT", "3D_RetTime"],
    "intensity": ["intensity", "Intensity", "intb", "M", "3D_Intensity", "3D_Mobility_Mass"],
    "rimp1": [
        "first_isotopic_peak_ratio",
        "riMp1",
        "rimp1",
        "Mp1int",
        "Ap1intNorm",
    ],
    "m_intensity": ["M", "intensity", "Intensity", "intb", "3D_Intensity", "3D_Mobility_Mass"],
    "mp1_intensity": ["M+1", "intensity_mp1", "M+1_intensity"],
}


class ProcessingError(RuntimeError):
    """Raised when an input table cannot be processed safely."""


@dataclass(frozen=True)
class ProcessingConfig:
    input_path: Path
    output_path: Path
    sheet_name: str | int = 0
    ion_mode: str = "POS"
    rimp1_method: str = "auto"
    mass_calibration_method: str = "none"
    lock_mass_value: float | None = None
    lock_mass_tolerance_ppm: float = 1000.0
    first_calibration_method: str = "cims_reduced_ccs"
    second_calibration_method: str = "none"
    mz_column: str | None = None
    ccs_column: str | None = None
    tc_column: str | None = None
    charge_column: str | None = None
    default_charge: int = 1
    rt_column: str | None = None
    intensity_column: str | None = None
    rimp1_column: str | None = None
    m_intensity_column: str | None = None
    mp1_intensity_column: str | None = None
    standards_path: Path | None = None
    standards_sheet_name: str | int = 0
    standards_tc_column: str | None = None
    standards_mz_column: str | None = None
    standards_ccs_column: str | None = None
    standards_charge_column: str | None = None
    mass_standards_path: Path | None = None
    mass_standards_sheet_name: str | int = 0
    mass_observed_mz_column: str | None = None
    mass_exact_mz_column: str | None = None
    prediction_model: str = "none"
    screening_strategy: str = "model_only"
    rimp1_uncertainty_percent: float = 20.0
    ccs_uncertainty_percent: float = 3.0
    minimum_m_intensity: float = 0.0
    priority_top_n: int = 200
    library_matching_enabled: bool = False
    database_path: Path | None = None
    database_sheet_name: str | int = 0
    database_mass_tolerance_ppm: float = 5.0
    database_ccs_tolerance_percent: float = 5.0
    isotope_ratio_matching: bool = False
    isotope_ratio_tolerance_percent: float = 20.0
    mc_iterations: int = 200
    blank_subtraction_enabled: bool = False
    blank_paths: tuple[Path, ...] = ()
    blank_sheet_name: str | int = 0
    blank_mz_tolerance_ppm: float = 5.0
    blank_rt_tolerance_min: float = 0.10
    blank_ccs_tolerance_percent: float = 3.0
    blank_sigma_multiplier: float = 3.0


@dataclass(frozen=True)
class ProcessingSummary:
    input_rows: int
    predictable_rows: int
    level_counts: dict[str, int]
    rimp1_method: str
    first_calibration_formula: str
    first_calibration_r2: float | None
    second_calibration_applied: bool
    prediction_model: str
    prediction_model_path: Path | None
    output_path: Path


@dataclass(frozen=True)
class BatchProcessingSummary:
    sample_count: int
    succeeded_count: int
    failed_count: int
    output_directory: Path
    summary_path: Path
    results: tuple[ProcessingSummary, ...]


def _resource_root() -> Path:
    if getattr(sys, "frozen", False):
        return Path(getattr(sys, "_MEIPASS"))
    return Path(__file__).resolve().parents[1]


def model_path(mode: str, prediction_model: str = LEGACY_RF_MC_MODEL) -> Path:
    mode = mode.upper()
    prediction_model = prediction_model.strip().lower()
    registry = _model_registry()
    record = registry.get(prediction_model)
    if record is None:
        available = ", ".join(item["label"] for item in registry.values()) or "none"
        raise ProcessingError(
            f"Prediction model is not available: {prediction_model}. "
            f"Available installed models: {available}"
        )
    paths = record["paths"]
    path = paths.get(mode) if isinstance(paths, dict) else None
    if isinstance(path, Path) and path.exists():
        return path
    raise ProcessingError(
        f"Model file for {prediction_model} {mode} mode was not found."
    )


def default_database_path() -> Path:
    """Return the bundled PFAS database path."""
    return _resource_root() / DEFAULT_DATABASE_FILENAME


def _models_root() -> Path:
    return _resource_root() / "models"


def _resolve_model_file(value: object) -> Path:
    text = str(value).strip()
    if not text:
        raise ProcessingError("Model manifest contains a blank model file path.")
    path = Path(text)
    if not path.is_absolute():
        path = _models_root() / path
    return path


def _split_feature_columns(value: object) -> list[str]:
    if value is None or pd.isna(value):
        return list(FEATURE_COLS)
    text = str(value).strip()
    if not text:
        return list(FEATURE_COLS)
    return [column.strip() for column in re.split(r"[;,]", text) if column.strip()]


def _model_registry() -> dict[str, dict[str, object]]:
    registry: dict[str, dict[str, object]] = {}
    legacy_paths: dict[str, Path] = {}
    for mode in SUPPORTED_ION_MODES:
        joblib_filename = f"rf_sklearn_bundle_{mode}.joblib"
        pickle_filename = f"rf_sklearn_bundle_{mode}.pkl"
        candidates = [
            _models_root() / joblib_filename,
            _models_root() / pickle_filename,
            _resource_root()
            / "outputs_20260604_reanalysis_8"
            / "models"
            / pickle_filename,
        ]
        legacy_path = next((candidate for candidate in candidates if candidate.exists()), None)
        if legacy_path is not None:
            legacy_paths[mode] = legacy_path
    if SUPPORTED_ION_MODES.issubset(legacy_paths):
        registry[LEGACY_RF_MC_MODEL] = {
            "id": LEGACY_RF_MC_MODEL,
            "label": "RF-MC riMp1 +/-20% CCS +/-3%",
            "description": MODEL_DESCRIPTIONS[LEGACY_RF_MC_MODEL],
            "paths": legacy_paths,
            "feature_columns": list(FEATURE_COLS),
            "rimp1_uncertainty_percent": 20.0,
            "ccs_uncertainty_percent": 3.0,
            "thresholds": LEGACY_MODEL_THRESHOLDS,
        }

    manifest_path = _models_root() / "model_manifest.csv"
    if not manifest_path.exists():
        return registry
    try:
        manifest = pd.read_csv(manifest_path)
    except Exception as exc:
        raise ProcessingError(f"Cannot read model manifest: {manifest_path}: {exc}") from exc

    required = {"model_id", "label", "pos_path", "neg_path"}
    missing = sorted(required - set(manifest.columns))
    if missing:
        raise ProcessingError(
            "Model manifest is missing required columns: " + ", ".join(missing)
        )
    for row_index, row in manifest.iterrows():
        model_id = str(row["model_id"]).strip().lower()
        if not model_id:
            raise ProcessingError(f"Model manifest row {row_index + 2} has no model_id.")
        paths = {
            "POS": _resolve_model_file(row["pos_path"]),
            "NEG": _resolve_model_file(row["neg_path"]),
        }
        if not all(path.exists() for path in paths.values()):
            continue
        thresholds: dict[str, dict[str, float]] = {}
        for mode, prefix in [("POS", "pos"), ("NEG", "neg")]:
            candidate = pd.to_numeric(
                row.get(f"{prefix}_candidate_threshold", np.nan),
                errors="coerce",
            )
            high_confidence = pd.to_numeric(
                row.get(f"{prefix}_high_confidence_threshold", np.nan),
                errors="coerce",
            )
            if pd.notna(candidate) and pd.notna(high_confidence):
                thresholds[mode] = {
                    "candidate": float(candidate),
                    "high_confidence": float(high_confidence),
                }
        registry[model_id] = {
            "id": model_id,
            "label": str(row["label"]).strip() or model_id,
            "description": str(row.get("description", "")).strip(),
            "paths": paths,
            "feature_columns": _split_feature_columns(row.get("feature_cols")),
            "rimp1_uncertainty_percent": row.get("rimp1_uncertainty_percent", np.nan),
            "ccs_uncertainty_percent": row.get("ccs_uncertainty_percent", np.nan),
            "thresholds": thresholds,
        }
    return registry


def available_prediction_models() -> list[dict[str, object]]:
    """Return installed prediction models that the app can actually execute."""
    return list(_model_registry().values())


def load_bundle(mode: str, prediction_model: str = LEGACY_RF_MC_MODEL) -> dict:
    mode = mode.upper()
    if mode not in SUPPORTED_ION_MODES:
        raise ProcessingError(f"Unsupported ion mode: {mode}")
    path = model_path(mode, prediction_model)
    if path.suffix.lower() == ".joblib":
        return joblib.load(path)
    with path.open("rb") as handle:
        return pickle.load(handle)


def list_excel_sheets(path: Path) -> list[str]:
    try:
        return pd.ExcelFile(path).sheet_names
    except Exception as exc:
        raise ProcessingError(f"Cannot read Excel workbook: {exc}") from exc


def read_input_table(path: Path, sheet_name: str | int = 0) -> pd.DataFrame:
    try:
        if path.suffix.lower() == ".csv":
            return pd.read_csv(path)
        return pd.read_excel(path, sheet_name=sheet_name)
    except Exception as exc:
        raise ProcessingError(f"Cannot read input table: {exc}") from exc


def _normalized(value: object) -> str:
    return "".join(str(value).strip().lower().split())


def detect_column(columns: list[object], field: str, mode: str | None = None) -> str | None:
    names = [str(column) for column in columns]
    normalized = {_normalized(column): column for column in names}
    aliases = list(COLUMN_ALIASES[field])

    if field == "mz" and mode:
        aliases = (["mz_M+H"] if mode == "POS" else ["mz_M-H"]) + aliases
    if field == "ccs" and mode:
        aliases = (
            ["CCS_M+H_ALLCCS2", "CCS_M+H"] if mode == "POS"
            else ["CCS_M-H_ALLCCS2", "CCS_M-H"]
        ) + aliases

    for alias in aliases:
        match = normalized.get(_normalized(alias))
        if match is not None:
            return match
    return None


def suggested_columns(df: pd.DataFrame, mode: str) -> dict[str, str | None]:
    columns = list(df.columns)
    return {field: detect_column(columns, field, mode) for field in COLUMN_ALIASES}


def _selected_or_detected(
    df: pd.DataFrame,
    selected: str | None,
    field: str,
    mode: str,
    required: bool = True,
) -> str | None:
    if selected == "__autodetected__":
        selected = None
    if selected and selected in df.columns:
        return selected
    detected = detect_column(list(df.columns), field, mode)
    if required and detected is None:
        raise ProcessingError(
            f"Could not identify the {field} column. Select it in Advanced column mapping."
        )
    return detected


def _numeric(series: pd.Series) -> pd.Series:
    return pd.to_numeric(series, errors="coerce")


def parse_molecular_formula(formula: object) -> dict[str, int]:
    """Parse a simple molecular formula such as C8H5ClF16O3S."""
    text = str(formula).strip()
    if not text or text.lower() == "nan":
        raise ValueError("Molecular formula is blank.")
    tokens = re.findall(r"([A-Z][a-z]?)(\d*)", text)
    if not tokens or "".join(element + count for element, count in tokens) != text:
        raise ValueError(f"Unsupported molecular formula format: {text}")
    composition: dict[str, int] = {}
    for element, count_text in tokens:
        count = int(count_text) if count_text else 1
        composition[element] = composition.get(element, 0) + count
    return composition


def theoretical_isotope_ratios(
    formula: object,
    hydrogen_adjustment: int = 0,
    max_order: int = 4,
) -> tuple[dict[int, float], str]:
    """Calculate nominal M+n/M ratios from natural isotope abundances."""
    try:
        composition = parse_molecular_formula(formula)
    except ValueError as exc:
        return {}, str(exc)

    if hydrogen_adjustment:
        composition["H"] = composition.get("H", 0) + hydrogen_adjustment
        if composition["H"] < 0:
            return {}, "Formula does not contain enough hydrogen for [M-H]-."
        if composition["H"] == 0:
            composition.pop("H")

    unsupported = sorted(set(composition) - set(ISOTOPE_ABUNDANCES))
    if unsupported:
        return {}, "Unsupported isotope elements: " + ", ".join(unsupported)

    distribution = np.zeros(max_order + 1, dtype=float)
    distribution[0] = 1.0
    for element, count in composition.items():
        atom_profile = ISOTOPE_ABUNDANCES[element]
        for _ in range(count):
            updated = np.zeros(max_order + 1, dtype=float)
            for current_order, current_probability in enumerate(distribution):
                if current_probability == 0:
                    continue
                for shift, abundance in atom_profile.items():
                    if current_order + shift <= max_order:
                        updated[current_order + shift] += (
                            current_probability * abundance
                        )
            distribution = updated

    if distribution[0] <= 0:
        return {}, "Monoisotopic abundance is zero under the configured profile."
    ratios = {
        order: float(100.0 * distribution[order] / distribution[0])
        for order in range(1, max_order + 1)
    }
    return ratios, "calculated"


def _database_column(
    data: pd.DataFrame,
    aliases: list[str],
    required: bool = True,
) -> str | None:
    normalized = {_normalized(column): str(column) for column in data.columns}
    for alias in aliases:
        match = normalized.get(_normalized(alias))
        if match is not None:
            return match
    if required:
        raise ProcessingError(
            "PFAS database is missing a required column. Expected one of: "
            + ", ".join(aliases)
        )
    return None


def prepare_pfas_database(data: pd.DataFrame) -> pd.DataFrame:
    """Standardize the Chemical List PFAS database schema."""
    mass_column = _database_column(
        data,
        ["MONOISOTOPIC MASS", "ExactMonoisotopicMass", "MonoisotopicMass"],
    )
    formula_column = _database_column(
        data,
        ["MOLECULAR FORMULA", "Molecula_formula2", "Molecula_formula"],
    )
    name_column = _database_column(
        data,
        ["PREFERRED NAME", "Chemical Name", "name"],
        required=False,
    )
    smiles_column = _database_column(
        data,
        ["SMILES", "SMILES_conical", "SMILES_canon", "QSAR Ready SMILES"],
        required=False,
    )
    ccs_column = _database_column(
        data,
        [
            "CCS",
            "Reference CCS",
            "Standard CCS",
            "Predicted CCS",
            "library_CCS",
            "CCS_N2",
            "CCS (N2)",
            "CCS_M+H",
            "CCS_M-H",
        ],
        required=False,
    )
    dtxsid_column = _database_column(data, ["DTXSID"], required=False)
    cas_column = _database_column(data, ["CASRN", "CAS.."], required=False)
    inchikey_column = _database_column(data, ["INCHIKEY", "InChIKey"], required=False)

    database = pd.DataFrame(
        {
            "database_row": np.arange(2, len(data) + 2),
            "neutral_monoisotopic_mass": _numeric(data[mass_column]),
            "formula": data[formula_column].astype("string"),
            "name": (
                data[name_column].astype("string")
                if name_column
                else pd.Series(pd.NA, index=data.index, dtype="string")
            ),
            "smiles": (
                data[smiles_column].astype("string")
                if smiles_column
                else pd.Series(pd.NA, index=data.index, dtype="string")
            ),
            "library_CCS": (
                _numeric(data[ccs_column])
                if ccs_column
                else pd.Series(np.nan, index=data.index, dtype=float)
            ),
            "dtxsid": (
                data[dtxsid_column].astype("string")
                if dtxsid_column
                else pd.Series(pd.NA, index=data.index, dtype="string")
            ),
            "casrn": (
                data[cas_column].astype("string")
                if cas_column
                else pd.Series(pd.NA, index=data.index, dtype="string")
            ),
            "inchikey": (
                data[inchikey_column].astype("string")
                if inchikey_column
                else pd.Series(pd.NA, index=data.index, dtype="string")
            ),
        }
    )
    database = database.dropna(subset=["neutral_monoisotopic_mass"]).copy()
    database = database[database["neutral_monoisotopic_mass"] > 0].copy()

    theory_rows: list[dict[str, object]] = []
    for formula in database["formula"]:
        ratios, status = theoretical_isotope_ratios(formula)
        theory_rows.append(
            {
                "theoretical_rMp1": ratios.get(1, np.nan),
                "theoretical_rMp2": ratios.get(2, np.nan),
                "theoretical_rMp3": ratios.get(3, np.nan),
                "theoretical_rMp4": ratios.get(4, np.nan),
                "theoretical_isotope_status": status,
            }
        )
    database = pd.concat(
        [database.reset_index(drop=True), pd.DataFrame(theory_rows)],
        axis=1,
    )
    return database.sort_values(
        "neutral_monoisotopic_mass", kind="mergesort"
    ).reset_index(drop=True)


def prioritize_pfas_peaks(
    data: pd.DataFrame,
    minimum_intensity: float,
    top_n: int,
) -> pd.DataFrame:
    """Select the most intense final-screen-positive monoisotopic candidates."""
    candidate = (
        _numeric(data["PFAS_final_candidate"]) == 1
        if "PFAS_final_candidate" in data
        else _numeric(data["PFAS_level"]) >= 1
    )
    intensity = _numeric(data["PFAS_M_intensity"])
    isotope_role = (
        data["PFAS_isotope_role"].astype(str)
        if "PFAS_isotope_role" in data
        else pd.Series("M", index=data.index)
    )
    eligible = (
        candidate
        & (isotope_role == "M")
        & intensity.notna()
        & (intensity >= minimum_intensity)
    )
    prioritized = data.loc[eligible].copy()
    prioritized = prioritized.sort_values(
        ["PFAS_M_intensity", "PFAS_prob_mean"],
        ascending=[False, False],
        kind="mergesort",
    )
    if top_n > 0:
        prioritized = prioritized.head(top_n)
    prioritized.insert(
        0,
        "PFAS_intensity_rank",
        np.arange(1, len(prioritized) + 1),
    )
    return prioritized


def match_pfas_database(
    peaks: pd.DataFrame,
    database: pd.DataFrame,
    mode: str,
    tolerance_ppm: float,
    ccs_tolerance_percent: float,
    isotope_ratio_matching: bool,
    isotope_ratio_tolerance_percent: float,
) -> pd.DataFrame:
    """Match prioritized peaks to a PFAS library by m/z, optional CCS, and optional isotope ratios."""
    if peaks.empty or database.empty:
        return pd.DataFrame()

    mode = mode.upper()
    if mode == "POS":
        hypotheses = [
            ("[M+H]+", -PROTON_MASS_DA, 0),
            ("M+", 0.0, 0),
        ]
    elif mode == "NEG":
        hypotheses = [
            ("[M-H]-", PROTON_MASS_DA, -1),
            ("M-", 0.0, 0),
        ]
    else:
        raise ProcessingError(f"Unsupported ion mode for database matching: {mode}")

    ccs_tolerance_percent = abs(float(ccs_tolerance_percent))
    isotope_ratio_tolerance_percent = abs(float(isotope_ratio_tolerance_percent))
    db_mass = database["neutral_monoisotopic_mass"].to_numpy(float)
    matches: list[dict[str, object]] = []
    for peak_index, peak in peaks.iterrows():
        observed_mz = float(peak["PFAS_mz_used"])
        experimental_ccs = peak.get("PFAS_CCS_stage1", np.nan)
        experimental_ccs = (
            float(experimental_ccs) if pd.notna(experimental_ccs) else np.nan
        )
        for ion_hypothesis, neutral_adjustment, hydrogen_adjustment in hypotheses:
            target_neutral_mass = observed_mz + neutral_adjustment
            mass_window = target_neutral_mass * tolerance_ppm / 1e6
            start = int(
                np.searchsorted(
                    db_mass,
                    target_neutral_mass - mass_window,
                    side="left",
                )
            )
            stop = int(
                np.searchsorted(
                    db_mass,
                    target_neutral_mass + mass_window,
                    side="right",
                )
            )
            for database_index in range(start, stop):
                candidate = database.iloc[database_index]
                neutral_mass = float(candidate["neutral_monoisotopic_mass"])
                theoretical_mz = (
                    neutral_mass + PROTON_MASS_DA
                    if ion_hypothesis == "[M+H]+"
                    else neutral_mass - PROTON_MASS_DA
                    if ion_hypothesis == "[M-H]-"
                    else neutral_mass
                )
                mass_error_da = target_neutral_mass - neutral_mass
                ppm_error = mass_error_da / neutral_mass * 1e6
                library_ccs = candidate.get("library_CCS", np.nan)
                library_ccs = float(library_ccs) if pd.notna(library_ccs) else np.nan
                if pd.notna(library_ccs) and pd.notna(experimental_ccs):
                    ccs_error = experimental_ccs - library_ccs
                    ccs_error_percent = ccs_error / library_ccs * 100.0
                    ccs_match = abs(ccs_error_percent) <= ccs_tolerance_percent
                    ccs_match_status = (
                        "matched"
                        if ccs_match
                        else f"outside +/-{ccs_tolerance_percent:g}% CCS tolerance"
                    )
                    if not ccs_match:
                        continue
                else:
                    ccs_error = np.nan
                    ccs_error_percent = np.nan
                    ccs_match = pd.NA
                    ccs_match_status = "library or experimental CCS not available"

                ratios, theory_status = theoretical_isotope_ratios(
                    candidate["formula"],
                    hydrogen_adjustment=hydrogen_adjustment,
                )
                if not ratios:
                    ratios = {
                        order: candidate.get(f"theoretical_rMp{order}", np.nan)
                        for order in ISOTOPE_ORDERS
                    }
                    theory_status = str(
                        candidate.get("theoretical_isotope_status", theory_status)
                    )

                isotope_available = 0
                isotope_passed = 0
                isotope_failures: list[str] = []
                isotope_ci: dict[str, object] = {}
                for order in ISOTOPE_ORDERS:
                    theory_value = ratios.get(order, np.nan)
                    experimental_value = peak.get(
                        f"PFAS_riMp{order}",
                        peak.get(f"PFAS_rMp{order}", np.nan),
                    )
                    if pd.notna(theory_value):
                        ci_low = float(theory_value) * (
                            1.0 - isotope_ratio_tolerance_percent / 100.0
                        )
                        ci_high = float(theory_value) * (
                            1.0 + isotope_ratio_tolerance_percent / 100.0
                        )
                    else:
                        ci_low = np.nan
                        ci_high = np.nan
                    within_ci = pd.NA
                    if pd.notna(experimental_value) and pd.notna(theory_value):
                        isotope_available += 1
                        within_ci = ci_low <= float(experimental_value) <= ci_high
                        if within_ci:
                            isotope_passed += 1
                        else:
                            isotope_failures.append(f"rMp{order}")
                    isotope_ci[f"theoretical_rMp{order}"] = theory_value
                    isotope_ci[f"theoretical_rMp{order}_CI_low"] = ci_low
                    isotope_ci[f"theoretical_rMp{order}_CI_high"] = ci_high
                    isotope_ci[f"rMp{order}_within_CI"] = within_ci
                    isotope_ci[f"delta_rMp{order}"] = (
                        float(experimental_value) - float(theory_value)
                        if pd.notna(experimental_value) and pd.notna(theory_value)
                        else np.nan
                    )
                isotope_match = (
                    isotope_available > 0 and isotope_available == isotope_passed
                )
                if isotope_ratio_matching and not isotope_match:
                    continue
                if not isotope_ratio_matching:
                    isotope_status = "not requested"
                elif isotope_match:
                    isotope_status = (
                        f"matched {isotope_passed}/{isotope_available} available ratios"
                    )
                elif isotope_available == 0:
                    isotope_status = "no comparable experimental/theoretical ratios"
                else:
                    isotope_status = "outside CI: " + ", ".join(isotope_failures)

                row = {
                    "PFAS_peak_output_index": int(peak_index),
                    "PFAS_intensity_rank": peak.get("PFAS_intensity_rank", np.nan),
                    "PFAS_ion_mode": mode,
                    "PFAS_ion_hypothesis": ion_hypothesis,
                    "PFAS_observed_mz": observed_mz,
                    "PFAS_experimental_mz": observed_mz,
                    "PFAS_experimental_CCS": experimental_ccs,
                    "PFAS_library_CCS": library_ccs,
                    "PFAS_mass_shift_to_library_Da": neutral_adjustment,
                    "PFAS_target_neutral_mass": target_neutral_mass,
                    "PFAS_theoretical_mz": theoretical_mz,
                    "PFAS_mass_error_Da": mass_error_da,
                    "PFAS_mass_error_ppm": ppm_error,
                    "PFAS_CCS_error": ccs_error,
                    "PFAS_CCS_error_percent": ccs_error_percent,
                    "PFAS_CCS_match": ccs_match,
                    "PFAS_CCS_match_status": ccs_match_status,
                    "PFAS_M_intensity": peak.get("PFAS_M_intensity", np.nan),
                    "PFAS_level": peak.get("PFAS_level", np.nan),
                    "PFAS_level_name": peak.get("PFAS_level_name", ""),
                    "PFAS_prob_mean": peak.get("PFAS_prob_mean", np.nan),
                    "PFAS_prob_p05": peak.get("PFAS_prob_p05", np.nan),
                    "experimental_riMp1": peak.get(
                        "PFAS_riMp1", peak.get("PFAS_rMp1", np.nan)
                    ),
                    "experimental_riMp2": peak.get(
                        "PFAS_riMp2", peak.get("PFAS_rMp2", np.nan)
                    ),
                    "experimental_riMp3": peak.get(
                        "PFAS_riMp3", peak.get("PFAS_rMp3", np.nan)
                    ),
                    "experimental_riMp4": peak.get(
                        "PFAS_riMp4", peak.get("PFAS_rMp4", np.nan)
                    ),
                    "experimental_rMp1": peak.get(
                        "PFAS_riMp1", peak.get("PFAS_rMp1", np.nan)
                    ),
                    "experimental_rMp2": peak.get(
                        "PFAS_riMp2", peak.get("PFAS_rMp2", np.nan)
                    ),
                    "experimental_rMp3": peak.get(
                        "PFAS_riMp3", peak.get("PFAS_rMp3", np.nan)
                    ),
                    "experimental_rMp4": peak.get(
                        "PFAS_riMp4", peak.get("PFAS_rMp4", np.nan)
                    ),
                    "database_row": candidate["database_row"],
                    "DTXSID": candidate["dtxsid"],
                    "CASRN": candidate["casrn"],
                    "preferred_name": candidate["name"],
                    "molecular_formula": candidate["formula"],
                    "neutral_monoisotopic_mass": neutral_mass,
                    "SMILES": candidate["smiles"],
                    "InChIKey": candidate["inchikey"],
                    "theoretical_isotope_status": theory_status,
                    "PFAS_isotope_ratio_matching_requested": isotope_ratio_matching,
                    "PFAS_isotope_ratio_tolerance_percent": isotope_ratio_tolerance_percent,
                    "PFAS_isotope_ratio_match": isotope_match,
                    "PFAS_isotope_ratio_match_status": isotope_status,
                    "PFAS_isotope_ratio_available_count": isotope_available,
                    "PFAS_isotope_ratio_pass_count": isotope_passed,
                }
                row.update(isotope_ci)
                matches.append(row)

    if not matches:
        return pd.DataFrame()
    result = pd.DataFrame(matches)
    result["PFAS_absolute_mass_error_ppm"] = result[
        "PFAS_mass_error_ppm"
    ].abs()
    isotope_delta_columns = [
        f"delta_rMp{order}" for order in ISOTOPE_ORDERS
    ]
    result["PFAS_mean_absolute_isotope_delta"] = (
        result[isotope_delta_columns].abs().mean(axis=1)
    )
    result["PFAS_absolute_CCS_error_percent"] = result[
        "PFAS_CCS_error_percent"
    ].abs()
    return result.sort_values(
        [
            "PFAS_intensity_rank",
            "PFAS_absolute_mass_error_ppm",
            "PFAS_absolute_CCS_error_percent",
            "PFAS_mean_absolute_isotope_delta",
        ],
        kind="mergesort",
        na_position="last",
    ).reset_index(drop=True)


def pair_first_isotopic_peaks(
    data: pd.DataFrame,
    mz_column: str,
    rt_column: str,
    tc_column: str,
    intensity_column: str,
    progress: Callable[[str], None] | None = None,
) -> pd.DataFrame:
    """Build 3D_TC-resolved M+1 to M+4 isotope envelopes."""
    result = data.copy()
    result["_PFAS_input_row"] = np.arange(1, len(result) + 1)
    result["_PFAS_mz"] = _numeric(result[mz_column])
    result["_PFAS_rt"] = _numeric(result[rt_column])
    result["_PFAS_tc"] = _numeric(result[tc_column])
    result["_PFAS_intensity"] = _numeric(result[intensity_column])
    result = result.dropna(
        subset=["_PFAS_mz", "_PFAS_rt", "_PFAS_tc", "_PFAS_intensity"]
    ).copy()
    result = result.sort_values("_PFAS_mz", kind="mergesort").reset_index(drop=True)

    mz_values = result["_PFAS_mz"].to_numpy(float)
    rt_values = result["_PFAS_rt"].to_numpy(float)
    tc_values = result["_PFAS_tc"].to_numpy(float)
    intensity_values = result["_PFAS_intensity"].to_numpy(float)
    row_count = len(result)

    if progress:
        progress("Building mutually exclusive M+1 to M+4 isotope envelopes...")

    envelope_base = np.full(row_count, -1, dtype=int)
    isotope_order = np.zeros(row_count, dtype=int)
    mass_offset = np.full(row_count, np.nan)
    rt_error = np.full(row_count, np.nan)
    tc_error = np.full(row_count, np.nan)

    # A feature assigned as M+n cannot seed another envelope in the same
    # 3D_TC-resolved trace. Equal-mass, coeluting peaks at distinct 3D_TC
    # values remain independent M features.
    for candidate_index in range(row_count):
        candidate_mz = mz_values[candidate_index]
        assignments: list[
            tuple[float, float, float, int, int, float, float, float]
        ] = []
        for order in ISOTOPE_ORDERS:
            expected_base_mz = candidate_mz - float(order)
            start = int(
                np.searchsorted(
                    mz_values,
                    expected_base_mz - HIGHER_ISOTOPE_WINDOW_DA,
                    side="left",
                )
            )
            stop = int(
                np.searchsorted(
                    mz_values,
                    expected_base_mz + HIGHER_ISOTOPE_WINDOW_DA,
                    side="right",
                )
            )
            for base_index in range(start, min(stop, candidate_index)):
                if isotope_order[base_index] != 0:
                    continue
                observed_mass_offset = candidate_mz - mz_values[base_index]
                observed_mass_error = observed_mass_offset - float(order)
                observed_rt_error = rt_values[candidate_index] - rt_values[base_index]
                observed_tc_error = tc_values[candidate_index] - tc_values[base_index]
                if (
                    abs(observed_mass_error) <= HIGHER_ISOTOPE_WINDOW_DA
                    and abs(observed_rt_error) <= RT_TOLERANCE_MIN
                    and abs(observed_tc_error) <= TC_TOLERANCE
                ):
                    assignments.append(
                        (
                            abs(observed_rt_error),
                            abs(observed_tc_error),
                            abs(observed_mass_error),
                            base_index,
                            order,
                            observed_mass_offset,
                            observed_rt_error,
                            observed_tc_error,
                        )
                    )
        if not assignments:
            continue
        (
            _,
            _,
            _,
            selected_base,
            selected_order,
            selected_mass_offset,
            selected_rt_error,
            selected_tc_error,
        ) = min(assignments)
        envelope_base[candidate_index] = selected_base
        isotope_order[candidate_index] = selected_order
        mass_offset[candidate_index] = selected_mass_offset
        rt_error[candidate_index] = selected_rt_error
        tc_error[candidate_index] = selected_tc_error

    result["PFAS_isotope_envelope_base_sorted_index"] = envelope_base
    result["PFAS_isotope_order"] = isotope_order
    result["PFAS_isotope_mass_offset_Da"] = mass_offset
    result["PFAS_isotope_rt_error_min"] = rt_error
    result["PFAS_isotope_tc_error"] = tc_error
    result["PFAS_isotope_role"] = np.where(
        isotope_order == 0,
        "M",
        np.char.add("M+", isotope_order.astype(str)),
    )

    denominator = pd.Series(intensity_values).replace(0, np.nan).to_numpy(float)
    for order in ISOTOPE_ORDERS:
        intensity_sum = np.full(row_count, np.nan)
        candidate_count = np.zeros(row_count, dtype=int)
        mz_min = np.full(row_count, np.nan)
        mz_max = np.full(row_count, np.nan)
        offset_min = np.full(row_count, np.nan)
        offset_max = np.full(row_count, np.nan)
        rt_min = np.full(row_count, np.nan)
        rt_max = np.full(row_count, np.nan)
        tc_min = np.full(row_count, np.nan)
        tc_max = np.full(row_count, np.nan)

        assigned = np.where(isotope_order == order)[0]
        for base_index in np.unique(envelope_base[assigned]):
            indices = assigned[envelope_base[assigned] == base_index]
            intensity_sum[base_index] = np.sum(intensity_values[indices])
            candidate_count[base_index] = len(indices)
            mz_min[base_index] = np.min(mz_values[indices])
            mz_max[base_index] = np.max(mz_values[indices])
            offsets = mz_values[indices] - mz_values[base_index]
            rt_errors = rt_values[indices] - rt_values[base_index]
            tc_errors = tc_values[indices] - tc_values[base_index]
            offset_min[base_index] = np.min(offsets)
            offset_max[base_index] = np.max(offsets)
            rt_min[base_index] = np.min(rt_errors)
            rt_max[base_index] = np.max(rt_errors)
            tc_min[base_index] = np.min(tc_errors)
            tc_max[base_index] = np.max(tc_errors)

        prefix = f"PFAS_M_plus_{order}"
        result[f"{prefix}_intensity_sum"] = intensity_sum
        result[f"{prefix}_candidate_count"] = candidate_count
        result[f"{prefix}_mz_min"] = mz_min
        result[f"{prefix}_mz_max"] = mz_max
        result[f"{prefix}_mass_offset_min_Da"] = offset_min
        result[f"{prefix}_mass_offset_max_Da"] = offset_max
        result[f"{prefix}_rt_error_min"] = rt_min
        result[f"{prefix}_rt_error_max"] = rt_max
        result[f"{prefix}_tc_error_min"] = tc_min
        result[f"{prefix}_tc_error_max"] = tc_max
        result[f"{prefix}_ambiguous"] = candidate_count > 1
        result[f"PFAS_rMp{order}"] = 100.0 * intensity_sum / denominator

    carbon_mz = np.full(row_count, np.nan)
    carbon_intensity = np.full(row_count, np.nan)
    carbon_mass_error = np.full(row_count, np.nan)
    carbon_rt_error = np.full(row_count, np.nan)
    carbon_tc_error = np.full(row_count, np.nan)
    carbon_candidate_count = np.zeros(row_count, dtype=int)
    assigned_m1 = np.where(isotope_order == 1)[0]
    for base_index in np.unique(envelope_base[assigned_m1]):
        indices = assigned_m1[envelope_base[assigned_m1] == base_index]
        errors = (
            mz_values[indices]
            - mz_values[base_index]
            - FIRST_ISOTOPE_SPACING_DA
        )
        eligible = indices[np.abs(errors) <= MASS_TOLERANCE_DA]
        carbon_candidate_count[base_index] = len(eligible)
        if len(eligible) == 0:
            continue
        selected = min(
            eligible,
            key=lambda index: (
                abs(
                    mz_values[index]
                    - mz_values[base_index]
                    - FIRST_ISOTOPE_SPACING_DA
                ),
                abs(rt_values[index] - rt_values[base_index]),
                abs(tc_values[index] - tc_values[base_index]),
            ),
        )
        carbon_mz[base_index] = mz_values[selected]
        carbon_intensity[base_index] = intensity_values[selected]
        carbon_mass_error[base_index] = (
            mz_values[selected] - mz_values[base_index] - FIRST_ISOTOPE_SPACING_DA
        )
        carbon_rt_error[base_index] = rt_values[selected] - rt_values[base_index]
        carbon_tc_error[base_index] = tc_values[selected] - tc_values[base_index]

    result["PFAS_first_13C_peak_mz"] = carbon_mz
    result["PFAS_first_13C_peak_intensity"] = carbon_intensity
    result["PFAS_first_13C_peak_mass_error_Da"] = carbon_mass_error
    result["PFAS_first_13C_peak_rt_error_min"] = carbon_rt_error
    result["PFAS_first_13C_peak_tc_error"] = carbon_tc_error
    result["PFAS_first_13C_peak_candidate_count"] = carbon_candidate_count
    result["PFAS_first_13C_peak_ratio"] = 100.0 * carbon_intensity / denominator

    # The deployed model expects the complete nominal M+1 cluster ratio.
    result["first_isotopic_peak_ratio"] = result["PFAS_rMp1"]
    return result


def _calculate_first_isotopic_peak_ratio(
    df: pd.DataFrame,
    config: ProcessingConfig,
    mz_column: str,
    progress: Callable[[str], None] | None,
) -> tuple[pd.DataFrame, str]:
    mode = config.ion_mode.upper()
    method = config.rimp1_method.lower()
    rimp1_column = _selected_or_detected(
        df, config.rimp1_column, "rimp1", mode, required=False
    )
    m_column = _selected_or_detected(
        df, config.m_intensity_column, "m_intensity", mode, required=False
    )
    mp1_column = _selected_or_detected(
        df, config.mp1_intensity_column, "mp1_intensity", mode, required=False
    )

    if method == "auto":
        if rimp1_column and _numeric(df[rimp1_column]).notna().any():
            method = "existing"
        elif m_column and mp1_column:
            method = "abundance"
        else:
            method = "pairing"

    result = df.copy()
    if method == "existing":
        if not rimp1_column:
            raise ProcessingError(
                "Existing first isotopic peak ratio was selected, but no compatible "
                "column was found."
            )
        result["first_isotopic_peak_ratio"] = _numeric(result[rimp1_column])
        result["PFAS_rMp1"] = result["first_isotopic_peak_ratio"]
        result["PFAS_riMp1_source"] = f"Existing column: {rimp1_column}"
    elif method == "abundance":
        if not m_column or not mp1_column:
            raise ProcessingError(
                "Calculating the first isotopic peak ratio requires both intensity columns."
            )
        denominator = _numeric(result[m_column]).replace(0, np.nan)
        mp1_intensity = _numeric(result[mp1_column])
        result["first_isotopic_peak_ratio"] = (
            100.0 * mp1_intensity / denominator
        )
        result["PFAS_M_intensity"] = _numeric(result[m_column])
        result["PFAS_M_plus_1_intensity_sum"] = mp1_intensity
        result["PFAS_rMp1"] = result["first_isotopic_peak_ratio"]
        result["PFAS_riMp1_source"] = (
            f"Calculated from intensity columns: {mp1_column} / {m_column}"
        )
    elif method == "pairing":
        rt_column = _selected_or_detected(df, config.rt_column, "rt", mode)
        tc_column = _selected_or_detected(df, config.tc_column, "tc", mode)
        intensity_column = _selected_or_detected(
            df, config.intensity_column, "intensity", mode
        )
        result = pair_first_isotopic_peaks(
            result,
            mz_column,
            rt_column,
            tc_column,
            intensity_column,
            progress,
        )
    else:
        raise ProcessingError(
            f"Unsupported first isotopic peak ratio method: {config.rimp1_method}"
        )
    return result, method


def fit_mass_calibration(
    standards: pd.DataFrame,
    observed_mz_column: str,
    exact_mz_column: str,
    method: str,
) -> tuple[Callable[[pd.Series], pd.Series], str, float | None, pd.DataFrame, float, float]:
    calibration = standards.copy()
    calibration["PFAS_mass_observed_mz"] = _numeric(
        calibration[observed_mz_column]
    )
    calibration["PFAS_mass_exact_mz"] = _numeric(calibration[exact_mz_column])
    calibration = calibration.dropna(
        subset=["PFAS_mass_observed_mz", "PFAS_mass_exact_mz"]
    ).copy()
    calibration = calibration[
        (calibration["PFAS_mass_observed_mz"] > 0)
        & (calibration["PFAS_mass_exact_mz"] > 0)
    ].copy()
    if calibration.empty:
        raise ProcessingError(
            "Mass calibration requires standards with observed and exact m/z values."
        )
    calibration["PFAS_mass_error_ppm_observed"] = (
        (calibration["PFAS_mass_observed_mz"] - calibration["PFAS_mass_exact_mz"])
        / calibration["PFAS_mass_exact_mz"]
        * 1e6
    )
    method = method.lower()
    if method == "lock":
        ppm_offset = float(calibration["PFAS_mass_error_ppm_observed"].mean())

        def apply(observed: pd.Series) -> pd.Series:
            return _numeric(observed) / (1.0 + ppm_offset / 1e6)

        calibration["PFAS_mass_fitted_error_ppm"] = ppm_offset
        formula = f"lock mass: corrected_mz = observed_mz / (1 + {ppm_offset:.10g}/1e6)"
        r2 = None
    elif method == "linear":
        if len(calibration) < 2:
            raise ProcessingError(
                "Linear mass calibration requires at least two standards."
            )
        if calibration["PFAS_mass_observed_mz"].nunique() < 2:
            raise ProcessingError(
                "Linear mass calibration requires at least two distinct observed m/z values."
            )
        slope, intercept = np.polyfit(
            calibration["PFAS_mass_observed_mz"].to_numpy(float),
            calibration["PFAS_mass_error_ppm_observed"].to_numpy(float),
            1,
        )

        def apply(observed: pd.Series) -> pd.Series:
            observed_numeric = _numeric(observed)
            fitted_ppm = slope * observed_numeric + intercept
            return observed_numeric / (1.0 + fitted_ppm / 1e6)

        calibration["PFAS_mass_fitted_error_ppm"] = (
            slope * calibration["PFAS_mass_observed_mz"] + intercept
        )
        total = (
            calibration["PFAS_mass_error_ppm_observed"]
            - calibration["PFAS_mass_error_ppm_observed"].mean()
        )
        residual_ppm = (
            calibration["PFAS_mass_error_ppm_observed"]
            - calibration["PFAS_mass_fitted_error_ppm"]
        )
        denominator = float(np.square(total).sum())
        r2 = (
            1.0 - float(np.square(residual_ppm).sum()) / denominator
            if denominator
            else 1.0
        )
        formula = (
            f"linear ppm: ppm_error = {slope:.10g} * observed_mz + "
            f"{intercept:.10g}; corrected_mz = observed_mz / (1 + ppm_error/1e6)"
        )
    else:
        raise ProcessingError(f"Unsupported mass calibration method: {method}")

    calibration["PFAS_mass_corrected_mz"] = apply(
        calibration["PFAS_mass_observed_mz"]
    )
    calibration["PFAS_mass_residual_ppm"] = (
        (calibration["PFAS_mass_corrected_mz"] - calibration["PFAS_mass_exact_mz"])
        / calibration["PFAS_mass_exact_mz"]
        * 1e6
    )
    calibration = calibration.sort_values("PFAS_mass_observed_mz").reset_index(
        drop=True
    )
    mz_min = float(calibration["PFAS_mass_observed_mz"].min())
    mz_max = float(calibration["PFAS_mass_observed_mz"].max())
    return apply, formula, r2, calibration, mz_min, mz_max


def apply_mass_calibration(
    source: pd.DataFrame,
    config: ProcessingConfig,
    mode: str,
    mz_column: str,
) -> tuple[pd.DataFrame, str, float | None, pd.DataFrame | None]:
    result = source.copy()
    method = config.mass_calibration_method.lower()
    result["PFAS_mz_raw"] = _numeric(result[mz_column])
    if method == "none":
        result["PFAS_mz_calibrated"] = result["PFAS_mz_raw"]
        result["PFAS_mass_error_ppm_estimated"] = np.nan
        result["PFAS_mass_calibration_extrapolated"] = False
        result["PFAS_mass_calibration_formula"] = "None"
        return result, "None", None, None
    if method == "lock":
        if config.lock_mass_value is None or not np.isfinite(config.lock_mass_value):
            raise ProcessingError("Enter a valid positive lock mass value.")
        lock_exact = float(config.lock_mass_value)
        if lock_exact <= 0:
            raise ProcessingError("Lock mass value must be positive.")
        lock_tolerance_ppm = float(config.lock_mass_tolerance_ppm)
        if not np.isfinite(lock_tolerance_ppm) or lock_tolerance_ppm <= 0:
            raise ProcessingError("Lock mass search tolerance must be positive.")
        valid_mz = result["PFAS_mz_raw"].dropna()
        valid_mz = valid_mz[valid_mz > 0]
        if valid_mz.empty:
            raise ProcessingError("No valid observed m/z value is available for lock-mass calibration.")
        nearest_index = (valid_mz - lock_exact).abs().idxmin()
        lock_observed = float(valid_mz.loc[nearest_index])
        lock_error_ppm = (lock_observed - lock_exact) / lock_exact * 1e6
        if abs(lock_error_ppm) > lock_tolerance_ppm:
            raise ProcessingError(
                f"No observed peak was found sufficiently close to lock mass {lock_exact:.8g}. "
                f"Nearest observed m/z is {lock_observed:.8g} ({lock_error_ppm:+.1f} ppm), "
                f"outside the selected +/-{lock_tolerance_ppm:g} ppm tolerance."
            )
        correction_factor = lock_exact / lock_observed
        result["PFAS_mz_calibrated"] = result["PFAS_mz_raw"] * correction_factor
        result["PFAS_mass_error_ppm_estimated"] = lock_error_ppm
        result["PFAS_mass_calibration_extrapolated"] = False
        formula = (
            f"lock mass: corrected_mz = observed_mz * {lock_exact:.10g} / "
            f"{lock_observed:.10g}"
        )
        result["PFAS_mass_calibration_formula"] = formula
        calibration_table = pd.DataFrame(
            {
                "PFAS_mass_observed_mz": [lock_observed],
                "PFAS_mass_exact_mz": [lock_exact],
                "PFAS_mass_error_ppm_observed": [lock_error_ppm],
                "PFAS_mass_correction_factor": [correction_factor],
                "PFAS_mass_corrected_mz": [lock_observed * correction_factor],
                "PFAS_mass_residual_ppm": [0.0],
                "PFAS_source_row": [int(nearest_index) + 2],
            }
        )
        return result, formula, None, calibration_table
    if config.mass_standards_path is None or not config.mass_standards_path.exists():
        raise ProcessingError(
            "Choose a Mass calibration standards file for linear correction."
        )
    standards = read_input_table(
        config.mass_standards_path, config.mass_standards_sheet_name
    )
    observed_column = _selected_or_detected(
        standards, config.mass_observed_mz_column, "observed_mz", mode
    )
    exact_column = _selected_or_detected(
        standards, config.mass_exact_mz_column, "exact_mz", mode
    )
    apply, formula, r2, calibration_table, mz_min, mz_max = fit_mass_calibration(
        standards, observed_column, exact_column, method
    )
    result["PFAS_mz_calibrated"] = apply(result["PFAS_mz_raw"])
    result["PFAS_mass_error_ppm_estimated"] = (
        (result["PFAS_mz_raw"] - result["PFAS_mz_calibrated"])
        / result["PFAS_mz_calibrated"]
        * 1e6
    )
    result["PFAS_mass_calibration_extrapolated"] = (
        (result["PFAS_mz_raw"] < mz_min) | (result["PFAS_mz_raw"] > mz_max)
    )
    result["PFAS_mass_calibration_formula"] = formula
    return result, formula, r2, calibration_table


def _validated_charge(
    values: pd.Series,
    field_name: str,
) -> pd.Series:
    charge = _numeric(values).abs()
    invalid = charge.notna() & (charge <= 0)
    if invalid.any():
        raise ProcessingError(f"{field_name} must contain non-zero charge values.")
    return charge


def reduced_mass(
    mz: pd.Series | np.ndarray,
    charge: pd.Series | np.ndarray,
) -> np.ndarray:
    mz_values = np.asarray(mz, dtype=float)
    charge_values = np.abs(np.asarray(charge, dtype=float))
    ion_mass = mz_values * charge_values
    return ion_mass * NITROGEN_MASS_DA / (ion_mass + NITROGEN_MASS_DA)


def fit_cims_calibration(
    standards: pd.DataFrame,
    tc_column: str,
    mz_column: str,
    ccs_column: str,
    charge_column: str | None = None,
    default_charge: int = 1,
) -> tuple[float, float, float, pd.DataFrame]:
    calibration = standards.copy()
    calibration["PFAS_standard_tc"] = _numeric(calibration[tc_column])
    calibration["PFAS_standard_mz"] = _numeric(calibration[mz_column])
    calibration["PFAS_standard_CCS"] = _numeric(calibration[ccs_column])
    if charge_column:
        calibration["PFAS_standard_charge"] = _validated_charge(
            calibration[charge_column], "Standard charge"
        )
    else:
        if default_charge == 0:
            raise ProcessingError("Default charge must be non-zero.")
        calibration["PFAS_standard_charge"] = abs(default_charge)
    calibration = calibration.dropna(
        subset=[
            "PFAS_standard_tc",
            "PFAS_standard_mz",
            "PFAS_standard_CCS",
            "PFAS_standard_charge",
        ]
    ).copy()
    calibration = calibration[
        (calibration["PFAS_standard_tc"] > 0)
        & (calibration["PFAS_standard_mz"] > 0)
        & (calibration["PFAS_standard_CCS"] > 0)
        & (calibration["PFAS_standard_charge"] > 0)
    ].copy()
    if len(calibration) < 3:
        raise ProcessingError(
            "Cyclic/TWIMS power-law CCS calibration requires at least three "
            "standards with positive numeric tc, m/z, CCS, and charge."
        )
    if calibration["PFAS_standard_tc"].nunique() < 3:
        raise ProcessingError(
            "Cyclic/TWIMS power-law calibration requires at least three distinct "
            "positive tc values."
        )
    calibration["PFAS_reduced_mass_N2"] = reduced_mass(
        calibration["PFAS_standard_mz"],
        calibration["PFAS_standard_charge"],
    )
    calibration["PFAS_standard_reduced_CCS"] = (
        calibration["PFAS_standard_CCS"]
        * np.sqrt(calibration["PFAS_reduced_mass_N2"])
        / calibration["PFAS_standard_charge"]
    )
    exponent, log_amplitude = np.polyfit(
        np.log(calibration["PFAS_standard_tc"].to_numpy(float)),
        np.log(calibration["PFAS_standard_reduced_CCS"].to_numpy(float)),
        1,
    )
    amplitude = float(np.exp(log_amplitude))
    exponent = float(exponent)
    if not (np.isfinite(amplitude) and amplitude > 0 and np.isfinite(exponent)):
        raise ProcessingError("Could not fit a valid cyclic/TWIMS CCS power law.")
    predicted_reduced = amplitude * np.power(
        calibration["PFAS_standard_tc"], exponent
    )
    residual_reduced = (
        calibration["PFAS_standard_reduced_CCS"] - predicted_reduced
    )
    total = (
        calibration["PFAS_standard_reduced_CCS"]
        - calibration["PFAS_standard_reduced_CCS"].mean()
    )
    denominator = float(np.square(total).sum())
    r2 = (
        1.0 - float(np.square(residual_reduced).sum()) / denominator
        if denominator
        else 1.0
    )
    calibration["PFAS_fitted_reduced_CCS"] = predicted_reduced
    calibration["PFAS_power_law_A"] = amplitude
    calibration["PFAS_power_law_B"] = exponent
    calibration["PFAS_back_calculated_CCS"] = (
        predicted_reduced
        * calibration["PFAS_standard_charge"]
        / np.sqrt(calibration["PFAS_reduced_mass_N2"])
    )
    calibration["PFAS_CCS_residual"] = (
        calibration["PFAS_standard_CCS"]
        - calibration["PFAS_back_calculated_CCS"]
    )
    calibration["PFAS_CCS_relative_error_percent"] = (
        100.0
        * calibration["PFAS_CCS_residual"]
        / calibration["PFAS_standard_CCS"].replace(0, np.nan)
    )
    calibration = calibration.sort_values("PFAS_standard_tc").reset_index(drop=True)
    return amplitude, exponent, float(r2), calibration


def add_calibration_chart(
    worksheet,
    calibration_table: pd.DataFrame,
    r2: float,
) -> None:
    row_count = len(calibration_table)
    if row_count < 2:
        return
    tc_column = calibration_table.columns.get_loc("PFAS_standard_tc") + 1
    ccs_column = (
        calibration_table.columns.get_loc("PFAS_standard_reduced_CCS") + 1
    )
    fitted_column = (
        calibration_table.columns.get_loc("PFAS_fitted_reduced_CCS") + 1
    )
    x_values = Reference(
        worksheet, min_col=tc_column, min_row=2, max_row=row_count + 1
    )

    chart = ScatterChart()
    chart.title = f"cIMS/TWIMS power-law CCS calibration (R2 = {r2:.5f})"
    chart.x_axis.title = "Corrected arrival time, tc"
    chart.y_axis.title = "Reduced CCS (CCS * sqrt(mu) / |z|)"
    chart.height = 9
    chart.width = 16

    measured_values = Reference(
        worksheet, min_col=ccs_column, min_row=2, max_row=row_count + 1
    )
    measured = Series(measured_values, x_values, title="Standards")
    measured.marker.symbol = "circle"
    measured.marker.size = 7
    measured.graphicalProperties.line.noFill = True
    chart.series.append(measured)

    fitted_values = Reference(
        worksheet, min_col=fitted_column, min_row=2, max_row=row_count + 1
    )
    fitted = Series(fitted_values, x_values, title="Power-law fit")
    fitted.marker.symbol = "none"
    fitted.graphicalProperties.line.solidFill = "D62728"
    fitted.graphicalProperties.line.width = 22000
    chart.series.append(fitted)
    worksheet.add_chart(chart, f"A{row_count + 4}")


def assign_plot_confidence(
    data: pd.DataFrame,
    strategy: str,
) -> pd.DataFrame:
    """Assign a four-level display class without changing screening results."""
    strategy = strategy.lower()
    if strategy not in SUPPORTED_SCREENING_STRATEGIES:
        raise ProcessingError(f"Unsupported screening strategy: {strategy}")

    model_rank = (
        _numeric(data.get("PFAS_level", pd.Series(0, index=data.index)))
        .fillna(0)
        .clip(0, 3)
        .astype(int)
        .to_numpy()
    )
    linear_probability = _numeric(
        data.get(
            "PFAS_linear_pass_probability",
            pd.Series(np.nan, index=data.index),
        )
    ).to_numpy(float)
    linear_rank = np.select(
        [
            linear_probability >= LINEAR_HIGH_CONFIDENCE_PROBABILITY,
            linear_probability >= LINEAR_MEDIUM_CONFIDENCE_PROBABILITY,
            linear_probability >= LINEAR_CANDIDATE_PROBABILITY,
        ],
        [3, 2, 1],
        default=0,
    ).astype(int)

    if strategy == "model_only":
        rank = model_rank
    elif strategy == "linear_only":
        rank = linear_rank
    elif strategy == "intersection":
        rank = np.minimum(model_rank, linear_rank)
    else:
        rank = np.maximum(model_rank, linear_rank)

    selected = _numeric(
        data.get("PFAS_final_candidate", pd.Series(0, index=data.index))
    ).fillna(0).to_numpy() == 1
    isotope_role = data.get(
        "PFAS_isotope_role", pd.Series("M", index=data.index)
    ).astype(str).to_numpy()
    rank = np.where(selected & (isotope_role == "M"), rank, 0).astype(int)

    result = pd.DataFrame(index=data.index)
    result["PFAS_plot_confidence_rank"] = rank
    result["PFAS_plot_confidence_class"] = PLOT_CONFIDENCE_LABELS[rank]
    result["PFAS_plot_color_hex"] = PLOT_CONFIDENCE_COLORS[rank]
    return result


def add_pfas_scatter_plots(
    workbook,
    data: pd.DataFrame,
    strategy: str,
) -> None:
    """Add native Excel scatter plots for the filtered monoisotopic features."""
    worksheet = workbook.create_sheet("PFAS scatter plots")
    worksheet["A1"] = "PFAS screening scatter plots"
    worksheet["A2"] = "Screening strategy"
    worksheet["B2"] = strategy
    worksheet["A3"] = "Displayed rows"
    worksheet["B3"] = (
        "Monoisotopic M features after calibration and optional blank subtraction; "
        "rejected features remain visible in grey."
    )
    worksheet["A4"] = "Confidence combination"
    worksheet["B4"] = (
        "MLM level for MLM only; linear MC pass probability for linear only; "
        "lower confidence for intersection; higher confidence for union."
    )

    for offset, (label, color) in enumerate(
        zip(PLOT_CONFIDENCE_LABELS[::-1], PLOT_CONFIDENCE_COLORS[::-1]),
        start=1,
    ):
        worksheet.cell(row=offset, column=10, value=str(label))
        worksheet.cell(row=offset, column=11, value=f"#{color}")
        worksheet.cell(row=offset, column=11).fill = PatternFill(
            fill_type="solid", fgColor=str(color)
        )

    plot_data = data.copy()
    role = plot_data.get(
        "PFAS_isotope_role", pd.Series("M", index=plot_data.index)
    ).astype(str)
    plot_data = plot_data.loc[role == "M"].copy()
    for column in ["PFAS_mz_used", "PFAS_CCS_model_input", "PFAS_riMp1"]:
        plot_data[column] = _numeric(plot_data[column])
    plot_data = plot_data.dropna(
        subset=["PFAS_mz_used", "PFAS_CCS_model_input", "PFAS_riMp1"]
    )

    header_row = 32
    data_row = header_row + 1
    charts = [
        ("m/z vs CCS after PFAS screening", "m/z", "PFAS_mz_used", "A7"),
        ("riMp1 vs CCS after PFAS screening", "riMp1", "PFAS_riMp1", "J7"),
    ]
    for chart_index, (title, x_title, x_column, anchor) in enumerate(charts):
        chart = ScatterChart()
        chart.title = title
        chart.x_axis.title = x_title
        chart.y_axis.title = "CCS"
        chart.height = 12
        chart.width = 16
        for rank in range(4):
            column = 1 + chart_index * 8 + rank * 2
            label = str(PLOT_CONFIDENCE_LABELS[rank])
            color = str(PLOT_CONFIDENCE_COLORS[rank])
            worksheet.cell(row=header_row, column=column, value=f"{x_title}: {label}")
            worksheet.cell(row=header_row, column=column + 1, value=f"CCS: {label}")
            subset = plot_data.loc[
                _numeric(plot_data["PFAS_plot_confidence_rank"]).fillna(0).astype(int)
                == rank
            ]
            if subset.empty:
                continue
            for row_offset, (_, point) in enumerate(subset.iterrows()):
                worksheet.cell(
                    row=data_row + row_offset,
                    column=column,
                    value=float(point[x_column]),
                )
                worksheet.cell(
                    row=data_row + row_offset,
                    column=column + 1,
                    value=float(point["PFAS_CCS_model_input"]),
                )
            max_row = data_row + len(subset) - 1
            x_values = Reference(
                worksheet, min_col=column, min_row=data_row, max_row=max_row
            )
            y_values = Reference(
                worksheet, min_col=column + 1, min_row=data_row, max_row=max_row
            )
            series = Series(y_values, x_values, title=label)
            series.marker.symbol = "circle"
            series.marker.size = 6
            series.marker.graphicalProperties.solidFill = color
            series.marker.graphicalProperties.line.solidFill = color
            series.graphicalProperties.line.noFill = True
            chart.series.append(series)
        worksheet.add_chart(chart, anchor)

    worksheet.freeze_panes = f"A{data_row}"


def apply_first_calibration(
    source: pd.DataFrame,
    config: ProcessingConfig,
    mode: str,
    mz_column: str,
) -> tuple[pd.DataFrame, str, float | None, pd.DataFrame | None, str]:
    result = source.copy()
    method = config.first_calibration_method.lower()
    if method == "cims_reduced_ccs":
        if config.standards_path is None or not config.standards_path.exists():
            raise ProcessingError(
                "Choose a standards file for IMS CCS calibration."
            )
        sample_tc_column = _selected_or_detected(
            source, config.tc_column, "tc", mode
        )
        standards = read_input_table(
            config.standards_path, config.standards_sheet_name
        )
        standards_tc_column = _selected_or_detected(
            standards, config.standards_tc_column, "tc", mode
        )
        standards_mz_column = _selected_or_detected(
            standards, config.standards_mz_column, "mz", mode
        )
        standards_ccs_column = _selected_or_detected(
            standards, config.standards_ccs_column, "ccs", mode
        )
        standards_charge_column = _selected_or_detected(
            standards,
            config.standards_charge_column,
            "charge",
            mode,
            required=False,
        )
        amplitude, exponent, r2, calibration_table = fit_cims_calibration(
            standards,
            standards_tc_column,
            standards_mz_column,
            standards_ccs_column,
            standards_charge_column,
            config.default_charge,
        )
        sample_charge_column = _selected_or_detected(
            source, config.charge_column, "charge", mode, required=False
        )
        result["PFAS_tc_input"] = _numeric(result[sample_tc_column])
        if sample_charge_column:
            result["PFAS_charge_used"] = _validated_charge(
                result[sample_charge_column], "Sample charge"
            ).fillna(abs(config.default_charge))
        else:
            if config.default_charge == 0:
                raise ProcessingError("Default charge must be non-zero.")
            result["PFAS_charge_used"] = abs(config.default_charge)
        result["PFAS_reduced_mass_N2"] = reduced_mass(
            _numeric(result[mz_column]),
            result["PFAS_charge_used"],
        )
        valid_tc = result["PFAS_tc_input"] > 0
        result["PFAS_reduced_CCS_fitted"] = np.nan
        result.loc[valid_tc, "PFAS_reduced_CCS_fitted"] = amplitude * np.power(
            result.loc[valid_tc, "PFAS_tc_input"], exponent
        )
        result["PFAS_CCS_stage1"] = (
            result["PFAS_reduced_CCS_fitted"]
            * result["PFAS_charge_used"]
            / np.sqrt(result["PFAS_reduced_mass_N2"])
        )
        tc_min = float(calibration_table["PFAS_standard_tc"].min())
        tc_max = float(calibration_table["PFAS_standard_tc"].max())
        mz_min = float(calibration_table["PFAS_standard_mz"].min())
        mz_max = float(calibration_table["PFAS_standard_mz"].max())
        sample_mz = _numeric(result[mz_column])
        result["PFAS_calibration_tc_extrapolated"] = (
            (result["PFAS_tc_input"] < tc_min)
            | (result["PFAS_tc_input"] > tc_max)
        )
        result["PFAS_calibration_mz_extrapolated"] = (
            (sample_mz < mz_min) | (sample_mz > mz_max)
        )
        result["PFAS_calibration_extrapolated"] = (
            result["PFAS_calibration_tc_extrapolated"]
            | result["PFAS_calibration_mz_extrapolated"]
        )
        formula = (
            f"reduced_CCS = {amplitude:.10g} * tc^{exponent:.10g}; "
            "CCS = reduced_CCS * |z| / sqrt(mu_N2)"
        )
        result["PFAS_CCS_stage1_formula"] = formula
        return result, formula, r2, calibration_table, "PFAS_CCS_stage1"
    if method == "existing_ccs":
        if config.ccs_column == "__calibrated__":
            raise ProcessingError(
                "The Calibrated CCS mapping is produced by the cIMS/TWIMS "
                "power-law calibration from standards. "
                "Select that calibration method, or map an existing CCS worksheet column."
            )
        ccs_column = _selected_or_detected(source, config.ccs_column, "ccs", mode)
        result["PFAS_CCS_stage1"] = _numeric(result[ccs_column])
        formula = f"Existing CCS column: {ccs_column}"
        result["PFAS_CCS_stage1_formula"] = formula
        return result, formula, None, None, ccs_column
    raise ProcessingError(
        f"Unsupported CCS calibration method: {config.first_calibration_method}"
    )


def _bundle_feature_columns(bundle: dict) -> list[str]:
    feature_columns = bundle.get("feature_cols")
    if feature_columns is None:
        model_feature_names = getattr(bundle.get("model"), "feature_names_in_", None)
        if model_feature_names is not None:
            feature_columns = list(model_feature_names)
    if feature_columns is None:
        feature_columns = FEATURE_COLS
    return [str(column) for column in feature_columns]


def _required_raw_columns(feature_columns: list[str]) -> set[str]:
    requirements: set[str] = set()
    for column in feature_columns:
        normalized = _normalized(column)
        if normalized in {"mz", "mzpercentile"} or normalized.startswith("mover"):
            requirements.add("mz")
        if normalized == "ccs" or normalized.startswith("ccsover"):
            requirements.add("CCS")
        if "mp1" in normalized or "rimp1" in normalized or "rmp1" in normalized:
            requirements.add("riMp1")
        if "rimp2" in normalized or "rmp2" in normalized:
            requirements.add("riMp2")
        if "rimp3" in normalized or "rmp3" in normalized:
            requirements.add("riMp3")
        if "rimp4" in normalized or "rmp4" in normalized:
            requirements.add("riMp4")
    return requirements


def make_features(
    raw: pd.DataFrame,
    medians: pd.Series,
    train_mz_sorted: np.ndarray,
    feature_columns: list[str] | None = None,
) -> pd.DataFrame:
    selected_features = list(feature_columns or FEATURE_COLS)
    d = pd.DataFrame(index=raw.index)
    for column in ["mz", "CCS", "riMp1", "riMp2", "riMp3", "riMp4"]:
        d[column] = _numeric(raw[column]) if column in raw else np.nan
    d["rMp1"] = d["riMp1"]
    d["rMp2"] = d["riMp2"]
    d["rMp3"] = d["riMp3"]
    d["rMp4"] = d["riMp4"]
    eps = 1e-12
    denominator = max(len(train_mz_sorted) - 1, 1)
    d["mz_percentile"] = (
        np.searchsorted(train_mz_sorted, d["mz"].to_numpy(float), side="right")
        / denominator
    )
    d["CCS_over_m"] = d["CCS"] / (d["mz"] + eps)
    d["CCS_over_Mp1"] = d["CCS"] / (d["riMp1"] + eps)
    d["M_over_Mp1"] = d["mz"] / (d["riMp1"] + eps)
    d["CCS_over_sqrtm"] = d["CCS"] / (np.sqrt(d["mz"]) + eps)
    d["CCS_over_m23"] = d["CCS"] / ((d["mz"] ** (2 / 3)) + eps)
    missing = [column for column in selected_features if column not in d]
    if missing:
        raise ProcessingError(
            "Selected model requires unsupported feature columns: "
            + ", ".join(missing)
        )
    return d[selected_features].replace([np.inf, -np.inf], np.nan).fillna(medians)


def _predict_probability(bundle: dict, raw: pd.DataFrame) -> np.ndarray:
    features = make_features(
        raw,
        bundle["medians"],
        bundle["train_mz_sorted"],
        _bundle_feature_columns(bundle),
    )
    return bundle["model"].predict_proba(features)[:, 1]


def mc_predict(
    bundle: dict,
    raw: pd.DataFrame,
    iterations: int,
    rimp1_rel_noise: float | None = None,
    ccs_rel_noise: float | None = None,
    progress: Callable[[str], None] | None = None,
    chunk_size: int = 8000,
) -> tuple[np.ndarray, np.ndarray, np.ndarray]:
    if iterations < 1:
        raise ProcessingError("Monte Carlo iterations must be at least 1.")
    rng = np.random.default_rng(42)
    n = len(raw)
    pmean = np.full(n, np.nan)
    pstd = np.full(n, np.nan)
    p05 = np.full(n, np.nan)
    rimp1_noise = (
        abs(float(rimp1_rel_noise))
        if rimp1_rel_noise is not None
        else abs(float(bundle.get("riMp1_rel_noise", 0.20)))
    )
    ccs_noise = (
        abs(float(ccs_rel_noise))
        if ccs_rel_noise is not None
        else abs(float(bundle.get("ccs_rel_noise", 0.03)))
    )

    for start in range(0, n, chunk_size):
        end = min(start + chunk_size, n)
        base = raw.iloc[start:end].reset_index(drop=True)
        probabilities = np.zeros((iterations, len(base)), dtype=np.float32)
        for index in range(iterations):
            perturbed = base.copy()
            for ratio_column in ["riMp1", "riMp2", "riMp3", "riMp4"]:
                if ratio_column not in perturbed:
                    continue
                perturbed[ratio_column] = np.clip(
                    perturbed[ratio_column].to_numpy(float)
                    * (1 + rng.uniform(-rimp1_noise, rimp1_noise, len(base))),
                    1e-12,
                    None,
                )
            perturbed["CCS"] = np.clip(
                perturbed["CCS"].to_numpy(float)
                * (1 + rng.uniform(-ccs_noise, ccs_noise, len(base))),
                1e-12,
                None,
            )
            probabilities[index] = _predict_probability(bundle, perturbed)
            if progress and (index + 1) % max(iterations // 5, 1) == 0:
                progress(
                    f"Predicting rows {start + 1}-{end}: "
                    f"Monte Carlo {index + 1}/{iterations}"
                )
        pmean[start:end] = probabilities.mean(axis=0)
        pstd[start:end] = probabilities.std(axis=0)
        p05[start:end] = np.quantile(probabilities, 0.05, axis=0)
    return pmean, pstd, p05


def assign_tier(
    pmean: np.ndarray,
    p05: np.ndarray,
    t_low: float,
    t_high: float,
) -> tuple[np.ndarray, np.ndarray]:
    level = np.zeros(len(pmean), dtype=int)
    level[(pmean >= t_high) & (pmean < t_low)] = 1
    level[(pmean >= t_low) & (p05 < t_low)] = 2
    level[p05 >= t_low] = 3
    names = np.array(
        [
            "Level 0 Not PFAS",
            "Level 1 Candidate PFAS",
            "Level 2 Medium confidence",
            "Level 3 High-confidence PFAS",
        ],
        dtype=object,
    )
    return level, names[level]


def calculate_linear_screening(raw: pd.DataFrame, mode: str) -> pd.DataFrame:
    """Apply the ion-mode-specific PFAS linear CCS boundaries."""
    mode = mode.upper()
    if mode == "POS":
        upper_slope, upper_intercept = 0.2537, 84.701
        lower_slope, lower_intercept = 4.657, 98.570
    elif mode == "NEG":
        upper_slope, upper_intercept = 0.24284, 84.432
        lower_slope, lower_intercept = 4.415, 97.440
    else:
        raise ProcessingError(f"Unsupported ion mode for linear screening: {mode}")

    mz = _numeric(raw["mz"])
    ccs = _numeric(raw["CCS"])
    rimp1 = _numeric(raw["riMp1"])
    upper_boundary = upper_slope * mz + upper_intercept
    lower_boundary = lower_slope * rimp1 + lower_intercept
    evaluable = mz.notna() & ccs.notna() & rimp1.notna()
    passed = evaluable & (ccs < upper_boundary) & (ccs > lower_boundary)

    result = pd.DataFrame(index=raw.index)
    result["PFAS_linear_upper_CCS_boundary"] = upper_boundary
    result["PFAS_linear_lower_CCS_boundary"] = lower_boundary
    result["PFAS_linear_upper_residual"] = ccs - upper_boundary
    result["PFAS_linear_lower_residual"] = ccs - lower_boundary
    result["PFAS_linear_screen_evaluable"] = evaluable
    result["PFAS_linear_screen_pass"] = np.where(evaluable, passed.astype(int), np.nan)
    result["PFAS_linear_screen_label"] = np.where(
        ~evaluable,
        "Not evaluable",
        np.where(passed, "Linear PFAS", "Linear non-PFAS"),
    )
    result["PFAS_linear_screen_formula"] = (
        f"CCS < {upper_slope:g}*mz + {upper_intercept:g} and "
        f"CCS > {lower_slope:g}*riMp1 + {lower_intercept:g}"
    )
    return result


def mc_linear_screening(
    raw: pd.DataFrame,
    mode: str,
    iterations: int,
    rimp1_rel_noise: float,
    ccs_rel_noise: float,
    progress: Callable[[str], None] | None = None,
    chunk_size: int = 8000,
) -> pd.DataFrame:
    """Propagate riMp1 and CCS uncertainty through the linear PFAS screen."""
    if iterations < 1:
        raise ProcessingError("Monte Carlo iterations must be at least 1.")
    deterministic = calculate_linear_screening(raw, mode)
    result = deterministic.copy()
    result["PFAS_linear_deterministic_pass"] = result[
        "PFAS_linear_screen_pass"
    ]
    evaluable = result["PFAS_linear_screen_evaluable"].astype(bool)
    pass_count = np.full(len(raw), np.nan)
    pass_probability = np.full(len(raw), np.nan)
    upper_mean = np.full(len(raw), np.nan)
    upper_p05 = np.full(len(raw), np.nan)
    upper_p95 = np.full(len(raw), np.nan)
    lower_mean = np.full(len(raw), np.nan)
    lower_p05 = np.full(len(raw), np.nan)
    lower_p95 = np.full(len(raw), np.nan)
    rng = np.random.default_rng(42)
    rimp1_noise = abs(float(rimp1_rel_noise))
    ccs_noise = abs(float(ccs_rel_noise))

    for start in range(0, len(raw), chunk_size):
        end = min(start + chunk_size, len(raw))
        chunk_valid = evaluable.iloc[start:end].to_numpy(bool)
        if not chunk_valid.any():
            continue
        mz = _numeric(raw["mz"].iloc[start:end]).to_numpy(float)
        ccs = _numeric(raw["CCS"].iloc[start:end]).to_numpy(float)
        rimp1 = _numeric(raw["riMp1"].iloc[start:end]).to_numpy(float)
        upper_boundary = deterministic[
            "PFAS_linear_upper_CCS_boundary"
        ].iloc[start:end].to_numpy(float)
        upper_residuals = np.empty((iterations, end - start), dtype=np.float32)
        lower_residuals = np.empty((iterations, end - start), dtype=np.float32)
        passes = np.zeros((iterations, end - start), dtype=bool)
        for index in range(iterations):
            perturbed_rimp1 = rimp1 * (
                1.0 + rng.uniform(-rimp1_noise, rimp1_noise, end - start)
            )
            perturbed_ccs = ccs * (
                1.0 + rng.uniform(-ccs_noise, ccs_noise, end - start)
            )
            if mode.upper() == "POS":
                lower_boundary = 4.657 * perturbed_rimp1 + 98.570
            else:
                lower_boundary = 4.415 * perturbed_rimp1 + 97.440
            upper_residuals[index] = perturbed_ccs - upper_boundary
            lower_residuals[index] = perturbed_ccs - lower_boundary
            passes[index] = (
                chunk_valid
                & (upper_residuals[index] < 0)
                & (lower_residuals[index] > 0)
            )
        target = slice(start, end)
        pass_count[target] = passes.sum(axis=0)
        pass_probability[target] = passes.mean(axis=0)
        upper_mean[target] = upper_residuals.mean(axis=0)
        upper_p05[target] = np.quantile(upper_residuals, 0.05, axis=0)
        upper_p95[target] = np.quantile(upper_residuals, 0.95, axis=0)
        lower_mean[target] = lower_residuals.mean(axis=0)
        lower_p05[target] = np.quantile(lower_residuals, 0.05, axis=0)
        lower_p95[target] = np.quantile(lower_residuals, 0.95, axis=0)
        if progress:
            progress(f"Linear Monte Carlo rows {start + 1}-{end}: complete")

    result["PFAS_linear_MC_iterations"] = iterations
    result["PFAS_linear_pass_count"] = pass_count
    result["PFAS_linear_pass_probability"] = pass_probability
    result["PFAS_linear_upper_residual_mean"] = upper_mean
    result["PFAS_linear_upper_residual_p05"] = upper_p05
    result["PFAS_linear_upper_residual_p95"] = upper_p95
    result["PFAS_linear_lower_residual_mean"] = lower_mean
    result["PFAS_linear_lower_residual_p05"] = lower_p05
    result["PFAS_linear_lower_residual_p95"] = lower_p95
    result["PFAS_linear_screen_pass"] = np.where(
        evaluable,
        (
            result["PFAS_linear_pass_probability"]
            >= LINEAR_CANDIDATE_PROBABILITY
        ).astype(int),
        np.nan,
    )
    result["PFAS_linear_screen_label"] = np.select(
        [
            ~evaluable,
            result["PFAS_linear_pass_probability"]
            >= LINEAR_HIGH_CONFIDENCE_PROBABILITY,
            result["PFAS_linear_pass_probability"]
            >= LINEAR_CANDIDATE_PROBABILITY,
        ],
        [
            "Not evaluable",
            "High-confidence linear PFAS",
            "Uncertain linear PFAS candidate",
        ],
        default="Linear non-PFAS",
    )
    return result


def combine_screening_results(
    model_candidate: pd.Series,
    linear_candidate: pd.Series,
    strategy: str,
) -> pd.Series:
    """Combine model and linear labels according to the selected strategy."""
    strategy = strategy.lower()
    if strategy not in SUPPORTED_SCREENING_STRATEGIES:
        raise ProcessingError(f"Unsupported screening strategy: {strategy}")
    model_pass = _numeric(model_candidate).fillna(0).astype(bool)
    linear_pass = _numeric(linear_candidate).fillna(0).astype(bool)
    if strategy == "model_only":
        selected = model_pass
    elif strategy == "linear_only":
        selected = linear_pass
    elif strategy == "intersection":
        selected = model_pass & linear_pass
    else:
        selected = model_pass | linear_pass
    return selected.astype(int)


def _blank_intensity_columns(
    frame: pd.DataFrame,
    config: ProcessingConfig,
    mode: str,
    *,
    require_primary: bool,
) -> dict[str, str]:
    primary = _selected_or_detected(
        frame,
        config.intensity_column,
        "intensity",
        mode,
        required=False,
    )
    m_column = _selected_or_detected(
        frame,
        config.m_intensity_column,
        "m_intensity",
        mode,
        required=False,
    )
    mp1_column = _selected_or_detected(
        frame,
        config.mp1_intensity_column,
        "mp1_intensity",
        mode,
        required=False,
    )
    primary = primary or m_column
    if require_primary and not primary:
        raise ProcessingError(
            "Blank subtraction requires an intensity column in every sample and blank file."
        )
    result: dict[str, str] = {}
    if primary:
        result["signal"] = primary
    if m_column:
        result["m"] = m_column
    if mp1_column:
        result["mp1"] = mp1_column
    return result


def _prepare_blank_replicate(
    path: Path,
    config: ProcessingConfig,
    mode: str,
) -> pd.DataFrame:
    sheet: str | int = 0 if path.suffix.lower() == ".csv" else config.blank_sheet_name
    source = read_input_table(path, sheet)
    if source.empty:
        raise ProcessingError(f"Blank file contains no rows: {path}")
    original_mz_column = _selected_or_detected(source, config.mz_column, "mz", mode)
    mass_source, _, _, _ = apply_mass_calibration(
        source, config, mode, original_mz_column
    )
    calibrated, _, _, _, _ = apply_first_calibration(
        mass_source, config, mode, "PFAS_mz_calibrated"
    )
    rt_column = _selected_or_detected(calibrated, config.rt_column, "rt", mode)
    intensity_columns = _blank_intensity_columns(
        calibrated, config, mode, require_primary=True
    )
    result = pd.DataFrame(
        {
            "mz": _numeric(calibrated["PFAS_mz_calibrated"]),
            "rt": _numeric(calibrated[rt_column]),
            "CCS": _numeric(calibrated["PFAS_CCS_stage1"]),
        }
    )
    for key, column in intensity_columns.items():
        result[key] = _numeric(calibrated[column]).clip(lower=0)
    result = result.dropna(subset=["mz", "rt", "CCS", "signal"])
    result = result[(result["mz"] > 0) & (result["CCS"] > 0)].copy()
    return result.sort_values("mz", kind="mergesort").reset_index(drop=True)


def apply_blank_subtraction(
    sample: pd.DataFrame,
    blank_replicates: list[pd.DataFrame],
    config: ProcessingConfig,
    mode: str,
) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Subtract aligned mean blank intensity and apply the mean + k*SD rule."""
    if len(blank_replicates) < 2:
        raise ProcessingError(
            "Blank subtraction requires at least two blank replicate files."
        )
    mz_tolerance = float(config.blank_mz_tolerance_ppm)
    rt_tolerance = float(config.blank_rt_tolerance_min)
    ccs_tolerance = float(config.blank_ccs_tolerance_percent)
    sigma_multiplier = float(config.blank_sigma_multiplier)
    if min(mz_tolerance, rt_tolerance, ccs_tolerance) <= 0:
        raise ProcessingError("Blank matching tolerances must be positive.")
    if sigma_multiplier < 0:
        raise ProcessingError("The blank SD multiplier cannot be negative.")

    result = sample.copy()
    sample_rt_column = _selected_or_detected(result, config.rt_column, "rt", mode)
    sample_intensity_columns = _blank_intensity_columns(
        result, config, mode, require_primary=True
    )
    sample_mz = _numeric(result["PFAS_mz_calibrated"]).to_numpy(float)
    sample_rt = _numeric(result[sample_rt_column]).to_numpy(float)
    sample_ccs = _numeric(result["PFAS_CCS_stage1"]).to_numpy(float)
    row_count = len(result)
    replicate_count = len(blank_replicates)
    matched_indices = np.full((replicate_count, row_count), -1, dtype=int)

    for replicate_index, blank in enumerate(blank_replicates):
        blank_mz = blank["mz"].to_numpy(float)
        blank_rt = blank["rt"].to_numpy(float)
        blank_ccs = blank["CCS"].to_numpy(float)
        for row_index, (mz, rt, ccs) in enumerate(
            zip(sample_mz, sample_rt, sample_ccs)
        ):
            if not (np.isfinite(mz) and mz > 0 and np.isfinite(rt) and np.isfinite(ccs) and ccs > 0):
                continue
            mz_delta = mz * mz_tolerance / 1e6
            start = int(np.searchsorted(blank_mz, mz - mz_delta, side="left"))
            stop = int(np.searchsorted(blank_mz, mz + mz_delta, side="right"))
            if start >= stop:
                continue
            candidates = np.arange(start, stop)
            rt_error = np.abs(blank_rt[candidates] - rt)
            ccs_error_percent = np.abs(blank_ccs[candidates] - ccs) / ccs * 100.0
            eligible = candidates[
                (rt_error <= rt_tolerance)
                & (ccs_error_percent <= ccs_tolerance)
            ]
            if len(eligible) == 0:
                continue
            mass_error_ppm = np.abs(blank_mz[eligible] - mz) / mz * 1e6
            normalized_score = (
                (mass_error_ppm / mz_tolerance) ** 2
                + (np.abs(blank_rt[eligible] - rt) / rt_tolerance) ** 2
                + (
                    (np.abs(blank_ccs[eligible] - ccs) / ccs * 100.0)
                    / ccs_tolerance
                )
                ** 2
            )
            matched_indices[replicate_index, row_index] = int(
                eligible[int(np.argmin(normalized_score))]
            )

    corrected_by_key: dict[str, np.ndarray] = {}
    original_by_column = {
        column: _numeric(result[column]).to_numpy(float)
        for column in set(sample_intensity_columns.values())
    }
    corrected_by_column: dict[str, np.ndarray] = {}
    primary_blank_values = np.zeros((replicate_count, row_count), dtype=float)
    matched_count = (matched_indices >= 0).sum(axis=0)
    for key, sample_column in sample_intensity_columns.items():
        blank_values = np.zeros((replicate_count, row_count), dtype=float)
        for replicate_index, blank in enumerate(blank_replicates):
            if key not in blank:
                continue
            indices = matched_indices[replicate_index]
            present = indices >= 0
            blank_values[replicate_index, present] = (
                _numeric(blank[key]).fillna(0).to_numpy(float)[indices[present]]
            )
        if key == "signal":
            primary_blank_values = blank_values
        sample_values = original_by_column[sample_column]
        corrected = np.maximum(sample_values - blank_values.mean(axis=0), 0.0)
        corrected_by_key[key] = corrected
        corrected_by_column[sample_column] = corrected
        result[f"PFAS_{key}_intensity_raw"] = sample_values
        result[f"PFAS_blank_mean_{key}_intensity"] = blank_values.mean(axis=0)
        result[f"PFAS_blank_corrected_{key}_intensity"] = corrected
    for sample_column, corrected in corrected_by_column.items():
        result[sample_column] = corrected

    blank_mean = primary_blank_values.mean(axis=0)
    blank_sd = primary_blank_values.std(axis=0, ddof=1)
    threshold = blank_mean + sigma_multiplier * blank_sd
    sample_raw = _numeric(result["PFAS_signal_intensity_raw"]).to_numpy(float)
    passed = np.isfinite(sample_raw) & (sample_raw >= threshold)
    result["PFAS_blank_subtraction_applied"] = True
    result["PFAS_blank_matched_replicates"] = matched_count
    result["PFAS_blank_total_replicates"] = replicate_count
    result["PFAS_blank_mean_intensity"] = blank_mean
    result["PFAS_blank_sd_intensity"] = blank_sd
    result["PFAS_blank_threshold_mean_plus_3sd"] = threshold
    result["PFAS_blank_corrected_intensity"] = corrected_by_key["signal"]
    result["PFAS_blank_subtraction_pass"] = passed
    result["PFAS_blank_subtraction_status"] = np.where(
        passed, "Retained: sample intensity >= blank mean + 3 SD", "Removed by blank subtraction"
    )
    removed = result.loc[~passed].copy()
    retained = result.loc[passed].copy()
    return retained, removed


def process_batch(
    base_config: ProcessingConfig,
    sample_paths: list[Path] | tuple[Path, ...],
    output_directory: Path,
    progress: Callable[[str], None] | None = None,
) -> BatchProcessingSummary:
    paths = [Path(path) for path in sample_paths]
    if not paths:
        raise ProcessingError("Choose at least one sample file for batch processing.")
    output_directory.mkdir(parents=True, exist_ok=True)
    notify = progress or (lambda _message: None)
    rows: list[dict[str, object]] = []
    results: list[ProcessingSummary] = []
    used_names: set[str] = set()
    for index, input_path in enumerate(paths, start=1):
        stem = input_path.stem
        output_name = f"{stem}_PFAS_labeled.xlsx"
        if output_name.lower() in used_names:
            output_name = f"{stem}_{index}_PFAS_labeled.xlsx"
        used_names.add(output_name.lower())
        output_path = output_directory / output_name
        notify(f"Batch sample {index}/{len(paths)}: {input_path.name}")
        config = replace(
            base_config,
            input_path=input_path,
            output_path=output_path,
            sheet_name=(
                0 if input_path.suffix.lower() == ".csv" else base_config.sheet_name
            ),
        )
        try:
            summary = process_file(config, progress=notify)
            results.append(summary)
            rows.append(
                {
                    "sample_file": str(input_path),
                    "status": "Completed",
                    "input_rows": summary.input_rows,
                    "predictable_rows": summary.predictable_rows,
                    "output_file": str(summary.output_path),
                    "error": "",
                }
            )
        except Exception as exc:
            rows.append(
                {
                    "sample_file": str(input_path),
                    "status": "Failed",
                    "input_rows": np.nan,
                    "predictable_rows": np.nan,
                    "output_file": "",
                    "error": str(exc),
                }
            )
            notify(f"Failed: {input_path.name}: {exc}")
    summary_path = output_directory / "PFAS_batch_summary.xlsx"
    pd.DataFrame(rows).to_excel(summary_path, index=False)
    failed_count = sum(row["status"] == "Failed" for row in rows)
    return BatchProcessingSummary(
        sample_count=len(paths),
        succeeded_count=len(results),
        failed_count=failed_count,
        output_directory=output_directory,
        summary_path=summary_path,
        results=tuple(results),
    )


def process_file(
    config: ProcessingConfig,
    progress: Callable[[str], None] | None = None,
) -> ProcessingSummary:
    mode = config.ion_mode.upper()
    screening_strategy = config.screening_strategy.strip().lower()
    if screening_strategy not in SUPPORTED_SCREENING_STRATEGIES:
        raise ProcessingError(
            f"Unsupported screening strategy: {config.screening_strategy}"
        )
    uses_model = screening_strategy != "linear_only"
    prediction_model = config.prediction_model.strip().lower()
    if uses_model and (not prediction_model or prediction_model == "none"):
        raise ProcessingError(
                "No Machine Learning Model (MLM) was selected. "
                "Select an installed MLM before processing."
        )
    notify = progress or (lambda _message: None)
    notify("Reading input table...")
    source = read_input_table(config.input_path, config.sheet_name)
    if source.empty:
        raise ProcessingError("The selected sheet contains no rows.")

    original_mz_column = _selected_or_detected(source, config.mz_column, "mz", mode)
    notify("Applying optional mass calibration...")
    (
        mass_source,
        mass_formula,
        mass_r2,
        mass_calibration_table,
    ) = apply_mass_calibration(source, config, mode, original_mz_column)
    mz_column = "PFAS_mz_calibrated"
    notify("Applying CCS calibration...")
    (
        stage1_source,
        first_formula,
        first_r2,
        calibration_table,
        _ccs_source,
    ) = apply_first_calibration(mass_source, config, mode, mz_column)
    blank_removed = pd.DataFrame()
    if config.blank_subtraction_enabled:
        if len(config.blank_paths) < 2:
            raise ProcessingError(
                "Blank subtraction requires at least two blank replicate files."
            )
        notify(
            f"Preparing {len(config.blank_paths):,} blank replicates for subtraction..."
        )
        blank_replicates = [
            _prepare_blank_replicate(Path(path), config, mode)
            for path in config.blank_paths
        ]
        notify(
            "Matching sample peaks to blanks by calibrated m/z, RT, and calibrated CCS..."
        )
        stage1_source, blank_removed = apply_blank_subtraction(
            stage1_source, blank_replicates, config, mode
        )
        notify(
            f"Blank subtraction retained {len(stage1_source):,} rows and removed "
            f"{len(blank_removed):,} rows."
        )
        if stage1_source.empty:
            notify("All rows were removed by blank subtraction; writing audit output.")
            empty_output = stage1_source.copy()
            empty_output["PFAS_ion_mode"] = mode
            empty_output["PFAS_final_candidate"] = pd.Series(dtype=int)
            empty_output["PFAS_final_screen_label"] = pd.Series(dtype=str)
            summary_table = pd.DataFrame(
                [
                    ("Input file", str(config.input_path)),
                    ("Input sheet", str(config.sheet_name)),
                    ("Ion mode", mode),
                    ("Input rows", len(source)),
                    ("Output rows", 0),
                    ("Blank subtraction", True),
                    ("Blank replicate files", len(config.blank_paths)),
                    ("Blank detection threshold", "blank mean intensity + 3 SD"),
                    ("Rows removed by blank subtraction", len(blank_removed)),
                    ("Processing status", "All rows removed by blank subtraction"),
                ],
                columns=["Setting", "Value"],
            )
            config.output_path.parent.mkdir(parents=True, exist_ok=True)
            with pd.ExcelWriter(config.output_path, engine="openpyxl") as writer:
                empty_output.to_excel(writer, sheet_name="Labeled data", index=False)
                summary_table.to_excel(writer, sheet_name="Run summary", index=False)
                blank_removed.to_excel(
                    writer, sheet_name="Blank-subtracted rows", index=False
                )
                if mass_calibration_table is not None:
                    mass_calibration_table.to_excel(
                        writer, sheet_name="Mass calibration", index=False
                    )
                if calibration_table is not None:
                    calibration_table.to_excel(
                        writer, sheet_name="Calibration curve", index=False
                    )
            return ProcessingSummary(
                input_rows=len(source),
                predictable_rows=0,
                level_counts={
                    "Level 0 Not PFAS": 0,
                    "Level 1 Candidate PFAS": 0,
                    "Level 2 Medium confidence": 0,
                    "Level 3 High-confidence PFAS": 0,
                    "Not predictable": 0,
                },
                rimp1_method=config.rimp1_method,
                first_calibration_formula=first_formula,
                first_calibration_r2=first_r2,
                second_calibration_applied=False,
                prediction_model=config.prediction_model,
                prediction_model_path=None,
                output_path=config.output_path,
            )
    prepared, used_rimp1_method = _calculate_first_isotopic_peak_ratio(
        stage1_source,
        config,
        mz_column,
        notify,
    )

    if "_PFAS_mz" not in prepared:
        prepared["_PFAS_input_row"] = np.arange(1, len(prepared) + 1)
        prepared["_PFAS_mz"] = _numeric(prepared[mz_column])
    if "_PFAS_CCS_input" not in prepared:
        prepared["_PFAS_CCS_input"] = _numeric(prepared["PFAS_CCS_stage1"])

    second_calibration_applied = False
    prepared["PFAS_ion_mode"] = mode
    prepared["PFAS_mz_used"] = prepared["_PFAS_mz"]
    prepared["PFAS_CCS_stage1"] = prepared["_PFAS_CCS_input"]
    prepared["PFAS_CCS_model_input"] = prepared["PFAS_CCS_stage1"]
    prepared["PFAS_CCS_model_input_formula"] = (
        "CCS calibration result used directly as model input"
    )
    prepared["PFAS_first_isotopic_peak_ratio_used"] = _numeric(
        prepared["first_isotopic_peak_ratio"]
    )
    prepared["PFAS_first_isotopic_peak_ratio_method"] = used_rimp1_method
    if "PFAS_rMp1" not in prepared:
        prepared["PFAS_rMp1"] = prepared[
            "PFAS_first_isotopic_peak_ratio_used"
        ]
    for order in (2, 3, 4):
        target = f"PFAS_rMp{order}"
        if target in prepared:
            continue
        source_column = next(
            (
                column
                for column in [f"rMp{order}", f"riMp{order}"]
                if column in prepared
            ),
            None,
        )
        prepared[target] = (
            _numeric(prepared[source_column])
            if source_column
            else np.nan
        )
    for order in ISOTOPE_ORDERS:
        prepared[f"PFAS_riMp{order}"] = _numeric(prepared[f"PFAS_rMp{order}"])
        intensity_target = f"PFAS_M_plus_{order}_intensity_sum"
        if intensity_target not in prepared:
            prepared[intensity_target] = np.nan
    if "PFAS_isotope_role" not in prepared:
        prepared["PFAS_isotope_role"] = (
            prepared["isotope_role"].astype(str)
            if "isotope_role" in prepared
            else "M"
        )

    raw = pd.DataFrame(
        {
            "mz": prepared["PFAS_mz_used"],
            "CCS": prepared["PFAS_CCS_model_input"],
            "riMp1": prepared["PFAS_riMp1"],
            "riMp2": prepared["PFAS_riMp2"],
            "riMp3": prepared["PFAS_riMp3"],
            "riMp4": prepared["PFAS_riMp4"],
        }
    )
    notify("Propagating uncertainty through the linear PFAS screen...")
    prepared = prepared.join(
        mc_linear_screening(
            raw,
            mode,
            config.mc_iterations,
            abs(float(config.rimp1_uncertainty_percent)) / 100.0,
            abs(float(config.ccs_uncertainty_percent)) / 100.0,
            notify,
        )
    )
    prediction_model_path: Path | None = None
    bundle: dict | None = None
    model_feature_columns: list[str] = []
    validated_rimp1_uncertainty = np.nan
    validated_ccs_uncertainty = np.nan
    model_candidate_threshold = np.nan
    model_high_confidence_threshold = np.nan
    valid = pd.Series(False, index=prepared.index)
    validation_status = "Not applicable: model screening skipped"
    prepared["PFAS_processing_note"] = "Model screening skipped by selected strategy"

    if uses_model:
        prediction_model_path = model_path(mode, prediction_model)
        bundle = load_bundle(mode, prediction_model)
        model_high_confidence_threshold = float(bundle["t_low_fdr10"])
        model_candidate_threshold = float(
            bundle.get("t_high_fdr18", bundle["t_high_fdr20"])
        )
        model_record = _model_registry()[prediction_model]
        validated_rimp1_uncertainty = float(
            model_record.get("rimp1_uncertainty_percent", np.nan)
        )
        validated_ccs_uncertainty = float(
            model_record.get("ccs_uncertainty_percent", np.nan)
        )
        formal_validation_condition = (
            np.isfinite(validated_rimp1_uncertainty)
            and np.isfinite(validated_ccs_uncertainty)
            and np.isclose(
                abs(float(config.rimp1_uncertainty_percent)),
                abs(validated_rimp1_uncertainty),
            )
            and np.isclose(
                abs(float(config.ccs_uncertainty_percent)),
                abs(validated_ccs_uncertainty),
            )
        )
        validation_status = (
            "Validated formal mode"
            if formal_validation_condition
            else "Exploratory / unvalidated uncertainty setting"
        )
        model_feature_columns = _bundle_feature_columns(bundle)
        required_raw_columns = _required_raw_columns(model_feature_columns)
        unsupported_required = sorted(required_raw_columns - set(raw.columns))
        if unsupported_required:
            raise ProcessingError(
                "Selected model requires input columns that this app cannot prepare: "
                + ", ".join(unsupported_required)
            )
        requirement_mask = pd.Series(True, index=raw.index)
        for column in sorted(required_raw_columns):
            numeric = _numeric(raw[column])
            requirement_mask &= numeric.notna() & (numeric > 0)
        valid = requirement_mask & (prepared["PFAS_isotope_role"] == "M")
        required_label = ", ".join(sorted(required_raw_columns))
        prepared["PFAS_processing_note"] = np.where(
            prepared["PFAS_isotope_role"] != "M",
            "Assigned as an isotope peak; only monoisotopic M features are screened",
            np.where(
                valid,
                "",
                f"Missing or non-positive model input: {required_label}",
            ),
        )
        if "PFAS_calibration_extrapolated" in prepared:
            extrapolated = prepared["PFAS_calibration_extrapolated"].fillna(False).astype(bool)
            prepared.loc[
                valid & extrapolated, "PFAS_processing_note"
            ] = "Prediction used CCS calibration outside the standards' tc or m/z range"
        if "PFAS_mass_calibration_extrapolated" in prepared:
            mass_extrapolated = prepared[
                "PFAS_mass_calibration_extrapolated"
            ].fillna(False).astype(bool)
            current_note = prepared["PFAS_processing_note"].fillna("").astype(str)
            mass_note = "Prediction used mass calibration outside the standards' m/z range"
            selected = valid & mass_extrapolated
            prepared.loc[selected, "PFAS_processing_note"] = np.where(
                current_note[selected] == "",
                mass_note,
                current_note[selected] + "; " + mass_note,
            )
    prepared["PFAS_prediction_validation_status"] = validation_status

    pmean = np.full(len(prepared), np.nan)
    pstd = np.full(len(prepared), np.nan)
    p05 = np.full(len(prepared), np.nan)
    level = np.full(len(prepared), -1, dtype=int)
    level_name = np.full(len(prepared), "Not predictable", dtype=object)

    predictable_count = int(valid.sum())
    if predictable_count:
        notify(
            f"Running {mode} prediction model for "
            f"{predictable_count:,} predictable rows..."
        )
        valid_raw = raw.loc[valid].reset_index(drop=True)
        predicted_mean, predicted_std, predicted_p05 = mc_predict(
            bundle,
            valid_raw,
            config.mc_iterations,
            abs(float(config.rimp1_uncertainty_percent)) / 100.0,
            abs(float(config.ccs_uncertainty_percent)) / 100.0,
            notify,
        )
        valid_indices = np.where(valid.to_numpy())[0]
        pmean[valid_indices] = predicted_mean
        pstd[valid_indices] = predicted_std
        p05[valid_indices] = predicted_p05
        predicted_level, predicted_names = assign_tier(
            predicted_mean,
            predicted_p05,
            model_high_confidence_threshold,
            model_candidate_threshold,
        )
        level[valid_indices] = predicted_level
        level_name[valid_indices] = predicted_names

    prepared["PFAS_prob_mean"] = pmean
    prepared["PFAS_prob_std"] = pstd
    prepared["PFAS_prob_p05"] = p05
    prepared["PFAS_MC_uncertainty"] = pmean - p05
    prepared["PFAS_level"] = level
    prepared["PFAS_level_name"] = level_name
    prepared["PFAS_candidate"] = np.where(valid, (level >= 1).astype(int), np.nan)
    prepared["PFAS_high_confidence"] = np.where(
        valid, (level == 3).astype(int), np.nan
    )
    prepared["PFAS_screening_strategy"] = screening_strategy
    prepared["PFAS_final_candidate"] = combine_screening_results(
        prepared["PFAS_candidate"],
        prepared["PFAS_linear_screen_pass"],
        screening_strategy,
    )
    prepared.loc[
        prepared["PFAS_isotope_role"] != "M", "PFAS_final_candidate"
    ] = 0
    prepared["PFAS_final_screen_label"] = np.where(
        prepared["PFAS_final_candidate"] == 1,
        "Selected PFAS candidate",
        "Rejected",
    )
    plot_confidence = assign_plot_confidence(prepared, screening_strategy)
    for column in plot_confidence.columns:
        prepared[column] = plot_confidence[column]

    if "_PFAS_intensity" in prepared:
        prepared["PFAS_M_intensity"] = _numeric(prepared["_PFAS_intensity"])
    else:
        intensity_column = _selected_or_detected(
            prepared,
            config.intensity_column,
            "intensity",
            mode,
            required=False,
        )
        prepared["PFAS_M_intensity"] = (
            _numeric(prepared[intensity_column])
            if intensity_column
            else np.nan
        )

    prepared = prepared.drop(
        columns=[
            "_PFAS_input_row",
            "_PFAS_mz",
            "_PFAS_rt",
            "_PFAS_tc",
            "_PFAS_intensity",
            "_PFAS_CCS_input",
        ],
        errors="ignore",
    )

    prioritized = prioritize_pfas_peaks(
        prepared,
        minimum_intensity=float(config.minimum_m_intensity),
        top_n=int(config.priority_top_n),
    )
    prepared["PFAS_priority_selected"] = 0
    prepared["PFAS_intensity_rank"] = np.nan
    if not prioritized.empty:
        prepared.loc[prioritized.index, "PFAS_priority_selected"] = 1
        prepared.loc[prioritized.index, "PFAS_intensity_rank"] = prioritized[
            "PFAS_intensity_rank"
        ]

    database_matches = pd.DataFrame()
    database_path_used = config.database_path or default_database_path()
    if config.library_matching_enabled and not prioritized.empty:
        if not database_path_used.exists():
            raise ProcessingError(
                f"PFAS database file was not found: {database_path_used}"
            )
        notify(
            f"Matching {len(prioritized):,} prioritized peaks to "
            f"{database_path_used.name}..."
        )
        database_sheet = (
            0
            if database_path_used.suffix.lower() == ".csv"
            else config.database_sheet_name
        )
        database = prepare_pfas_database(
            read_input_table(database_path_used, database_sheet)
        )
        database_matches = match_pfas_database(
            prioritized,
            database,
            mode,
            float(config.database_mass_tolerance_ppm),
            float(config.database_ccs_tolerance_percent),
            bool(config.isotope_ratio_matching),
            abs(float(config.isotope_ratio_tolerance_percent)),
        )

    prepared["PFAS_database_match_count"] = 0
    if not database_matches.empty:
        match_counts = database_matches.groupby(
            "PFAS_peak_output_index"
        ).size()
        prepared.loc[
            match_counts.index,
            "PFAS_database_match_count",
        ] = match_counts.astype(int)
        prioritized["PFAS_database_match_count"] = (
            prioritized.index.to_series().map(match_counts).fillna(0).astype(int)
        )

        best = (
            database_matches.sort_values(
                [
                    "PFAS_peak_output_index",
                    "PFAS_absolute_mass_error_ppm",
                    "PFAS_mean_absolute_isotope_delta",
                ],
                kind="mergesort",
            )
            .drop_duplicates("PFAS_peak_output_index")
            .set_index("PFAS_peak_output_index")
        )
        best_columns = {
            "PFAS_ion_hypothesis": "PFAS_best_ion_hypothesis",
            "PFAS_mass_error_ppm": "PFAS_best_mass_error_ppm",
            "PFAS_CCS_error_percent": "PFAS_best_CCS_error_percent",
            "PFAS_isotope_ratio_match_status": "PFAS_best_isotope_ratio_match_status",
            "preferred_name": "PFAS_best_database_name",
            "molecular_formula": "PFAS_best_molecular_formula",
            "SMILES": "PFAS_best_SMILES",
            "theoretical_rMp1": "PFAS_best_theoretical_rMp1",
            "theoretical_rMp2": "PFAS_best_theoretical_rMp2",
            "theoretical_rMp3": "PFAS_best_theoretical_rMp3",
            "theoretical_rMp4": "PFAS_best_theoretical_rMp4",
        }
        prioritized = prioritized.join(
            best[list(best_columns)].rename(columns=best_columns),
            how="left",
        )
    else:
        prioritized["PFAS_database_match_count"] = 0

    counts = {
        name: int((level_name == name).sum())
        for name in [
            "Level 0 Not PFAS",
            "Level 1 Candidate PFAS",
            "Level 2 Medium confidence",
            "Level 3 High-confidence PFAS",
            "Not predictable",
        ]
    }
    summary_table = pd.DataFrame(
        [
            ("Input file", str(config.input_path)),
            ("Input sheet", str(config.sheet_name)),
            ("Ion mode", mode),
            ("Input rows", len(source)),
            ("Output rows", len(prepared)),
            ("Blank subtraction", config.blank_subtraction_enabled),
            ("Blank replicate files", len(config.blank_paths)),
            (
                "Blank matching m/z tolerance (+/- ppm)",
                config.blank_mz_tolerance_ppm,
            ),
            (
                "Blank matching RT tolerance (+/- min)",
                config.blank_rt_tolerance_min,
            ),
            (
                "Blank matching CCS tolerance (+/- %)",
                config.blank_ccs_tolerance_percent,
            ),
            ("Blank detection threshold", "blank mean intensity + 3 SD"),
            ("Rows removed by blank subtraction", len(blank_removed)),
            ("Predictable rows", predictable_count),
            ("Screening strategy", screening_strategy),
            (
                "Linear screening formula",
                prepared["PFAS_linear_screen_formula"].iloc[0],
            ),
            (
                "Linear-screen positive rows",
                int((_numeric(prepared["PFAS_linear_screen_pass"]) == 1).sum()),
            ),
            (
                "Linear candidate probability threshold",
                LINEAR_CANDIDATE_PROBABILITY,
            ),
            (
                "Linear high-confidence probability threshold",
                LINEAR_HIGH_CONFIDENCE_PROBABILITY,
            ),
            (
                "Linear riMp1 uncertainty (+/- %)",
                abs(float(config.rimp1_uncertainty_percent)),
            ),
            (
                "Linear CCS uncertainty (+/- %)",
                abs(float(config.ccs_uncertainty_percent)),
            ),
            (
                "Final selected PFAS rows",
                int((prepared["PFAS_final_candidate"] == 1).sum()),
            ),
            ("Priority minimum M intensity", config.minimum_m_intensity),
            ("Priority Top N", config.priority_top_n),
            ("Prioritized final-positive M peaks", len(prioritized)),
            ("Advanced PFAS library matching", config.library_matching_enabled),
            (
                "PFAS library",
                str(database_path_used) if config.library_matching_enabled else "Skipped",
            ),
            (
                "PFAS library mass tolerance (+/- ppm)",
                config.database_mass_tolerance_ppm,
            ),
            (
                "PFAS library CCS tolerance (+/- %)",
                config.database_ccs_tolerance_percent,
            ),
            ("PFAS library isotope ratio matching", config.isotope_ratio_matching),
            (
                "PFAS library mass hypotheses",
                (
                    "M and [M+H]+ (observed m/z - 1.007276466621 Da)"
                    if mode == "POS"
                    else "M and [M-H]- (observed m/z + 1.007276466621 Da)"
                ),
            ),
            (
                "PFAS library isotope ratio CI (+/- %)",
                abs(float(config.isotope_ratio_tolerance_percent)),
            ),
            ("PFAS library match rows", len(database_matches)),
            ("First isotopic peak ratio method", used_rimp1_method),
            ("Exact 13C spacing (Da)", FIRST_ISOTOPE_SPACING_DA),
            ("Exact 13C audit tolerance (+/- Da)", MASS_TOLERANCE_DA),
            (
                "M+1 to M+4 nominal cluster window (+/- Da)",
                HIGHER_ISOTOPE_WINDOW_DA,
            ),
            ("Isotope RT tolerance (+/- min)", RT_TOLERANCE_MIN),
            ("Isotope 3D_TC tolerance (+/-)", TC_TOLERANCE),
            ("Mass calibration", mass_formula),
            ("Lock mass search tolerance (+/- ppm)", config.lock_mass_tolerance_ppm),
            ("Mass calibration R2", mass_r2),
            ("CCS calibration", first_formula),
            ("CCS calibration R2", first_r2),
            (
                "Machine Learning Model (MLM)",
                MODEL_DESCRIPTIONS.get(prediction_model, prediction_model)
                if uses_model
                else "Skipped (linear screening only)",
            ),
            (
                "Prediction model path",
                str(prediction_model_path) if prediction_model_path else "Skipped",
            ),
            ("Prediction model feature columns", ", ".join(model_feature_columns)),
            (
                "Model validated riMp1 uncertainty (+/- %)",
                validated_rimp1_uncertainty,
            ),
            (
                "Model validated CCS uncertainty (+/- %)",
                validated_ccs_uncertainty,
            ),
            ("Prediction validation status", validation_status),
            ("MLM candidate probability threshold", model_candidate_threshold),
            (
                "MLM high-confidence probability threshold",
                model_high_confidence_threshold,
            ),
            ("Calibration gas", "Nitrogen"),
            ("Nitrogen mass (Da)", NITROGEN_MASS_DA),
            (
                "Calibration transformation",
                "reduced_CCS = CCS * sqrt(mu_N2) / |z|",
            ),
            ("CCS model input", prepared["PFAS_CCS_model_input_formula"].iloc[0]),
            ("riMp1 uncertainty (+/- %)", config.rimp1_uncertainty_percent),
            ("CCS uncertainty (+/- %)", config.ccs_uncertainty_percent),
            ("Monte Carlo iterations", config.mc_iterations),
            ("Level 0", counts["Level 0 Not PFAS"]),
            ("Level 1", counts["Level 1 Candidate PFAS"]),
            ("Level 2", counts["Level 2 Medium confidence"]),
            ("Level 3", counts["Level 3 High-confidence PFAS"]),
            ("Not predictable", counts["Not predictable"]),
        ],
        columns=["Item", "Value"],
    )

    notify("Writing labeled Excel output...")
    config.output_path.parent.mkdir(parents=True, exist_ok=True)
    try:
        with pd.ExcelWriter(config.output_path, engine="openpyxl") as writer:
            prepared.to_excel(writer, sheet_name="Labeled data", index=False)
            summary_table.to_excel(writer, sheet_name="Run summary", index=False)
            if config.blank_subtraction_enabled:
                blank_removed.to_excel(
                    writer, sheet_name="Blank-subtracted rows", index=False
                )
            if mass_calibration_table is not None:
                mass_calibration_table.to_excel(
                    writer, sheet_name="Mass calibration", index=False
                )
            if calibration_table is not None:
                calibration_table.to_excel(
                    writer, sheet_name="Calibration curve", index=False
                )
                add_calibration_chart(
                    writer.book["Calibration curve"],
                    calibration_table,
                    float(first_r2),
                )
            prioritized.to_excel(
                writer,
                sheet_name="Prioritized PFAS peaks",
                index=False,
            )
            add_pfas_scatter_plots(writer.book, prepared, screening_strategy)
            if config.library_matching_enabled:
                database_matches.to_excel(
                    writer,
                    sheet_name="PFAS database matches",
                    index=False,
                )
    except Exception as exc:
        raise ProcessingError(f"Cannot write output workbook: {exc}") from exc
    notify("Complete.")
    return ProcessingSummary(
        input_rows=len(source),
        predictable_rows=predictable_count,
        level_counts=counts,
        rimp1_method=used_rimp1_method,
        first_calibration_formula=first_formula,
        first_calibration_r2=first_r2,
        second_calibration_applied=second_calibration_applied,
        prediction_model=prediction_model,
        prediction_model_path=prediction_model_path,
        output_path=config.output_path,
    )
