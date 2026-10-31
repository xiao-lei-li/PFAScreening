from __future__ import annotations

import queue
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from .engine import (
    BatchProcessingSummary,
    ProcessingConfig,
    ProcessingError,
    available_prediction_models,
    default_database_path,
    list_excel_sheets,
    process_batch,
    process_file,
    read_input_table,
    suggested_columns,
)


APP_TITLE = "PFAS CCS Screening"
NONE_VALUE = "(not selected)"
AUTODETECTED_VALUE = "Autodetected"
CALIBRATED_CCS_VALUE = "Calibrated (from Processing CCS calibration)"

RIMP1_METHODS = {
    "Auto detect": "auto",
    "Use existing first isotopic peak ratio column": "existing",
    "Calculate from monoisotopic and first-isotope intensities": "abundance",
    "Pair M+1 to M+4 isotope peaks from m/z, RT, 3D_TC, intensity": "pairing",
}

MASS_CALIBRATION_METHODS = {
    "No mass calibration": "none",
    "Lock mass ppm offset": "lock",
    "Linear correction from standards": "linear",
}

FIRST_CALIBRATION_METHODS = {
    "cIMS/TWIMS power-law calibration (tC, m/z, CCS)": "cims_reduced_ccs",
    "Use an existing CCS column": "existing_ccs",
}

RIMP1_UNCERTAINTY_OPTIONS = ["-20%", "-15%", "-10%", "-5%", "5%", "10%", "15%", "20%"]
CCS_UNCERTAINTY_OPTIONS = ["-10%", "-5%", "-3%", "3%", "5%", "10%"]
MC_ITERATION_OPTIONS = ["200", "150", "100", "50"]

SCREENING_STRATEGIES = {
    "Machine Learning Model (MLM) only": "model_only",
    "Linear relationship only": "linear_only",
    "MLM AND linear relationship (intersection)": "intersection",
    "MLM OR linear relationship (union)": "union",
}

PFAS_LIBRARY_METHODS = {
    "PFAS master list": "master",
    "Customized library": "custom",
}

NO_MODEL_LABEL = "Select a Machine Learning Model (MLM)"


class ScreeningApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("1080x1100")
        self.minsize(820, 700)
        self.option_add("*Font", ("Segoe UI", 10))

        self.input_var = tk.StringVar()
        self.output_var = tk.StringVar()
        self.sample_count_var = tk.StringVar(value="No sample files selected.")
        self.sample_paths: list[Path] = []
        self.blank_subtraction_var = tk.BooleanVar(value=False)
        self.blank_files_var = tk.StringVar(value="No blank files selected.")
        self.blank_paths: list[Path] = []
        self.blank_mz_tolerance_var = tk.DoubleVar(value=5.0)
        self.blank_rt_tolerance_var = tk.DoubleVar(value=0.10)
        self.blank_ccs_tolerance_var = tk.DoubleVar(value=3.0)
        self.output_label_var = tk.StringVar(value="Output file")
        self.sheet_var = tk.StringVar()
        self.mode_var = tk.StringVar(value="POS")
        self.rimp1_method_var = tk.StringVar(value=next(iter(RIMP1_METHODS)))
        self.mass_calibration_var = tk.StringVar(
            value=next(iter(MASS_CALIBRATION_METHODS))
        )
        self.lock_mass_value_var = tk.StringVar()
        self.lock_mass_tolerance_var = tk.DoubleVar(value=1000.0)
        self.first_calibration_var = tk.StringVar(
            value=next(iter(FIRST_CALIBRATION_METHODS))
        )
        self.model_options = available_prediction_models()
        self.model_label_to_id = {
            str(option["label"]): str(option["id"]) for option in self.model_options
        }
        self.prediction_model_var = tk.StringVar(value=NO_MODEL_LABEL)
        self.screening_strategy_var = tk.StringVar(value=next(iter(SCREENING_STRATEGIES)))
        self.screening_info_var = tk.StringVar()
        self.model_info_var = tk.StringVar(
            value=(
                "No MLM selected. Installed MLM choices use m/z, CCS, and riMp1; "
                "they do not use riMp2-riMp4."
            )
        )
        self.standards_path_var = tk.StringVar()
        self.standards_sheet_var = tk.StringVar()
        self.standards_tc_var = tk.StringVar(value=NONE_VALUE)
        self.standards_mz_var = tk.StringVar(value=NONE_VALUE)
        self.standards_ccs_var = tk.StringVar(value=NONE_VALUE)
        self.standards_charge_var = tk.StringVar(value=NONE_VALUE)
        self.mass_standards_path_var = tk.StringVar()
        self.mass_standards_sheet_var = tk.StringVar()
        self.mass_observed_mz_var = tk.StringVar(value=NONE_VALUE)
        self.mass_exact_mz_var = tk.StringVar(value=NONE_VALUE)
        self.default_charge_var = tk.IntVar(value=1)
        self.rimp1_uncertainty_var = tk.StringVar(value="20%")
        self.ccs_uncertainty_var = tk.StringVar(value="3%")
        self.mc_var = tk.StringVar(value="200")
        self.minimum_intensity_var = tk.DoubleVar(value=0.0)
        self.priority_top_n_var = tk.IntVar(value=200)
        self.library_matching_var = tk.BooleanVar(value=False)
        self.library_method_var = tk.StringVar(value=next(iter(PFAS_LIBRARY_METHODS)))
        self.database_mz_tolerance_var = tk.DoubleVar(value=5.0)
        self.database_ccs_tolerance_var = tk.DoubleVar(value=5.0)
        self.isotope_ratio_matching_var = tk.BooleanVar(value=False)
        self.isotope_ratio_tolerance_var = tk.DoubleVar(value=20.0)
        self.database_path_var = tk.StringVar(value=str(default_database_path()))
        self.database_sheet_var = tk.StringVar(value="Worksheet1")
        self.status_var = tk.StringVar(value="Choose an Excel file to begin.")
        self.column_vars = {
            field: tk.StringVar(value=NONE_VALUE)
            for field in [
                "mz",
                "ccs",
                "tc",
                "charge",
                "rt",
                "intensity",
                "rimp1",
                "m_intensity",
                "mp1_intensity",
            ]
        }
        self.column_boxes: dict[str, ttk.Combobox] = {}
        self.event_queue: queue.Queue[tuple[str, object]] = queue.Queue()
        self._build_ui()
        self.after(100, self._drain_events)

    def _build_ui(self) -> None:
        root = ttk.Frame(self)
        root.pack(fill="both", expand=True)
        root.rowconfigure(0, weight=1)
        root.columnconfigure(0, weight=1)

        canvas = tk.Canvas(root, highlightthickness=0)
        scrollbar = ttk.Scrollbar(root, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.grid(row=0, column=0, sticky="nsew")
        scrollbar.grid(row=0, column=1, sticky="ns")

        outer = ttk.Frame(canvas, padding=18)
        scroll_window = canvas.create_window((0, 0), window=outer, anchor="nw")

        def _update_scroll_region(_event: tk.Event) -> None:
            canvas.configure(scrollregion=canvas.bbox("all"))

        def _fit_scroll_width(event: tk.Event) -> None:
            canvas.itemconfigure(scroll_window, width=event.width)

        def _mousewheel(event: tk.Event) -> None:
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

        outer.bind("<Configure>", _update_scroll_region)
        canvas.bind("<Configure>", _fit_scroll_width)
        canvas.bind_all("<MouseWheel>", _mousewheel)
        outer.columnconfigure(1, weight=1)

        bottom = ttk.Frame(root, padding=(18, 8, 18, 12))
        bottom.grid(row=1, column=0, columnspan=2, sticky="ew")
        bottom.columnconfigure(1, weight=1)

        title = ttk.Label(outer, text=APP_TITLE, font=("Segoe UI Semibold", 19))
        title.grid(row=0, column=0, columnspan=3, sticky="w")
        subtitle = ttk.Label(
            outer,
            text=(
                "Preprocess an Excel feature table, calculate experimental isotope "
                "ratios, calibrate CCS, and assign PFAS probability tiers."
            ),
            foreground="#4a5568",
        )
        subtitle.grid(row=1, column=0, columnspan=3, sticky="w", pady=(2, 18))

        ttk.Label(outer, text="Sample file(s)").grid(row=2, column=0, sticky="w", pady=5)
        ttk.Entry(outer, textvariable=self.input_var).grid(
            row=2, column=1, sticky="ew", padx=8
        )
        ttk.Button(outer, text="Choose sample(s)...", command=self._browse_input).grid(
            row=2, column=2, sticky="ew"
        )
        ttk.Label(
            outer, textvariable=self.sample_count_var, foreground="#4a5568"
        ).grid(row=3, column=1, columnspan=2, sticky="w", padx=8)

        ttk.Label(outer, text="Worksheet").grid(row=4, column=0, sticky="w", pady=5)
        self.sheet_box = ttk.Combobox(
            outer, textvariable=self.sheet_var, state="readonly"
        )
        self.sheet_box.grid(row=4, column=1, sticky="ew", padx=8)
        self.sheet_box.bind("<<ComboboxSelected>>", lambda _event: self._load_columns())

        blank = ttk.LabelFrame(outer, text="Batch blank subtraction", padding=12)
        blank.grid(row=5, column=0, columnspan=3, sticky="ew", pady=(10, 8))
        blank.columnconfigure(1, weight=1)
        ttk.Checkbutton(
            blank,
            text="Use multiple blank replicates (blank mean + 3 SD)",
            variable=self.blank_subtraction_var,
            command=self._update_blank_mode,
        ).grid(row=0, column=0, columnspan=4, sticky="w", pady=(0, 6))
        self.blank_files_label = ttk.Label(blank, text="Blank files")
        self.blank_files_label.grid(row=1, column=0, sticky="w", pady=3)
        self.blank_files_entry = ttk.Entry(
            blank, textvariable=self.blank_files_var, state="readonly"
        )
        self.blank_files_entry.grid(row=1, column=1, columnspan=2, sticky="ew", padx=8)
        self.blank_browse_button = ttk.Button(
            blank, text="Choose blanks...", command=self._browse_blanks
        )
        self.blank_browse_button.grid(row=1, column=3)
        self.blank_mz_label = ttk.Label(blank, text="m/z tolerance (ppm)")
        self.blank_mz_label.grid(row=2, column=0, sticky="w", pady=3)
        self.blank_mz_entry = ttk.Entry(
            blank, textvariable=self.blank_mz_tolerance_var, width=12
        )
        self.blank_mz_entry.grid(row=2, column=1, sticky="w", padx=8)
        self.blank_rt_label = ttk.Label(blank, text="RT tolerance (min)")
        self.blank_rt_label.grid(row=2, column=2, sticky="w", padx=(18, 0), pady=3)
        self.blank_rt_entry = ttk.Entry(
            blank, textvariable=self.blank_rt_tolerance_var, width=12
        )
        self.blank_rt_entry.grid(row=2, column=3, sticky="w", padx=8)
        self.blank_ccs_label = ttk.Label(blank, text="CCS tolerance (%)")
        self.blank_ccs_label.grid(row=3, column=0, sticky="w", pady=3)
        self.blank_ccs_entry = ttk.Entry(
            blank, textvariable=self.blank_ccs_tolerance_var, width=12
        )
        self.blank_ccs_entry.grid(row=3, column=1, sticky="w", padx=8)
        ttk.Label(
            blank,
            text=(
                "For each sample peak, the nearest blank peak within all three "
                "tolerances is used. Missing blank peaks count as zero intensity."
            ),
            foreground="#4a5568",
            wraplength=900,
        ).grid(row=4, column=0, columnspan=4, sticky="w", pady=(6, 0))
        self.blank_controls = [
            self.blank_files_label,
            self.blank_files_entry,
            self.blank_browse_button,
            self.blank_mz_label,
            self.blank_mz_entry,
            self.blank_rt_label,
            self.blank_rt_entry,
            self.blank_ccs_label,
            self.blank_ccs_entry,
        ]

        settings = ttk.LabelFrame(outer, text="Processing settings", padding=12)
        settings.grid(row=6, column=0, columnspan=3, sticky="ew", pady=(14, 8))
        settings.columnconfigure(1, weight=1)
        settings.columnconfigure(3, weight=1)

        ttk.Label(settings, text="Ion mode").grid(row=0, column=0, sticky="w", pady=5)
        mode_frame = ttk.Frame(settings)
        mode_frame.grid(row=0, column=1, sticky="w")
        for value, text in [("POS", "Positive [M+H]+"), ("NEG", "Negative [M-H]-")]:
            ttk.Radiobutton(
                mode_frame,
                text=text,
                value=value,
                variable=self.mode_var,
                command=self._on_ion_mode_changed,
            ).pack(side="left", padx=(0, 18))

        ttk.Label(settings, text="rMp1 source").grid(
            row=1, column=0, sticky="w", pady=5
        )
        ttk.Combobox(
            settings,
            textvariable=self.rimp1_method_var,
            values=list(RIMP1_METHODS),
            state="readonly",
        ).grid(row=1, column=1, sticky="ew", padx=(10, 0))

        ttk.Label(settings, text="CCS calibration").grid(
            row=2, column=0, sticky="w", pady=5
        )
        self.first_calibration_box = ttk.Combobox(
            settings,
            textvariable=self.first_calibration_var,
            values=list(FIRST_CALIBRATION_METHODS),
            state="readonly",
        )
        self.first_calibration_box.grid(row=2, column=1, sticky="ew", padx=(10, 0))
        self.first_calibration_box.bind(
            "<<ComboboxSelected>>", lambda _event: self._update_ccs_mapping_mode()
        )

        ttk.Label(settings, text="Mass calibration").grid(
            row=3, column=0, sticky="w", pady=5
        )
        self.mass_calibration_box = ttk.Combobox(
            settings,
            textvariable=self.mass_calibration_var,
            values=list(MASS_CALIBRATION_METHODS),
            state="readonly",
        )
        self.mass_calibration_box.grid(row=3, column=1, sticky="ew", padx=(10, 0))
        self.mass_calibration_box.bind(
            "<<ComboboxSelected>>", lambda _event: self._update_mass_calibration_mode()
        )
        ttk.Label(settings, text="Lock mass exact m/z").grid(
            row=3, column=2, sticky="w", padx=(24, 0), pady=5
        )
        self.lock_mass_entry = ttk.Entry(
            settings, textvariable=self.lock_mass_value_var, width=14
        )
        self.lock_mass_entry.grid(row=3, column=3, sticky="ew", padx=(10, 0))

        ttk.Label(settings, text="Lock mass search tolerance (ppm)").grid(
            row=4, column=2, sticky="w", padx=(24, 0), pady=5
        )
        self.lock_mass_tolerance_entry = ttk.Entry(
            settings, textvariable=self.lock_mass_tolerance_var, width=14
        )
        self.lock_mass_tolerance_entry.grid(
            row=4, column=3, sticky="ew", padx=(10, 0)
        )
        ttk.Label(settings, text="PFAS screening method").grid(
            row=4, column=0, sticky="w", pady=5
        )
        self.screening_strategy_box = ttk.Combobox(
            settings,
            textvariable=self.screening_strategy_var,
            values=list(SCREENING_STRATEGIES),
            state="readonly",
        )
        self.screening_strategy_box.grid(
            row=4, column=1, sticky="ew", padx=(10, 0)
        )
        self.screening_strategy_box.bind(
            "<<ComboboxSelected>>",
            lambda _event: self._update_screening_strategy(),
        )

        ttk.Label(settings, text="Machine Learning Model (MLM)").grid(
            row=5, column=0, sticky="w", pady=5
        )
        self.model_box = ttk.Combobox(
            settings,
            textvariable=self.prediction_model_var,
            values=[NO_MODEL_LABEL] + list(self.model_label_to_id),
            state="readonly",
        )
        self.model_box.grid(row=5, column=1, columnspan=3, sticky="ew", padx=(10, 0))
        self.model_box.bind(
            "<<ComboboxSelected>>", lambda _event: self._on_model_selected()
        )
        ttk.Label(
            settings,
            textvariable=self.model_info_var,
            foreground="#8a4b00",
            wraplength=780,
        ).grid(row=6, column=0, columnspan=4, sticky="w", pady=(0, 6))

        ttk.Label(settings, text="MC iterations").grid(
            row=7, column=0, sticky="w", pady=5
        )
        self.mc_box = ttk.Combobox(
            settings,
            textvariable=self.mc_var,
            values=MC_ITERATION_OPTIONS,
            state="readonly",
            width=10,
        )
        self.mc_box.grid(row=7, column=1, sticky="w", padx=(10, 0))
        ttk.Label(settings, text="riMp1 uncertainty (%)").grid(
            row=7, column=2, sticky="w", padx=(24, 0), pady=5
        )
        self.rimp1_uncertainty_box = ttk.Combobox(
            settings,
            textvariable=self.rimp1_uncertainty_var,
            values=RIMP1_UNCERTAINTY_OPTIONS,
            state="readonly",
            width=12,
        )
        self.rimp1_uncertainty_box.grid(
            row=7, column=3, sticky="ew", padx=(10, 0)
        )

        ttk.Label(settings, text="CCS uncertainty (%)").grid(
            row=8, column=0, sticky="w", pady=5
        )
        self.ccs_uncertainty_box = ttk.Combobox(
            settings,
            textvariable=self.ccs_uncertainty_var,
            values=CCS_UNCERTAINTY_OPTIONS,
            state="readonly",
            width=12,
        )
        self.ccs_uncertainty_box.grid(
            row=8, column=1, sticky="w", padx=(10, 0)
        )
        ttk.Label(settings, text="Minimum M intensity").grid(
            row=8, column=2, sticky="w", padx=(24, 0), pady=5
        )
        ttk.Entry(
            settings,
            textvariable=self.minimum_intensity_var,
            width=14,
        ).grid(row=8, column=3, sticky="ew", padx=(10, 0))

        ttk.Label(settings, text="Default |charge|").grid(
            row=9, column=0, sticky="w", pady=5
        )
        ttk.Spinbox(
            settings,
            from_=1,
            to=20,
            textvariable=self.default_charge_var,
            width=10,
        ).grid(row=9, column=1, sticky="w", padx=(10, 0))
        ttk.Label(settings, text="Top N intense PFAS peaks").grid(
            row=9, column=2, sticky="w", padx=(24, 0), pady=5
        )
        ttk.Spinbox(
            settings,
            from_=1,
            to=100000,
            textvariable=self.priority_top_n_var,
            width=14,
        ).grid(row=9, column=3, sticky="ew", padx=(10, 0))

        ttk.Label(
            settings,
            textvariable=self.screening_info_var,
            foreground="#174a7e",
            wraplength=980,
            justify="left",
        ).grid(row=10, column=0, columnspan=4, sticky="w", pady=(8, 2))

        mapping = ttk.LabelFrame(outer, text="Advanced column mapping", padding=12)
        mapping.grid(row=9, column=0, columnspan=3, sticky="ew", pady=8)
        mapping.columnconfigure(1, weight=1)
        labels = [
            ("mz", "m/z"),
            ("ccs", "Existing CCS / calibrated result"),
            ("tc", "3D_TC / corrected arrival time"),
            ("charge", "Charge (optional; default = 1)"),
            ("rt", "Retention time"),
            ("intensity", "Base peak intensity"),
            ("rimp1", "Existing rMp1 / first isotopic peak ratio"),
            ("m_intensity", "Monoisotopic peak intensity"),
            ("mp1_intensity", "First isotopic peak intensity"),
        ]
        for row, (field, label) in enumerate(labels):
            ttk.Label(mapping, text=label).grid(row=row, column=0, sticky="w", pady=3)
            box = ttk.Combobox(
                mapping, textvariable=self.column_vars[field], state="readonly"
            )
            box.grid(row=row, column=1, sticky="ew", padx=(10, 0), pady=3)
            self.column_boxes[field] = box

        library = ttk.LabelFrame(
            outer, text="Advanced PFAS library matching", padding=12
        )
        library.grid(row=10, column=0, columnspan=3, sticky="ew", pady=8)
        library.columnconfigure(1, weight=1)
        ttk.Checkbutton(
            library,
            text="Use advanced PFAS library matching",
            variable=self.library_matching_var,
            command=self._update_library_mode,
        ).grid(row=0, column=0, columnspan=4, sticky="w", pady=(0, 6))
        library_controls: list[tk.Widget] = []
        library_label = ttk.Label(library, text="PFAS library")
        library_label.grid(row=1, column=0, sticky="w", pady=3)
        library_controls.append(library_label)
        library_mode = ttk.Frame(library)
        library_mode.grid(row=1, column=1, columnspan=2, sticky="w", padx=8)
        for label in PFAS_LIBRARY_METHODS:
            radio = ttk.Radiobutton(
                library_mode,
                text=label,
                value=label,
                variable=self.library_method_var,
                command=self._update_library_mode,
            )
            radio.pack(side="left", padx=(0, 18))
            library_controls.append(radio)
        file_label = ttk.Label(library, text="Library file")
        file_label.grid(row=2, column=0, sticky="w", pady=3)
        library_controls.append(file_label)
        self.database_entry = ttk.Entry(library, textvariable=self.database_path_var)
        self.database_entry.grid(row=2, column=1, sticky="ew", padx=8)
        database_browse = ttk.Button(
            library,
            text="Browse...",
            command=self._browse_database,
        )
        database_browse.grid(row=2, column=2)
        sheet_label = ttk.Label(library, text="Worksheet")
        sheet_label.grid(row=3, column=0, sticky="w", pady=3)
        self.database_sheet_box = ttk.Combobox(
            library,
            textvariable=self.database_sheet_var,
            state="readonly",
            values=["Worksheet1"],
        )
        self.database_sheet_box.grid(row=3, column=1, columnspan=2, sticky="ew", padx=8)
        mz_tolerance_label = ttk.Label(library, text="Mass tolerance (ppm)")
        mz_tolerance_label.grid(row=4, column=0, sticky="w", pady=3)
        mz_tolerance_entry = ttk.Entry(
            library, textvariable=self.database_mz_tolerance_var, width=12
        )
        mz_tolerance_entry.grid(row=4, column=1, sticky="w", padx=8)
        ccs_tolerance_label = ttk.Label(library, text="CCS tolerance (%)")
        ccs_tolerance_label.grid(row=4, column=2, sticky="w", padx=(18, 0), pady=3)
        ccs_tolerance_entry = ttk.Entry(
            library, textvariable=self.database_ccs_tolerance_var, width=12
        )
        ccs_tolerance_entry.grid(row=4, column=3, sticky="w", padx=8)
        isotope_matching_check = ttk.Checkbutton(
            library,
            text="Match by isotope ratios after library matching",
            variable=self.isotope_ratio_matching_var,
            command=self._update_library_mode,
        )
        isotope_matching_check.grid(row=5, column=0, columnspan=4, sticky="w", pady=(6, 0))
        isotope_tolerance_label = ttk.Label(
            library, text="Isotope ratio tolerance (%)"
        )
        isotope_tolerance_label.grid(row=6, column=0, sticky="w", pady=3)
        self.isotope_tolerance_entry = ttk.Entry(
            library, textvariable=self.isotope_ratio_tolerance_var, width=12
        )
        self.isotope_tolerance_entry.grid(row=6, column=1, sticky="w", padx=8)
        library_controls.extend(
            [
                file_label,
                self.database_entry,
                database_browse,
                sheet_label,
                self.database_sheet_box,
                mz_tolerance_label,
                mz_tolerance_entry,
                ccs_tolerance_label,
                ccs_tolerance_entry,
                isotope_matching_check,
                isotope_tolerance_label,
                self.isotope_tolerance_entry,
            ]
        )
        self.isotope_tolerance_controls = [
            isotope_tolerance_label,
            self.isotope_tolerance_entry,
        ]
        self.library_controls = library_controls

        mass_standards = ttk.LabelFrame(
            outer, text="Mass calibration standards", padding=12
        )
        self.mass_standards_frame = mass_standards
        mass_standards.grid(row=7, column=0, columnspan=3, sticky="ew", pady=8)
        mass_standards.columnconfigure(1, weight=1)
        ttk.Label(mass_standards, text="Mass standards file").grid(
            row=0, column=0, sticky="w", pady=3
        )
        ttk.Entry(mass_standards, textvariable=self.mass_standards_path_var).grid(
            row=0, column=1, sticky="ew", padx=8
        )
        ttk.Button(
            mass_standards, text="Browse...", command=self._browse_mass_standards
        ).grid(row=0, column=2)
        ttk.Label(mass_standards, text="Worksheet").grid(
            row=1, column=0, sticky="w", pady=3
        )
        self.mass_standards_sheet_box = ttk.Combobox(
            mass_standards,
            textvariable=self.mass_standards_sheet_var,
            state="readonly",
        )
        self.mass_standards_sheet_box.grid(row=1, column=1, sticky="ew", padx=8)
        self.mass_standards_sheet_box.bind(
            "<<ComboboxSelected>>", lambda _event: self._load_mass_standard_columns()
        )
        ttk.Label(mass_standards, text="Observed m/z column").grid(
            row=2, column=0, sticky="w", pady=3
        )
        self.mass_observed_mz_box = ttk.Combobox(
            mass_standards, textvariable=self.mass_observed_mz_var, state="readonly"
        )
        self.mass_observed_mz_box.grid(row=2, column=1, sticky="ew", padx=8)
        ttk.Label(mass_standards, text="Exact m/z column").grid(
            row=3, column=0, sticky="w", pady=3
        )
        self.mass_exact_mz_box = ttk.Combobox(
            mass_standards, textvariable=self.mass_exact_mz_var, state="readonly"
        )
        self.mass_exact_mz_box.grid(row=3, column=1, sticky="ew", padx=8)

        standards = ttk.LabelFrame(
            outer, text="CCS calibration standards", padding=12
        )
        self.ccs_standards_frame = standards
        standards.grid(row=8, column=0, columnspan=3, sticky="ew", pady=8)
        standards.columnconfigure(1, weight=1)
        ttk.Label(standards, text="Standards file").grid(
            row=0, column=0, sticky="w", pady=3
        )
        ttk.Entry(standards, textvariable=self.standards_path_var).grid(
            row=0, column=1, sticky="ew", padx=8
        )
        ttk.Button(
            standards, text="Browse...", command=self._browse_standards
        ).grid(row=0, column=2)
        ttk.Label(standards, text="Worksheet").grid(
            row=1, column=0, sticky="w", pady=3
        )
        self.standards_sheet_box = ttk.Combobox(
            standards, textvariable=self.standards_sheet_var, state="readonly"
        )
        self.standards_sheet_box.grid(row=1, column=1, sticky="ew", padx=8)
        self.standards_sheet_box.bind(
            "<<ComboboxSelected>>", lambda _event: self._load_standard_columns()
        )
        ttk.Label(standards, text="Standard tc column").grid(
            row=2, column=0, sticky="w", pady=3
        )
        self.standards_tc_box = ttk.Combobox(
            standards, textvariable=self.standards_tc_var, state="readonly"
        )
        self.standards_tc_box.grid(row=2, column=1, sticky="ew", padx=8)
        ttk.Label(standards, text="Standard m/z column").grid(
            row=3, column=0, sticky="w", pady=3
        )
        self.standards_mz_box = ttk.Combobox(
            standards, textvariable=self.standards_mz_var, state="readonly"
        )
        self.standards_mz_box.grid(row=3, column=1, sticky="ew", padx=8)
        ttk.Label(standards, text="Reference CCS column").grid(
            row=4, column=0, sticky="w", pady=3
        )
        self.standards_ccs_box = ttk.Combobox(
            standards, textvariable=self.standards_ccs_var, state="readonly"
        )
        self.standards_ccs_box.grid(row=4, column=1, sticky="ew", padx=8)
        ttk.Label(standards, text="Standard charge (optional)").grid(
            row=5, column=0, sticky="w", pady=3
        )
        self.standards_charge_box = ttk.Combobox(
            standards, textvariable=self.standards_charge_var, state="readonly"
        )
        self.standards_charge_box.grid(row=5, column=1, sticky="ew", padx=8)
        ttk.Label(
            standards,
            text=(
                "Uses CCS' = CCS*sqrt(mu)/|z| and fits CCS' = A*tC^B. "
                "Requires at least 3 standards with distinct positive corrected tC values."
            ),
            foreground="#4a5568",
            wraplength=900,
        ).grid(row=6, column=0, columnspan=3, sticky="w", pady=(6, 0))

        ttk.Label(bottom, textvariable=self.output_label_var).grid(
            row=0, column=0, sticky="w", pady=3
        )
        ttk.Entry(bottom, textvariable=self.output_var).grid(
            row=0, column=1, sticky="ew", padx=8
        )
        ttk.Button(bottom, text="Browse...", command=self._browse_output).grid(
            row=0, column=2, sticky="ew"
        )

        action_frame = ttk.Frame(bottom)
        action_frame.grid(row=1, column=0, columnspan=3, sticky="ew", pady=(8, 4))
        action_frame.columnconfigure(0, weight=1)
        self.progress = ttk.Progressbar(action_frame, mode="indeterminate")
        self.progress.grid(row=0, column=0, sticky="ew", padx=(0, 12))
        self.run_button = ttk.Button(
            action_frame, text="Process and write output", command=self._start_processing
        )
        self.run_button.grid(row=0, column=1)

        ttk.Label(bottom, textvariable=self.status_var, foreground="#2d5f8b").grid(
            row=2, column=0, columnspan=3, sticky="w", pady=(2, 4)
        )
        self.log = tk.Text(
            bottom,
            height=4,
            wrap="word",
            state="disabled",
            background="#f7f9fb",
            relief="solid",
            borderwidth=1,
        )
        self.log.grid(row=3, column=0, columnspan=3, sticky="ew")
        self._update_library_mode()
        self._update_mass_calibration_mode()
        self._update_screening_strategy()
        self._update_ccs_mapping_mode()
        self._update_blank_mode()

    def _browse_input(self) -> None:
        paths = filedialog.askopenfilenames(
            title="Choose one or more sample feature tables",
            filetypes=[
                ("Excel workbooks", "*.xlsx *.xlsm *.xls"),
                ("CSV files", "*.csv"),
                ("All files", "*.*"),
            ],
        )
        if not paths:
            return
        self.sample_paths = [Path(path) for path in paths]
        source = self.sample_paths[0]
        self.input_var.set(str(source))
        if len(self.sample_paths) == 1:
            self.sample_count_var.set(f"1 sample selected: {source.name}")
            self.output_label_var.set("Output file")
            self.output_var.set(
                str(source.with_name(f"{source.stem}_PFAS_labeled.xlsx"))
            )
        else:
            self.sample_count_var.set(
                f"{len(self.sample_paths)} samples selected; column mapping is read from {source.name}."
            )
            self.output_label_var.set("Batch output folder")
            self.output_var.set(str(source.parent / "PFAS_batch_output"))
        try:
            if source.suffix.lower() == ".csv":
                sheets = ["CSV"]
            else:
                sheets = list_excel_sheets(source)
            self.sheet_box["values"] = sheets
            self.sheet_var.set(sheets[0])
            self._load_columns()
            self.status_var.set(f"Loaded {source.name}")
        except ProcessingError as exc:
            messagebox.showerror(APP_TITLE, str(exc))

    def _browse_blanks(self) -> None:
        paths = filedialog.askopenfilenames(
            title="Choose at least two blank replicate tables",
            filetypes=[
                ("Excel workbooks", "*.xlsx *.xlsm *.xls"),
                ("CSV files", "*.csv"),
                ("All files", "*.*"),
            ],
        )
        if not paths:
            return
        self.blank_paths = [Path(path) for path in paths]
        names = ", ".join(path.name for path in self.blank_paths[:3])
        if len(self.blank_paths) > 3:
            names += f", and {len(self.blank_paths) - 3} more"
        self.blank_files_var.set(f"{len(self.blank_paths)} blanks: {names}")

    def _update_blank_mode(self) -> None:
        state = "normal" if self.blank_subtraction_var.get() else "disabled"
        for widget in self.blank_controls:
            try:
                widget.configure(state=state)
            except tk.TclError:
                pass
        if self.blank_subtraction_var.get():
            self.blank_files_entry.configure(state="readonly")

    def _browse_output(self) -> None:
        if len(self.sample_paths) > 1:
            initial = Path(self.output_var.get()) if self.output_var.get() else None
            path = filedialog.askdirectory(
                title="Choose batch output folder",
                initialdir=str(initial.parent) if initial else None,
            )
            if path:
                self.output_var.set(path)
            return
        initial = Path(self.output_var.get()) if self.output_var.get() else None
        path = filedialog.asksaveasfilename(
            title="Save labeled output",
            defaultextension=".xlsx",
            initialdir=str(initial.parent) if initial else None,
            initialfile=initial.name if initial else "PFAS_labeled.xlsx",
            filetypes=[("Excel workbook", "*.xlsx")],
        )
        if path:
            self.output_var.set(path)

    def _browse_standards(self) -> None:
        path = filedialog.askopenfilename(
            title="Choose cIMS calibration standards table",
            filetypes=[
                ("Excel workbooks", "*.xlsx *.xlsm *.xls"),
                ("CSV files", "*.csv"),
                ("All files", "*.*"),
            ],
        )
        if not path:
            return
        self.standards_path_var.set(path)
        source = Path(path)
        try:
            sheets = (
                ["CSV"]
                if source.suffix.lower() == ".csv"
                else list_excel_sheets(source)
            )
            self.standards_sheet_box["values"] = sheets
            self.standards_sheet_var.set(sheets[0])
            self._load_standard_columns()
        except ProcessingError as exc:
            messagebox.showerror(APP_TITLE, str(exc))

    def _browse_database(self) -> None:
        path = filedialog.askopenfilename(
            title="Choose PFAS chemical database",
            filetypes=[
                ("Excel workbooks", "*.xlsx *.xlsm *.xls"),
                ("CSV files", "*.csv"),
                ("All files", "*.*"),
            ],
        )
        if not path:
            return
        self.library_method_var.set("Customized library")
        self.database_path_var.set(path)
        source = Path(path)
        try:
            sheets = (
                ["CSV"]
                if source.suffix.lower() == ".csv"
                else list_excel_sheets(source)
            )
            self.database_sheet_box["values"] = sheets
            self.database_sheet_var.set(sheets[0])
        except ProcessingError as exc:
            messagebox.showerror(APP_TITLE, str(exc))

    def _update_library_mode(self) -> None:
        if PFAS_LIBRARY_METHODS[self.library_method_var.get()] == "master":
            master_path = default_database_path()
            self.database_path_var.set(str(master_path))
            self.database_sheet_box["values"] = ["Worksheet1"]
            self.database_sheet_var.set("Worksheet1")
        enabled = bool(self.library_matching_var.get())
        for widget in self.library_controls:
            try:
                widget.configure(state="normal" if enabled else "disabled")
            except tk.TclError:
                pass
        if enabled:
            self.database_sheet_box.configure(state="readonly")
        isotope_enabled = enabled and bool(self.isotope_ratio_matching_var.get())
        for widget in self.isotope_tolerance_controls:
            widget.configure(state="normal" if isotope_enabled else "disabled")

    def _update_mass_calibration_mode(self) -> None:
        method = MASS_CALIBRATION_METHODS[self.mass_calibration_var.get()]
        self.lock_mass_entry.configure(state="normal" if method == "lock" else "disabled")
        self.lock_mass_tolerance_entry.configure(
            state="normal" if method == "lock" else "disabled"
        )
        if method == "linear":
            self.mass_standards_frame.grid()
        else:
            self.mass_standards_frame.grid_remove()

    def _update_ccs_mapping_mode(self) -> None:
        method = FIRST_CALIBRATION_METHODS[self.first_calibration_var.get()]
        current = self.column_vars["ccs"].get()
        if method == "cims_reduced_ccs":
            self.column_vars["ccs"].set(CALIBRATED_CCS_VALUE)
            self.ccs_standards_frame.grid()
        elif current == CALIBRATED_CCS_VALUE:
            self.column_vars["ccs"].set(NONE_VALUE)
            self.ccs_standards_frame.grid_remove()
        else:
            self.ccs_standards_frame.grid_remove()

    def _update_model_info(self) -> None:
        selected = self.prediction_model_var.get()
        model_id = self.model_label_to_id.get(selected)
        if not model_id:
            self.model_info_var.set(
                "No MLM selected. Processing will stop before MLM prediction."
            )
            self._update_screening_info()
            return
        option = next(
            (item for item in self.model_options if item["id"] == model_id),
            None,
        )
        if option is None:
            self.model_info_var.set("Selected MLM metadata was not found.")
            self._update_screening_info()
            return
        features = ", ".join(str(value) for value in option["feature_columns"])
        validated_rimp1 = option.get("rimp1_uncertainty_percent", "unknown")
        validated_ccs = option.get("ccs_uncertainty_percent", "unknown")
        self.model_info_var.set(
            f"{option['description']} Features: {features}. "
            f"Validated formal condition: riMp1 +/-{validated_rimp1}%, "
            f"CCS +/-{validated_ccs}%. Other uncertainty settings are exploratory."
        )
        self._update_screening_info()

    def _on_model_selected(self) -> None:
        model_id = self.model_label_to_id.get(self.prediction_model_var.get())
        option = next(
            (item for item in self.model_options if item["id"] == model_id),
            None,
        )
        if option is not None:
            try:
                rimp1 = abs(float(option["rimp1_uncertainty_percent"]))
                ccs = abs(float(option["ccs_uncertainty_percent"]))
                self.rimp1_uncertainty_var.set(f"{rimp1:g}%")
                self.ccs_uncertainty_var.set(f"{ccs:g}%")
            except (KeyError, TypeError, ValueError):
                pass
        self._update_model_info()

    def _on_ion_mode_changed(self) -> None:
        self._load_columns()
        self._update_screening_info()

    def _update_screening_info(self) -> None:
        strategy = SCREENING_STRATEGIES[self.screening_strategy_var.get()]
        mode = self.mode_var.get().upper()
        if mode == "POS":
            linear_formula = (
                "POS linear: CCS < 0.2537*mz + 84.701 AND "
                "CCS > 4.657*riMp1 + 98.570."
            )
        else:
            linear_formula = (
                "NEG linear: CCS < 0.24284*mz + 84.432 AND "
                "CCS > 4.415*riMp1 + 97.440."
            )
        linear_basis = (
            f"{linear_formula} Linear MC: pass probability >=80% = high-confidence; "
            "50% to <80% = uncertain candidate; <50% = non-PFAS."
        )

        selected_model_id = self.model_label_to_id.get(
            self.prediction_model_var.get()
        )
        option = next(
            (item for item in self.model_options if item["id"] == selected_model_id),
            None,
        )
        thresholds = option.get("thresholds", {}) if option else {}
        mode_thresholds = thresholds.get(mode, {}) if isinstance(thresholds, dict) else {}
        candidate = mode_thresholds.get("candidate")
        high = mode_thresholds.get("high_confidence")
        if candidate is None or high is None:
            mlm_basis = "MLM basis: select an MLM to display its probability thresholds."
        else:
            mlm_basis = (
                f"MLM {mode}: Level 0 if pmean < {candidate:g}; Level 1 if "
                f"{candidate:g} <= pmean < {high:g}; Level 2 if pmean >= {high:g} "
                f"but p05 < {high:g}; Level 3 if p05 >= {high:g}."
            )

        if strategy == "model_only":
            text = mlm_basis
        elif strategy == "linear_only":
            text = linear_basis
        elif strategy == "intersection":
            text = f"Intersection requires both to pass. {mlm_basis} {linear_basis}"
        else:
            text = f"Union requires either one to pass. {mlm_basis} {linear_basis}"
        self.screening_info_var.set(text)

    def _update_screening_strategy(self) -> None:
        strategy = SCREENING_STRATEGIES[self.screening_strategy_var.get()]
        if strategy == "linear_only":
            self.model_box.configure(state="disabled")
            self.mc_box.configure(state="readonly")
            self.rimp1_uncertainty_box.configure(state="readonly")
            self.ccs_uncertainty_box.configure(state="readonly")
            self.model_info_var.set(
                "Linear-only screening uses the ion-mode CCS boundaries and skips "
                "MLM prediction. MC iterations and riMp1/CCS uncertainty are used "
                "to calculate the linear pass probability."
            )
        else:
            self.model_box.configure(state="readonly")
            self.mc_box.configure(state="readonly")
            self.rimp1_uncertainty_box.configure(state="readonly")
            self.ccs_uncertainty_box.configure(state="readonly")
            self._update_model_info()
        self._update_screening_info()

    def _percent_value(self, value: str) -> float:
        text = str(value).strip().replace("%", "")
        return abs(float(text))

    def _browse_mass_standards(self) -> None:
        path = filedialog.askopenfilename(
            title="Choose mass calibration standards table",
            filetypes=[
                ("Excel workbooks", "*.xlsx *.xlsm *.xls"),
                ("CSV files", "*.csv"),
                ("All files", "*.*"),
            ],
        )
        if not path:
            return
        self.mass_standards_path_var.set(path)
        source = Path(path)
        try:
            sheets = (
                ["CSV"]
                if source.suffix.lower() == ".csv"
                else list_excel_sheets(source)
            )
            self.mass_standards_sheet_box["values"] = sheets
            self.mass_standards_sheet_var.set(sheets[0])
            self._load_mass_standard_columns()
        except ProcessingError as exc:
            messagebox.showerror(APP_TITLE, str(exc))

    def _load_mass_standard_columns(self) -> None:
        path_text = self.mass_standards_path_var.get().strip()
        if not path_text or not Path(path_text).exists():
            return
        try:
            path = Path(path_text)
            sheet: str | int = (
                0
                if path.suffix.lower() == ".csv"
                else self.mass_standards_sheet_var.get()
            )
            frame = read_input_table(path, sheet)
            values = [NONE_VALUE] + [str(column) for column in frame.columns]
            suggestions = suggested_columns(frame, self.mode_var.get())
            self.mass_observed_mz_box["values"] = values
            self.mass_exact_mz_box["values"] = values
            self.mass_observed_mz_var.set(
                suggestions.get("observed_mz") or NONE_VALUE
            )
            self.mass_exact_mz_var.set(suggestions.get("exact_mz") or NONE_VALUE)
        except ProcessingError as exc:
            messagebox.showerror(APP_TITLE, str(exc))

    def _load_standard_columns(self) -> None:
        path_text = self.standards_path_var.get().strip()
        if not path_text or not Path(path_text).exists():
            return
        try:
            path = Path(path_text)
            sheet: str | int = (
                0 if path.suffix.lower() == ".csv" else self.standards_sheet_var.get()
            )
            frame = read_input_table(path, sheet)
            values = [NONE_VALUE] + [str(column) for column in frame.columns]
            suggestions = suggested_columns(frame, self.mode_var.get())
            self.standards_tc_box["values"] = values
            self.standards_mz_box["values"] = values
            self.standards_ccs_box["values"] = values
            self.standards_charge_box["values"] = values
            self.standards_tc_var.set(suggestions.get("tc") or NONE_VALUE)
            self.standards_mz_var.set(suggestions.get("mz") or NONE_VALUE)
            self.standards_ccs_var.set(suggestions.get("ccs") or NONE_VALUE)
            self.standards_charge_var.set(
                suggestions.get("charge") or NONE_VALUE
            )
        except ProcessingError as exc:
            messagebox.showerror(APP_TITLE, str(exc))

    def _load_columns(self) -> None:
        path_text = self.input_var.get().strip()
        if not path_text or not Path(path_text).exists():
            return
        try:
            path = Path(path_text)
            sheet: str | int = 0 if path.suffix.lower() == ".csv" else self.sheet_var.get()
            frame = read_input_table(path, sheet)
            values = [NONE_VALUE] + [str(column) for column in frame.columns]
            suggestions = suggested_columns(frame, self.mode_var.get())
            for field, box in self.column_boxes.items():
                if field == "ccs":
                    field_values = [NONE_VALUE, CALIBRATED_CCS_VALUE] + [
                        str(column) for column in frame.columns
                    ]
                elif field in {"rimp1", "mp1_intensity"}:
                    field_values = [NONE_VALUE, AUTODETECTED_VALUE] + [
                        str(column) for column in frame.columns
                    ]
                else:
                    field_values = values
                box["values"] = field_values
                if field == "ccs" and FIRST_CALIBRATION_METHODS[self.first_calibration_var.get()] == "cims_reduced_ccs":
                    self.column_vars[field].set(CALIBRATED_CCS_VALUE)
                elif field in {"rimp1", "mp1_intensity"}:
                    self.column_vars[field].set(AUTODETECTED_VALUE)
                else:
                    self.column_vars[field].set(suggestions.get(field) or NONE_VALUE)
            self._append_log(
                f"Detected {len(frame):,} rows and {len(frame.columns)} columns."
            )
        except ProcessingError as exc:
            self.status_var.set(str(exc))

    def _selected_column(self, field: str) -> str | None:
        value = self.column_vars[field].get()
        if field == "ccs" and value == CALIBRATED_CCS_VALUE:
            return "__calibrated__"
        if field in {"rimp1", "mp1_intensity"} and value == AUTODETECTED_VALUE:
            return "__autodetected__"
        return None if value == NONE_VALUE or not value else value

    def _start_processing(self) -> None:
        input_path = Path(self.input_var.get().strip())
        output_path = Path(self.output_var.get().strip())
        sample_paths = (
            list(self.sample_paths)
            if self.sample_paths and self.sample_paths[0] == input_path
            else [input_path]
        )
        missing_samples = [path for path in sample_paths if not path.exists()]
        if missing_samples:
            messagebox.showerror(
                APP_TITLE,
                f"Sample file was not found: {missing_samples[0]}",
            )
            return
        if not self.output_var.get().strip():
            messagebox.showerror(APP_TITLE, "Choose an output file.")
            return
        library_matching_enabled = bool(self.library_matching_var.get())
        blank_subtraction_enabled = bool(self.blank_subtraction_var.get())
        try:
            mc_iterations = int(self.mc_var.get())
            default_charge = int(self.default_charge_var.get())
            priority_top_n = int(self.priority_top_n_var.get())
            minimum_intensity = float(self.minimum_intensity_var.get())
            rimp1_uncertainty = self._percent_value(self.rimp1_uncertainty_var.get())
            ccs_uncertainty = self._percent_value(self.ccs_uncertainty_var.get())
            database_mz_tolerance = (
                float(self.database_mz_tolerance_var.get())
                if library_matching_enabled
                else 5.0
            )
            database_ccs_tolerance = (
                float(self.database_ccs_tolerance_var.get())
                if library_matching_enabled
                else 5.0
            )
            isotope_ratio_tolerance = (
                float(self.isotope_ratio_tolerance_var.get())
                if library_matching_enabled
                and bool(self.isotope_ratio_matching_var.get())
                else 20.0
            )
            lock_mass_value = (
                float(self.lock_mass_value_var.get())
                if MASS_CALIBRATION_METHODS[self.mass_calibration_var.get()] == "lock"
                else None
            )
            lock_mass_tolerance = (
                float(self.lock_mass_tolerance_var.get())
                if MASS_CALIBRATION_METHODS[self.mass_calibration_var.get()] == "lock"
                else 1000.0
            )
            blank_mz_tolerance = (
                float(self.blank_mz_tolerance_var.get())
                if blank_subtraction_enabled
                else 5.0
            )
            blank_rt_tolerance = (
                float(self.blank_rt_tolerance_var.get())
                if blank_subtraction_enabled
                else 0.10
            )
            blank_ccs_tolerance = (
                float(self.blank_ccs_tolerance_var.get())
                if blank_subtraction_enabled
                else 3.0
            )
        except (TypeError, ValueError):
            messagebox.showerror(
                APP_TITLE,
                "Check MC iterations, charge, intensity threshold, Top N, uncertainty, blank tolerances, and library tolerance values.",
            )
            return
        if default_charge < 1:
            messagebox.showerror(APP_TITLE, "Default charge must be at least 1.")
            return
        if minimum_intensity < 0:
            messagebox.showerror(APP_TITLE, "Minimum M intensity cannot be negative.")
            return
        if priority_top_n < 1:
            messagebox.showerror(APP_TITLE, "Top N must be at least 1.")
            return
        if rimp1_uncertainty <= 0 or ccs_uncertainty <= 0:
            messagebox.showerror(APP_TITLE, "MLM uncertainty values must be positive.")
            return
        if library_matching_enabled and database_mz_tolerance <= 0:
            messagebox.showerror(APP_TITLE, "Library mass tolerance must be positive.")
            return
        if library_matching_enabled and database_ccs_tolerance <= 0:
            messagebox.showerror(APP_TITLE, "Library CCS tolerance must be positive.")
            return
        if (
            library_matching_enabled
            and bool(self.isotope_ratio_matching_var.get())
            and isotope_ratio_tolerance <= 0
        ):
            messagebox.showerror(APP_TITLE, "Isotope ratio tolerance must be positive.")
            return
        if lock_mass_value is not None and lock_mass_tolerance <= 0:
            messagebox.showerror(APP_TITLE, "Lock mass search tolerance must be positive.")
            return
        if blank_subtraction_enabled:
            if len(self.blank_paths) < 2:
                messagebox.showerror(
                    APP_TITLE,
                    "Blank subtraction requires at least two blank replicate files.",
                )
                return
            missing_blanks = [path for path in self.blank_paths if not path.exists()]
            if missing_blanks:
                messagebox.showerror(
                    APP_TITLE,
                    f"Blank file was not found: {missing_blanks[0]}",
                )
                return
            if min(
                blank_mz_tolerance,
                blank_rt_tolerance,
                blank_ccs_tolerance,
            ) <= 0:
                messagebox.showerror(
                    APP_TITLE, "All blank matching tolerances must be positive."
                )
                return
        mass_method = MASS_CALIBRATION_METHODS[self.mass_calibration_var.get()]
        if mass_method == "linear":
            mass_standards_path = Path(self.mass_standards_path_var.get().strip())
            if not mass_standards_path.exists():
                messagebox.showerror(
                    APP_TITLE,
                    "Choose a valid Mass calibration standards file.",
                )
                return
        ccs_method = FIRST_CALIBRATION_METHODS[self.first_calibration_var.get()]
        if ccs_method == "cims_reduced_ccs":
            ccs_standards_path = Path(self.standards_path_var.get().strip())
            if not ccs_standards_path.exists():
                messagebox.showerror(
                    APP_TITLE,
                    "Choose a valid CCS calibration standards file.",
                )
                return
        screening_strategy = SCREENING_STRATEGIES[self.screening_strategy_var.get()]
        prediction_model = self.model_label_to_id.get(
            self.prediction_model_var.get()
        )
        if screening_strategy != "linear_only" and not prediction_model:
            messagebox.showerror(
                APP_TITLE,
                (
                    "Select a Machine Learning Model (MLM) before processing.\n\n"
                    "The installed MLM choices use m/z, CCS, and riMp1 only. "
                    "No riMp1-riMp4 prediction model is currently installed."
                ),
            )
            return
        prediction_model = prediction_model or "none"
        library_method = PFAS_LIBRARY_METHODS[self.library_method_var.get()]
        database_path = (
            default_database_path()
            if library_method == "master"
            else Path(self.database_path_var.get().strip())
        )
        if library_matching_enabled and not database_path.exists():
            messagebox.showerror(APP_TITLE, "Choose a valid PFAS library file.")
            return

        sheet: str | int = (
            0 if input_path.suffix.lower() == ".csv" else self.sheet_var.get()
        )
        config = ProcessingConfig(
            input_path=input_path,
            output_path=output_path,
            sheet_name=sheet,
            ion_mode=self.mode_var.get(),
            rimp1_method=RIMP1_METHODS[self.rimp1_method_var.get()],
            mass_calibration_method=MASS_CALIBRATION_METHODS[
                self.mass_calibration_var.get()
            ],
            lock_mass_value=lock_mass_value,
            lock_mass_tolerance_ppm=lock_mass_tolerance,
            first_calibration_method=FIRST_CALIBRATION_METHODS[
                self.first_calibration_var.get()
            ],
            second_calibration_method="none",
            mz_column=self._selected_column("mz"),
            ccs_column=self._selected_column("ccs"),
            tc_column=self._selected_column("tc"),
            charge_column=self._selected_column("charge"),
            default_charge=default_charge,
            rt_column=self._selected_column("rt"),
            intensity_column=self._selected_column("intensity"),
            rimp1_column=self._selected_column("rimp1"),
            m_intensity_column=self._selected_column("m_intensity"),
            mp1_intensity_column=self._selected_column("mp1_intensity"),
            standards_path=(
                Path(self.standards_path_var.get().strip())
                if self.standards_path_var.get().strip()
                else None
            ),
            standards_sheet_name=(
                0
                if self.standards_path_var.get().lower().endswith(".csv")
                else self.standards_sheet_var.get()
            ),
            standards_tc_column=(
                None
                if self.standards_tc_var.get() == NONE_VALUE
                else self.standards_tc_var.get()
            ),
            standards_mz_column=(
                None
                if self.standards_mz_var.get() == NONE_VALUE
                else self.standards_mz_var.get()
            ),
            standards_ccs_column=(
                None
                if self.standards_ccs_var.get() == NONE_VALUE
                else self.standards_ccs_var.get()
            ),
            standards_charge_column=(
                None
                if self.standards_charge_var.get() == NONE_VALUE
                else self.standards_charge_var.get()
            ),
            mass_standards_path=(
                Path(self.mass_standards_path_var.get().strip())
                if self.mass_standards_path_var.get().strip()
                else None
            ),
            mass_standards_sheet_name=(
                0
                if self.mass_standards_path_var.get().lower().endswith(".csv")
                else self.mass_standards_sheet_var.get()
            ),
            mass_observed_mz_column=(
                None
                if self.mass_observed_mz_var.get() == NONE_VALUE
                else self.mass_observed_mz_var.get()
            ),
            mass_exact_mz_column=(
                None
                if self.mass_exact_mz_var.get() == NONE_VALUE
                else self.mass_exact_mz_var.get()
            ),
            prediction_model=prediction_model,
            screening_strategy=screening_strategy,
            rimp1_uncertainty_percent=rimp1_uncertainty,
            ccs_uncertainty_percent=ccs_uncertainty,
            minimum_m_intensity=minimum_intensity,
            priority_top_n=priority_top_n,
            library_matching_enabled=library_matching_enabled,
            database_path=database_path,
            database_sheet_name=(
                0
                if database_path.suffix.lower() == ".csv"
                else self.database_sheet_var.get()
            ),
            database_mass_tolerance_ppm=database_mz_tolerance,
            database_ccs_tolerance_percent=database_ccs_tolerance,
            isotope_ratio_matching=(
                library_matching_enabled
                and bool(self.isotope_ratio_matching_var.get())
            ),
            isotope_ratio_tolerance_percent=isotope_ratio_tolerance,
            mc_iterations=mc_iterations,
            blank_subtraction_enabled=blank_subtraction_enabled,
            blank_paths=tuple(self.blank_paths) if blank_subtraction_enabled else (),
            blank_sheet_name=sheet,
            blank_mz_tolerance_ppm=blank_mz_tolerance,
            blank_rt_tolerance_min=blank_rt_tolerance,
            blank_ccs_tolerance_percent=blank_ccs_tolerance,
            blank_sigma_multiplier=3.0,
        )
        self.run_button.configure(state="disabled")
        self.progress.start(12)
        self.status_var.set("Processing...")
        self._append_log("Starting screening run.")
        threading.Thread(
            target=self._worker,
            args=(config, sample_paths, output_path),
            daemon=True,
        ).start()

    def _worker(
        self,
        config: ProcessingConfig,
        sample_paths: list[Path],
        output_path: Path,
    ) -> None:
        try:
            progress = lambda message: self.event_queue.put(("progress", message))
            if len(sample_paths) > 1:
                summary = process_batch(
                    config,
                    sample_paths,
                    output_path,
                    progress=progress,
                )
            else:
                summary = process_file(config, progress=progress)
            self.event_queue.put(("complete", summary))
        except Exception as exc:
            self.event_queue.put(("error", exc))

    def _drain_events(self) -> None:
        try:
            while True:
                event, payload = self.event_queue.get_nowait()
                if event == "progress":
                    self.status_var.set(str(payload))
                    self._append_log(str(payload))
                elif event == "complete":
                    self.progress.stop()
                    self.run_button.configure(state="normal")
                    summary = payload
                    if isinstance(summary, BatchProcessingSummary):
                        self.status_var.set(
                            f"Batch complete: {summary.succeeded_count}/"
                            f"{summary.sample_count} samples succeeded."
                        )
                        self._append_log(
                            f"Saved batch summary to {summary.summary_path}"
                        )
                        messagebox.showinfo(
                            APP_TITLE,
                            (
                                "Batch screening complete.\n\n"
                                f"Succeeded: {summary.succeeded_count}\n"
                                f"Failed: {summary.failed_count}\n\n"
                                f"Batch summary:\n{summary.summary_path}"
                            ),
                        )
                        continue
                    self.status_var.set(
                        f"Complete: {summary.predictable_rows:,}/{summary.input_rows:,} "
                        "rows predicted."
                    )
                    self._append_log(
                        f"Saved labeled workbook to {summary.output_path}"
                    )
                    messagebox.showinfo(
                        APP_TITLE,
                        (
                            f"Screening complete.\n\n"
                            f"Predictable rows: {summary.predictable_rows:,}/"
                            f"{summary.input_rows:,}\n"
                            f"Level 3: "
                            f"{summary.level_counts['Level 3 High-confidence PFAS']:,}\n\n"
                            f"Output:\n{summary.output_path}"
                        ),
                    )
                elif event == "error":
                    self.progress.stop()
                    self.run_button.configure(state="normal")
                    message = (
                        str(payload)
                        if isinstance(payload, ProcessingError)
                        else f"Unexpected error: {payload}"
                    )
                    self.status_var.set(message)
                    self._append_log(message)
                    messagebox.showerror(APP_TITLE, message)
        except queue.Empty:
            pass
        self.after(100, self._drain_events)

    def _append_log(self, message: str) -> None:
        self.log.configure(state="normal")
        self.log.insert("end", message.rstrip() + "\n")
        self.log.see("end")
        self.log.configure(state="disabled")


def main() -> None:
    app = ScreeningApp()
    app.mainloop()
