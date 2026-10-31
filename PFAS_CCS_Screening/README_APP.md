# PFAS CCS Screening App

This desktop app processes Excel or CSV feature tables for PFAS screening. It can calibrate experimental CCS from cIMS standards, calculate or import the `first_isotopic_peak_ratio`, extract an experimental M+1 to M+4 isotope envelope, run the trained positive- or negative-mode model, and export a labeled Excel workbook.

## Batch Processing And Blank Subtraction

`Choose sample(s)...` accepts one or more sample files. All files in a batch use the selected worksheet, column mapping, mass/CCS calibration, model, linear-screening, and library-matching settings. Each sample receives its own `sample_name_PFAS_labeled.xlsx` output, and the selected output folder receives `PFAS_batch_summary.xlsx`. A failed sample is recorded in that summary while the remaining samples continue.

Blank subtraction requires at least two blank replicate files. Samples and blanks use the same worksheet name and column mapping. Mass and CCS calibration are applied before peaks are aligned using all three user-defined tolerances: calibrated m/z in ppm, RT in minutes, and calibrated CCS in percent. If several peaks in one blank satisfy the limits, the app uses the peak with the smallest normalized combined m/z/RT/CCS error. A missing peak in a blank replicate contributes zero intensity.

```text
blank threshold = blank mean intensity + 3 * blank sample SD
corrected sample intensity = max(sample raw intensity - blank mean intensity, 0)
```

Only peaks with `sample raw intensity >= blank threshold` are retained. Corrected intensity is used by downstream automatic isotope pairing, riMp1-riMp4 calculation, the minimum-intensity filter, and screening. If the user explicitly selects an existing riMp1 column, that supplied ratio is not recalculated.

`Labeled data` records the raw intensity, blank mean, blank SD, threshold, corrected intensity, matched-blank count, and pass status. Removed peaks are retained for audit in `Blank-subtracted rows`. An audit workbook is still written when every peak in a sample is removed.

## What The App Does

For each input feature, the app can:

1. Read `m/z`, corrected arrival time `tc` or existing CCS, and isotope information.
2. Optionally correct measured `m/z` using mass calibration standards.
3. Convert `tc` to CCS using a calibration curve built from standards.
4. Use the first calibrated experimental CCS directly as model input. The old fixed POS/NEG second CCS transformation has been removed from the app workflow.
5. Calculate or use `rMp1`; automatic pairing also calculates experimental `rMp2`, `rMp3`, and `rMp4`.
6. Predict PFAS probability only after the user explicitly selects an installed prediction model.
7. Select the most intense likely PFAS peaks and optionally match them to a PFAS library by m/z, CCS, and isotope ratios.
8. Assign a PFAS confidence level:

- `Level 0 Not PFAS`
- `Level 1 Candidate PFAS`
- `Level 2 Medium confidence`
- `Level 3 High-confidence PFAS`
- `Not predictable`

## Start The App

Use the portable Windows build:

```text
dist\PFAS_CCS_Screening\PFAS_CCS_Screening.exe
```

Keep the whole `dist\PFAS_CCS_Screening` folder together. Do not move only the `.exe`, because the models and Python runtime are stored in the same folder.

You can also run the Python entry point during development:

```powershell
python launch_pfas_app.py
```

## Input Files

### Sample Feature Table

The sample table should contain, at minimum:

- `m/z`
- Either corrected arrival time `tc` or an existing calibrated CCS column
- A way to provide `rMp1`

Optional but recommended:

- Charge column, if ions are not singly charged
- Retention time, `3D_TC`, and intensity columns for automatic envelope pairing

If no charge column is selected, the app uses the `Default |charge|` setting. The default is `1`, which matches the `[M+H]+` and `[M-H]-` workflows.

### Standards Table For CCS Calibration

For cIMS calibration, provide a standards table with:

- Standard `tc`
- Standard `m/z`
- Reference CCS in nitrogen
- Charge, optional

Use standards measured under the same ion mode and instrument conditions as the samples. The standards should cover the sample `tc` and `m/z` range as much as possible.

### Mass Calibration Standards

Mass calibration is optional. If used, provide a standards table with:

- Observed or measured `m/z`
- Known exact `m/z`
- Optional compound name or note columns

These standards should be measured in the same run or under the same acquisition conditions as the samples.

## Optional Mass Calibration

Mass calibration is applied before CCS calibration, isotope pairing, and model prediction. The output always keeps both raw and calibrated m/z. For `Lock mass ppm offset`, enter the known exact lock-mass m/z in `Lock mass exact m/z`. The app finds the nearest observed sample peak within the user-selected `Lock mass search tolerance (ppm)` and applies the single-point multiplicative correction to every m/z value. The default search tolerance is `1000 ppm`. The `Mass calibration standards` section is shown and required only when `Linear correction from standards` is selected.

### No Mass Calibration

The app uses the input `m/z` directly.

### Lock Mass PPM Offset

Enter the known exact lock-mass m/z. The app locates the closest observed peak in
the sample table and applies a one-point multiplicative correction:

```text
correction_factor = lock_exact_mz / lock_observed_mz
corrected_mz = observed_mz * correction_factor
```

The closest peak must be within the selected lock-mass search tolerance;
otherwise the app stops with an error instead of using an unrelated peak. The selected lock
peak, its original error, correction factor, and residual are recorded in the
`Mass calibration` sheet.

### Linear Correction From Standards

The app fits mass error as a function of observed m/z:

```text
ppm_error = a * observed_mz + b
corrected_mz = observed_mz / (1 + ppm_error / 1e6)
```

This is the recommended option when multiple standards span the sample m/z range.

Rows outside the mass standards' observed m/z range are still processed, but they are flagged as extrapolated.

## CCS Calibration

### IMS Standards To Experimental CCS

The app fits a cyclic/TWIMS power-law calibration from standards using reduced
CCS. The supplied `tc` must already be the corrected arrival time, with the
instrument-external and m/z-dependent transfer-time correction completed.
The `CCS calibration standards` section is shown and required only when
`cIMS/TWIMS power-law calibration (tC, m/z, CCS)` is selected. It is hidden when
an existing CCS column is used.

For each standard:

```text
m_ion = (m/z) * |z|
mu = m_ion * m_N2 / (m_ion + m_N2)
m_N2 = 28.0134 Da
reduced_CCS = CCS * sqrt(mu) / |z|
```

At least three standards with distinct positive `tc` values are required. The
fitted calibration is:

```text
reduced_CCS = A * tc^B
```

For each sample:

```text
CCS_stage1 = A * tc^B * |z| / sqrt(mu)
```

This is a no-time-offset power law because the input is corrected `tc`. Raw
arrival time must not be supplied to this mode without first applying its time
and m/z-dependent transfer correction.

Rows outside the standards' `tc` or `m/z` range are still predicted when possible, but they are flagged as extrapolated in the output.

### Use Existing CCS Instead

If your sample table already contains calibrated experimental CCS, choose:

```text
CCS calibration: Use an existing CCS column
```

Then map the CCS column in the advanced column mapping section. When `cIMS/TWIMS power-law calibration` is selected, the `Existing CCS / calibrated result` mapping shows `Calibrated (from Processing CCS calibration)`; this means the CCS produced during Processing is passed to prediction and library matching.

In Advanced column mapping, `Existing rMp1 / first isotopic peak ratio` and
`First isotopic peak intensity` include an explicit `Autodetected` option. This
selection is resolved during Processing using the supported column aliases; it
is not treated as a worksheet column name.

### Model Input CCS

The current app workflow does not apply a second CCS correction. The model input CCS is the calibrated or selected experimental CCS:

- CCS calculated from standards using `tc`, `m/z`, and reference CCS; or
- an existing calibrated CCS column selected by the user.

The old fixed POS/NEG transformations are no longer exposed in the GUI and are not applied during processing.

## M+1 Isotope Ratio Options

The app supports three ways to provide `rMp1`. The historical external names
`first_isotopic_peak_ratio` and `riMp1` remain accepted for compatibility.

### 1. Use an Existing First Isotopic Peak Ratio

Select a column that already contains `first_isotopic_peak_ratio`. Legacy
`riMp1` columns are accepted as input aliases.

### 2. Calculate From Monoisotopic and First-Isotope Intensities

Select the monoisotopic and first isotopic peak intensity columns. The app calculates:

```text
first_isotopic_peak_ratio =
    100 * I(first isotopic peak) / I(monoisotopic peak)
```

### 3. Pair the Experimental Isotope Envelope Automatically

Select `m/z`, retention time, `3D_TC`, and intensity. CCS is not used for
isotope pairing. The app builds mutually exclusive singly charged M+1 to M+4
envelopes using:

```text
nominal mass shifts: +1, +2, +3, and +4 Da
nominal cluster window: +/-0.020 Da
RT difference: +/-0.020 min
3D_TC difference: +/-0.060
```

Each feature can belong to only one isotope envelope, and an assigned M+n peak
cannot seed another envelope. Assignment conflicts are resolved by absolute RT
error, absolute `3D_TC` error, and absolute nominal mass error, in that order.
All assigned fine-structure peaks within each nominal cluster are summed:

```text
rMp1 = 100 * sum(I(M+1 fine-structure candidates)) / I(M)
rMp2 = 100 * sum(I(M+2 fine-structure candidates)) / I(M)
rMp3 = 100 * sum(I(M+3 fine-structure candidates)) / I(M)
rMp4 = 100 * sum(I(M+4 fine-structure candidates)) / I(M)
```

The RF-MC model uses `rMp1`. The app also reports `first_13C_peak_ratio`, which
selects the assigned M+1 peak nearest `+1.00335483507 Da` within `+/-0.010 Da`.
This single-peak value is an audit measurement and is not the model input.

Candidate counts, m/z ranges, mass-offset ranges, RT error ranges, `3D_TC`
error ranges, isotope roles, and ambiguity flags are retained for auditing.
The M+2 to M+4 ratios support preliminary isotope-pattern or elemental-class
assessment; they do not constitute structural identification.

## Prediction Model

The app does not silently choose a model. In `Processing settings`, select a
prediction model before processing. If no model is selected, processing stops
with an error.

The app includes two RF-MC model choices:

```text
RF-MC riMp1 +/-20% CCS +/-3%
RF-MC riMp1 +/-5% CCS +/-3%
```

Model files:

```text
models\rf_sklearn_bundle_POS.joblib
models\rf_sklearn_bundle_NEG.joblib
models\rf_sklearn_bundle_rimp1_5_ccs_3_POS.joblib
models\rf_sklearn_bundle_rimp1_5_ccs_3_NEG.joblib
```

Both are random-forest Monte Carlo models and use this feature set:

```text
mz_percentile, mz, CCS, riMp1, CCS_over_m, CCS_over_Mp1,
M_over_Mp1, CCS_over_sqrtm, CCS_over_m23
```

It does not use `rMp2`, `rMp3`, or `rMp4` for prediction. Those ratios are used
for output reporting and optional PFAS library isotope-ratio matching. A new
prediction model that directly uses `rMp1-rMp4` is not currently installed in
this app folder.

The older POS/NEG pair is validated for `riMp1 +/-20%` and `CCS +/-3%`. The
second pair is validated for `riMp1 +/-5%` and `CCS +/-3%`. Selecting a model
automatically sets the matching uncertainty controls. Other uncertainty values
remain available for sensitivity analysis, but the output is marked
`Exploratory / unvalidated uncertainty setting` and must not be interpreted
using either model's formal validation claim.

### Adding Future Uncertainty-Specific Models

Future models can be added without changing the GUI code. Put the model files in
the `models` folder and create or update:

```text
models\model_manifest.csv
```

Required columns:

```text
model_id,label,pos_path,neg_path,feature_cols
```

Recommended optional columns:

```text
description,rimp1_uncertainty_percent,ccs_uncertainty_percent,
pos_candidate_threshold,pos_high_confidence_threshold,
neg_candidate_threshold,neg_high_confidence_threshold
```

Example:

```csv
model_id,label,pos_path,neg_path,feature_cols,description,rimp1_uncertainty_percent,ccs_uncertainty_percent,pos_candidate_threshold,pos_high_confidence_threshold,neg_candidate_threshold,neg_high_confidence_threshold
rf_rimp5_ccs3,RF-MC riMp1 +/-5% CCS +/-3%,rf_rimp5_ccs3_POS.joblib,rf_rimp5_ccs3_NEG.joblib,"mz_percentile,mz,CCS,riMp1,CCS_over_m,CCS_over_Mp1,M_over_Mp1,CCS_over_sqrtm,CCS_over_m23",Model trained for riMp1 +/-5% and CCS +/-3%,5,3,0.555,0.947,0.55875,0.976
```

Paths may be file names relative to the `models` folder or absolute paths. The
app only lists manifest models when both POS and NEG files exist.

If a selected model requires unsupported feature columns, processing stops with
an explicit error. Supported direct isotope-ratio input feature names are
`riMp1`, `riMp2`, `riMp3`, `riMp4` and aliases `rMp1`, `rMp2`, `rMp3`, `rMp4`.

## Linear-Relationship PFAS Screening

The `PFAS screening method` setting provides four strategies:

- `Machine Learning Model (MLM) only`: use `PFAS Level >=1` from the selected MLM.
- `Linear relationship only`: skip model prediction and use only the equations below.
- `MLM AND linear relationship (intersection)`: require both methods to pass.
- `MLM OR linear relationship (union)`: accept a peak when either method passes.

A blue criteria panel below the processing settings updates whenever the user
changes the ion mode, screening method, or MLM. It displays the active POS/NEG
linear equations, linear probability levels, MLM `pmean`/`p05` thresholds, and
the selected intersection or union rule.

When `Linear relationship only` is selected, only the MLM selector is disabled.
MC iterations, riMp1 uncertainty, and CCS uncertainty remain editable and are
used to propagate measurement uncertainty through the linear equations.

The linear screen uses calibrated m/z, final CCS, and final experimental riMp1:

```text
POS: CCS < 0.2537*mz + 84.701 AND CCS > 4.657*riMp1 + 98.570
NEG: CCS < 0.24284*mz + 84.432 AND CCS > 4.415*riMp1 + 97.440
```

Both comparisons are strict, so a value exactly on a boundary does not pass.
The output retains both boundaries, the pass label, and signed residuals:

```text
upper residual = CCS - upper CCS boundary   (must be < 0 to pass)
lower residual = CCS - lower CCS boundary   (must be > 0 to pass)
```

For each Monte Carlo iteration, riMp1 and CCS are independently perturbed by the
selected relative uncertainties and the two strict inequalities are evaluated.
The linear pass probability is the fraction of iterations that pass both. A
probability `>=0.80` is labeled high-confidence, `0.50 <= p < 0.80` is an uncertain
candidate, and `<0.50` is linear non-PFAS. Both high-confidence and uncertain
candidates have `PFAS_linear_screen_pass=1`.

`PFAS_final_candidate` is the combined result selected by the user. Intensity
prioritization, Top N selection, and optional library matching use this final
candidate label.

## Intensity Prioritization And PFAS Library Matching

The app can prioritize the most abundant likely PFAS features using both:

- `Minimum M intensity`
- `Top N intense PFAS peaks`

Only monoisotopic `M` features with `PFAS_final_candidate=1` are eligible. They
are ranked by M intensity and then PFAS probability; linear-only runs have no
model probability and are therefore ordered by M intensity.

Advanced PFAS library matching is optional. Enable `Use advanced PFAS library matching` to run it; leave the option off to skip all database reading and matching. When enabled, the bundled default library is `Chemical List PFAS.xlsx`. Choose either:

- `PFAS master list`: the bundled `Chemical List PFAS.xlsx`.
- `Customized library`: a user-selected Excel or CSV file.

The library must contain molecular formula and neutral monoisotopic mass. Name,
SMILES, DTXSID, CASRN, InChIKey, and CCS are optional but are included in the
output when present. The default mass tolerance is `+/-5 ppm`; CCS tolerance is
entered as a percent. The bundled master list currently does not contain CCS, so
CCS filtering is only applied when the selected library has a CCS column.

Two ion interpretations are searched:

```text
POS: search the library with observed M and observed M - 1.007276466621 Da
NEG: search the library with observed M and observed M + 1.007276466621 Da
proton mass: 1.007276466621 Da
```

The selected mass tolerance is applied independently to both mass hypotheses.
The selected CCS tolerance is also enforced when both experimental and library
CCS values are available.

For every candidate, the app reports the preferred name, DTXSID, CASRN,
molecular formula, library neutral monoisotopic mass, library CCS when available,
experimental m/z, experimental CCS, ion hypothesis, mass shift used for the
library search, target neutral mass, mass error, and CCS error.

If `Match by isotope ratios after library matching` is enabled, the app calculates
theoretical nominal `rMp1` to `rMp4` from the molecular formula and compares them
with experimental `rMp1` to `rMp4`. The confidence interval is based on the
independent `Isotope ratio tolerance (%)`; for example, `20%` gives a theoretical
`rMpN +/-20%` interval. Candidates outside the isotope-ratio confidence interval
are removed from the PFAS library match sheet.

SMILES represents a possible database structure. A mass or molecular formula
may match multiple structures, so the app retains every candidate within the
selected tolerance and does not declare a unique identification.

## Step-By-Step Use

1. Open `PFAS_CCS_Screening.exe`.
2. Choose the input Excel or CSV feature table.
3. Choose the worksheet.
4. Select the ion mode:
   - `Positive [M+H]+`
   - `Negative [M-H]-`
5. Choose the first isotopic peak ratio source.
6. Choose the optional mass calibration method:
   - `No mass calibration`
   - `Lock mass ppm offset`
   - `Linear correction from standards`
7. Choose the first CCS calibration method:
   - `cIMS/TWIMS power-law calibration (tC, m/z, CCS)`
   - `Use an existing CCS column`
8. Choose `MLM only`, `Linear relationship only`, `MLM AND linear`, or `MLM OR linear`.
9. For lock-mass calibration, enter `Lock mass exact m/z`. For linear calibration, use the displayed Mass calibration standards section and map observed and exact `m/z`.
10. If using IMS CCS calibration, use the displayed CCS calibration standards section and map standard `tc`, `m/z`, CCS, and optional charge.
11. Confirm or adjust the Advanced column mapping below the calibration standards. `Calibrated (from Processing CCS calibration)` uses the CCS calculated from the IMS standards.
12. Enable advanced PFAS library matching only when database matching is required.
13. Choose an output file.
14. Click `Process and write output`.

## Output Workbook

The app writes an Excel workbook with these sheets:

### Labeled data

Contains the original input rows plus app-generated columns, including:

- `PFAS_riMp1` through `PFAS_riMp4`, plus backward-compatible `PFAS_rMp1` through `PFAS_rMp4`
- `PFAS_M_intensity` and `PFAS_M_plus_1_intensity_sum` through `PFAS_M_plus_4_intensity_sum`

The canonical `PFAS_riMp1-riMp4` values are also passed to compatible prediction
models and optional isotope-ratio library matching.

- `PFAS_ion_mode`
- `PFAS_mz_raw`
- `PFAS_mz_calibrated`
- `PFAS_mz_used`
- `PFAS_mass_error_ppm_estimated`
- `PFAS_mass_calibration_extrapolated`
- `PFAS_tc_input`
- `PFAS_charge_used`
- `PFAS_CCS_stage1`
- `PFAS_CCS_model_input`
- `PFAS_calibration_extrapolated`
- `PFAS_first_isotopic_peak_ratio_used`
- `PFAS_isotope_role`
- `PFAS_isotope_order`
- `PFAS_isotope_envelope_base_sorted_index`
- `PFAS_first_13C_peak_ratio`
- `PFAS_first_13C_peak_mass_error_Da`
- `PFAS_first_13C_peak_rt_error_min`
- `PFAS_first_13C_peak_tc_error`
- `PFAS_rMp1`
- `PFAS_rMp2`
- `PFAS_rMp3`
- `PFAS_rMp4`
- `PFAS_M_plus_1_*` through `PFAS_M_plus_4_*` audit fields
- `PFAS_prob_mean`
- `PFAS_prob_std`
- `PFAS_prob_p05`
- `PFAS_MC_uncertainty`
- `PFAS_level`
- `PFAS_level_name`
- `PFAS_candidate`
- `PFAS_high_confidence`
- `PFAS_processing_note`

### Prioritized PFAS peaks

Contains the final-screen-positive monoisotopic M features that pass the M-intensity
threshold and Top N setting. It includes experimental `rMp1` to `rMp4`, model
scores, intensity rank, database match count, and the best mass-ranked database
candidate for quick review.

### PFAS scatter plots

Every result workbook contains two native Excel scatter plots: `m/z vs CCS` and
`riMp1 vs CCS`. They show monoisotopic M features after calibration and optional
blank subtraction. Most likely PFAS are red, likely PFAS are orange, possible
PFAS are green, and rejected/non-PFAS features are grey. The `Labeled data`
sheet records the corresponding rank, class, and hex color in
`PFAS_plot_confidence_rank`, `PFAS_plot_confidence_class`, and
`PFAS_plot_color_hex`.

For MLM-only screening, colors follow MLM Level 0-3. For linear-only screening,
the linear Monte Carlo pass-probability bands are `<0.50`, `0.50-<0.65`,
`0.65-<0.80`, and `>=0.80`. Intersection uses the lower of the MLM and linear
display confidence; union uses the higher. This color calculation is for
visualization only and does not modify the screening result.

### PFAS database matches

This sheet is generated only when advanced PFAS library matching is enabled. It contains all database candidates within the selected ppm tolerance. It includes
the ion hypothesis, mass error, formula, SMILES, experimental and theoretical
`rMp1` to `rMp4`, and their differences.

### Run summary

Records the run settings, calibration formula, R2, model mode, number of predictable rows, and level counts.

### Calibration curve

Created when cIMS calibration standards are used. It includes:

- Standard `tc`
- Standard `m/z`
- Reference CCS
- Reduced CCS
- Fitted reduced CCS
- Back-calculated CCS
- Residuals
- A calibration chart

### Mass calibration

Created when lock-mass or linear mass calibration is used. It includes:

- Observed m/z
- Exact m/z
- Observed ppm error
- Fitted ppm error
- Corrected m/z
- Residual ppm error

## Troubleshooting

If rows are `Not predictable`, check that `m/z`, final CCS, and the first
isotopic peak ratio are present and positive.

If many rows are extrapolated, use standards that better cover the sample `tc` and `m/z` ranges.

If many rows are mass-calibration extrapolated during linear calibration, use mass standards that cover the sample `m/z` range. Lock-mass calibration uses one global correction factor and is not marked as extrapolated.

If the output probabilities look unexpected, first confirm the POS/NEG mode, selected MLM, and CCS calibration source.

If the app does not open, keep the whole portable folder together and launch:

```text
dist\PFAS_CCS_Screening\PFAS_CCS_Screening.exe
```
