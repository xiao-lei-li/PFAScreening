# PFAS CCS Screening App 中文说明

该应用读取 Excel 或 CSV 特征表，完成质量校正、CCS 校正、第一同位素峰比值计算以及 PFAS 分级预测。

## 批处理与 blank subtraction

点击 `Choose sample(s)...` 可一次选择一个或多个样品文件。多个样品共用当前界面中的 worksheet、列映射、质量/CCS 校正、模型、线性筛选和 library matching 设置。批处理为每个样品生成独立的 `样品名_PFAS_labeled.xlsx`，并在所选输出文件夹生成 `PFAS_batch_summary.xlsx`。若某个样品失败，其错误会记录在批次汇总中，其余样品继续处理。

勾选 `Use multiple blank replicates` 后必须选择至少两个 blank 文件。样品和 blank 使用相同的 worksheet 名称及列映射。程序先分别完成质量和 CCS 校正，再按以下三个条件同时匹配样品峰和各 blank 中的峰：

- 校正后 m/z：用户设置的 ppm 容差；
- RT：用户设置的绝对分钟容差；
- 校正后 CCS：用户设置的百分比容差。

同一个 blank 中存在多个候选时，选择 m/z、RT、CCS 归一化综合偏差最小的峰；某个 blank 未检出匹配峰时，该 blank replicate 的强度按 0 计。对每个样品峰计算：

```text
blank threshold = blank mean intensity + 3 * blank sample SD
corrected sample intensity = max(sample raw intensity - blank mean intensity, 0)
```

仅保留 `sample raw intensity >= blank threshold` 的峰。扣空白后的强度进入后续自动同位素配对、riMp1-riMp4 计算、Minimum M intensity 和筛选流程。若用户明确选择使用输入文件中已有的 riMp1，该已有比值不会被程序重新计算。

输出的 `Labeled data` 包含原始强度、blank mean、blank SD、阈值、校正后强度、匹配 blank 数量和通过状态；被剔除的峰保存在 `Blank-subtracted rows` 工作表中。即使一个样品的全部峰均被剔除，程序仍会写出审计结果。

## 可选质量校正

质量校正在 CCS 校正、同位素峰计算和模型预测之前执行。输出会同时保留原始 m/z 和校正后的 m/z。

- `No mass calibration`：直接使用输入 m/z。
- `Lock mass ppm offset`：在 `Lock mass exact m/z` 输入已知 lock mass 的理论 m/z，并在 `Lock mass search tolerance (ppm)` 设置最近峰搜索容差（默认 `1000 ppm`）。程序在样品表中寻找容差内最接近的实测峰，并使用 `校正后 m/z = 实测 m/z × 理论 lock m/z / 实测 lock m/z` 对所有 m/z 作单点校正。没有峰落入所选容差时，程序停止并报错；此模式不需要上传质量标准物文件。
- `Linear correction from standards`：上传含实测 m/z 和理论 m/z 的标准物表，拟合随 m/z 变化的 ppm 误差；此模式至少需要两个不同的实测 m/z。只有选择该模式时，界面才显示并要求填写 `Mass calibration standards`。

Lock mass 或 linear 校正的拟合信息和校正结果会写入 `Mass calibration` 工作表。

## 第一同位素峰比值

应用支持三种输入方式：

- 使用已有的 `first_isotopic_peak_ratio`；旧字段 `riMp1` 仍可作为输入别名。
- 使用单同位素峰和第一同位素峰强度计算：

```text
first_isotopic_peak_ratio =
    100 * I(第一同位素峰) / I(单同位素峰)
```

- 根据 m/z、保留时间、`3D_TC` 和强度自动建立同位素包络。CCS 不参与配对。

自动配对假定离子为单电荷，参数为：

```text
名义质量偏移：+1、+2、+3、+4 Da
质量窗口：+/-0.020 Da
保留时间容差：+/-0.020 min
3D_TC 容差：+/-0.060
```

每个峰只能属于一个同位素包络，已经分配为 M+n 的峰不能再作为新的 M。候选冲突依次按照保留时间误差、`3D_TC` 误差和名义质量误差处理。

### M+1 至 M+4 实验同位素分布

自动配对模式提取 M+1、M+2、M+3 和 M+4 名义同位素簇。各比值均相对于单同位素峰 M：

```text
rMp1 = 100 * sum(I(M+1 候选峰)) / I(M)
rMp2 = 100 * sum(I(M+2 候选峰)) / I(M)
rMp3 = 100 * sum(I(M+3 候选峰)) / I(M)
rMp4 = 100 * sum(I(M+4 候选峰)) / I(M)
```

同一名义同位素簇内满足条件的精细同位素候选峰强度会被求和，同时输出候选数、m/z 范围、质量偏移范围、RT 误差范围、`3D_TC` 误差范围和歧义标记。

Advanced column mapping 中的 `Existing rMp1 / first isotopic peak ratio` 和 `First isotopic peak intensity` 均提供 `Autodetected`。选择后，程序会在 Processing 时按支持的列名别名检测真实输入列，不会把 `Autodetected` 当作 Excel 列名。

最终 `Labeled data` 保留规范字段 `PFAS_riMp1` 至 `PFAS_riMp4`、兼容字段 `PFAS_rMp1` 至 `PFAS_rMp4`，以及 `PFAS_M_intensity`、`PFAS_M_plus_1_intensity_sum` 至 `PFAS_M_plus_4_intensity_sum`。这些规范 riMp 值同时传给支持相应特征的预测模型和可选 isotope-ratio library matching。

当前 RF-MC 模型使用完整 M+1 簇比值 `rMp1`。另外输出 `first_13C_peak_ratio`，它是在已分配 M+1 簇中选择最接近 `+1.00335483507 Da` 的单峰，仅用于审计，不作为模型输入。M+2 至 M+4 用于初步同位素模式或元素组成类别判断，不能单独作为结构鉴定结果。

## 预测模型

应用现在不会静默选择模型。必须在 `Processing settings` 中显式选择 prediction model，否则点击 processing 后会直接停止并报错。

当前内置两个 RF-MC 模型供用户选择：

```text
RF-MC riMp1 +/-20% CCS +/-3%
RF-MC riMp1 +/-5% CCS +/-3%
```

模型文件：

```text
models\rf_sklearn_bundle_POS.joblib
models\rf_sklearn_bundle_NEG.joblib
models\rf_sklearn_bundle_rimp1_5_ccs_3_POS.joblib
models\rf_sklearn_bundle_rimp1_5_ccs_3_NEG.joblib
```

两个模型均为 RF-MC 模型，使用相同的特征：

```text
mz_percentile, mz, CCS, riMp1, CCS_over_m, CCS_over_Mp1,
M_over_Mp1, CCS_over_sqrtm, CCS_over_m23
```

它不使用 `rMp2`、`rMp3`、`rMp4` 进行预测。这些字段目前只进入输出和可选的 PFAS library isotope-ratio matching。也就是说，真正使用 `rMp1-rMp4` 的新预测模型目前还没有安装到 app 中。

旧 POS/NEG 模型的正式验证条件为 `riMp1 +/-20%` 和 `CCS +/-3%`；新增 POS/NEG 模型的正式验证条件为 `riMp1 +/-5%` 和 `CCS +/-3%`。选择模型后，界面会自动设置与该模型相符的 uncertainty。用户仍可改为其他 uncertainty 做敏感性分析，但输出会标记为 `Exploratory / unvalidated uncertainty setting`，不能沿用任一模型的正式验证结论。

### 后续加入不同 uncertainty 模型

后续模型不需要再改 GUI 代码。把模型文件放进 `models` 文件夹，并创建或更新：

```text
models\model_manifest.csv
```

必需列：

```text
model_id,label,pos_path,neg_path,feature_cols
```

建议增加的可选列：

```text
description,rimp1_uncertainty_percent,ccs_uncertainty_percent,
pos_candidate_threshold,pos_high_confidence_threshold,
neg_candidate_threshold,neg_high_confidence_threshold
```

示例：

```csv
model_id,label,pos_path,neg_path,feature_cols,description,rimp1_uncertainty_percent,ccs_uncertainty_percent,pos_candidate_threshold,pos_high_confidence_threshold,neg_candidate_threshold,neg_high_confidence_threshold
rf_rimp5_ccs3,RF-MC riMp1 +/-5% CCS +/-3%,rf_rimp5_ccs3_POS.joblib,rf_rimp5_ccs3_NEG.joblib,"mz_percentile,mz,CCS,riMp1,CCS_over_m,CCS_over_Mp1,M_over_Mp1,CCS_over_sqrtm,CCS_over_m23",Model trained for riMp1 +/-5% and CCS +/-3%,5,3,0.555,0.947,0.55875,0.976
```

`pos_path` 和 `neg_path` 可以是相对于 `models` 文件夹的文件名，也可以是绝对路径。只有 POS 和 NEG 两个文件都存在时，GUI 才会显示该模型。

如果模型要求 app 不能生成的特征列，处理会直接停止并报错。当前支持的 isotope-ratio 输入特征名包括 `riMp1`、`riMp2`、`riMp3`、`riMp4`，以及别名 `rMp1`、`rMp2`、`rMp3`、`rMp4`。

## 线性关系 PFAS 筛选

`PFAS screening method` 提供四种策略：

- `Machine Learning Model (MLM) only`：只使用 MLM 的 `PFAS Level >=1` 标签。
- `Linear relationship only`：完全跳过模型，只使用下述线性关系。
- `MLM AND linear relationship (intersection)`：MLM 和线性关系必须同时通过。
- `MLM OR linear relationship (union)`：MLM 或线性关系任一通过即可。

Processing settings 下方的蓝色依据说明会随 ion mode、screening method 和 MLM 选择即时更新，显示当前 POS/NEG 线性公式、线性概率分级、MLM 的 `pmean`/`p05` 阈值，以及交集或并集规则。

选择 `Linear relationship only` 时，只有 MLM 选择会变为不可编辑。MC iterations、riMp1 uncertainty 和 CCS uncertainty 保持可编辑，并用于将测量不确定度传播到线性关系判定。

线性筛选使用校正后的 m/z、最终 CCS 和最终实验 riMp1：

```text
POS：CCS < 0.2537*mz + 84.701 且 CCS > 4.657*riMp1 + 98.570
NEG：CCS < 0.24284*mz + 84.432 且 CCS > 4.415*riMp1 + 97.440
```

两个条件均为严格不等式，刚好位于边界时不通过。输出保留上下边界、线性标签和带方向的残差：

```text
upper residual = CCS - upper CCS boundary，必须 < 0
lower residual = CCS - lower CCS boundary，必须 > 0
```

每次 Monte Carlo 迭代都会按照所选相对 uncertainty 独立扰动 riMp1 和 CCS，再判断两条严格不等式。`PFAS_linear_pass_probability` 是同时满足两条关系的迭代比例：`>=0.80` 标记为高置信线性 PFAS，`0.50 <= p < 0.80` 标记为不确定线性 PFAS 候选，`<0.50` 标记为线性非 PFAS。前两级的 `PFAS_linear_screen_pass` 均为 1。

`PFAS_final_candidate` 是用户所选策略合并后的最终标签。Minimum M intensity、Top N 和可选 PFAS library matching 都使用这个最终标签，不再固定只看模型 Level。

## 强度筛选与 PFAS library 匹配

应用使用 `Minimum M intensity` 和 `Top N intense PFAS peaks` 筛选浓度较高的候选峰。只有 `isotope_role=M` 且 `PFAS_final_candidate=1` 的峰参与筛选，并按照 M 峰强度和 PFAS 概率排序；linear-only 模式没有模型概率时按 M 峰强度排序。

`Use advanced PFAS library matching` 是总开关。勾选后才会读取 library 并执行匹配；不勾选时完全跳过匹配，也不会生成 `PFAS database matches` 工作表。启用后可以选择：

- `PFAS master list`：使用应用自带的 `Chemical List PFAS.xlsx`。
- `Customized library`：用户自己选择 Excel 或 CSV library。

library 必须包含分子式和中性单同位素质量。名称、SMILES、DTXSID、CASRN、InChIKey 和 CCS 为可选字段；如果存在，会写入输出文件。默认 mass tolerance 为 `+/-5 ppm`；CCS tolerance 使用百分比。当前自带 master list 不包含 CCS 列，因此只有用户选择的 library 含 CCS 列时才会进行 CCS tolerance 过滤。

数据库匹配同时考虑：

```text
POS：分别使用观测 M 和“观测 M - 1.007276466621 Da”查库
NEG：分别使用观测 M 和“观测 M + 1.007276466621 Da”查库
质子精确质量：1.007276466621 Da
```

用户选择的 mass tolerance 会分别应用于两条质量匹配路径。实验 CCS 与 library CCS 都存在时，还会同时执行用户选择的 CCS tolerance；超出任一有效容差的候选不会进入匹配结果。

每个候选会输出名称、DTXSID、CASRN、分子式、library 中性单同位素质量、library CCS、筛选实验 PFAS 的 m/z/CCS、离子解释、查库所用质量变换量、目标中性质量、质量误差和 CCS 误差。

如果勾选 `Match by isotope ratios after library matching`，应用会根据 PFAS library 中的 molecular formula 计算理论 `rMp1`、`rMp2`、`rMp3`、`rMp4`，并与实验 `rMp1` 至 `rMp4` 匹配。置信区间使用独立的 `Isotope ratio tolerance (%)`，不再复用模型的 `riMp1 uncertainty (%)`；例如选择 `20%` 表示理论 `rMpN +/-20%`。超出 isotope ratio 置信区间的候选不会进入最终 `PFAS database matches` 表。

SMILES表示数据库中的可能结构。同一质量或分子式可能对应多个结构，因此应用保留容差内的全部候选，不自动指定唯一结构。

## CCS 校正

CCS calibration 可以使用标准物的校正到达时间、m/z 和参考 CCS 建立 cIMS/TWIMS 幂函数校正，也可以直接使用已有实验 CCS。只有选择 `cIMS/TWIMS power-law calibration (tC, m/z, CCS)` 时，界面才显示并要求填写 `CCS calibration standards`；选择已有 CCS 列时该区域隐藏且不要求标准物文件。

这里的 `tC` 必须已经扣除仪器外飞行时间及 m/z-dependent transfer time。程序首先计算：

```text
m_ion = (m/z) * |z|
mu = m_ion * m_N2 / (m_ion + m_N2)
reduced_CCS = CCS * sqrt(mu) / |z|
```

至少需要 3 个具有不同正 `tC` 的标准点，拟合：

```text
reduced_CCS = A * tC^B
```

样品 CCS 按下式反算：

```text
CCS = A * tC^B * |z| / sqrt(mu)
```

因为输入是已经校正的 `tC`，这里不再拟合额外的时间偏移 `t0`。未经时间与质量相关传输校正的原始 arrival time 不能直接输入这一模式。

`Mass calibration standards` 和 `CCS calibration standards` 均显示在 `Advanced column mapping` 上方，便于先完成校正标准物设置，再检查样品数据列映射。

在 Advanced column mapping 的 `Existing CCS / calibrated result` 中选择 `Calibrated (from Processing CCS calibration)`，表示后续模型预测和 library matching 使用 Processing 阶段由 cIMS 标准物校正得到的 CCS。若选择 `Use an existing CCS column`，则必须改为选择输入工作表里真实存在的 CCS 列。

当前流程已经删除二级 CCS 校正。模型输入 CCS 直接使用第一阶段得到的实验 CCS：

- 由标准物 `tC`、`m/z`、参考 CCS 建立曲线后计算得到的 CCS；或
- 用户在输入表中选择的已有实验 CCS 列。

旧的 POS/NEG 固定转换不再出现在 GUI 中，也不会在处理时应用。

## 输出

- `Labeled data`：原始字段、质量和 CCS 校正结果、第一同位素峰配对信息、预测概率及分级标签。
- `Run summary`：运行参数、校正公式、拟合结果和各 Level 数量。
- `Calibration curve`：使用 cIMS 标准物校正时生成。
- `Prioritized PFAS peaks`：达到强度阈值并位于 Top N 内的最终筛选阳性单同位素 M 峰。
- `PFAS scatter plots`：包含 `m/z vs CCS` 和 `riMp1 vs CCS` 两张 Excel 原生散点图。图中红色表示最可能 PFAS，橙色和绿色表示依次降低的候选置信度，灰色表示未通过筛选或非 PFAS。散点使用校正及可选 blank subtraction 后的单同位素 M 峰；`Labeled data` 同时保留绘图置信等级、类别和颜色代码。MLM-only 使用模型 Level 0-3；linear-only 按线性 Monte Carlo 通过概率 `<0.50`、`0.50-<0.65`、`0.65-<0.80`、`>=0.80` 分色；交集取两种方法中较低的显示置信度，并集取较高者。分色只用于可视化，不会改变筛选结果。
- `PFAS database matches`：仅在启用 advanced PFAS library matching 时生成，包含全部质量候选、离子解释、分子式、SMILES及实验/理论同位素分布。

模型内部仍将该比值映射到训练时使用的历史特征名 `riMp1`，但新版对外字段统一使用 `first_isotopic_peak_ratio`。

## 运行

```powershell
python launch_pfas_app.py
```

便携版位于：

```text
dist\PFAS_CCS_Screening\PFAS_CCS_Screening.exe
```
