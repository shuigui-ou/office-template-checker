---
name: office-template-checker
description: 对照模板（docx/xlsx/pptx/eml 成品）校验派生文件格式一致性，输出精确到位置的偏差报告。当用户给出模板和若干派生文件、要核对"有没有偏离模板"（字体/字号/颜色/字距/边距/对齐/位置大小/数字格式/邮件 HTML 内联样式等）时使用。基于全量属性对称 diff，参数树可自定义、可按维度-参数-小参数逐层下钻；已内置邮件(.eml)适配器。
---

# Office / 邮件 模板一致性校验器

## 何时用
- 用户给了一份**模板（已跑完并保存的成品）**和一份或多份**派生文件**，要核对格式是否偏离模板。
- 典型诉求："对照这份模板看看这份文件哪里不一样""这批文件是不是都套了模板格式""这封邮件和模板邮件字体/颜色一致吗"。
- 支持格式：**docx / xlsx / pptx / .eml**（邮件）。`.msg` 待 olefile 后续接入。

## 核心事实（务必照此口径）
1. **全量枚举 + 对称 diff**，不是固定清单。模板与派生各抽全部属性（Office 的 XML 属性 / 邮件的 header + HTML 内联样式），任何一边有另一边没有自动暴露——「参数不全」在底层物理上不可能发生。
2. **比较 key = (part, path, element, attr)，path 含位置索引**（`p[1]/r[1]/rPr[1]/sz[1]`，邮件 `html[1]/body[1]/p[2]/...`）。同参数不同位置天然不同 key，**永不合并**——句1 的 A 和 句C 的 A 算两个。
3. **默认交付物是成品**，按「全量可识别」处理，不追问宏是否跑过。仅 volatile 函数 / 外部引用 / 缓存域 / 未替换占位符 `{{}}` 需在报告标红（❌ 真缺口）。
4. **参数树（维度→参数→小参数）是过滤器/聚合层**，5 种匹配模式：精确 / 解算RGB / 按主题槽 / 容差±ε / 忽略。邮件属性归入 `D7_email` 维度。

## 运行方式（引擎在本 Skill 目录内）
Python 解释器用托管版本：
`C:\Users\199720.PC2775\.workbuddy\binaries\python\versions\3.13.12\python.exe`

- **建基线（模板只抽一次）**：
  `python checker.py --baseline <模板.docx|.eml> --out baseline.json`
- **单文件比对**：
  `python checker.py --template <模板> --derived <派生> --config param_tree.json --out report.md --json report.json`
- **批量目录比对**（docx/xlsx/pptx/eml 混放亦可）：
  `python checker.py --template <模板> --derived-dir <派生目录> --config param_tree.json --out-dir reports/`

运行目录请 `cd` 到本 Skill 目录（含 checker.py / office_adapter.py / email_adapter.py / adapter_interface.py / param_tree.json）。

## 输出
- `report.md`：按维度分组的偏差表（类型 位置 属性 模板值 派生值 层 模式）+ ❌ 人工复核清单。
- `report.json`：结构化偏差 + 红标。
- 批量时 `reports/_summary.json` 汇总每份文件的偏差数。

## 自定义参数树
编辑 `param_tree.json`：
- 维度 `D1–D6`（Office）与 `D7_email`（邮件）可整体开关（`"enabled": false` 整维不比）。
- 小参数键用 `element/attr` 斜杠形式（如 `sz/val`、`color/themeColor`、`style/font-family`），`match` 取五种模式之一。
- `ignore_attr_patterns`：噪声属性通配（默认已含 `rsid*`、`Ignorable`、`docId`、`#text`）。
- `tolerances`：数值类容差（字号默认 1 半磅、边距 5 twip、邮件字号 1 px、行高 0.1 等）。
- `path_overrides`：按路径正则强制忽略（如图形装饰节点）。
- 邮件特有：颜色自动归一化为 `#RRGGBB`（命名色/rgb()/3 位 hex 均归一），字号归一化为 px 数值（pt×1.333、em/rem×16），行高归一化为比值，便于 tolerance/rgb 模式生效。

## 报告解读注意（边界）
- ❌ 标红 = 渲染层/版本默认/跨格式/隐式默认值/外部目标/未替换占位符等真缺口，需人工复核，不要静默放过。
- 颜色「按主题槽」模式用于反向校验"禁止硬编码、必须用主题色"；「解算RGB」用于"最终颜色一样就放过"（邮件 color 走此模式）。
- 模板内部不同位置允许不同值（句1 与 句C 可各异）；只有显式聚合规则才会跨位置合并。
- 邮件正文占位符 `{{客户名}}` 若未替换，会在 ❌ 区列出「unresolved_placeholder」，便于发现漏填。

## 扩展：更多格式
适配器模式已验证（docx/xlsx/pptx + eml）。新增格式（如 .msg、PDF 表单）只需实现 `FormatAdapter` 子类（`extract/is_instance/anchors/theme_map`）并 `AdapterRegistry.register()`，引擎零改动。
