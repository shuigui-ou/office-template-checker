# office-template-checker

**对照模板校验派生文件格式一致性** —— 支持 Word / Excel / PowerPoint / 邮件(.eml)。

基于「全量属性对称 diff + 参数树匹配」：不是拿一份固定清单去比，而是把模板与派生文件各自的
**每一个声明式属性**都抽出来做对称比对，任何一边有、另一边没有都会自动暴露。因此「参数不全」
在底层物理上不可能发生。

## 为什么比人手更可靠

办公高手打开文件、进字体对话框、用格式刷去「看」，本质也是在读同样的属性——只是会疲劳、会漏、
给不出精确值。本工具做同一件事的穷举版，并精确到「第几段第几个 run」。

## 特性

- **全量枚举 + 对称 diff**：不依赖人工清单，覆盖冷门属性（行距规则、表格固定列宽、项目符号、
  主题色 tint、占位符等）。
- **位置索引比较 key**：`p[1]/r[1]/rPr[1]/sz[1]` 与 `p[2]/.../sz[1]` 是两条独立 key，**永不合并**
  —— 句1 的 A 参数与句C 的 A 参数算两个，绝不偷偷归一。
- **参数树（维度→参数→小参数）作为过滤器**：5 种匹配模式——精确 / 解算RGB / 按主题槽 /
  容差±ε / 忽略。可整体开关某一维、可下钻到单个小参数。
- **真缺口标红（❌）**：volatile 函数(TODAY/NOW/RAND…)、外部引用、缓存域、未替换占位符 `{{}}`
  会在报告中单列，需人工复核，绝不静默放过。
- **适配器架构**：docx/xlsx/pptx/eml 已实现；新格式只需实现 `FormatAdapter` 子类并 `register`，
  引擎零改动接入。

## 支持格式

| 格式 | 适配器 | 状态 |
|---|---|---|
| .docx | OfficeAdapter | ✅ |
| .xlsx | OfficeAdapter | ✅ |
| .pptx | OfficeAdapter | ✅ |
| .eml  | EmailAdapter | ✅ |
| .msg  | (需 olefile)   | ⏳ TODO |

## 快速开始

Python 解释器用 **3.8+**；本工具**仅用标准库，零第三方依赖**。

```bash
# 1) 建基线（模板只抽一次，可复用）
python checker.py --baseline template.docx --out baseline.json

# 2) 单文件比对
python checker.py --template template.docx --derived derived.docx \
                  --config param_tree.json --out report.md --json report.json

# 3) 批量目录比对（docx/xlsx/pptx/eml 混放亦可）
python checker.py --template template.docx --derived-dir ./deriveds/ \
                  --config param_tree.json --out-dir reports/
```

运行目录请 `cd` 到本 Skill 目录（含 checker.py / office_adapter.py / email_adapter.py /
adapter_interface.py / param_tree.json）。

## 输出

- `report.md`：按维度分组的偏差表（类型 / 位置 / 属性 / 模板值 / 派生值 / 层 / 模式）+ ❌ 人工复核清单。
- `report.json`：结构化偏差 + 红标。
- 批量时 `reports/_summary.json` 汇总每份文件的偏差数。

## 自定义参数树

编辑 `param_tree.json`：

- 维度 `D1–D6`（Office）与 `D7_email`（邮件）可整体开关（`"enabled": false` 整维不比）。
- 小参数键用 `element/attr` 斜杠形式（如 `sz/val`、`color/themeColor`、`style/font-family`），
  `match` 取五种模式之一。
- `ignore_attr_patterns`：噪声属性通配（默认已含 `rsid*`、`Ignorable`、`docId`、`#text`）。
- `tolerances`：数值类容差（字号默认 1 半磅、边距 5 twip、邮件字号 1 px、行高 0.1 等）。
- `path_overrides`：按路径正则强制忽略（如图形装饰节点）。

## 架构

```
模板 ─┐
      ├─► adapter.extract() ─► Prop(part,path,element,attr,value)  ─┐
派生 ─┘                                                           │
                                                                  ▼
                                          对称 diff（位置索引 key，互不合并）
                                                                  │
                                                                  ▼
                                          参数树 5 模式判定（过滤器/聚合层）
                                                                  │
                                                                  ▼
                                          偏差报告（md+json）+ ❌ 红标
```

## 扩展新格式

适配器模式已验证（docx/xlsx/pptx + eml）。新增格式（如 .msg、PDF 表单）只需实现 `FormatAdapter`
子类（`extract / is_instance / anchors / theme_map`）并 `AdapterRegistry.register()`，引擎零改动。

## 业务场景

同一套逻辑可复用于任何「模板/基线 vs 实例/派生」一致性校验场景：合同文书、配置基线、IaC 样板、
设计令牌、数据字典、i18n 翻译、监管报表等。详见《业务场景扩展评估.md》。

## 许可

MIT —— Copyright (c) 2026 shuigui-ou
