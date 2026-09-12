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

> 本包不含二进制样例。需生成演示用 `.docx` 样例，运行：
> `python build_sample.py`（产出 template.docx / derived.docx，用于验证引擎）。
> `.eml` 样例已随包附带（template.eml / derived.eml，纯文本可直接查看）。

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

## 快速演示（一键跑，无需准备样例）

不确定怎么用？先跑内置演示，自动生成模板/派生样例并打印标准偏差报告：

```bash
python checker.py --demo
```

演示会生成「模板第2段字号=24、派生被改为=40」的样例，输出一份标准报告（见下「报告样例」）。
想用你自己的文件？照上面「快速开始」里的命令替换 --template / --derived 即可。

## 常见问题（FAQ / 避坑）

- **报错「文件已损坏或不是有效的 Office 文档」？** 扩展名虽是 Office，但内部不是合法 zip 包——常被改名（实际是 PDF/图片）、加密文档或下载不完整。请用原生 Office 重新另存为对应格式。
- **报错「不支持的文件格式」？** 仅支持 .docx/.xlsx/.pptx/.eml。`.msg`（Outlook）暂不支持，请先在 Outlook「另存为」.eml；PDF 扫描件/图片不支持（无可抽取的声明式属性）。
- **报告里差异很多，正常吗？** 差异是派生相对模板真不一样的地方。差异过多可能是模板与派生本就不同基准，或想忽略某些属性——用 `ignore_attr_patterns` 或把对应维度 `"enabled": false`。
- **宏 / VBA 生成的格式能识别吗？** 只要已落盘为成品属性（跑完保存的文件）就能识别；未落盘 / P-code / 运行时求值的格式无法读取，会在报告标红（volatile / 外部引用 / 缓存域 / 未替换占位符 {{}}）。
- **中文报错乱码？** 确保终端 UTF-8（PowerShell 默认多为 UTF-8；cmd 用 `chcp 65001`）。

## 报告样例

单文件比对生成 `report.md`，形如（节选自内置演示）：

~~~
# 模板一致性偏差报告
- 模板：`.../template.docx`
- 派生：`.../derived.docx`
- 偏差总数：1　|　❌需人工复核：0

## D4 维度（1 处）
| 类型 | 位置 | 属性 | 模板 | 派生 | 层 | 模式 |
|---|---|---|---|---|---|---|
| changed | document.xml:document[1]/body[1]/p[2]/r[1]/rPr[1]/sz[1] | sz/val | 24 | 40 | content | tolerance |
~~~

`❌ 需人工复核` 区单列 volatile / 外部链接 / 缓存域 / 未替换占位符等真缺口。

## 能力边界

- **支持**：.docx / .xlsx / .pptx / .eml。`.msg` 待 olefile 后续接入。
- **不处理**：未落盘 / P-code / 运行时求值；PDF 扫描件、图片、加密文档。
- **批量**：`--derived-dir` 扫描目录下所有 .docx/.xlsx/.pptx/.eml，其他文件自动跳过；单次为 1 模板 vs N 派生，N 无硬性上限（受内存）。
- **性能**：耗时随属性总数增长；超长文档仍可行，但属性量很大时建议把无关维度 `enabled:false`。
- **角色**：覆盖「识别 + 校验」；不生成文档、不渲染、不修改文件，只产出偏差报告。

## 许可

MIT —— Copyright (c) 2026 shuigui-ou
