# 模板一致性偏差报告

- 模板：`C:\Users\199720.PC2775\.workbuddy\skills\office-template-checker\samples\template.docx`
- 派生：`C:\Users\199720.PC2775\.workbuddy\skills\office-template-checker\samples\derived.docx`
- 偏差总数：21　|　❌需人工复核：0

## D4 维度（6 处）

| 类型 | 位置 | 属性 | 模板 | 派生 | 层 | 模式 |
|---|---|---|---|---|---|---|
| added | document.xml:document0body0tbl0tr3tc0p0r0rPr0sz0 | sz/val | None | 20 | content | tolerance |
| added | document.xml:document0body0tbl0tr3tc1p0r0rPr0sz0 | sz/val | None | 20 | content | tolerance |
| changed | document.xml:document0body0tbl1tr3tc2p0r0rPr0sz0 | sz/val | 20 | 40 | content | tolerance |
| added | document.xml:document0body0tbl1tr4tc0p0r0rPr0sz0 | sz/val | None | 20 | content | tolerance |
| added | document.xml:document0body0tbl1tr4tc1p0r0rPr0sz0 | sz/val | None | 20 | content | tolerance |
| added | document.xml:document0body0tbl1tr4tc2p0r0rPr0sz0 | sz/val | None | 20 | content | tolerance |

## other 维度（15 处）

| 类型 | 位置 | 属性 | 模板 | 派生 | 层 | 模式 |
|---|---|---|---|---|---|---|
| added | document.xml:document0body0tbl0tr3tc0tcPr0tcW0 | tcW/w | None | 2000 | content | exact |
| added | document.xml:document0body0tbl0tr3tc0tcPr0tcW0 | tcW/type | None | dxa | content | exact |
| added | document.xml:document0body0tbl0tr3tc0p0r0t0 | t/space | None | preserve | content | exact |
| added | document.xml:document0body0tbl0tr3tc1tcPr0tcW0 | tcW/w | None | 2000 | content | exact |
| added | document.xml:document0body0tbl0tr3tc1tcPr0tcW0 | tcW/type | None | dxa | content | exact |
| added | document.xml:document0body0tbl0tr3tc1p0r0t0 | t/space | None | preserve | content | exact |
| added | document.xml:document0body0tbl1tr4tc0tcPr0tcW0 | tcW/w | None | 2000 | content | exact |
| added | document.xml:document0body0tbl1tr4tc0tcPr0tcW0 | tcW/type | None | dxa | content | exact |
| added | document.xml:document0body0tbl1tr4tc0p0r0t0 | t/space | None | preserve | content | exact |
| added | document.xml:document0body0tbl1tr4tc1tcPr0tcW0 | tcW/w | None | 2000 | content | exact |
| added | document.xml:document0body0tbl1tr4tc1tcPr0tcW0 | tcW/type | None | dxa | content | exact |
| added | document.xml:document0body0tbl1tr4tc1p0r0t0 | t/space | None | preserve | content | exact |
| added | document.xml:document0body0tbl1tr4tc2tcPr0tcW0 | tcW/w | None | 2000 | content | exact |
| added | document.xml:document0body0tbl1tr4tc2tcPr0tcW0 | tcW/type | None | dxa | content | exact |
| added | document.xml:document0body0tbl1tr4tc2p0r0t0 | t/space | None | preserve | content | exact |

## 偏差类型语义（如何解读本报告）

你**基于模板手动编辑**（改内容 / 格式 / 内容逻辑）后另存为新文件，因此两类差异性质不同：

- **格式偏离 `changed`**：同一**相对位置**上的格式属性值不同（字号、颜色、对齐、边距等）。这是审查重点——往往是编辑时不小心改坏了模板格式，或应当保持一致却漏改。
- **结构变动 `added`**：派生比模板多出的相对位置（如你插入了一行/列/段落）。属内容编辑的预期动作；本工具按相对位置对齐后**只把新增部分单列**，不会牵连后续行整片误报，请确认是否故意。
- **结构变动 `removed`**：派生比模板缺失的位置（如你删了一行/列）。可能是合理精简，也可能误删，需你确认。

> 内容文本（你填的字）只用作「相对位置」对齐锚点，**不参与格式比较**——所以你改内容不会凭空产生「格式偏离」；只有格式属性本身被改才会记为 `changed`。

---
> 核心保证（确定性，不调用任何大模型）：**按相对位置比对**——先按结构树逐层（正文 / 表 / 行 / 单元格）用贪心一步前瞻按 (结构, 内容锚点) 对齐，再比较「同一相对位置上的同一属性值是否相同」。模板某处插入一行/一列时，该新增行列被识别为 added/removed，其余行列仍按对齐后的相对序号一一对应，不会因绝对索引平移而整片误报。报告定位串（表N·行M·列K）中的序号即相对序号，非原始绝对索引。报告仅含参数树未忽略的差异；渲染层/版本默认/隐式默认值不在范围内。

## 引用安全声明（给大模型使用方 · 防止误引）

仅核对**格式属性层**（字体/字号/颜色/字距/边距/对齐/位置大小/数字格式/邮件头与 HTML 内联样式）。本工具是确定性引擎，不调用任何大模型。

**本工具未查 / 查不了（引用方不得越界断言）：**
- 内容语义与文字对错（你填的字只作相对位置对齐锚点，不参与格式比较）
- 内容逻辑（如公式引用关系、交叉引用、编号连续性的正确性）
- .doc / .pdf / .msg 等非 OOXML / 非 eml 格式（需另接适配器）
- 低于结构对齐阈值的零散块（相似度不足、无法与模板对应的片段）
- 模板本身的设计意图是否合理（本工具不判断模板好坏，只比对偏离）

**引用纪律（四不准）：**
- 引用须带定位(path 或 part:path 片段)与维度标签(D1-D7)，且限定在本报告已列出的偏差条目内。
- 不得对报告未列出的属性做「一致/不一致」断言；未列出 ≠ 一致（可能只是被参数树忽略或不在范围）。
- 带 ❌ 的偏差（volatile/外部链接/缓存域/未替换占位符）引用时须加「需人工复核」，不得直接判定合格。
- 不得编造报告中没有的页码、文件名或判定；如需补充，须说明「超出本工具检查范围，需人工/源文件确认」。