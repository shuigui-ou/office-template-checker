#!/usr/bin/env python3
# 症状验证：在「插入行+格式偏离+内容编辑」已覆盖之外，补测更多真实编辑症状，
# 重点验证相对位置对齐引擎对「删除/插入列/合并单元格/正文插段/多格式维度」的表现。
# 运行：python tests/test_symptoms.py
import os, zipfile, sys, re, json

SKILL = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, SKILL)
import checker as C

CT = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
      '<Default Extension="xml" ContentType="application/xml"/>'
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      '</Types>')
RELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
        '</Relationships>')


def cell(sz, text):
    return (f'<w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr>'
            f'<w:p><w:r><w:rPr><w:sz w:val="{sz}"/></w:rPr>'
            f'<w:t xml:space="preserve">{text}</w:t></w:r></w:p></w:tc>')


def cell_ex(sz, text, color=None):
    rpr = f'<w:rPr><w:sz w:val="{sz}"/>'
    if color:
        rpr += f'<w:color w:val="{color}"/>'
    rpr += '</w:rPr>'
    return (f'<w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr>'
            f'<w:p><w:r>{rpr}<w:t xml:space="preserve">{text}</w:t></w:r></w:p></w:tc>')


def cell_merge(text):
    # 跨 2 列的合并单元格（gridSpan）
    return (f'<w:tc><w:tcPr><w:gridSpan w:val="2"/><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>'
            f'<w:p><w:r><w:rPr><w:sz w:val="20"/></w:rPr>'
            f'<w:t xml:space="preserve">{text}</w:t></w:r></w:p></w:tc>')


def row(cells):
    return '<w:tr>' + ''.join(cells) + '</w:tr>'


def table(rows):
    return ('<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr>'
            + ''.join(rows) + '</w:tbl>')


def para(sz, text):
    return (f'<w:p><w:r><w:rPr><w:sz w:val="{sz}"/></w:rPr>'
            f'<w:t xml:space="preserve">{text}</w:t></w:r></w:p>')


def build(path, body_xml):
    doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
           '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
           f'<w:body>{body_xml}</w:body></w:document>')
    with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml', CT)
        z.writestr('_rels/.rels', RELS)
        z.writestr('word/document.xml', doc)


cfg = json.load(open(os.path.join(SKILL, "param_tree.json"), encoding="utf-8"))
pt = C.ParamTree(cfg)
TMP = os.path.join(SKILL, "samples", "_sym")
os.makedirs(TMP, exist_ok=True)


def run(tpl, drv):
    tp = os.path.join(TMP, "t.docx")
    dp = os.path.join(TMP, "d.docx")
    build(tp, tpl)
    build(dp, drv)
    devs, red = C.compare(tp, dp, cfg, pt)
    return devs


def idxs(devs, key):
    out = set()
    for d in devs:
        m = re.search(key + r'(\d+)', d['path'])
        if m:
            out.add(int(m.group(1)))
    return out


def kinds(devs):
    c = {"changed": 0, "added": 0, "removed": 0}
    for d in devs:
        c[d['kind']] = c.get(d['kind'], 0) + 1
    return c


results = []


def check(name, ok, detail):
    results.append((name, ok, detail))
    print(("PASS " if ok else "FAIL ") + name + " :: " + detail)


# ===== S1 删除中间行（模板表：制造商/型号/额定容量 → 派生删型号行）=====
tpl = table([row([cell(20, "制造商"), cell(20, "深圳市某电池")]),
             row([cell(20, "型号"), cell(20, "ABC-123")]),
             row([cell(20, "额定容量"), cell(20, "3000mAh")])])
drv = table([row([cell(20, "制造商"), cell(20, "深圳市某电池")]),
             row([cell(20, "额定容量"), cell(20, "3000mAh")])])
d = run(tpl, drv)
k = kinds(d)
ok = (len(idxs(d, "tr")) == 1) and k["changed"] == 0
check("S1 删除中间行", ok,
      f"removed行索引集合={idxs(d,'tr')} kinds={k}（期望 仅1行removed、0 changed、无级联）")

# ===== S2 删除中间列（表：[A,B,C] → [A,C]，每行都删 B）=====
tpl = table([row([cell(20, "A"), cell(20, "B"), cell(20, "C")]),
             row([cell(20, "1"), cell(20, "2"), cell(20, "3")]),
             row([cell(20, "4"), cell(20, "5"), cell(20, "6")])])
drv = table([row([cell(20, "A"), cell(20, "C")]),
             row([cell(20, "1"), cell(20, "3")]),
             row([cell(20, "4"), cell(20, "6")])])
d = run(tpl, drv)
k = kinds(d)
ok = (len(idxs(d, "tc")) == 1) and k["changed"] == 0 and k["added"] == 0
check("S2 删除中间列", ok,
      f"removed列索引集合={idxs(d,'tc')} kinds={k}（期望 仅1列removed、0 changed、无级联）")

# ===== S3 插入列（表：[A,B] → [A,B,C]）=====
tpl = table([row([cell(20, "A"), cell(20, "B")]),
             row([cell(20, "1"), cell(20, "2")])])
drv = table([row([cell(20, "A"), cell(20, "B"), cell(20, "C")]),
             row([cell(20, "1"), cell(20, "2"), cell(20, "3")])])
d = run(tpl, drv)
k = kinds(d)
ok = (idxs(d, "tc") == {2}) and k["changed"] == 0 and k["removed"] == 0
check("S3 插入列", ok,
      f"added列索引={idxs(d,'tc')} kinds={k}（期望 added仅tc2、0 changed/removed）")

# ===== S4 正文插段（body: p0, tbl0, tbl1 → p0, tbl0, pNEW, tbl1）=====
tpl = para(24, "标题段") + table([row([cell(20, "a"), cell(20, "b")])]) + \
      table([row([cell(20, "x"), cell(20, "y")])])
drv = para(24, "标题段") + table([row([cell(20, "a"), cell(20, "b")])]) + \
      para(20, "中段插入说明（内容编辑）") + table([row([cell(20, "x"), cell(20, "y")])])
d = run(tpl, drv)
k = kinds(d)
ok = k["changed"] == 0 and k["removed"] == 0 and k["added"] >= 1
check("S4 正文插段", ok,
      f"kinds={k}（期望 仅 added≥1、表内容0偏差、无级联）")

# ===== S5 多格式维度改动（同单元格 sz 20→40 且 color 000000→FF0000）=====
tpl = table([row([cell_ex(20, "样本", "000000")])])
drv = table([row([cell_ex(40, "样本", "FF0000")])])
d = run(tpl, drv)
k = kinds(d)
attrs = {x['element'] for x in d if x['kind'] == 'changed'}
ok = k["changed"] == 2 and attrs == {"sz", "color"}
check("S5 多格式维度改动", ok,
      f"changed={k['changed']} attrs={attrs}（期望 2 处 changed：sz 与 color）")

# ===== S6 合并单元格（模板 [A,B] → 派生 [A(gridSpan=2)]，探索性）=====
tpl = table([row([cell(20, "A"), cell(20, "B")])])
drv = table([row([cell_merge("A合并B")])])
d = run(tpl, drv)
k = kinds(d)
print(f"\n[S6 合并单元格 · 探索] kinds={k} 偏差={len(d)}")
for x in d[:8]:
    print(f"   {x['kind']:7} {x['path'][:54]:54} {x['element']}/{x['attr']} "
          f"t={x['template']} d={x['derived']}")
check("S6 合并单元格(探索)", True, "见上方明细，不强制断言")

# ===== S7 单元格内容对调（[甲,乙] → [乙,甲]，字号不变）=====
tpl = table([row([cell(20, "甲"), cell(20, "乙")])])
drv = table([row([cell(20, "乙"), cell(20, "甲")])])
d = run(tpl, drv)
ok = len(d) == 0
check("S7 单元格内容对调", ok,
      f"偏差数={len(d)}（期望 0：内容不参与格式比较）")

# ===== S8 同时插入行 + 内容编辑 =====
tpl = table([row([cell(20, "制造商"), cell(20, "X公司")]),
             row([cell(20, "型号"), cell(20, "ABC-123")])])
drv = table([row([cell(20, "制造商"), cell(20, "X公司改")]),
             row([cell(20, "型号"), cell(20, "ABC-123")]),
             row([cell(20, "生产日期"), cell(20, "2026-09")])])
d = run(tpl, drv)
k = kinds(d)
ok = (idxs(d, "tr") == {2}) and k["changed"] == 0 and k["removed"] == 0
check("S8 插行+内容编辑", ok,
      f"added行索引={idxs(d,'tr')} kinds={k}（期望 added仅tr2、0 changed/removed）")

# ===== S9 删除整张表（body: p0,tbl0,tbl1 → p0,tbl1，探索性）=====
tpl = para(24, "标题") + table([row([cell(20, "a"), cell(20, "b")])]) + \
      table([row([cell(20, "x"), cell(20, "y")])])
drv = para(24, "标题") + table([row([cell(20, "x"), cell(20, "y")])])
d = run(tpl, drv)
k = kinds(d)
# 被删的应是「恰好一张表」的整棵子树（removed 偏差只属于同一个相对 tbl 索引），
# 留存的表 0 偏差。相对标注下被删表会标到序列尾，故断言：removed 仅来自 1 个 tbl 索引。
rem_tbl = idxs(d, "tbl")
ok = (len(rem_tbl) == 1) and k["changed"] == 0 and k["added"] == 0
print(f"\n[S9 删除整表 · 探索] kinds={k} 总偏差={len(d)}，removed所属tbl索引={rem_tbl}")
check("S9 删除整表(探索)", ok, "仅1张表整棵被removed、留存表0偏差、无changed/added")

print("\n==== 汇总 ====")
np = sum(1 for _, o, _ in results if o)
print(f"{np}/{len(results)} 项通过")
sys.exit(0 if np == len(results) else 1)
