#!/usr/bin/env python3
# 回归测试：验证「相对位置」对齐。
# 模板表2 有 3 行(A/B/C)，派生在表2 第1行前插入一行 NEW。
# 期望：NEW 被识别为 added；A/B/C 仍按相对位置正确对应（无级联误报）。
# 另含：无插入、但 B 字号 22→40 → 仅 1 条 changed。
#
# 运行：python tests/test_relative_position.py
import os, zipfile, sys, json, tempfile

SKILL_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, SKILL_DIR)
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

def row(sz, text):
    return ('<w:tr><w:tc><w:p><w:r><w:rPr><w:sz w:val="%s"/></w:rPr>'
            '<w:t>%s</w:t></w:r></w:p></w:tc></w:tr>' % (sz, text))

def make_docx(path, table2_rows):
    tbl1 = '<w:tbl>' + row('36', 'T1') + '</w:tbl>'
    tbl2 = '<w:tbl>' + ''.join(table2_rows) + '</w:tbl>'
    body = '<w:body>%s%s</w:body>' % (tbl1, tbl2)
    doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
           '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
           + body + '</w:document>')
    with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml', CT)
        z.writestr('_rels/.rels', RELS)
        z.writestr('word/document.xml', doc)

tmp = tempfile.mkdtemp(prefix="otc_rel_")
T = os.path.join(tmp, 'tpl_rel.docx')
D1 = os.path.join(tmp, 'drv_insert.docx')     # 表2 前插入 NEW
D2 = os.path.join(tmp, 'drv_change.docx')     # 无插入，但 B 字号改成 40
make_docx(T, [row('20', 'A'), row('22', 'B'), row('24', 'C')])
make_docx(D1, [row('20', 'NEW'), row('20', 'A'), row('22', 'B'), row('24', 'C')])
make_docx(D2, [row('20', 'A'), row('40', 'B'), row('24', 'C')])

cfg = json.load(open(os.path.join(SKILL_DIR, 'param_tree.json'), encoding='utf-8'))
pt = C.ParamTree(cfg)

def run(label, drv):
    devs, red = C.compare(T, drv, cfg, pt)
    print('\n==== %s ==== 偏差 %d 条' % (label, len(devs)))
    for d in devs:
        print('  [%s] 表/行/列相对定位=%s | %s/%s | 模板=%s 派生=%s'
              % (d['kind'], d['path'], d['element'], d['attr'], d['template'], d['derived']))
    return devs

r1 = run('场景1：表2 第1行前插入 NEW（期望：仅 NEW 被 added，A/B/C 无级联）', D1)
r2 = run('场景2：无插入，但 B 字号 22→40（期望：仅 B 命中 changed）', D2)

# 断言
assert len([d for d in r1 if d['kind'] != 'added']) == 0, '场景1 不应有 changed/removed 级联'
assert len([d for d in r1 if d['kind'] == 'added']) >= 1, '场景1 应至少 1 条 added(NEW)'
assert len([d for d in r2 if d['kind'] == 'changed']) == 1, '场景2 应恰好 1 条 changed(B)'
assert len([d for d in r2 if d['kind'] != 'changed']) == 0, '场景2 不应有 added/removed'
print('\n全部断言通过：相对位置对齐正确，无级联误报，且真实格式改动仍被命中。')
