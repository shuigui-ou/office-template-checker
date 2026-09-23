#!/usr/bin/env python3
# 端到端验证样本构造：模拟真实工作流
#   模板 = 一份电池 UN38.3 检测报告模板
#   新文件 = 以模板为基础手动编辑：① 插入行(结构变动) ② 改字号(格式偏离)
#           ③ 改正文/单元格文字(内容编辑) ④ 新增样品信息行
# 构造后自动跑 compare + render_report，写出 samples/report.md 供人工核对可读性。
import os, zipfile, sys, json

SKILL = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, SKILL)
import checker as C

SAMPLES = os.path.join(SKILL, "samples")
os.makedirs(SAMPLES, exist_ok=True)

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


def cell(sz, text, bold=False):
    rpr = f'<w:rPr><w:sz w:val="{sz}"/></w:rPr>'
    return (f'<w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr>'
            f'<w:p><w:r>{rpr}<w:t xml:space="preserve">{text}</w:t></w:r></w:p></w:tc>')


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


# ---------- 模板 ----------
tpl_body = (
    para(36, "电池 UN38.3 检测报告") +
    para(24, "本报告依据联合国《试验和标准手册》第3部分38.3节出具。") +
    table([
        row([cell(20, "制造商"), cell(20, "深圳市某电池有限公司")]),
        row([cell(20, "型号"),   cell(20, "ABC-123")]),
        row([cell(20, "额定容量"), cell(20, "3000mAh")]),
    ]) +
    table([
        row([cell(20, "测试项目"), cell(20, "结果"), cell(20, "判定")]),
        row([cell(20, "T.1 高度模拟"), cell(20, "通过"), cell(20, "合格")]),
        row([cell(20, "T.2 热冲击"),  cell(20, "通过"), cell(20, "合格")]),
        row([cell(20, "T.3 振动"),    cell(20, "通过"), cell(20, "合格")]),
    ])
)

# ---------- 新文件（基于模板手动编辑）----------
# 编辑1：正文段落文字被改（内容编辑，应不报格式偏离）
# 编辑2：型号单元格文字 ABC-123 → ABC-123X（表格内内容编辑，观察表现）
# 编辑3：样品信息表 末尾插入「生产日期 | 2026-09」行（结构变动 → added）
# 编辑4：测试结果 T.3 判定单元格字号 20→40（格式偏离 → changed）
# 编辑5：测试结果表 末尾插入「T.4 撞击 | 通过 | 合格」行（结构变动 → added）
drv_body = (
    para(36, "电池 UN38.3 检测报告") +
    para(24, "本报告依据联合国《试验和标准手册》第3部分38.3节及客户补充要求出具。") +
    table([
        row([cell(20, "制造商"), cell(20, "深圳市某电池有限公司")]),
        row([cell(20, "型号"),   cell(20, "ABC-123X")]),
        row([cell(20, "额定容量"), cell(20, "3000mAh")]),
        row([cell(20, "生产日期"), cell(20, "2026-09")]),
    ]) +
    table([
        row([cell(20, "测试项目"), cell(20, "结果"), cell(20, "判定")]),
        row([cell(20, "T.1 高度模拟"), cell(20, "通过"), cell(20, "合格")]),
        row([cell(20, "T.2 热冲击"),  cell(20, "通过"), cell(20, "合格")]),
        # T.3 判定字号改成 40（文本 合格 不变）
        row([cell(20, "T.3 振动"), cell(20, "通过"), cell(40, "合格")]),
        row([cell(20, "T.4 撞击"), cell(20, "通过"), cell(20, "合格")]),
    ])
)

T = os.path.join(SAMPLES, "template.docx")
D = os.path.join(SAMPLES, "derived.docx")
build(T, tpl_body)
build(D, drv_body)

cfg = json.load(open(os.path.join(SKILL, "param_tree.json"), encoding="utf-8"))
pt = C.ParamTree(cfg)
devs, red = C.compare(T, D, cfg, pt)
md = C.render_report(T, D, devs, red, pt)
out = os.path.join(SAMPLES, "report.md")
open(out, "w", encoding="utf-8").write(md)

print(f"样本已生成：\n  {T}\n  {D}\n  {out}")
print(f"偏差 {len(devs)} 条，❌ {len(red)} 条\n")
print(md)
