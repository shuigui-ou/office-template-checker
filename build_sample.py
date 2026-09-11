import zipfile

CT = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>'''

RELS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>'''


def para(sz_val, text):
    return f'''    <w:p>
      <w:pPr><w:spacing w:before="240" w:after="120" w:line="360" w:lineRule="auto"/></w:pPr>
      <w:r><w:rPr><w:sz w:val="{sz_val}"/></w:rPr><w:t>{text}</w:t></w:r>
    </w:p>'''


def doc(body):
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
{body}
  </w:body>
</w:document>'''


def write(name, xml):
    with zipfile.ZipFile(name, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml', CT)
        z.writestr('_rels/.rels', RELS)
        z.writestr('word/document.xml', xml)


# 模板：句1 sz=36，句C(第2段) sz=24 —— 两者都有 A 参数(sz)，但值不同
write('template.docx', doc(para(36, 'Sentence-1') + para(24, 'Sentence-C')))
# 派生：句1 不变(sz=36)，句C 被改成 sz=40 —— 用来验证只报句C的变化
write('derived.docx', doc(para(36, 'Sentence-1') + para(40, 'Sentence-C')))
print('template.docx + derived.docx written')
