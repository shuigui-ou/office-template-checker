"""构造邮件格式校验样例：template.eml + derived.eml。

模板：两段 HTML，字体/颜色/对齐不同（对应"句1/句C"模型）。
派生：改第二段 font-family(color) 触发偏差；第一段留未替换占位符 {{客户名}} 触发 ❌。
"""
TEMPLATE = """\
From: noreply@example.com
To: user@example.com
Subject: 账户通知
MIME-Version: 1.0
Content-Type: text/html; charset=utf-8

<html><body>
<p style="font-family: Arial; font-size: 14px; color: #333333; text-align: left;">尊敬的客户：</p>
<p style="font-family: 'Microsoft YaHei'; font-size: 16px; color: #FF0000; text-align: center;">这是第二段正文。</p>
</body></html>
"""

DERIVED = """\
From: noreply@example.com
To: user@example.com
Subject: 账户通知
MIME-Version: 1.0
Content-Type: text/html; charset=utf-8

<html><body>
<p style="font-family: Arial; font-size: 14px; color: #333333; text-align: left;">尊敬的{{客户名}}：</p>
<p style="font-family: SimSun; font-size: 16px; color: #0000FF; text-align: center;">这是第二段正文。</p>
</body></html>
"""

with open("template.eml", "w", encoding="utf-8") as f:
    f.write(TEMPLATE)
with open("derived.eml", "w", encoding="utf-8") as f:
    f.write(DERIVED)
print("written: template.eml, derived.eml")
