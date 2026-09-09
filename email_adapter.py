"""邮件格式适配器：把 .eml 归一化为与 docx/xlsx/pptx 同构的 Prop 流。

设计目标（与 office_adapter 完全一致）：
- extract() 产出带位置索引的 Prop，path 形如
    html[1]/body[1]/div[1]/p[2]/... （对应"句1/句C"模型）
  同参数不同位置天然不同 key，永不合并。
- 邮件头 From/To/Subject/Content-Type(charset) 作为 part="header" 的 Prop。
- HTML 正文内联样式 style="font-family:...; color:..." 拆成多个 Prop：
    element="style", attr=css 属性名（font-family/font-size/color/line-height/text-align...）
- 颜色/字号/行高做归一化，使参数树 tolerance/rgb 模式可用：
    color   -> #RRGGBB（命名色/ rgb()/ 3位hex 均归一）
    font-size -> 纯数值（px 基准：pt*1.3333, em/rem*16）
    line-height -> 纯数值（150% -> 1.5）
- 占位符 {{客户名}} 在正文文本里 -> 产出 attr="#text" 的 Prop，供引擎标红
  "未替换占位符"（与 docx 的 volatile 同一 red_flags 通道）。
- 易变头字段（Date/Message-ID/Received/Return-Path/boundary）默认不产出，
  避免噪声淹没信号。

零第三方依赖：email（标准库）+ html.parser（标准库）。
.msg（Outlook OLE）需 olefile，暂未实现，is_instance 只认 .eml。
"""
import email
import re
from email.header import decode_header, make_header
from html.parser import HTMLParser

from adapter_interface import FormatAdapter, Prop, StructuralAnchor, AdapterRegistry

EMAIL_EXT = (".eml",)

# 常见 CSS 命名色 -> hex（覆盖 16 标准色 + 少量常用，够日常邮件）
NAMED_COLORS = {
    "black": "#000000", "white": "#FFFFFF", "red": "#FF0000", "green": "#008000",
    "lime": "#00FF00", "blue": "#0000FF", "yellow": "#FFFF00", "aqua": "#00FFFF",
    "cyan": "#00FFFF", "magenta": "#FF00FF", "fuchsia": "#FF00FF", "gray": "#808080",
    "grey": "#808080", "silver": "#C0C0C0", "maroon": "#800000", "olive": "#808000",
    "purple": "#800080", "teal": "#008080", "navy": "#000080", "orange": "#FFA500",
    "pink": "#FFC0CB", "brown": "#A52A2A", "gold": "#FFD700", "darkgray": "#A9A9A9",
    "lightgray": "#D3D3D3", "darkblue": "#00008B", "darkgreen": "#006400",
    "darkred": "#8B0000", "indigo": "#4B0082", "violet": "#EE82EE",
}

# HTML void 元素（无结束标签，路径栈不压）
VOID_TAGS = {
    "area", "base", "br", "col", "embed", "hr", "img", "input",
    "link", "meta", "param", "source", "track", "wbr",
}
# 跳过其文本内容的块（CSS 块 / 脚本，不做内联样式解析）
SKIP_TEXT_TAGS = {"style", "script", "head"}

# 邮件头：只产出这些有意义的字段（其余易变/噪声默认不抽）
HEADER_FIELDS = ["From", "To", "Cc", "Subject", "Reply-To", "Return-Path"]


# ---------- 归一化 ----------
def norm_color(v: str) -> str:
    v = (v or "").strip()
    if not v:
        return v
    m = re.fullmatch(r"#([0-9a-fA-F]{3})", v)
    if m:
        return "#" + "".join(c * 2 for c in m.group(1)).upper()
    m = re.fullmatch(r"#([0-9a-fA-F]{6})", v)
    if m:
        return "#" + m.group(1).upper()
    m = re.fullmatch(r"rgba?\(\s*([\d.]+)\s*,\s*([\d.]+)\s*,\s*([\d.]+)", v)
    if m:
        r, g, b = (int(round(float(x))) for x in m.groups())
        return f"#{r:02X}{g:02X}{b:02X}"
    name = v.lower()
    if name in NAMED_COLORS:
        return NAMED_COLORS[name]
    return v


def norm_size(v: str) -> str:
    v = (v or "").strip().lower()
    m = re.fullmatch(r"([\d.]+)px", v)
    if m:
        return f"{float(m.group(1)):.2f}"
    m = re.fullmatch(r"([\d.]+)pt", v)
    if m:
        return f"{float(m.group(1)) * 1.3333:.2f}"
    m = re.fullmatch(r"([\d.]+)em", v)
    if m:
        return f"{float(m.group(1)) * 16:.2f}"
    m = re.fullmatch(r"([\d.]+)rem", v)
    if m:
        return f"{float(m.group(1)) * 16:.2f}"
    return v  # 无法归一则保留原值


def norm_line_height(v: str) -> str:
    v = (v or "").strip().lower()
    m = re.fullmatch(r"([\d.]+)%", v)
    if m:
        return f"{float(m.group(1)) / 100:.3f}"
    m = re.fullmatch(r"([\d.]+)", v)
    if m:
        return f"{float(m.group(1)):.3f}"
    return v


def norm_family(v: str) -> str:
    parts = [p.strip().strip('"').strip("'") for p in (v or "").split(",")]
    parts = [p for p in parts if p]
    return parts[0] if parts else (v or "")


def parse_style(css: str):
    """拆 style="a:b; c:d" -> [(prop, value), ...]，并对关键属性归一化。"""
    out = []
    if not css:
        return out
    for decl in css.split(";"):
        if ":" not in decl:
            continue
        k, val = decl.split(":", 1)
        k = k.strip().lower()
        val = val.strip()
        if not k:
            continue
        if k == "font-family":
            val = norm_family(val)
        elif k == "color":
            val = norm_color(val)
        elif k == "background-color":
            val = norm_color(val)
        elif k == "font-size":
            val = norm_size(val)
        elif k == "line-height":
            val = norm_line_height(val)
        out.append((k, val))
    return out


# ---------- HTML 路径追踪 walker ----------
class EmailHTMLWalker(HTMLParser):
    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.props: list = []
        self.placeholders: list = []   # [(name, path), ...]
        self.stack: list = []          # [(tag, idx), ...] 当前路径（不含正在处理的元素）
        self.counts: dict = {}         # path_key -> {tag: n}
        self.skip_depth = 0            # style/script 内不抽文本

    def _path_key(self):
        return tuple(t for t, _ in self.stack)

    def _cur_path(self):
        return [f"{t}[{i}]" for t, i in self.stack]

    def handle_starttag(self, tag, attrs):
        tag = tag.lower()
        if tag in SKIP_TEXT_TAGS:
            self.skip_depth += 1
        key = self._path_key()
        d = self.counts.setdefault(key, {})
        d[tag] = d.get(tag, 0) + 1
        idx = d[tag]
        path = self._cur_path() + [f"{tag}[{idx}]"]
        p = "/".join(path)
        for k, v in attrs:
            if v is None:
                continue
            kl = k.lower()
            if kl == "style":
                for prop, val in parse_style(v):
                    self.props.append(Prop(part="body.html", path=p,
                                           element="style", attr=prop, value=val))
            elif kl in ("class", "id"):
                self.props.append(Prop(part="body.html", path=p,
                                       element=tag, attr=kl, value=v))
        if tag not in VOID_TAGS and tag not in SKIP_TEXT_TAGS:
            self.stack.append((tag, idx))

    def handle_endtag(self, tag):
        tag = tag.lower()
        if tag in SKIP_TEXT_TAGS and self.skip_depth > 0:
            self.skip_depth -= 1
        if self.stack and self.stack[-1][0] == tag:
            self.stack.pop()

    def handle_data(self, data):
        if self.skip_depth > 0:
            return
        if "{{" not in data:
            return
        p = "/".join(self._cur_path())
        for name in re.findall(r"\{\{[^}]+\}\}", data):
            self.placeholders.append((name, p))
        self.props.append(Prop(part="body.html", path=p,
                               element="#text", attr="#text", value=data.strip()))


def _decode_hdr(val: str) -> str:
    if not val:
        return ""
    try:
        return str(make_header(decode_header(val)))
    except Exception:
        return val


class EmailAdapter(FormatAdapter):
    """邮件(.eml)格式适配器。.msg 待 olefile 后续接入。"""

    fmt = "email"

    def is_instance(self, path: str) -> bool:
        if path.lower().endswith(EMAIL_EXT):
            return True
        # 也识别"无扩展名但确实是 MIME 文本"的兜底（可选）
        try:
            with open(path, "rb") as f:
                head = f.read(1024)
            return bool(re.search(rb"^(From|To|Subject|Content-Type|MIME-Version):",
                                   head, re.M))
        except Exception:
            return False

    def theme_map(self, path: str) -> dict:
        # 邮件无主题色板
        return {}

    def extract(self, path: str) -> list:
        props = []
        with open(path, "rb") as f:
            msg = email.message_from_binary_file(f)

        # ---- 邮件头（仅有意义字段）----
        for h in HEADER_FIELDS:
            val = msg.get(h)
            if val:
                props.append(Prop(part="header", path="header",
                                  element=h, attr="value",
                                  value=_decode_hdr(val)))
        ct = (msg.get_content_type() or "").lower()
        props.append(Prop(part="header", path="header",
                          element="content-type", attr="mime", value=ct))
        charset = (msg.get_content_charset() or "").lower()
        if charset:
            props.append(Prop(part="header", path="header",
                              element="content-type", attr="charset", value=charset))

        # ---- HTML 正文 ----
        html = self._get_html(msg)
        if html is not None:
            walker = EmailHTMLWalker()
            try:
                walker.feed(html)
            except Exception:
                pass
            props.extend(walker.props)
        return props

    def anchors(self, path: str) -> list:
        """结构锚点：未替换占位符 {{...}}。"""
        out = []
        with open(path, "rb") as f:
            msg = email.message_from_binary_file(f)
        html = self._get_html(msg)
        if html is None:
            return out
        walker = EmailHTMLWalker()
        try:
            walker.feed(html)
        except Exception:
            pass
        for name, p in walker.placeholders:
            out.append(StructuralAnchor(kind="placeholder", name=name, path=p))
        return out

    # ---- 内部工具 ----
    def _get_html(self, msg) -> str:
        charset = msg.get_content_charset() or "utf-8"
        if msg.is_multipart():
            for part in msg.walk():
                if part.get_content_type() == "text/html":
                    payload = part.get_payload(decode=True)
                    if payload is not None:
                        return payload.decode(charset, errors="replace")
            # 回退：取第一个 text/plain 也行（但无内联样式）
            return None
        if msg.get_content_type() == "text/html":
            payload = msg.get_payload(decode=True)
            if payload is not None:
                return payload.decode(charset, errors="replace")
        return None


AdapterRegistry.register(EmailAdapter())
