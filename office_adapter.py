"""Office 适配器：把 docx/xlsx/pptx 归一化为统一的 Prop 流。

- extract() 用通用 XML 遍历，产出带位置索引的 Prop（path 含 [n] 索引，
  保证同参数不同位置是不同 key，永不合并）。
- anchors() 返回结构锚点（样式名/书签/占位符），供绝对位置失效时按语义对齐。
- theme_map() 解析主题色板，供 RGB 解算使用。
- 捕获公式(<f>)与域代码(<instrText>)的文本，供 volatile/外部链接标红。
"""
import re
import zipfile
import xml.etree.ElementTree as ET

from adapter_interface import FormatAdapter, Prop, StructuralAnchor, AdapterRegistry

OFFICE_EXT = (".docx", ".xlsx", ".pptx", ".dotx", ".xltx", ".potx")
# 仅对这些元素的文本做捕获（避免把全文 t 文本都卷进来）
TEXT_CAPTURE = {"f", "instrText"}


def _ln(tag: str) -> str:
    return tag.split("}", 1)[1] if "}" in tag else tag


def _local_attr(el, name: str):
    """按本地名取属性值（忽略命名空间前缀）。"""
    for k, v in el.attrib.items():
        if _ln(k) == name:
            return v
    return None


class OfficeAdapter(FormatAdapter):
    fmt = "office"

    # ---- 引擎侧钩子 ----
    def is_instance(self, path: str) -> bool:
        if not path.lower().endswith(OFFICE_EXT):
            return False
        try:
            with zipfile.ZipFile(path) as z:
                return "[Content_Types].xml" in z.namelist()
        except Exception:
            return False

    def extract(self, path: str) -> list:
        props = []
        ss = self._load_shared_strings(path)
        with zipfile.ZipFile(path) as z:
            for n in z.namelist():
                if not (n.endswith(".xml") or n.endswith(".rels")):
                    continue
                try:
                    root = ET.fromstring(z.read(n))
                except Exception:
                    continue
                self._walk(root, [], 1, props, n, None, ss)
        return props

    def _load_shared_strings(self, path: str) -> dict:
        """xlsx 单元格文本常存于 xl/sharedStrings.xml（t="s"，<v> 为索引）。
        解析为 {索引: 文本}，供单元格内容锚点使用。"""
        out = {}
        try:
            with zipfile.ZipFile(path) as z:
                for n in z.namelist():
                    if n.endswith("sharedStrings.xml"):
                        out = self._shared_strings(z.read(n))
                        break
        except Exception:
            pass
        return out

    @staticmethod
    def _shared_strings(data: bytes) -> dict:
        ss = {}
        try:
            root = ET.fromstring(data)
        except Exception:
            return ss
        for i, si in enumerate(root):
            texts = [t.text for t in si.iter()
                     if _ln(t.tag) == "t" and t.text]
            ss[str(i)] = "".join(texts)
        return ss

    def _walk(self, el, pathlst, idx, props, part, para=None, ss=None):
        tag = _ln(el.tag)
        p = pathlst + [f"{tag}[{idx}]"]
        if tag in ("p", "c"):
            # 进入段落(docx/pptx) / 单元格(xlsx)：开启内容锚点上下文
            # （用于相对位置对齐，不参与格式比较）
            para = {"path": "/".join(p), "text": []}
        for k, v in el.attrib.items():
            props.append(Prop(part=part, path="/".join(p),
                              element=tag, attr=_ln(k), value=v))
        # 白名单元素（公式 / 域代码）文本仍按 #text 捕获，供 ❌ 标红
        if tag in TEXT_CAPTURE and el.text and el.text.strip():
            props.append(Prop(part=part, path="/".join(p),
                              element=tag, attr="#text", value=el.text.strip()))
        # 普通文本：累积到最近段落/单元格上下文，作为对齐锚点
        if tag == "t" and el.text and el.text.strip() and para is not None:
            para["text"].append(el.text.strip())
        child_idx = {}
        for c in el:
            ct = _ln(c.tag)
            child_idx[ct] = child_idx.get(ct, 0) + 1
            self._walk(c, p, child_idx[ct], props, part, para, ss)
        # 段落/单元格结束时，把累积文本作为 #ctext 锚点挂到该节点
        # （值稳定、用于对齐；xlsx 共享字符串 t="s" 时从 ss 表解析文本）
        if tag in ("p", "c") and para is not None:
            txt = "".join(para["text"])
            if not txt and tag == "c" and _local_attr(el, "t") == "s" and ss is not None:
                for ch in el:
                    if _ln(ch.tag) == "v" and ch.text is not None:
                        txt = ss.get(ch.text.strip(), "")
                        break
            if txt:
                props.append(Prop(part=part, path="/".join(p),
                                  element=tag, attr="#ctext", value=txt))

    def anchors(self, path: str) -> list:
        """结构锚点：样式名(pStyle)、书签(bookmarkStart)、占位符({{...}})。"""
        out = []
        with zipfile.ZipFile(path) as z:
            main = self._main_part(z)
            if not main:
                return out
            try:
                root = ET.fromstring(z.read(main))
            except Exception:
                return out

            def walk(el, pathlst, idx):
                tag = _ln(el.tag)
                p = pathlst + [f"{tag}[{idx}]"]
                if tag == "pStyle":
                    v = el.get("w:val") or el.get("val")
                    if v:
                        out.append(StructuralAnchor(kind="style", name=v,
                                                   path="/".join(p)))
                elif tag == "bookmarkStart":
                    v = el.get("w:name") or el.get("name")
                    if v:
                        out.append(StructuralAnchor(kind="bookmark", name=v,
                                                   path="/".join(p)))
                elif tag == "t" and el.text and "{{" in el.text:
                    for m in re.findall(r"\{\{[^}]+\}\}", el.text):
                        out.append(StructuralAnchor(kind="placeholder", name=m,
                                                   path="/".join(p)))
                child_idx = {}
                for c in el:
                    ct = _ln(c.tag)
                    child_idx[ct] = child_idx.get(ct, 0) + 1
                    walk(c, p, child_idx[ct])

            walk(root, [], 1)
        return out

    def theme_map(self, path: str) -> dict:
        m = {}
        try:
            with zipfile.ZipFile(path) as z:
                for n in z.namelist():
                    if re.search(r"theme/theme\d+\.xml$", n):
                        try:
                            root = ET.fromstring(z.read(n))
                        except Exception:
                            continue
                        for el in root.iter():
                            if _ln(el.tag) == "clrScheme":
                                for child in el:
                                    name = _ln(child.tag)
                                    for sub in child:
                                        if _ln(sub.tag) == "srgbClr":
                                            m[name] = sub.get("val")
                                        elif _ln(sub.tag) == "sysClr":
                                            m[name] = sub.get("lastClr")
                        break
        except Exception:
            pass
        return m

    # ---- 内部工具 ----
    def _main_part(self, z: zipfile.ZipFile):
        for n in z.namelist():
            if n.endswith("document.xml"):
                return n
            if n.endswith("presentation.xml"):
                return n
            if re.search(r"worksheets/sheet\d+\.xml$", n):
                return n
        return None


AdapterRegistry.register(OfficeAdapter())
