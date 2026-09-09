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
        with zipfile.ZipFile(path) as z:
            for n in z.namelist():
                if not (n.endswith(".xml") or n.endswith(".rels")):
                    continue
                try:
                    root = ET.fromstring(z.read(n))
                except Exception:
                    continue
                self._walk(root, [], 1, props, n)
        return props

    def _walk(self, el, pathlst, idx, props, part):
        tag = _ln(el.tag)
        p = pathlst + [f"{tag}[{idx}]"]
        for k, v in el.attrib.items():
            props.append(Prop(part=part, path="/".join(p),
                              element=tag, attr=_ln(k), value=v))
        # 仅对白名单元素捕获文本（公式 / 域代码），用于 ❌ 标红
        if tag in TEXT_CAPTURE and el.text and el.text.strip():
            props.append(Prop(part=part, path="/".join(p),
                              element=tag, attr="#text", value=el.text.strip()))
        child_idx = {}
        for c in el:
            ct = _ln(c.tag)
            child_idx[ct] = child_idx.get(ct, 0) + 1
            self._walk(c, p, child_idx[ct], props, part)

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
