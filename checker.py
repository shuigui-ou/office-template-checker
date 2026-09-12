"""模板一致性校验器主程序。

设计要点（已与用户共识）：
- 比对底层 = 对称全量枚举 + diff，非固定清单；参数树只是过滤器/聚合层。
- 比较 key = (part, path, element, attr)，path 含位置索引 → 同参数不同位置是不同 key，永不合并。
- 5 种匹配模式：exact / rgb / theme_slot / tolerance / ignore。
- 三档能力：识别(读 XML) > 校验(比对) > 生成；渲染层/版本默认/外部目标等为 ❌ 需标红。

用法：
  # 建基线
  python checker.py --baseline template.docx --out baseline.json
  # 单文件比对
  python checker.py --template template.docx --derived derived.docx \
        --config param_tree.json --out report.md --json report.json
  # 批量目录
  python checker.py --template template.docx --derived-dir ./deriveds/ \
        --config param_tree.json --out-dir ./reports/
"""
import argparse
import fnmatch
import hashlib
import json
import os
import re
import sys
import tempfile
import zipfile

from adapter_interface import AdapterRegistry, Prop
import office_adapter  # 注册 OfficeAdapter
import email_adapter   # 注册 EmailAdapter（.eml）

# 常见元素 → 维度标签（用于报告归类与"触碰样式层"判定）
ELEMENT_DIM = {
    "a:latin": "D1", "a:ea": "D1", "a:cs": "D1", "a:schemeClr": "D1",
    "pgMar": "D2", "pgSz": "D2", "cols": "D2",
    "jc": "D3", "ind": "D3", "spacing": "D3", "alignment": "D3",
    "rFonts": "D4", "sz": "D4", "color": "D4", "spacing": "D4",
    "b": "D4", "i": "D4", "u": "D4", "highlight": "D4", "shd": "D4",
    "srgbClr": "D5", "schemeClr": "D5", "ln": "D5", "xfrm": "D5",
    "numFmt": "D6",
    # 邮件层
    "style": "D4", "content-type": "D6",
    "line-height": "D3", "text-align": "D3", "margin": "D3",
    "font-weight": "D4", "#text": "D4",
}

# ❌ 真缺口：volatile / 外部引用 / 缓存域 / 未替换占位符 需人工复核
VOLATILE_RE = re.compile(r"\b(TODAY|NOW|RAND|OFFSET|INDIRECT)\b", re.I)
EXTERNAL_RE = re.compile(r"\[[^\]]+\][^!]*!")
FIELDCACHE_RE = re.compile(r"\b(TOC|REF|INDEX|SEQ|STYLEREF)\b", re.I)
PLACEHOLDER_RE = re.compile(r"\{\{[^}]+\}\}")


class CheckError(Exception):
    """对用户友好的错误：直接显示中文原因与修复建议，不抛栈。"""


def _fail(msg: str) -> None:
    print("✗ " + msg, file=sys.stderr)
    sys.exit(1)


def safe_extract(path: str, cfg=None):
    """带中文错误提示的抽取入口，统一兜底所有解析异常。

    返回 (adapter, props)；任何失败都以 CheckError 抛出可读中文信息，
    不让用户面对英文栈或闪退。
    """
    if not os.path.exists(path):
        raise CheckError(
            f"找不到文件：{path}\n"
            f"  请检查路径与文件名是否正确（中文路径需确保终端编码为 UTF-8）。")
    ext = os.path.splitext(path)[1].lower()
    office_ext = (".docx", ".xlsx", ".pptx", ".dotx", ".xltx", ".potx")
    if ext not in office_ext and ext != ".eml":
        raise CheckError(
            f"不支持的文件格式：{path}\n"
            f"  仅支持 .docx / .xlsx / .pptx / .eml。\n"
            f"  说明：.msg（Outlook）暂不支持，请先另存为 .eml；PDF 扫描件不支持。")
    # 扩展名合法，但适配器可能因文件损坏而拒绝 → 区分"损坏"与"不支持"
    try:
        adapter = AdapterRegistry.for_file(path)
    except ValueError:
        if ext in office_ext:
            raise CheckError(
                f"文件已损坏或不是有效的 Office 文档：{path}\n"
                f"  扩展名虽为 Office，但内部不是合法的 zip 包——可能被改名，"
                f"或实际是 PDF / 图片 / 加密文件。\n"
                f"  建议：用原生 Office 重新另存为 .docx / .xlsx / .pptx。")
        raise CheckError(
            f"不支持的文件格式：{path}\n"
            f"  仅支持 .docx / .xlsx / .pptx / .eml。")
    try:
        props = adapter.extract(path)
    except zipfile.BadZipFile:
        raise CheckError(
            f"文件已损坏或不是有效的 Office 文档：{path}\n"
            f"  扩展名虽为 Office，但内部不是合法的 zip 包——可能被改名，"
            f"或实际是 PDF / 图片 / 加密文件。\n"
            f"  建议：用原生 Office 重新另存为 .docx / .xlsx / .pptx。")
    except FileNotFoundError:
        raise CheckError(f"找不到文件：{path}")
    except Exception as e:
        raise CheckError(
            f"解析文件时出错：{path}\n"
            f"  原因：{e}\n"
            f"  建议：确认文件未损坏且为 Office 原生格式"
            f"（加密文档、PDF 扫描件、图片均不支持）。")
    return adapter, props


class ParamTree:
    def __init__(self, cfg: dict):
        self.cfg = cfg
        self.enabled_dims = {d: v.get("enabled", True)
                             for d, v in cfg.get("dimensions", {}).items()}
        # (element, attr) -> (dimension_label, mode_spec)
        self.rules = {}
        for dlabel, dv in cfg.get("dimensions", {}).items():
            if not dv.get("enabled", True):
                continue
            for p, pv in dv.get("params", {}).items():
                if not pv.get("enabled", True):
                    continue
                for sp, spv in pv.get("subparams", {}).items():
                    el, attr = (sp.split("/", 1) + [None])[:2]
                    self.rules[(el, attr)] = (dlabel, spv)
        self.ignore_attr = cfg.get("ignore_attr_patterns", [])
        self.theme_slot_attrs = set(cfg.get("theme_slot_attrs", []))
        self.tolerances = cfg.get("tolerances", {})
        self.path_overrides = cfg.get("path_overrides", [])

    def _is_ignored_path(self, prop: Prop) -> bool:
        for po in self.path_overrides:
            if re.search(po.get("path_regex", "$.^"), prop.path):
                return po.get("match") == "ignore"
        return False

    def is_noise(self, prop: Prop) -> bool:
        # 噪声忽略：属性名通配 + 维度被关
        for pat in self.ignore_attr:
            if fnmatch.fnmatch(prop.attr, pat) or \
               fnmatch.fnmatch(f"{prop.element}/{prop.attr}", pat):
                return True
        return False

    def mode_for(self, prop: Prop):
        for po in self.path_overrides:
            if re.search(po.get("path_regex", "$.^"), prop.path):
                return po.get("match", "exact"), None
        rule = self.rules.get((prop.element, prop.attr))
        if rule is None:
            rule = self.rules.get((prop.element, None))
        if rule is None:
            return "exact", None
        dlabel, spec = rule
        return spec.get("match", "exact"), spec


# ---------- 颜色解算 ----------
def resolve_rgb(value: str, theme_map: dict) -> str:
    v = (value or "").strip()
    if re.fullmatch(r"[0-9A-Fa-f]{6}", v):
        return v.upper()
    if re.fullmatch(r"[0-9A-Fa-f]{3}", v):
        return "".join(c * 2 for c in v).upper()
    if v in theme_map:
        return (theme_map[v] or "").upper()
    return v.upper()


def dim_of(element: str) -> str:
    return ELEMENT_DIM.get(element, "other")


def layer_of(part: str) -> str:
    """样式层 vs 内容/直接格式层（启发式）。"""
    if "styles.xml" in part or "numbering.xml" in part or "theme/" in part:
        return "style"
    return "content"


# ---------- 匹配 ----------
def values_match(mode, spec, tv, dv, t_theme, d_theme):
    if mode == "ignore":
        return True
    if mode == "exact":
        return str(tv) == str(dv)
    if mode == "rgb":
        return resolve_rgb(tv, t_theme) == resolve_rgb(dv, d_theme)
    if mode == "theme_slot":
        return str(tv) == str(dv)
    if mode == "tolerance":
        try:
            a, b = float(tv), float(dv)
        except Exception:
            return str(tv) == str(dv)
        tol = (spec or {}).get("tol", 0)
        return abs(a - b) <= tol
    return str(tv) == str(dv)


# ---------- 核心比对 ----------
def build_map(props):
    return {(p.part, p.path, p.element, p.attr): p.value for p in props}


def fingerprint(path: str) -> str:
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(8192), b""):
            h.update(chunk)
    return h.hexdigest()


def build_baseline(template_path, cfg, out_path):
    adapter, props = safe_extract(template_path, cfg)
    theme = adapter.theme_map(template_path)
    baseline = {
        "template": os.path.abspath(template_path),
        "template_version": fingerprint(template_path),
        "prop_count": len(props),
        "theme_map": theme,
        "props": [p.__dict__ for p in props],
    }
    with open(out_path, "w", encoding="utf-8") as f:
        json.dump(baseline, f, ensure_ascii=False, indent=2)
    print(f"[baseline] {template_path} -> {out_path} "
          f"({len(props)} props, version {baseline['template_version'][:12]})")


def scan_red_flags(props):
    flags = []
    for p in props:
        if p.attr != "#text":
            continue
        v = p.value or ""
        if VOLATILE_RE.search(v):
            flags.append({"type": "volatile", "loc": p.path,
                          "detail": VOLATILE_RE.search(v).group(1)})
        elif EXTERNAL_RE.search(v):
            flags.append({"type": "external_link", "loc": p.path, "detail": v[:60]})
        elif FIELDCACHE_RE.search(v):
            flags.append({"type": "cached_field", "loc": p.path,
                          "detail": FIELDCACHE_RE.search(v).group(1)})
        elif PLACEHOLDER_RE.search(v):
            flags.append({"type": "unresolved_placeholder", "loc": p.path,
                          "detail": PLACEHOLDER_RE.search(v).group(0)})
    return flags


def compare(template_path, derived_path, cfg, pt: ParamTree):
    adapter_t, props_t = safe_extract(template_path, cfg)
    adapter_d, props_d = safe_extract(derived_path, cfg)
    theme_t = adapter_t.theme_map(template_path)
    theme_d = adapter_d.theme_map(derived_path)
    mt, md = build_map(props_t), build_map(props_d)

    deviations = []
    seen = set()
    for key in sorted(set(mt) | set(md), key=lambda x: (x[0], x[1], x[2], x[3])):
        part, path, element, attr = key
        prop_t = Prop(part, path, element, attr, mt.get(key))
        prop_d = Prop(part, path, element, attr, md.get(key))
        # 噪声/忽略
        if pt._is_ignored_path(prop_t) or pt._is_ignored_path(prop_d):
            continue
        if pt.is_noise(prop_t) or pt.is_noise(prop_d):
            continue
        tv, dv = mt.get(key), md.get(key)
        if tv is not None and dv is not None:
            mode, spec = pt.mode_for(prop_t)
            ok = values_match(mode, spec, tv, dv, theme_t, theme_d)
            if ok:
                continue
            kind = "changed"
        elif tv is None:
            mode, spec = pt.mode_for(prop_d)
            if mode == "ignore":
                continue
            kind = "added"      # 模板没有、派生新增 → 偏离模板
        else:
            mode, spec = pt.mode_for(prop_t)
            if mode == "ignore":
                continue
            kind = "removed"    # 模板有、派生缺失
        deviations.append({
            "dimension": dim_of(element),
            "layer": layer_of(part),
            "part": part, "path": path,
            "element": element, "attr": attr,
            "mode": mode,
            "template": tv,
            "derived": dv,
            "kind": kind,
        })
    red = scan_red_flags(props_t) + scan_red_flags(props_d)
    return deviations, red


def render_report(template_path, derived_path, deviations, red, pt: ParamTree):
    lines = []
    lines.append(f"# 模板一致性偏差报告")
    lines.append("")
    lines.append(f"- 模板：`{os.path.abspath(template_path)}`")
    lines.append(f"- 派生：`{os.path.abspath(derived_path)}`")
    lines.append(f"- 偏差总数：{len(deviations)}　|　❌需人工复核：{len(red)}")
    lines.append("")
    by_dim = {}
    for d in deviations:
        by_dim.setdefault(d["dimension"], []).append(d)
    for dim in sorted(by_dim):
        lines.append(f"## {dim} 维度（{len(by_dim[dim])} 处）")
        lines.append("")
        lines.append("| 类型 | 位置 | 属性 | 模板 | 派生 | 层 | 模式 |")
        lines.append("|---|---|---|---|---|---|---|")
        for d in by_dim[dim]:
            loc = f"{d['part'].split('/')[-1]}:{d['path'][:48]}"
            lines.append(f"| {d['kind']} | {loc} | {d['element']}/{d['attr']} "
                         f"| {d['template']} | {d['derived']} | {d['layer']} | {d['mode']} |")
        lines.append("")
    if red:
        lines.append("## ❌ 需人工复核（真缺口：volatile/外部链接/缓存域/未替换占位符）")
        lines.append("")
        lines.append("| 类型 | 位置 | 说明 |")
        lines.append("|---|---|---|")
        for r in red:
            lines.append(f"| {r['type']} | {r['loc'][:60]} | {r['detail']} |")
        lines.append("")
    lines.append("---")
    lines.append("> 说明：比较 key 含位置索引，同参数不同位置互不合并；"
                 "报告仅含参数树未忽略的差异；渲染层/版本默认/隐式默认值不在范围内。")
    return "\n".join(lines)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--baseline", help="模板路径，仅建基线")
    ap.add_argument("--out", help="输出路径（baseline=json, 单文件=md）")
    ap.add_argument("--template", help="模板路径")
    ap.add_argument("--derived", help="派生文件")
    ap.add_argument("--derived-dir", help="派生文件目录（批量）")
    ap.add_argument("--config", default="param_tree.json")
    ap.add_argument("--json", help="额外输出结构化 JSON 路径")
    ap.add_argument("--out-dir", help="批量报告输出目录")
    ap.add_argument("--demo", action="store_true",
                   help="运行内置演示：自动生成模板/派生样例并比对，展示标准报告长什么样")
    args = ap.parse_args()

    if args.demo:
        run_demo()
        return

    try:
        with open(args.config, encoding="utf-8") as f:
            cfg = json.load(f)
    except FileNotFoundError:
        _fail(f"找不到参数树配置文件：{args.config}\n"
              f"  请确认 --config 指向正确的 param_tree.json"
              f"（默认即本目录下的 param_tree.json）。")
    except json.JSONDecodeError as e:
        _fail(f"参数树配置文件不是合法 JSON：{args.config}\n  原因：{e}")
    pt = ParamTree(cfg)

    if args.baseline:
        out = args.out or "baseline.json"
        build_baseline(args.baseline, cfg, out)
        return

    if not args.template or not (args.derived or args.derived_dir):
        ap.error("需提供 --template 与 --derived / --derived-dir")

    if args.derived:
        deviations, red = compare(args.template, args.derived, cfg, pt)
        md = render_report(args.template, args.derived, deviations, red, pt)
        out = args.out or "report.md"
        with open(out, "w", encoding="utf-8") as f:
            f.write(md)
        if args.json:
            with open(args.json, "w", encoding="utf-8") as f:
                json.dump({"deviations": deviations, "red_flags": red},
                          f, ensure_ascii=False, indent=2)
        print(f"[report] {out}  偏差={len(deviations)}  ❌={len(red)}")
        return

    if args.derived_dir:
        out_dir = args.out_dir or "reports"
        os.makedirs(out_dir, exist_ok=True)
        summary = []
        for fn in sorted(os.listdir(args.derived_dir)):
            if not fn.lower().endswith((".docx", ".xlsx", ".pptx", ".eml")):
                continue
            dp = os.path.join(args.derived_dir, fn)
            devs, red = compare(args.template, dp, cfg, pt)
            base = os.path.splitext(fn)[0]
            md = render_report(args.template, dp, devs, red, pt)
            with open(os.path.join(out_dir, base + ".md"), "w", encoding="utf-8") as f:
                f.write(md)
            summary.append({"file": fn, "deviations": len(devs), "red_flags": len(red)})
        with open(os.path.join(out_dir, "_summary.json"), "w", encoding="utf-8") as f:
            json.dump(summary, f, ensure_ascii=False, indent=2)
        print(f"[batch] {len(summary)} files -> {out_dir}")


# ---------- 内置演示 ----------
def _write_demo_docx(name: str, paras):
    """生成最小可用 .docx 演示文件（仅标准库，无第三方依赖）。"""
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
    body = ""
    for sz, text in paras:
        body += (f'    <w:p><w:r><w:rPr><w:sz w:val="{sz}"/></w:rPr>'
                 f'<w:t>{text}</w:t></w:r></w:p>')
    document = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
                '<w:body>' + body + '</w:body></w:document>')
    with zipfile.ZipFile(name, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CT)
        z.writestr("_rels/.rels", RELS)
        z.writestr("word/document.xml", document)


def run_demo():
    """一键演示：生成模板/派生样例 -> 比对 -> 打印标准偏差报告。"""
    here = os.path.dirname(os.path.abspath(__file__))
    cfg_path = os.path.join(here, "param_tree.json")
    if not os.path.exists(cfg_path):
        _fail(f"演示需要同目录的 param_tree.json，但未找到：{cfg_path}")
    tmp = tempfile.mkdtemp(prefix="otc_demo_")
    tpl = os.path.join(tmp, "template.docx")
    drv = os.path.join(tmp, "derived.docx")
    # 模板：两段都是模板样式；派生：仅第2段字号被改（模拟"偏离模板"）
    _write_demo_docx(tpl, [(36, "句一：模板标题（与派生一致）"),
                           (24, "句C：模板正文（字号 24）")])
    _write_demo_docx(drv, [(36, "句一：模板标题（与派生一致）"),
                           (40, "句C：派生被改成 40 磅（偏离！）")])
    cfg = json.load(open(cfg_path, encoding="utf-8"))
    pt = ParamTree(cfg)
    devs, red = compare(tpl, drv, cfg, pt)
    md = render_report(tpl, drv, devs, red, pt)
    out = os.path.join(tmp, "demo_report.md")
    with open(out, "w", encoding="utf-8") as f:
        f.write(md)
    print("=" * 64)
    print("演示：模板第2段字号=24，派生被改为=40，其余完全一致")
    print("=" * 64)
    print(md)
    print(f"\n✓ 演示完成。这就是一份标准『模板一致性偏差报告』。")
    print(f"  报告已写入：{out}")
    print(f"  想用你自己的文件？照『运行方式』里的命令替换 --template/--derived 即可。")


if __name__ == "__main__":
    try:
        main()
    except CheckError as e:
        print("✗ " + str(e), file=sys.stderr)
        sys.exit(1)
    except KeyboardInterrupt:
        sys.exit(130)
    except Exception as e:
        print("✗ 未预期错误：" + str(e), file=sys.stderr)
        print("  如能稳定复现，请把文件与命令发给我以便修复。", file=sys.stderr)
        sys.exit(2)
