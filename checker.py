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
    adapter = AdapterRegistry.for_file(template_path)
    props = adapter.extract(template_path)
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
    adapter_t = AdapterRegistry.for_file(template_path)
    adapter_d = AdapterRegistry.for_file(derived_path)
    props_t = adapter_t.extract(template_path)
    props_d = adapter_d.extract(derived_path)
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
    args = ap.parse_args()

    with open(args.config, encoding="utf-8") as f:
        cfg = json.load(f)
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


if __name__ == "__main__":
    main()
