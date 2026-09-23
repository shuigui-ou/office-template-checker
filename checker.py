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


# ---------- 相对位置对齐（核心：按结构相对位置比对，而非绝对索引） ----------
# 为什么不用绝对索引：绝对索引在「模板某表插入一行/列」后会整体平移，
# 导致后续所有行被误报为「新增/缺失」。相对位置 = 在同一父节点下，按子树特征
# （shape + 内容锚点）做贪心一步前瞻/等数顺序配对找到的对应关系，插入的行列被
# 识别为新增/缺失，其余项仍按对齐后的相对序号一一对应比较。这才是「相对位置
# 是否有相同的格式」。
_SEG = re.compile(r'([^/\[]+)(?:\[(\d+)\])?')


def parse_segments(path):
    return [(m.group(1), int(m.group(2)) if m.group(2) else None)
            for m in _SEG.finditer(path)]


_TAG_LABEL = {'w:body': '正文', 'w:tbl': '表', 'w:tr': '行', 'w:tc': '列',
              'w:p': '段', 'w:r': '字', 'a:tbl': '表', 'a:tr': '行', 'a:tc': '列',
              'a:p': '段', 'x:sheetData': '表', 'x:row': '行', 'x:c': '列'}


def _label(tag, rank):
    return '%s%d' % (_TAG_LABEL.get(tag, tag), rank)


def _skip_label(tag):
    # 部件名（word/document.xml 等）与 .rels 不作为定位标签，避免定位串冗余
    return '/' in tag or tag.endswith('.xml') or tag.endswith('.rels')


def build_tree(props):
    """把 Prop 流建成嵌套树：每个节点的 attrs=该元素自身的格式属性，
    children=按 (tag, 绝对索引) 排序的子节点；并自底向上计算局部内容锚点。"""
    root = {'attrs': [], 'children': {}, '_idx': 0, '_tag': 'root'}
    for p in props:
        segs = parse_segments(p.path)
        node = root
        for (tag, idx) in segs:
            node['children'].setdefault(tag, [])
            child = None
            for c in node['children'][tag]:
                if c['_idx'] == idx:
                    child = c
                    break
            if child is None:
                child = {'attrs': [], 'children': {}, '_idx': idx, '_tag': tag}
                node['children'][tag].append(child)
            node = child
        node['attrs'].append(p)
    _calc_anchor(root, False)
    return root


def _calc_anchor(node, in_tbl=False):
    """自底向上计算「局部内容锚点」：仅**表格内**的 w:p / w:tc / w:tr 携带文本锚点；
    正文段落（非表格内）锚点强制清空，按 shape/顺序对齐（纯字号改动等仍能命中 changed）。

    关键纪律：锚点只停在行/列这一层、**绝不上溢到文档/表层级**——
    否则插入一行会让整篇或整表签名变化，再次级联误报。"""
    for tag, kids in node['children'].items():
        child_in_tbl = in_tbl or (tag == 'tbl')
        for c in kids:
            _calc_anchor(c, child_in_tbl)
    tag = node.get('_tag')
    if tag == 'p':
        # 仅表格内的段落用文本做锚点；正文段落不锚定文本，避免文字改动破坏对齐
        node['anchor'] = (''.join(p.value for p in node['attrs'] if p.attr == '#ctext')
                          if in_tbl else '')
    elif tag in ('tc', 'tr'):
        node['anchor'] = ''.join(c.get('anchor', '')
                               for _, kids in node['children'].items()
                               for c in kids)
    else:
        node['anchor'] = ''


def _collect(node, out):
    out.extend(node['attrs'])
    for tag, kids in node['children'].items():
        for c in kids:
            _collect(c, out)


def shape_of(node):
    """子树格式属性 shape：只看「哪些 (element, attr) 存在」，与取值、行数、内容文本都无关。
    用于相对位置匹配——插入一行不会改变所在表/父节点的 shape。"""
    acc = []
    _collect(node, acc)
    return frozenset((p.element, p.attr) for p in acc if p.attr != "#ctext")


def node_sig(node):
    """子树签名（兼容旧接口）：shape 加可选内容锚点。"""
    s = shape_of(node)
    a = node.get("anchor")
    return s | frozenset([("#anchor", a)]) if a else s


def greedy_align(t_kids, d_kids):
    """相对位置对齐核心：按 (shape, 内容锚点) 做**贪心一步前瞻**匹配。

    相比 LCS，本算法对「内容编辑/对调」免疫、对「插入/删除行列」不级联，正是用户工作流所需。

    关键分派：
    - **同层节点数相等** → 不存在结构性增删，所有差异只能是「内容编辑/对调」(0 偏差)
      或「格式改动」(changed)。此时直接**按位置顺序配对**，不触发任何插入/删除逻辑。
      这是修复「单元格内容对调被误报为删除+插入」的核心：等数时顺序配对，
      内容不参与格式比较 → 0 偏差；而格式属性变动仍在逐属性比较中命中 changed。
    - **节点数不等** → 必含真实增删，走一步前瞻定位多出的那一个：
      若一步前瞻能解释对方为「插入」(派生多一个) 或「删除」(模板多一个)，
      优先吸收该增删，不牵连后续；其余按对齐后的相对序号一一对应。

    贪心保证：任何插入/删除最多错位一行/列即被前瞻吸收，**不会整片级联**。"""
    n, m = len(t_kids), len(d_kids)
    if n == m:
        # 等数 → 顺序配对（内容编辑/对调/格式改动都在此分支处理，绝不变出增删）
        return [("match", i, i) for i in range(n)]
    def sp(n):
        return (shape_of(n), n.get("anchor") or "")
    pairs = []
    i = j = 0
    while i < n and j < m:
        st, sa = sp(t_kids[i])
        sd, da = sp(d_kids[j])
        if st == sd and sa == da:
            pairs.append(("match", i, j)); i += 1; j += 1
        elif st == sd:
            # 同结构、仅内容(文本)不同 → 先尝试用一步前瞻解释对方为插入/删除，
            # 否则按「同一相对位置的纯内容编辑」匹配（内容不参与格式比较）。
            if j + 1 < m and sp(d_kids[j + 1]) == (st, sa):
                pairs.append(("add", -1, j)); j += 1        # 派生多一个 → 插入
            elif i + 1 < n and sp(t_kids[i + 1]) == (sd, da):
                pairs.append(("rem", i, -1)); i += 1        # 模板多一个 → 删除
            else:
                pairs.append(("match", i, j)); i += 1; j += 1
        elif j + 1 < m and sp(d_kids[j + 1]) == (st, sa):
            pairs.append(("add", -1, j)); j += 1            # 派生多一个 → 插入
        elif i + 1 < n and sp(t_kids[i + 1]) == (sd, da):
            pairs.append(("rem", i, -1)); i += 1            # 模板多一个 → 删除
        else:
            pairs.append(("rem", i, -1)); pairs.append(("add", -1, j)); i += 1; j += 1
    while i < n:
        pairs.append(("rem", i, -1)); i += 1
    while j < m:
        pairs.append(("add", -1, j)); j += 1
    return pairs


def _mk(part, loc, element, attr, mode, tv, dv, kind):
    return {"dimension": dim_of(element), "layer": layer_of(part),
            "part": part, "path": loc, "element": element, "attr": attr,
            "mode": mode, "template": tv, "derived": dv, "kind": kind}


def align_nodes(tn, dn, pt, theme_t, theme_d, rank_path, loc, deviations):
    """递归对齐两个（已匹配的）节点：先比本节点属性，再按 tag 分组对子节点做贪心一步前瞻对齐。"""
    t_map = {p.attr: p for p in tn['attrs']}
    d_map = {p.attr: p for p in dn['attrs']}
    for attr in set(t_map) | set(d_map):
        if attr == "#ctext":
            continue  # 内容锚点仅用于对齐，不参与格式比较
        tp = t_map.get(attr)
        dp = d_map.get(attr)
        if tp and dp:
            if pt._is_ignored_path(tp) or pt.is_noise(tp):
                continue
            mode, spec = pt.mode_for(tp)
            if mode == 'ignore':
                continue
            if not values_match(mode, spec, tp.value, dp.value, theme_t, theme_d):
                deviations.append(_mk(tp.part, loc, tp.element, attr, mode,
                                      tp.value, dp.value, 'changed'))
        elif tp is None:
            if pt._is_ignored_path(dp) or pt.is_noise(dp):
                continue
            if pt.mode_for(dp)[0] == 'ignore':
                continue
            deviations.append(_mk(dp.part, loc, dp.element, attr, pt.mode_for(dp)[0],
                                  None, dp.value, 'added'))
        else:
            if pt._is_ignored_path(tp) or pt.is_noise(tp):
                continue
            if pt.mode_for(tp)[0] == 'ignore':
                continue
            deviations.append(_mk(tp.part, loc, tp.element, attr, pt.mode_for(tp)[0],
                                  tp.value, None, 'removed'))

    for tag in set(tn['children']) | set(dn['children']):
        t_kids = tn['children'].get(tag, [])
        d_kids = dn['children'].get(tag, [])
        pairs = greedy_align(t_kids, d_kids)
        matched_t, matched_d = set(), set()
        rank_of_t, rank_of_d, r = {}, {}, 0
        for kind, a, b in pairs:
            if kind == 'match':
                matched_t.add(a); matched_d.add(b)
                rank_of_t[a] = rank_of_d[b] = r
                r += 1
        for kind, a, b in pairs:
            if kind == 'match':
                continue
            if kind == 'rem':
                rank_of_t[a] = r; r += 1
            else:
                rank_of_d[b] = r; r += 1
        for kind, a, b in pairs:
            if kind != 'match':
                continue
            lab = '' if _skip_label(tag) else _label(tag, rank_of_t[a])
            new_loc = (loc + lab) if (loc and lab) else (lab or loc)
            align_nodes(t_kids[a], d_kids[b], pt, theme_t, theme_d,
                        rank_path + [(tag, rank_of_t[a])], new_loc, deviations)
        for kind, a, b in pairs:
            if kind != 'rem':
                continue
            lab = '' if _skip_label(tag) else _label(tag, rank_of_t[a])
            sub_loc = (loc + lab) if (loc and lab) else (lab or loc)
            emit_subtree(t_kids[a], pt, sub_loc, 'removed', deviations)
        for kind, a, b in pairs:
            if kind != 'add':
                continue
            lab = '' if _skip_label(tag) else _label(tag, rank_of_d[b])
            sub_loc = (loc + lab) if (loc and lab) else (lab or loc)
            emit_subtree(d_kids[b], pt, sub_loc, 'added', deviations)


def emit_subtree(node, pt, loc, kind, deviations):
    """把未匹配子节点整棵子树作为 added/removed 偏差逐属性展开。"""
    for p in node['attrs']:
        if p.attr == "#ctext":
            continue  # 内容锚点仅用于对齐，不参与格式比较
        if pt._is_ignored_path(p) or pt.is_noise(p):
            continue
        if pt.mode_for(p)[0] == 'ignore':
            continue
        if kind == 'removed':
            deviations.append(_mk(p.part, loc, p.element, p.attr, pt.mode_for(p)[0],
                                  p.value, None, 'removed'))
        else:
            deviations.append(_mk(p.part, loc, p.element, p.attr, pt.mode_for(p)[0],
                                  None, p.value, 'added'))
    for tag, kids in node['children'].items():
        for ci, c in enumerate(kids):
            lab = '' if _skip_label(tag) else _label(tag, ci)
            sub_loc = (loc + lab) if (loc and lab) else (lab or loc)
            emit_subtree(c, pt, sub_loc, kind, deviations)


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
    """相对位置比对：先按结构树把模板/派生建成嵌套树，再逐层（正文/表/行/列）
    用贪心一步前瞻按 (结构, 内容锚点) 对齐，插入的行列被识别为 added/removed，
    其余按对齐后的相对序号一一对应比较。详见本文件顶部「相对位置对齐」段。"""
    adapter_t, props_t = safe_extract(template_path, cfg)
    adapter_d, props_d = safe_extract(derived_path, cfg)
    theme_t = adapter_t.theme_map(template_path)
    theme_d = adapter_d.theme_map(derived_path)
    root_t, root_d = build_tree(props_t), build_tree(props_d)

    deviations = []
    align_nodes(root_t, root_d, pt, theme_t, theme_d, [], '', deviations)

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
    lines.append("## 偏差类型语义（如何解读本报告）")
    lines.append("")
    lines.append("你**基于模板手动编辑**（改内容 / 格式 / 内容逻辑）后另存为新文件，因此两类差异性质不同：")
    lines.append("")
    lines.append("- **格式偏离 `changed`**：同一**相对位置**上的格式属性值不同（字号、颜色、对齐、边距等）。这是审查重点——往往是编辑时不小心改坏了模板格式，或应当保持一致却漏改。")
    lines.append("- **结构变动 `added`**：派生比模板多出的相对位置（如你插入了一行/列/段落）。属内容编辑的预期动作；本工具按相对位置对齐后**只把新增部分单列**，不会牵连后续行整片误报，请确认是否故意。")
    lines.append("- **结构变动 `removed`**：派生比模板缺失的位置（如你删了一行/列）。可能是合理精简，也可能误删，需你确认。")
    lines.append("")
    lines.append("> 内容文本（你填的字）只用作「相对位置」对齐锚点，**不参与格式比较**——所以你改内容不会凭空产生「格式偏离」；只有格式属性本身被改才会记为 `changed`。")
    lines.append("")
    lines.append("---")
    lines.append("> 核心保证（确定性，不调用任何大模型）：**按相对位置比对**——先按结构树逐层"
                 "（正文 / 表 / 行 / 单元格）用贪心一步前瞻按 (结构, 内容锚点) 对齐，再比较「同一相对位置上的同一属性值"
                 "是否相同」。模板某处插入一行/一列时，该新增行列被识别为 added/removed，"
                 "其余行列仍按对齐后的相对序号一一对应，不会因绝对索引平移而整片误报。"
                 "报告定位串（表N·行M·列K）中的序号即相对序号，非原始绝对索引。"
                 "报告仅含参数树未忽略的差异；渲染层/版本默认/隐式默认值不在范围内。")
    lines.append("")
    lines.append("## 引用安全声明（给大模型使用方 · 防止误引）")
    lines.append("")
    cb = (pt.cfg or {}).get("citation_boundary", {})
    lines.append(cb.get("scope",
        "**本工具是确定性引擎，不调用任何大模型；下列纪律用于约束引用方（agent），使其只能基于已锚定事实发言。**"))
    lines.append("")
    if cb.get("not_checked"):
        lines.append("**本工具未查 / 查不了（引用方不得越界断言）：**")
        for item in cb["not_checked"]:
            lines.append("- %s" % item)
        lines.append("")
    lines.append("**引用纪律（四不准）：**")
    discipline = cb.get("discipline") or [
        "引用须带定位(path 或 part:path 片段)与维度标签，且限定在本报告已列出的偏差条目内。",
        "不得对报告未列出的属性做一致/不一致断言；未列出 ≠ 一致。",
        "带 ❌ 的偏差引用时须加「需人工复核」，不得直接判定合格。",
        "不得编造报告中没有的页码、文件名或判定。",
    ]
    for item in discipline:
        lines.append("- %s" % item)
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
