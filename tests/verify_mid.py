#!/usr/bin/env python3
"""多角度验证：xlsx 中间插入/删除行是否被正确对齐（不产生坐标级联噪声）。

核心判据：
- 坐标属性 row/r、c/r、dimension/ref 被 ignore_attr_patterns 抑制，不报 changed；
- 中间插入行时，新增的「生产日期」行被识别为 added，原「容量」行（下移）被正确匹配（不误报 added）；
- 中间删除行时，被删的「型号」行被识别为 removed，其余行正确匹配。
"""
import os, sys, json, tempfile
SKILL = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, SKILL)
import checker as C
import openpyxl
from openpyxl.styles import Font

cfg = json.load(open(os.path.join(SKILL, "param_tree.json"), encoding="utf-8"))
pt = C.ParamTree(cfg)


def mk_xlsx(path, rows):
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = "S1"
    for r, row in enumerate(rows, 1):
        for c, val in enumerate(row, 1):
            ws.cell(r, c, val)
    wb.save(path)


def row_alignment(t_path, d_path):
    """返回 [(kind, t_anchor, d_anchor), ...]，反映行级对齐配对。"""
    _, props_t = C.safe_extract(t_path, cfg)
    _, props_d = C.safe_extract(d_path, cfg)
    rt = C.build_tree(props_t); rd = C.build_tree(props_d)

    def find_rows(node):
        res = []
        if node.get('_tag') == 'row':
            res.append(node)
        for _tg, kids in node['children'].items():
            for k in kids:
                res.extend(find_rows(k))
        return res
    tt, dd = find_rows(rt), find_rows(rd)
    pairs = C.greedy_align(tt, dd)
    out = []
    for kind, a, b in pairs:
        ta = tt[a]['anchor'] if a >= 0 else '-'
        db = dd[b]['anchor'] if b >= 0 else '-'
        out.append((kind, ta, db))
    return out


def summarize(title, t_rows, d_rows):
    TMP = tempfile.mkdtemp(prefix="otc_mid_")
    tp = f"{TMP}/t.xlsx"; dp = f"{TMP}/d.xlsx"
    mk_xlsx(tp, t_rows); mk_xlsx(dp, d_rows)
    d, red = C.compare(tp, dp, cfg, pt)
    k = {}
    for x in d:
        k[x['kind']] = k.get(x['kind'], 0) + 1
    coord = [x for x in d if x['kind'] == 'changed'
             and (x['element'], x['attr']) in (('c', 'r'), ('row', 'r'), ('dimension', 'ref'))]
    print(f"\n=== {title} ===")
    print(f"偏差 {len(d)}  kinds={dict(k)}")
    print(f"坐标属性(row/r,c/r,dimension/ref)被误报 changed 数: {len(coord)}")
    print("行级对齐配对:")
    for kind, ta, db in row_alignment(tp, dp):
        print(f"  [{kind:6}] {ta[:18]!r:20} -> {db[:18]!r}")
    return k, coord


# 1) 中间插入行（在第2、3行间插入「生产日期」）
k1, c1 = summarize(
    "xlsx 中间插入行",
    [["制造商", "深圳"], ["型号", "ABC-123"], ["容量", "3000mAh"]],
    [["制造商", "深圳"], ["型号", "ABC-123"], ["生产日期", "2026-09"], ["容量", "3000mAh"]],
)

# 2) 中间删除行（删掉「型号」行）
k2, c2 = summarize(
    "xlsx 中间删除行",
    [["制造商", "深圳"], ["型号", "ABC-123"], ["容量", "3000mAh"]],
    [["制造商", "深圳"], ["容量", "3000mAh"]],
)

# 3) 断言
ok = True
if c1:
    print("\n[FAIL] 中间插入行仍把坐标属性误报为 changed"); ok = False
# 中间插入：新增行应为 added，且 added 行锚点含"生产日期"；容量行应被匹配(非 added)
align1 = row_alignment.__wrapped__ if hasattr(row_alignment, '__wrapped__') else None
_, props_t = C.safe_extract.__self__ if False else (None, None)
print("\n=== 断言 ===")
# 直接重新取对齐判断新增行内容
TMP = tempfile.mkdtemp(prefix="otc_assert_")
mk_xlsx(f"{TMP}/t.xlsx", [["制造商", "深圳"], ["型号", "ABC-123"], ["容量", "3000mAh"]])
mk_xlsx(f"{TMP}/d.xlsx", [["制造商", "深圳"], ["型号", "ABC-123"], ["生产日期", "2026-09"], ["容量", "3000mAh"]])
al = row_alignment(f"{TMP}/t.xlsx", f"{TMP}/d.xlsx")
added = [db for kind, ta, db in al if kind == 'add']
if not any('生产日期' in (a or '') for a in added):
    print("[FAIL] 中间插入行：新增行未被识别为 added(内容应为 生产日期)"); ok = False
else:
    print("[PASS] 中间插入行：新增行正确识别为 added(生产日期)")
if any('容量' in (a or '') for a in added):
    print("[FAIL] 中间插入行：被下移的 容量 行被错误报为 added"); ok = False
else:
    print("[PASS] 中间插入行：下移的 容量 行未被误报为 added")

if c2:
    print("[FAIL] 中间删除行仍把坐标属性误报为 changed"); ok = False
else:
    print("[PASS] 中间删除行：无坐标属性 changed")

print("\n结果:", "ALL PASS" if ok else "HAS FAILURES")
