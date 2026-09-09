import json
from checker import ParamTree, values_match
from adapter_interface import Prop

cfg = json.load(open("param_tree.json", encoding="utf-8"))
pt = ParamTree(cfg)

def t(name, got, exp):
    print(f"  [{'OK' if got==exp else 'FAIL'}] {name}: got={got} exp={exp}")

print("== 匹配模式 ==")
t("tolerance 24 vs 25 (<=1)", values_match("tolerance", {"tol": 1}, "24", "25", {}, {}), True)
t("tolerance 24 vs 26 (>1)", values_match("tolerance", {"tol": 1}, "24", "26", {}, {}), False)
t("rgb 大小写同色", values_match("rgb", None, "FF0000", "ff0000", {}, {}), True)
t("theme_slot 同槽", values_match("theme_slot", None, "accent1", "accent1", {}, {}), True)
t("theme_slot 异槽", values_match("theme_slot", None, "accent1", "accent2", {}, {}), False)
t("exact 同", values_match("exact", None, "x", "x", {}, {}), True)
t("exact 异", values_match("exact", None, "x", "y", {}, {}), False)

print("== 噪声过滤 ==")
t("rsidR 忽略", pt.is_noise(Prop("a", "p", "r", "rsidR", "x")), True)
t("docId 忽略", pt.is_noise(Prop("a", "p", "r", "docId", "x")), True)
t("sz/val 不过滤", pt.is_noise(Prop("a", "p", "sz", "val", "24")), False)

print("== 模式查表 ==")
m, _ = pt.mode_for(Prop("a", "p", "sz", "val", "24"))
t("sz/val -> tolerance", m, "tolerance")
m, _ = pt.mode_for(Prop("a", "p", "color", "val", "#FF0000"))
t("color/val -> rgb", m, "rgb")
m, _ = pt.mode_for(Prop("a", "p", "color", "themeColor", "accent1"))
t("color/themeColor -> theme_slot", m, "theme_slot")
m, _ = pt.mode_for(Prop("a", "p", "jc", "val", "left"))
t("jc/val 未配置 -> exact", m, "exact")
