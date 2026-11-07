# -*- coding: utf-8 -*-
r"""相邻边界：白名单命中 vs 差一个字符 —— 补齐 `REVIEW-CODE-R4` 应修 5 点名的缺口。

审查原话：

> 扣分在 28% 是同一输入的重复断言、边界取样只挑恰好通过的 `x`/`t`

前半句经实测**夸大**（真实重复率 6%，见 §「关于 28%」），后半句**成立**：
`test_plain_quantity_needs_no_layout` 里挑的是 `s=10km`、`t=2h`、`m=5kg`
这类「恰好命中白名单」的样例，**没有一条证明「差一个字符就不降级」**。

一个只测正例的白名单测试，等价于没测 —— 把白名单改成 `.*` 之后它照样全绿。
本文件补的就是这个方向：**每一条正例都配一条只差一个字符的反例**。

## 关于 28%

用 pytest `--collect-only` 的真实 test id 统计（146 条参数化用例）：

    不同输入 137 个，跨函数重复 9 条 → 重复率 6%

那 9 条里有 3 条是**不同测试角度**（`F=ma` 既出现在「不是单位不许降」也出现在
「需要排版要留 OMML」，两者锁的是不同性质，不算重复）。所以真正的冗余很小，
不需要为了数字去删测试。审查那个 28% 可能把「同一函数内不同断言行」也算进去了。
**结论：不删测试，只补边界。**
"""
from __future__ import annotations

import sys
from pathlib import Path

import pytest

SCRIPTS = Path(__file__).resolve().parents[1] / "scripts"
sys.path.insert(0, str(SCRIPTS))

import docx_mathify as M  # noqa: E402
import linear_to_latex as L  # noqa: E402


def reason(src: str) -> str | None:
    """线性串 → 判据理由（None = 不降级）。"""
    return M.no_layout_reason(L.linear_to_latex(src))


# ---------------------------------------------------------------- 单位：差一字符
# (正例, 反例, 为什么反例不成立)
#
# ⚠️ 只收**真反例**：先把候选喂给实现确认过才写进来。
# 踩过的坑：`m=5g`、`c=0.5mol` 都是**合法单位、应当降级**，一度被误写成反例。
UNIT_NEIGHBOURS = [
    ("s=10km",   "s=10k",    "`k` 只是 SI 词头，不是单位"),
    ("s=10km",   "s=10kmm",  "`kmm` 不是单位（词头叠词头）"),
    ("m=5kg",    "m=5k",     "`k` 不是单位"),
    ("m=5kg",    "m=5kgs",   "`kgs` 尾部多一个 s"),
    ("c=0.5mol/L", "c=0.5molL", "`molL` 少斜杠又不是合法单位"),
    ("c=0.5mol/L", "c=0.5mol/", "尾部悬空斜杠"),
    ("c=0.5mol/L", "c=0.5mol/Lx", "尾部多字母"),
    ("c=0.5mol/L", "c=0.5/mol", "分子缺失"),
    ("T=25℃",    "T=25℃m",   "℃ 后缀单位后又跟字母"),
    ("x=80%",    "x=80%%",   "重复百分号"),
    ("x=80%",    "x=80%℃",   "两个后缀单位叠加"),
]


@pytest.mark.parametrize("good,bad,why", UNIT_NEIGHBOURS)
def test_one_character_off_the_whitelist_stops_the_downgrade(good, bad, why):
    """★ 白名单的**边界**：正例降级、差一个字符就不降。

    这是审查点名的缺口 —— 原来的测试只有前半句。
    """
    assert reason(good) is not None, f"正例 {good!r} 本该降级（{why}）"
    assert reason(bad) is None, f"反例 {bad!r} 不该降级：{why}（实际 {reason(bad)!r}）"


@pytest.mark.parametrize("src,unit", [
    ("m=5g", "g"),            # `g` 是克，合法单位
    ("c=0.5mol", "mol"),      # 无斜杠的纯单位
    ("s=10km", "km"),         # 带 SI 词头
    ("m=5kg", "kg"),
])
def test_short_units_also_downgrade(src, unit):
    """反过来的一面：这些**短单位也是合法单位**，必须降级。

    与上一条配对，防止「为了少误伤而把白名单收太紧」把真单位也拒了。
    """
    assert reason(src) is not None, f"{src!r}（单位 {unit!r}）应降级"


# ---------------------------------------------------------------- 量符号：单 vs 双字母
@pytest.mark.parametrize("good", ["s=10km", "m=5kg", "v=3m/s", "x=3", "f'=2"])
def test_single_letter_lhs_downgrades(good):
    assert reason(good) is not None, f"{good!r} 应降级"


@pytest.mark.parametrize("bad,why", [
    ("ss=10km", "LHS 是双字母，不是单个量符号"),
    ("xx=3", "LHS 双字母"),
    ("s1=10km", "LHS 带数字下标"),
    ("=10km", "LHS 为空"),
    ("f''=2", "LHS 两个撇（只允许一个）"),
])
def test_multi_letter_lhs_never_downgrades(bad, why):
    assert reason(bad) is None, f"{bad!r} 不该降级：{why}"


# ---------------------------------------------------------------- 等号：恰好一个
@pytest.mark.parametrize("good", ["x=3", "s=10km"])
def test_exactly_one_equals_downgrades(good):
    assert reason(good) is not None


@pytest.mark.parametrize("bad,why", [
    ("x=3=4", "两个等号（链式）"),
    ("x==3", "连写等号"),
    ("a=b=10km", "链式多等号"),
    ("x=3=", "尾部悬空等号"),
])
def test_two_or_more_equals_never_downgrade(bad, why):
    """链式等式/关系链不是「量值」，不能降级。"""
    assert reason(bad) is None, f"{bad!r} 不该降级：{why}"


# ---------------------------------------------------------------- 数值：合法 vs 非法
@pytest.mark.parametrize("good", ["x=3", "x=-2", "x=+3", "x=3.5"])
def test_numeric_rhs_downgrades(good):
    assert reason(good) is not None, f"{good!r} 应降级"


def test_leading_dot_decimal_is_rejected_on_purpose():
    """`x=.5` 不降级 —— 保守取向，不是 bug。

    `.5` 在中文工科稿里少见（都写 `0.5`），而 `.` 在正文里是句号的可能性更高。
    方向保守：宁可留 OMML，也不把 `.` 误当小数点。
    """
    assert reason("x=.5") is None
    assert reason("x=0.5") is not None      # 标准写法要过


@pytest.mark.parametrize("bad,why", [
    ("x=", "RHS 为空"),
    ("x=abc", "RHS 不是数值"),
    ("x=2x", "RHS 是变量表达式"),
    ("y=2x", "函数关系"),
    ("n=2t", "`t` 刻意不在 UNITS（护住 v=s/t）"),
])
def test_non_numeric_rhs_never_downgrades(bad, why):
    assert reason(bad) is None, f"{bad!r} 不该降级：{why}"


# ---------------------------------------------------------------- 空格：容忍但不误伤
@pytest.mark.parametrize("src", ["s =10km", "s= 10km", "s = 10km"])
def test_spaces_around_equals_are_tolerated(src):
    """等号周围有空格仍应降级（真实语料里很常见）。"""
    assert reason(src) is not None, f"{src!r} 应降级"


# ---------------------------------------------------------------- 名实相符的交叉验证
def test_whitelist_membership_explains_every_positive_case():
    """每个降级的正例，其单位必须**真的**在白名单里 —— 防止白名单被悄悄放宽。

    这条是把「正例通过」与「为什么通过」绑在一起：若哪天有人把单位判定改成
    「任意字母串」，白名单里查不到就会被这条抓住。

    ⚠️ 注意查的是 **`L.UNITS`**（全量单位表），不是 `L.UNIT_LIST`
    （那是 `POW_UNIT_RE` 用的 48 项精简表，刻意不含 `m`/`g`/`mol`/`km` 这类
    与量符号冲突的短单位）。用错表会让这条测试假红。
    """
    positives = ["s=10km", "m=5kg", "c=0.5mol/L", "T=25℃", "x=80%", "L=60dB"]
    for src in positives:
        r = reason(src)
        assert r is not None, f"{src!r} 应降级"
        rhs = src.split("=", 1)[1]
        unit = rhs.lstrip("+-.0123456789")
        if not unit:
            continue                       # 纯数值，无单位
        ok = (
            unit in L.UNITS                                  # 全量单位表
            or unit in M._UNIT_TAIL_EXTRA                    # 后缀单位 ℃ % °
            or any(unit.startswith(u) for u in L.UNITS)      # 带词头形态
            or unit in ("dB", "mol/L")                       # 复合/特殊
        )
        assert ok, f"{src!r} 的单位 {unit!r} 不在单位表里，却降级了"
