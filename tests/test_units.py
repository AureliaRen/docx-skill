# -*- coding: utf-8 -*-
"""单位斜杠（`/` 读作「每」）回归套件。

背景：`ρ=1.2kg/m³` 曾被写成 `\\frac{1.2kg}{m^{3}}` —— 物理语义错误。
`/` 在量纲串里是「每」，只有在量符号之间才是分数线。

本文件锁三件事：

  A. 量纲串一律不产 `\\frac`，且幂次落成真 OMML 上标（`m^{3}`）
  B. 真分式照旧转 —— `v=s/t`、`p=F/A`、`ρ=m/V`、`E=V/m` 这些
     「两边都是合法单位符号」的真分式是**最尖锐的反向用例**，一条都不许被吃掉
  C. `is_dimension_string()` 判定函数的边界

OMML 断言需要 `latex2mathml` + 本机 Office 的 `MML2OMML.XSL`；
缺任一项时整组 skip（与 `tests/test_formula.py` 同口径），CI 上仍跑纯字符串部分。

Run:  python -m pytest tests/test_units.py -v
"""
from __future__ import annotations

import sys
from pathlib import Path

import pytest

SCRIPTS = Path(__file__).resolve().parents[1] / "scripts"
sys.path.insert(0, str(SCRIPTS))

import linear_to_latex as L  # noqa: E402
import math_omml  # noqa: E402

# ---------------------------------------------------------------------------
# A. 量纲串：`/` = 每，不许出现 \frac
#    顺序对应 task-3 的覆盖清单：kg/m³ g/cm³ m/s m/s² mol/L g/mol kJ/mol
#    J/(kg·K) N·m A·h W/m² kPa mA/cm² Ω·cm km/h L/min r/min Ω·m V/m C/kg mol/(L·s)
# ---------------------------------------------------------------------------
UNIT_CASES = [
    # 密度 / 体积质量
    ("rho_kg_m3_literal",   "ρ=1.2kg/m³",          r"\rho =1.2\,\mathrm{kg/m^{3}}"),
    ("rho_kg_m3_script",    "ρ=1.2kg/m^{3}",       r"\rho =1.2\,\mathrm{kg/m^{3}}"),
    ("rho_kg_m3_spaced",    "ρ = 1.2 kg/m³",       r"\rho = 1.2\,\mathrm{kg/m^{3}}"),
    ("rho_g_cm3",           "ρ=1.2g/cm³",          r"\rho =1.2\,\mathrm{g/cm^{3}}"),
    ("kg_per_m_no_exp",     "λ=2kg/m",             r"\lambda =2\,\text{kg/m}"),
    # 速度 / 加速度
    ("speed_m_s",           "v=9.8m/s",            r"v=9.8\,\text{m/s}"),
    ("speed_m_s2_literal",  "v=9.8m/s²",           r"v=9.8\,\mathrm{m/s^{2}}"),
    ("speed_m_s2_script",   "v=9.8m/s^{2}",        r"v=9.8\,\mathrm{m/s^{2}}"),
    ("speed_m_s2_caret",    "v=9.8m/s^2",          r"v=9.8\,\mathrm{m/s^{2}}"),
    ("speed_km_h",          "v=60km/h",            r"v=60\,\text{km/h}"),
    ("speed_from_pow10",    "a = 3.0×10^{8} m/s²", r"a = 3.0\times 10^{8}\,\mathrm{m/s^{2}}"),
    # 物质的量浓度 / 摩尔质量 / 反应热
    ("conc_mol_L",          "c=0.5mol/L",          r"c=0.5\,\text{mol/L}"),
    ("molar_mass_g_mol",    "M=18g/mol",           r"M=18\,\text{g/mol}"),
    ("enthalpy_kJ_mol",     "ΔH=−285kJ/mol",       r"\Delta H=-285\,\text{kJ/mol}"),
    ("rate_mol_L_s",        "r=3mol/(L·s)",        r"r=3\,\text{mol/(L·s)}"),
    # 比热容 / 传热
    ("heat_cap_J_kgK",      "c=4180J/(kg·K)",      r"c=4180\,\text{J/(kg·K)}"),
    ("conduct_W_mK",        "λ=0.6W/(m·K)",        r"\lambda =0.6\,\text{W/(m·K)}"),
    ("flux_W_m2",           "q=1200W/m²",          r"q=1200\,\mathrm{W/m^{2}}"),
    # 力矩 / 电荷量
    ("torque_N_m",          "w=5N·m",              r"w=5\,\text{N·m}"),
    ("charge_A_h",          "Q=2A·h",              r"Q=2\,\text{A·h}"),
    # 压力（不带斜杠，走既有单位白名单）
    ("pressure_kPa",        "p=101kPa",            r"p=101\,\text{kPa}"),
    # 电流密度 / 电阻率
    ("current_density",     "j=5mA/cm²",           r"j=5\,\mathrm{mA/cm^{2}}"),
    ("current_density_mu",  "μ=1.2mA/cm²",         r"\mu =1.2\,\mathrm{mA/cm^{2}}"),
    ("resistivity_Ohm_cm",  "ρ=1.7Ω·cm",           r"\rho =1.7\,\text{Ω·cm}"),
    ("resistivity_Ohm_m",   "ρ=1.7Ω·m",            r"\rho =1.7\,\text{Ω·m}"),
    # 场强 / 剂量
    ("field_V_m",           "E=100V/m",            r"E=100\,\text{V/m}"),
    ("dose_C_kg",           "D=1.5C/kg",           r"D=1.5\,\text{C/kg}"),
    # 转速
    ("rev_per_min",         "n=3000r/min",         r"n=3000\,\text{r/min}"),
    ("volume_rate_L_min",   "Q=2.5L/min",          r"Q=2.5\,\text{L/min}"),
    # 不带斜杠的量纲串（既有能力不回归）
    ("power_kW_no_slash",   "P=2.5kW",             r"P=2.5\,\text{kW}"),
    ("energy_kWh_no_slash", "Q=5kWh",              r"Q=5\,\text{kWh}"),
    ("current_mA_no_slash", "I=2mA",               r"I=2\,\text{mA}"),
]

# ---------------------------------------------------------------------------
# B. 真分式：必须与改造前逐字一致（反向用例比正向更重要）
# ---------------------------------------------------------------------------
FRACTION_CASES = [
    # 两侧根本不是单位
    ("letters_a_b",         "a/b",        r"\frac{a}{b}"),
    ("numbers_1_2",         "1/2",        r"\frac{1}{2}"),
    ("mixed_10_10k",        "10/10k",     r"\frac{10}{10k}"),
    ("parens_6_2p2",        "6/(2+2)",    r"\frac{6}{2+2}"),
    # t / R / M / d / n 刻意不在单位词表里，所以这些真分式毫发无伤
    ("ohm_law_U_R",         "I=U/R",      r"I=\frac{U}{R}"),
    ("speed_s_over_t",      "v=s/t",      r"v=\frac{s}{t}"),
    ("power_W_over_t",      "P=W/t",      r"P=\frac{W}{t}"),
    ("molar_n_m_M",         "n=m/M",      r"n=\frac{m}{M}"),
    ("field_E_U_d",         "E=U/d",      r"E=\frac{U}{d}"),
    ("conc_c_n_V",          "c=n/V",      r"c=\frac{n}{V}"),
    # ★ 最尖锐：两边都是合法单位符号，但前面是 '='（量符号语境）→ 必须是分式
    ("pressure_F_over_A",   "p=F/A",      r"p=\frac{F}{A}"),
    ("density_m_over_V",    "ρ=m/V",      r"\rho =\frac{m}{V}"),
    ("field_E_V_over_m",    "E=V/m",      r"E=\frac{V}{m}"),
    ("reactance_omega_L",   "X=ωL/R",     r"X=\frac{\omega L}{R}"),
    # 复合结构
    ("two_fractions",       "U_{A} = (6/2 + 3)/(1/2 + 1/2)",
     r"U_{A} = \frac{\frac{6}{2} + 3}{\frac{1}{2} + \frac{1}{2}}"),
    ("nested_operands",     "I = U_{oc}/(R_{0} + R_{L})",
     r"I = \frac{U_{oc}}{R_{0} + R_{L}}"),
    ("limit_sinx_x",        "lim_{x→0} (sinx)/x = 1",
     r"\lim_{x\to 0} \frac{\sin x}{x} = 1"),
    ("integral_dx",         "∫dx/(1 + x^{2}) = arctanx + C",
     r"\int \frac{dx}{1 + x^{2}} = \arctan x + C"),
    # 无斜杠 → 行为不变
    ("newton_second_law",   "F=ma",       r"F=ma"),
    ("ohm_no_frac",         "U=IR",       r"U=IR"),
    ("kcl",                 "ΣI=0",       r"\sum I=0"),
    ("frequency",           "f=50Hz",     r"f=50\,\text{Hz}"),
    ("matrix_name_A",       "A^{T}A = E", r"A^{T}A = E"),
    ("power_value",         "x^{2} = 4",  r"x^{2} = 4"),
    # 单位白名单的老形态不许被新逻辑改写
    ("energy_J_s",          "h = 6.63×10^{−34} J·s",
     r"h = 6.63\times 10^{-34}\,\text{J·s}"),
    ("gas_constant",        "R = 8.31 J/(mol·K)",
     r"R = 8.31\,\text{J/(mol·K)}"),
]

# ---------------------------------------------------------------------------
# C. 判定函数边界
# ---------------------------------------------------------------------------
PREDICATE_CASES = [
    ("kg_per_m3",    "kg/m³",   True),
    ("mol_per_L",    "mol/L",   True),
    ("m_per_s",      "m/s",     True),
    ("N_dot_m",      "N·m",     True),
    ("A_dot_h",      "A·h",     True),
    ("V_per_m",      "V/m",     True),
    ("s_over_t",     "v=s/t",   False),
    ("F_over_A",     "F/A",     False),
    ("m_over_V",     "m/V",     False),
    ("one_over_two", "1/2",     False),
    ("a_over_b",     "a/b",     False),
    ("E_eq_V_per_m", "E=V/m",   False),
    ("ten_over_10k", "10/10k",  False),
]


# OMML 可用性判定收在 `tests/_omml.py` 一处（REVIEW-CODE-R4 应修 5）。
from _omml import needs_omml  # noqa: E402

_M = math_omml.M_NS


def _frac_types(omml) -> list:
    """OMML 里所有分式的类型；`m:type` 缺省即 bar（堆叠分数线）。"""
    out = []
    for f in omml.findall(".//{%s}f" % _M):
        e = f.find("{%s}fPr/{%s}type" % (_M, _M))
        out.append((e.get("{%s}val" % _M) if e is not None else None) or "bar")
    return out


def _omml_text(omml) -> str:
    return "".join(t.text or "" for t in omml.iter("{%s}t" % _M))


def _flat(s: str) -> str:
    """去掉斜杠与空白：`lin` 型分式不把 `/` 落成 m:t，比对时忽略它。"""
    for ch in ("/", " ", "\u00a0", "\t"):
        s = s.replace(ch, "")
    return s


# ---------------------------------------------------------------------------
# A. 量纲串
# ---------------------------------------------------------------------------
@pytest.mark.parametrize("linear,expected", [c[1:] for c in UNIT_CASES],
                         ids=[c[0] for c in UNIT_CASES])
def test_unit_slash_is_per_not_fraction(linear, expected):
    assert L.linear_to_latex(linear) == expected


@pytest.mark.parametrize("linear,_expected", [c[1:] for c in UNIT_CASES],
                         ids=[c[0] for c in UNIT_CASES])
def test_unit_cases_never_produce_frac(linear, _expected):
    out = L.linear_to_latex(linear)
    assert "\\frac" not in out
    if "/" in linear:
        # 斜杠保留成线性外观，没有被换成分数线（幂次另走 \mathrm，无幂次走 \text）
        assert "/" in out
        assert r"\mathrm{" in out or r"\text{" in out


# ---------------------------------------------------------------------------
# B. 真分式不许被吃掉
# ---------------------------------------------------------------------------
@pytest.mark.parametrize("linear,expected", [c[1:] for c in FRACTION_CASES],
                         ids=[c[0] for c in FRACTION_CASES])
def test_real_fractions_are_untouched(linear, expected):
    assert L.linear_to_latex(linear) == expected


@pytest.mark.parametrize("linear", [
    "v=s/t", "p=F/A", "ρ=m/V", "E=V/m", "c=n/V", "I=U/R",
])
def test_two_unit_lookalikes_stay_fractions(linear):
    """★ 反向要点：两边都是合法单位符号，也不能被读成量纲串。"""
    assert "\\frac" in L.linear_to_latex(linear)


SLASH_CASES = [c for c in UNIT_CASES if "/" in c[1]]


@pytest.mark.parametrize("unit_case", [c[1] for c in SLASH_CASES],
                         ids=[c[0] for c in SLASH_CASES])
def test_unit_strings_are_not_fraction_only_because_of_the_slash(unit_case):
    """量纲串的 `/` 两侧都必须是单位 —— 换成量符号就该退回分式。"""
    broken = unit_case.split("=", 1)[-1]
    assert "\\frac" not in L.linear_to_latex("x=" + broken)


# ---------------------------------------------------------------------------
# C. 判定函数 / 词表
# ---------------------------------------------------------------------------
@pytest.mark.parametrize("text,expected", [c[1:] for c in PREDICATE_CASES],
                         ids=[c[0] for c in PREDICATE_CASES])
def test_is_dimension_string(text, expected):
    assert L.is_dimension_string(text) is expected


def test_unit_slash_chains_reports_raw_fragments():
    assert L.unit_slash_chains("ρ=1.2kg/m³") == ["kg/m³"]
    assert L.unit_slash_chains("v=s/t") == []
    assert L.unit_slash_chains("c=4180J/(kg·K)") == ["J/(kg·K)"]
    assert L.unit_slash_chains("q=1200W/m²；R=8.31J/(mol·K)") == \
        ["W/m²", "J/(mol·K)"]
    assert L.unit_slash_chains("p=101kPa") == []


def test_unit_vocabulary_contains_the_required_symbols():
    """task-3 点名要覆盖的单位符号，必须真在词表里。"""
    required = ["kg", "g", "cm", "m", "s", "mol", "L", "J", "K", "W", "V",
                "A", "N", "C", "Ω", "kPa", "Pa", "mA", "kJ", "min", "h", "r"]
    missing = [u for u in required if u not in L.UNITS]
    assert not missing, "missing units: %s" % missing


def test_ambiguous_variable_letters_are_deliberately_excluded():
    """t/d/c/a/M/U/R 这些当量符号的概率高于当单位，刻意不入表 —— 这是 v=s/t 的护栏。"""
    for sym in ["t", "d", "c", "a", "M", "U", "Q", "R", "X"]:
        assert sym not in L.UNITS
    assert "min" in L.UNITS          # 分钟是整词登记，不是 m+in 拼出来的


def test_compound_whitelist_covers_bare_unit_strings():
    for c in ["m/s", "km/h", "mol/L", "g/mol", "kJ/mol", "J/(kg·K)",
              "W/m²".replace("²", "^2"), "mA/cm^2", "Ω·m", "V/m", "C/kg",
              "N·m", "A·h", "L/min", "r/min", "kg/m^3", "mol/(L·s)"]:
        assert c in L.COMPOUND_UNITS, c


# ---------------------------------------------------------------------------
# D. OMML 结构：幂次要落成真上标，且不能是堆叠分数线
# ---------------------------------------------------------------------------
@needs_omml
@pytest.mark.parametrize("linear", [c[1] for c in UNIT_CASES],
                         ids=[c[0] for c in UNIT_CASES])
def test_unit_omml_never_uses_a_stacked_bar_fraction(linear):
    omml = math_omml.latex_to_omml_lxml(L.linear_to_latex(linear))
    assert "bar" not in _frac_types(omml)


@needs_omml
@pytest.mark.parametrize("linear,needle", [
    ("ρ=1.2kg/m³", "kg/m"),
    ("v=9.8m/s²", "m/s"),
    ("c=4180J/(kg·K)", "J/(kg·K)"),
    ("r=3mol/(L·s)", "mol/(L·s)"),
    ("v=60km/h", "km/h"),
])
def test_unit_omml_keeps_the_linear_look(linear, needle):
    omml = math_omml.latex_to_omml_lxml(L.linear_to_latex(linear))
    assert _flat(needle) in _flat(_omml_text(omml))


@needs_omml
@pytest.mark.parametrize("linear,exp", [
    ("ρ=1.2kg/m³", "3"),
    ("v=9.8m/s²", "2"),
    ("j=5mA/cm²", "2"),
    ("q=1200W/m²", "2"),
])
def test_unit_exponent_becomes_a_real_omml_superscript(linear, exp):
    omml = math_omml.latex_to_omml_lxml(L.linear_to_latex(linear))
    sups = omml.findall(".//{%s}sup" % _M)
    assert len(sups) == 1, "expected exactly one m:sup"
    assert _omml_text(sups[0]) == exp


@needs_omml
def test_a_real_fraction_still_becomes_a_stacked_bar_fraction():
    """反证：bar 型 m:f 的检查不是空转 —— `a/b` 必须产出它。"""
    omml = math_omml.latex_to_omml_lxml(L.linear_to_latex("a/b"))
    assert _frac_types(omml) == ["bar"]


@needs_omml
def test_field_strength_lookalike_is_still_a_stacked_fraction():
    omml = math_omml.latex_to_omml_lxml(L.linear_to_latex("E=V/m"))
    assert _frac_types(omml) == ["bar"]
