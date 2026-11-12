# -*- coding: utf-8 -*-
"""``fix_zoom.fix_docx`` 的返回码，以及 ``docx_create`` 是否把它透传出去。

背景（2026-11-11 实测）：``docx_create`` 在 ``doc.save()`` 之后调 ``fix_docx()``
补 schema 强制的 ``w:percent``，但**丢弃返回值**。``fix_docx`` 对「不是可读 zip」
返回 ``2``，于是那条失败路径会以 ``{"ok": true}`` + exit 0 收场 —— 与「回显一个
从未生效的 ``--index``」属同一类缺陷：告诉调用方成功，而它要的东西没落地。

两条断言方向，缺一不可：

* **负向**：``fix_docx`` 对坏包必须返 ``2`` 且不写盘（不然透传的是 0）；
* **正向**：``docx_create`` 正常出包仍然是 ``0``（透传不能把成功也变成失败）。

这些用例**不碰** ``latex2mathml`` / ``MML2OMML.XSL`` / Office，所以没有
``needs_omml`` 标记 —— 在无 Office 的 CI runner 上是 pass，不是 skip。
"""
from __future__ import annotations

import json
import subprocess
import sys
import zipfile
from pathlib import Path

import pytest

SCRIPTS = Path(__file__).resolve().parent.parent / "scripts"
sys.path.insert(0, str(SCRIPTS))

from fix_zoom import fix_docx, fix_settings_xml  # noqa: E402


# --------------------------------------------------------------------------
# fix_docx —— 它自己能报的唯一失败：输入不是可读的 zip
# --------------------------------------------------------------------------
def test_fix_docx_returns_2_for_a_non_zip(tmp_path):
    target = tmp_path / "not-a-package.docx"
    target.write_bytes(b"this is not a zip at all")
    assert fix_docx(target) == 2


def test_fix_docx_returns_2_and_leaves_the_file_untouched(tmp_path):
    """拒绝必须是完整的：坏包原样留着，不能被截断成半成品。"""
    target = tmp_path / "broken.docx"
    payload = b"\x00\x01definitely not a zip"
    target.write_bytes(payload)
    before = target.read_bytes()

    assert fix_docx(target) == 2
    assert target.read_bytes() == before, "refused input must not be rewritten"


def test_fix_docx_returns_2_for_a_missing_path(tmp_path):
    assert fix_docx(tmp_path / "nope.docx") == 2


def test_fix_docx_returns_0_when_settings_has_no_zoom(tmp_path):
    """没有 ``word/settings.xml`` 时无事可做，仍是成功 —— 透传不能把它当失败。"""
    pkg = tmp_path / "no-settings.docx"
    with zipfile.ZipFile(pkg, "w") as z:
        z.writestr("word/document.xml", "<w:document/>")

    assert fix_docx(pkg) == 0


def test_fix_docx_patches_zoom_and_returns_0(tmp_path):
    pkg = tmp_path / "needs-percent.docx"
    settings = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        '<w:zoom w:val="bestFit"/>'
        "</w:settings>"
    )
    with zipfile.ZipFile(pkg, "w") as z:
        z.writestr("word/settings.xml", settings)

    assert fix_docx(pkg) == 0
    with zipfile.ZipFile(pkg) as z:
        patched = z.read("word/settings.xml").decode("utf-8")
    assert 'w:percent="100"' in patched


# --------------------------------------------------------------------------
# 纯函数层：已经合法的 zoom 不被二次改写（幂等）
# --------------------------------------------------------------------------
def test_fix_settings_xml_is_idempotent_on_a_valid_zoom():
    xml = '<w:settings><w:zoom w:val="bestFit" w:percent="100"/></w:settings>'
    assert fix_settings_xml(xml, 100)[1] == 0


def test_fix_settings_xml_repairs_a_bare_self_closing_zoom():
    """裸 ``<w:zoom/>`` 同样违反 schema（缺 ``w:percent``），必须被补上。"""
    xml = "<w:settings><w:zoom/></w:settings>"
    patched, touched = fix_settings_xml(xml, 100)
    assert touched == 1
    assert 'w:percent="100"' in patched


# --------------------------------------------------------------------------
# docx_create 端到端：成功路径必须仍是 0 + ok:true（透传不退化成误报）
# --------------------------------------------------------------------------
SPEC = {
    "blocks": [{"type": "paragraph", "runs": [{"text": "hello"}]}],
}


def _run_create(tmp_path, spec_path, out_path):
    return subprocess.run(
        [sys.executable, str(SCRIPTS / "docx_create.py"),
         str(spec_path), str(out_path)],
        capture_output=True, text=True,
    )


def test_docx_create_still_exits_0_and_reports_ok(tmp_path):
    spec = tmp_path / "spec.json"
    spec.write_text(json.dumps(SPEC), encoding="utf-8")
    out = tmp_path / "out.docx"

    proc = _run_create(tmp_path, spec, out)

    assert proc.returncode == 0, proc.stderr
    payload = json.loads(proc.stdout)
    assert payload["ok"] is True
    assert out.is_file()


def test_docx_create_output_carries_a_valid_zoom(tmp_path):
    """fix_docx 的成果必须真的落在产物里，而不只是 rc=0。"""
    spec = tmp_path / "spec.json"
    spec.write_text(json.dumps(SPEC), encoding="utf-8")
    out = tmp_path / "out.docx"

    assert _run_create(tmp_path, spec, out).returncode == 0

    with zipfile.ZipFile(out) as z:
        settings = z.read("word/settings.xml").decode("utf-8")
    assert "w:percent" in settings


# --------------------------------------------------------------------------
# 透传本身：fix_docx 报非零时 docx_create 必须一起报
#
# 这一条是变异测试逼出来的。端到端只能覆盖「出包成功」那条路（rc 必然 0），
# 而 fix_docx 的失败分支在真实出包里几乎不可达 —— 于是「把守卫删掉」这种
# 回归，在没有本条时**测试全绿**（实测 M1：9 passed）。所以失败路径必须靠
# monkeypatch 造，否则这条修复等于没被任何用例锁住。
# --------------------------------------------------------------------------
def test_docx_create_propagates_a_nonzero_fix_zoom_result(tmp_path, monkeypatch):
    import docx_create

    spec = tmp_path / "spec.json"
    spec.write_text(json.dumps(SPEC), encoding="utf-8")
    out = tmp_path / "out.docx"

    def fake_fix_docx(path, dst=None, percent=100):
        return 2

    monkeypatch.setattr(docx_create, "fix_docx", fake_fix_docx)
    monkeypatch.setattr(sys, "argv",
                        ["docx_create.py", str(spec), str(out)])

    assert docx_create.main() == 2


def test_docx_create_keeps_exit_0_when_fix_zoom_succeeds(tmp_path, monkeypatch):
    """反向：fix_docx 返 0 时不能被这条守卫误伤。"""
    import docx_create

    spec = tmp_path / "spec.json"
    spec.write_text(json.dumps(SPEC), encoding="utf-8")
    out = tmp_path / "out.docx"

    monkeypatch.setattr(docx_create, "fix_docx",
                        lambda path, dst=None, percent=100: 0)
    monkeypatch.setattr(sys, "argv",
                        ["docx_create.py", str(spec), str(out)])

    assert docx_create.main() == 0
