# -*- coding: utf-8 -*-
"""tests/run_all.py — make-ppt engine 回归测试。

用法：
    python tests/run_all.py            # 全量
    python tests/run_all.py --quick    # 只跑单元自测（units + edit）

示例语料是可选的：设置环境变量 MAKEPPT_CORPUS（目录，或其中一份 .pptx），
并放入 deck-a.pptx、deck-b.pptx、deck-merged.pptx 三份文件。未设置或文件不齐时
打印 SKIP，只跑不依赖语料的自测（units、build 的无模板分支），不会抛异常。

设计原则：**测试本身要能失败**。每个模块的 --self-test 里都带坏样本；
有语料时再跑一遍真实文件，确认"正常文件 0 ERROR"不是靠跳过检查得来的。
"""
import os
import sys
import time
import traceback

HERE = os.path.dirname(os.path.abspath(__file__))
ENGINE = os.path.abspath(os.path.join(HERE, "..", "engine"))
OUT = os.path.abspath(os.path.join(HERE, "..", "out"))
sys.path.insert(0, ENGINE)
os.makedirs(OUT, exist_ok=True)

CORPUS_ROOT = os.environ.get("MAKEPPT_CORPUS") or None
if CORPUS_ROOT and os.path.isfile(CORPUS_ROOT):
    CORPUS_ROOT = os.path.dirname(os.path.abspath(CORPUS_ROOT))
CORPUS = {
    "deck-a.pptx": os.path.join(CORPUS_ROOT, "deck-a.pptx") if CORPUS_ROOT else "",
    "deck-b.pptx": os.path.join(CORPUS_ROOT, "deck-b.pptx") if CORPUS_ROOT else "",
    "合并稿.pptx": (os.path.join(CORPUS_ROOT, "deck-merged.pptx")
                     if CORPUS_ROOT else ""),
}
HAS_CORPUS = bool(CORPUS_ROOT) and all(
    p and os.path.isfile(p) for p in CORPUS.values())

RESULTS = []


def run(name, fn, *args, **kwargs):
    t0 = time.time()
    try:
        msg = fn(*args, **kwargs)
        RESULTS.append((name, True, msg or "", time.time() - t0))
        print("  PASS  %-28s %5.1fs  %s" % (name, time.time() - t0, msg or ""))
        return True
    except Exception as e:
        RESULTS.append((name, False, "%s: %s" % (type(e).__name__, e),
                        time.time() - t0))
        print("  FAIL  %-28s %5.1fs  %s" % (name, time.time() - t0, e))
        traceback.print_exc(limit=3)
        return False


def skip(name, why):
    print("  SKIP  %-28s %s" % (name, why))


# --------------------------------------------------------------------------
def t_units():
    import units
    rc = units._self_test()
    assert rc == 0, "units 自测失败"
    return "含坏样本证伪"


def t_read():
    import read
    for label, path in CORPUS.items():
        if not os.path.exists(path):
            raise FileNotFoundError(path)
        d = read.Deck.load(path)
        assert d.n > 0, label
        # 不变量：每页有部件、有版式
        for s in d.slides:
            assert s.part and s.layout, "%s 第 %d 页缺部件或版式" % (label, s.index)
    return "%d 份语料可读" % len(CORPUS)


def t_build():
    import build
    assert build._self_test() == 0
    return "模板/无模板各建一份（无语料时只建无模板）"


def t_merge():
    import merge
    assert merge._self_test() == 0
    return "跨母版合并 2 页"


def t_edit():
    import edit
    if not hasattr(edit, "_self_test"):
        raise AssertionError("edit.py 没有 _self_test")
    assert edit._self_test() == 0
    return "改造链路（复制/删页/改字/备注）"


def t_check():
    import check
    fn = getattr(check, "_self_test", None) or getattr(check, "self_test", None)
    assert fn, "check.py 没有 _self_test / self_test"
    assert fn() == 0
    # deck-a.pptx 已修复双图标，现在应 0 ERROR
    rep = check.check(CORPUS["deck-a.pptx"])
    assert rep.ok, "deck-a.pptx 已修复，仍有 ERROR：%s" % rep.errors[:3]
    # 修复前备份应仍被检出 4 条 inherit/DUP_ICON（检查器灵敏度）
    bak = os.path.join(CORPUS_ROOT, "deck-a-before-fix.pptx")
    if os.path.exists(bak):
        rb = check.check(bak)
        assert len(rb.errors) == 4, "备份预期 4 ERROR，实际 %d" % len(rb.errors)
        assert all("inherit" in e for e in rb.errors), "备份的 ERROR 应全是 inherit"
    # 成品必须 0 ERROR
    for label in ("deck-b.pptx", "合并稿.pptx"):
        rep2 = check.check(CORPUS[label])
        assert rep2.ok, "%s 有 ERROR：%s" % (label, rep2.errors[:3])
    return "deck-a.pptx 0 ERROR；备份仍检出 4 条双图标；成品 0 ERROR"


def t_preview():
    import preview
    if not hasattr(preview, "_self_test"):
        raise AssertionError("preview.py 没有 _self_test")
    assert preview._self_test() == 0
    return "渲染 + 拼图"


def t_end_to_end():
    """全链路：读 deck-b.pptx → 复制一页 → 改字 → 写备注 → 查 → 渲染。"""
    import edit, check, preview
    src = CORPUS["deck-b.pptx"]
    out = os.path.join(OUT, "e2e.pptx")
    ed = edit.Editor.open(src)
    n0 = ed.deck.n
    ed.duplicate_slide(2)
    ed.set_notes(n0 + 1, "回归测试：这一页是复制出来的")
    ed.save(out)
    import read
    d = read.Deck.load(out)
    assert d.n == n0 + 1, "页数 %d != %d" % (d.n, n0 + 1)
    assert d.slide(n0 + 1).notes, "备注没写进去"
    rep = check.check(out)
    assert rep.ok, rep.errors[:3]
    pngs = preview.render(out, os.path.join(OUT, "e2e_preview"),
                          pages=[n0 + 1], width=1100)
    assert pngs and os.path.getsize(pngs[0]) > 5000
    return "复制页+备注+查+渲染通过"


# --------------------------------------------------------------------------
def _reload_engine():
    """同一解释器进程内重复导入时，不 reload 会测到旧代码。
    环境变量 RUN_ALL_NO_RELOAD=1 可跳过（调试用）。"""
    if os.environ.get("RUN_ALL_NO_RELOAD"):
        return
    import importlib
    for name in ("units", "read", "build", "merge", "edit", "check", "preview"):
        try:
            importlib.reload(importlib.import_module(name))
        except Exception:
            pass


def main():
    _reload_engine()
    quick = "--quick" in sys.argv
    print("engine:", ENGINE)
    if HAS_CORPUS:
        print("corpus:", CORPUS_ROOT)
    else:
        print("corpus: 未设置 MAKEPPT_CORPUS 或文件不齐 → SKIP 依赖语料的测试")
    print("=" * 62)
    run("units", t_units)
    if not quick:
        run("build", t_build)
    if HAS_CORPUS:
        run("edit", t_edit)
    else:
        skip("edit", "语料不可用（设置 MAKEPPT_CORPUS）")
    if not quick:
        for name, fn in (("read", t_read), ("merge", t_merge),
                         ("check", t_check), ("preview", t_preview),
                         ("end2end", t_end_to_end)):
            if HAS_CORPUS:
                run(name, fn)
            else:
                skip(name, "语料不可用（设置 MAKEPPT_CORPUS）")
    print("=" * 62)
    bad = [r for r in RESULTS if not r[1]]
    print("%d 项通过，%d 项失败" % (len(RESULTS) - len(bad), len(bad)))
    if bad:
        for name, _ok, msg, _t in bad:
            print("  FAIL", name, "—", msg)
        return 1
    print("ALL PASS")
    return 0


if __name__ == "__main__":
    sys.exit(main())
