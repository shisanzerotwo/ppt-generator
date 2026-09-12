"""S1 防回归用例：`export_pages()` 只允许 Close **我们自己打开**的那份稿。

背景（真机实测见 `tools/probes/s1_probe_v2.py`）：
    PowerPoint 对**已打开的同一路径**返回既有实例，`Presentations.Count` 不增加。
    早期实现里 `finally: pres.Close()` 无条件执行 —— 用户若正开着这份稿，会被关掉
    （未保存修改即丢失）。修复：用「Open 前后 Count 是否增加」判定归属，只关自己那份。

这两个用例用打桩 COM 把两种归属分开验证，**不启动 PowerPoint**：
    - 用户已打开（Count 不增）→ 绝不能 Close，绝不能 Quit
    - 我们自己打开（Count +1）→ 必须 Close；关完 Count==0 时按守卫 Quit
"""

import sys
import types

import pptx_io

PNG = b"\x89PNG\r\n\x1a\n" + b"0" * 64


class _FakeSlide:
    def __init__(self, png, on_export):
        self._png, self._on_export = png, on_export

    def Export(self, path, fmt, w, h):                      # noqa: N802
        with open(path, "wb") as f:
            f.write(self._png)
        self._on_export()


class _FakeSlides:
    Count = 1

    def __init__(self, png, on_export):
        self._png, self._on_export = png, on_export

    def __call__(self, i):
        return _FakeSlide(self._png, self._on_export)


class _FakePres:
    def __init__(self, png, on_export, on_close):
        self.PageSetup = types.SimpleNamespace(SlideWidth=12192000, SlideHeight=6858000)
        self.Slides = _FakeSlides(png, on_export)
        self._on_close = on_close
        self.closed = 0

    def Close(self):                                        # noqa: N802
        self.closed += 1
        self._on_close()


class _FakePresentations:
    """`Open` 之后 Count 变成 `post`；`Close` 会让它减一（模拟真实语义）。"""

    def __init__(self, pre, post, png, on_export):
        self._count, self._post, self._png = pre, post, png
        self._on_export = on_export
        self.opened = None

    @property
    def Count(self):
        return self._count

    def Open(self, *a, **kw):                               # noqa: N802
        self._count = self._post
        self.opened = _FakePres(self._png, self._on_export, self._dec)
        return self.opened

    def _dec(self):
        self._count -= 1


class _FakeApp:
    def __init__(self, pre, post, png, on_export):
        self.Presentations = _FakePresentations(pre, post, png, on_export)
        self.quit_calls = 0

    def Quit(self):                                         # noqa: N802
        self.quit_calls += 1


def _install(monkeypatch, *, pre, post):
    exported = []
    app = _FakeApp(pre, post, PNG, lambda: exported.append(1))

    fake_pythoncom = types.ModuleType("pythoncom")
    fake_pythoncom.CoInitialize = lambda: None
    fake_pythoncom.CoUninitialize = lambda: None
    fake_client = types.ModuleType("win32com.client")
    fake_client.DispatchEx = lambda _progid: app
    fake_win32com = types.ModuleType("win32com")
    fake_win32com.client = fake_client

    monkeypatch.setitem(sys.modules, "pythoncom", fake_pythoncom)
    monkeypatch.setitem(sys.modules, "win32com", fake_win32com)
    monkeypatch.setitem(sys.modules, "win32com.client", fake_client)
    monkeypatch.setattr(pptx_io, "powerpoint_available", lambda: True)
    monkeypatch.setattr(pptx_io, "_powerpoint_running", lambda: False)
    return app, exported


def _deck(tmp_path):
    p = tmp_path / "deck.pptx"
    p.write_bytes(b"x")
    return str(p)


def test_user_owned_deck_is_not_closed(tmp_path, monkeypatch):
    """用户已开着同一份稿（Open 后 Count 不增）→ 不能 Close，也不能 Quit。"""
    app, exported = _install(monkeypatch, pre=1, post=1)

    shots = pptx_io.export_pages(_deck(tmp_path), str(tmp_path / "bg"), 1280)

    assert exported == [1], "导出本身应当成功"
    assert app.Presentations.opened.closed == 0, "用户已打开的稿被 Close 了 —— S1 回归！"
    assert app.quit_calls == 0, "用户稿仍开着，Quit 守卫应挡住"


def test_own_deck_is_closed_and_then_quit(tmp_path, monkeypatch):
    """稿是我们打开的（Count +1）→ 必须 Close；关完 Count==0 → 按守卫 Quit。"""
    app, exported = _install(monkeypatch, pre=0, post=1)

    shots = pptx_io.export_pages(_deck(tmp_path), str(tmp_path / "bg"), 1280)

    assert exported == [1]
    assert app.Presentations.opened.closed == 1, "我们自己打开的稿必须关掉"
    assert app.quit_calls == 1, "结束后 Count==0 且事先无 PowerPoint → 应 Quit"


def test_failed_open_does_not_close_or_crash(tmp_path, monkeypatch):
    """Open 抛异常时：不应 Close（pres is None），且错误要包装成 PptxError。

    pre=1 表示「用户那边已经有稿」：此时 Quit 守卫必须挡住（Count != 0），
    否则 Open 失败反而会把用户的 PowerPoint 关掉。
    """
    app, _ = _install(monkeypatch, pre=1, post=1)

    def boom(*a, **kw):
        raise RuntimeError("PowerPoint 忙")

    app.Presentations.Open = boom

    try:
        pptx_io.export_pages(_deck(tmp_path), str(tmp_path / "bg"), 1280)
        raise AssertionError("应当抛出 PptxError")
    except pptx_io.PptxError as exc:
        assert exc.code == "COM_EXPORT_FAILED"
    assert app.quit_calls == 0, "用户仍有稿在场（Count=1），Quit 守卫必须挡住"
