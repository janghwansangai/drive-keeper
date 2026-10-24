"""실기 테스트용 Windows 도우미: 프로세스·창 찾기, 최소 docx 만들기, COM 재시도."""
from __future__ import annotations

import csv
import io
import subprocess
import threading
import time
import zipfile
from pathlib import Path

import win32con
import win32gui
import win32process

RPC_E_CALL_REJECTED = -2147418111
RPC_E_SERVERCALL_RETRYLATER = -2147417846


def pids_of(exe: str) -> set[int]:
    out = subprocess.run(["tasklist", "/FI", f"IMAGENAME eq {exe}", "/FO", "CSV", "/NH"],
                         capture_output=True, text=True, encoding="mbcs", errors="replace").stdout
    pids = set()
    for row in csv.reader(io.StringIO(out)):
        if len(row) >= 2 and row[0].lower() == exe.lower():
            pids.add(int(row[1]))
    return pids


def windows_of(pid: int) -> list[tuple[int, bool, str, str]]:
    found = []

    def cb(h, acc):
        _, p = win32process.GetWindowThreadProcessId(h)
        if p == pid:
            acc.append((h, bool(win32gui.IsWindowVisible(h)), win32gui.GetClassName(h), win32gui.GetWindowText(h)))
        return True
    win32gui.EnumWindows(cb, found)
    return found


def child_texts(hwnd: int) -> list[tuple[int, str, str]]:
    out = []

    def cb(c, acc):
        acc.append((c, win32gui.GetClassName(c), win32gui.GetWindowText(c)))
        return True
    try:
        win32gui.EnumChildWindows(hwnd, cb, out)
    except Exception:
        pass
    return out


class DialogDismisser:
    """지정한 프로세스에 뜬 대화상자 중 규칙에 맞는 것의 버튼을 눌러 준다.

    rules: [(본문에 포함된 글자, 누를 버튼 글자의 시작)]
    설정을 바꾸지 않는 쪽 버튼(아니요/확인)만 쓰도록 호출하는 쪽에서 정한다.
    """

    def __init__(self, pids_fn, rules: list[tuple[str, str]]) -> None:
        self.pids_fn = pids_fn
        self.rules = rules
        self.clicked: list[str] = []
        self._stop = threading.Event()
        self._t = threading.Thread(target=self._run, daemon=True)
        self._t.start()

    def _run(self) -> None:
        while not self._stop.is_set():
            try:
                for pid in self.pids_fn():
                    for h, visible, cls, _title in windows_of(pid):
                        if not visible or cls != "#32770":
                            continue
                        kids = child_texts(h)
                        body = " ".join(t for _, c, t in kids if c == "Static")
                        for needle, button in self.rules:
                            if needle in body:
                                for c, ccls, text in kids:
                                    if ccls == "Button" and text.replace("&", "").startswith(button):
                                        win32gui.PostMessage(c, win32con.BM_CLICK, 0, 0)
                                        self.clicked.append(f"{needle} -> {button}")
                                        time.sleep(0.5)
                                        break
            except Exception:
                pass
            self._stop.wait(0.3)

    def stop(self) -> None:
        self._stop.set()


def com_retry(fn, timeout: float = 40.0):
    """Office가 바쁠 때(대화상자·초기화 중) 거부된 COM 호출을 다시 시도."""
    import pywintypes
    end = time.monotonic() + timeout
    while True:
        try:
            return fn()
        except pywintypes.com_error as e:
            if e.hresult in (RPC_E_CALL_REJECTED, RPC_E_SERVERCALL_RETRYLATER) and time.monotonic() < end:
                time.sleep(0.3)
                continue
            raise


def close_process_windows(pid: int, wait: float = 15.0) -> bool:
    """보이는 최상위 창에 닫기(WM_CLOSE)를 보내고 프로세스가 끝나길 기다린다."""
    for h, visible, cls, title in windows_of(pid):
        if visible and cls != "#32770":
            win32gui.PostMessage(h, win32con.WM_CLOSE, 0, 0)
    end = time.monotonic() + wait
    while time.monotonic() < end:
        if pid not in pids_of_pid(pid):
            return True
        time.sleep(0.3)
    return False


def pids_of_pid(pid: int) -> set[int]:
    out = subprocess.run(["tasklist", "/FI", f"PID eq {pid}", "/FO", "CSV", "/NH"],
                         capture_output=True, text=True, encoding="mbcs", errors="replace").stdout
    return {pid} if f'"{pid}"' in out else set()


def kill(pid: int) -> None:
    subprocess.run(["taskkill", "/PID", str(pid), "/F"], capture_output=True)


def make_docx_with_word(path: Path, text: str) -> None:
    """MS Word로 진짜 docx를 만든다 (한워드는 손으로 만든 최소 docx를 열지 못함).

    이 PC는 Word가 기본 프로그램이 아니라 시작할 때 '기본 프로그램으로 설정할까요?'를 묻는다.
    설정을 바꾸지 않는 '아니요'만 누른다. 이 함수가 띄운 Word만 정리한다.
    """
    import pythoncom
    import win32com.client
    pythoncom.CoInitialize()
    path.parent.mkdir(parents=True, exist_ok=True)
    before = pids_of("WINWORD.EXE")
    dis = DialogDismisser(lambda: pids_of("WINWORD.EXE") - before, [("기본 프로그램", "아니요")])
    word = win32com.client.DispatchEx("Word.Application")
    try:
        com_retry(lambda: setattr(word, "Visible", False))
        word = com_retry(lambda: win32com.client.gencache.EnsureDispatch(word))
        d = com_retry(lambda: word.Documents.Add())
        com_retry(lambda: d.Range().InsertAfter(text))
        com_retry(lambda: d.SaveAs2(str(path), 16))
        d.Close(False)
        word.Quit(0)
    finally:
        time.sleep(2)
        for pid in pids_of("WINWORD.EXE") - before:
            kill(pid)
        dis.stop()


def make_min_docx(path: Path, text: str) -> None:
    """Word가 열 수 있는 최소 docx (한워드는 못 엶 → make_docx_with_word 사용)."""
    path.parent.mkdir(parents=True, exist_ok=True)
    ct = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
          '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
          '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
          '<Default Extension="xml" ContentType="application/xml"/>'
          '<Override PartName="/word/document.xml" ContentType="application/'
          'vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>')
    rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/'
            'relationships/officeDocument" Target="word/document.xml"/></Relationships>')
    doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
           '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
           f'<w:body><w:p><w:r><w:t>{text}</w:t></w:r></w:p></w:body></w:document>')
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", ct)
        z.writestr("_rels/.rels", rels)
        z.writestr("word/document.xml", doc)
