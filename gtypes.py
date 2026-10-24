"""드라이브 항목 종류 판별과 표시 이름 (구글 라이브러리 없이 동작)."""
from __future__ import annotations

FOLDER_MIME = "application/vnd.google-apps.folder"
SHORTCUT_MIME = "application/vnd.google-apps.shortcut"

# Google 문서/시트/슬라이드는 원본 파일이 없어서 오피스 형식으로 '내보내기'해야 열 수 있다
GOOGLE_EXPORT = {
    "application/vnd.google-apps.document": (
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document", ".docx"),
    "application/vnd.google-apps.spreadsheet": (
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", ".xlsx"),
    "application/vnd.google-apps.presentation": (
        "application/vnd.openxmlformats-officedocument.presentationml.presentation", ".pptx"),
}

_GOOGLE_KIND = {
    "application/vnd.google-apps.document": ("Google 문서", "gdoc"),
    "application/vnd.google-apps.spreadsheet": ("Google 스프레드시트", "gsheet"),
    "application/vnd.google-apps.presentation": ("Google 프레젠테이션", "gslides"),
    "application/vnd.google-apps.form": ("Google 설문지", "gother"),
    "application/vnd.google-apps.drawing": ("Google 드로잉", "gother"),
    "application/vnd.google-apps.site": ("Google 사이트", "gother"),
    "application/vnd.google-apps.jam": ("Jamboard", "gother"),
    "application/vnd.google-apps.script": ("Apps Script", "gother"),
    "application/vnd.google-apps.map": ("Google 지도", "gother"),
}

_EXT_KIND: dict[str, tuple[str, str]] = {}


def _reg(label: str, icon: str, *exts: str) -> None:
    for e in exts:
        _EXT_KIND[e] = (label, icon)


_reg("Word 문서", "word", ".doc", ".docx", ".docm", ".dot", ".dotx", ".odt")
_reg("서식 있는 텍스트", "word", ".rtf")
_reg("Excel 워크시트", "excel", ".xls", ".xlsx", ".xlsm", ".xlsb", ".xlt", ".xltx", ".ods")
_reg("CSV 파일", "excel", ".csv")
_reg("PowerPoint 프레젠테이션", "ppt", ".ppt", ".pptx", ".pptm", ".pps", ".ppsx", ".odp")
_reg("한글 문서", "hwp", ".hwp", ".hwpx", ".hwt")
_reg("한셀 문서", "excel", ".cell")
_reg("한쇼 문서", "ppt", ".show")
_reg("PDF 문서", "pdf", ".pdf")
_reg("이미지", "image", ".jpg", ".jpeg", ".png", ".gif", ".bmp", ".webp", ".heic", ".tif",
     ".tiff", ".svg", ".ico")
_reg("동영상", "video", ".mp4", ".mov", ".avi", ".mkv", ".wmv", ".m4v", ".webm")
_reg("오디오", "audio", ".mp3", ".wav", ".m4a", ".aac", ".flac", ".ogg", ".wma")
_reg("압축 파일", "zip", ".zip", ".7z", ".rar", ".alz", ".egg", ".tar", ".gz")
_reg("텍스트 문서", "text", ".txt", ".md", ".log", ".json", ".xml")


def is_folder(meta: dict) -> bool:
    return meta.get("mimeType") == FOLDER_MIME


def is_shortcut(meta: dict) -> bool:
    return meta.get("mimeType") == SHORTCUT_MIME


def is_folderish(meta: dict) -> bool:
    """폴더이거나 폴더를 가리키는 바로가기 (정렬할 때 위쪽에 둠)."""
    if is_folder(meta):
        return True
    return is_shortcut(meta) and (meta.get("shortcutDetails") or {}).get("targetMimeType") == FOLDER_MIME


def is_google_native(meta: dict) -> bool:
    mime = meta.get("mimeType", "")
    return mime.startswith("application/vnd.google-apps") and mime not in (FOLDER_MIME, SHORTCUT_MIME)


def is_exportable(meta: dict) -> bool:
    return meta.get("mimeType") in GOOGLE_EXPORT


def is_drive_root(meta: dict) -> bool:
    return is_folder(meta) and bool(meta.get("driveId")) and meta.get("driveId") == meta.get("id")


def local_name(meta: dict) -> str:
    """작업 폴더에 저장할 이름 (Google 문서는 내보낼 확장자를 붙임)."""
    name = meta.get("name") or "이름 없음"
    mime = meta.get("mimeType", "")
    if mime in GOOGLE_EXPORT:
        ext = GOOGLE_EXPORT[mime][1]
        if not name.lower().endswith(ext):
            name += ext
    return name


def _ext(name: str) -> str:
    dot = name.rfind(".")
    return name[dot:].lower() if dot > 0 else ""


def kind_label(meta: dict) -> str:
    if is_folder(meta):
        return "공유 드라이브" if is_drive_root(meta) else "폴더"
    if is_shortcut(meta):
        return "바로가기"
    mime = meta.get("mimeType", "")
    if mime in _GOOGLE_KIND:
        return _GOOGLE_KIND[mime][0]
    if is_google_native(meta):
        return "Google 파일"
    ext = _ext(meta.get("name", ""))
    if ext in _EXT_KIND:
        return _EXT_KIND[ext][0]
    return f"{ext[1:].upper()} 파일" if ext else "파일"


def icon_key(meta: dict) -> str:
    if is_folder(meta):
        return "drive" if is_drive_root(meta) else "folder"
    if is_shortcut(meta):
        return "sc_folder" if is_folderish(meta) else "sc_file"
    mime = meta.get("mimeType", "")
    if mime in _GOOGLE_KIND:
        return _GOOGLE_KIND[mime][1]
    if is_google_native(meta):
        return "gother"
    return _EXT_KIND.get(_ext(meta.get("name", "")), ("", "file"))[1]
