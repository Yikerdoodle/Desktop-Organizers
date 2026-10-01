r"""
FILE FACTS - everything plain code (no AI) can learn about a file, so the sorter decides with more to go on.

  every file   where it was downloaded from (Windows quietly remembers the website: "Zone.Identifier")
  text         the text itself (web pages: tags stripped)
  Word/Excel/PowerPoint/LibreOffice   the text + title, subject, keywords, author
  old .doc / .rtf                     readable text pulled out of the file
  e-books (.epub)                     title, author, subject + the first chapters
  PDFs         text of the first pages + title/author/subject (pypdf);
               scanned PDFs (pages that are pictures): the page picture is read with Windows OCR
  pictures     text inside the picture (Windows OCR), camera, date taken, size, "looks like a screenshot"
  video/audio  length, recording date, recording device, song/album tags (ffprobe)
  zips         the names of the files inside
  installers   program name, maker, version
  shortcuts    what they point to (.lnk target / .url web address)

Everything runs locally. Results are cached per file (size + modified time), so each file is only read once.
"""
from __future__ import annotations

import ctypes
import html
import json
import logging
import os
import re
import shutil
import subprocess
import tempfile
import zipfile
from ctypes import wintypes
from dataclasses import dataclass, field
from pathlib import Path
from urllib.parse import urlparse

VIDEO = {".mov", ".mp4", ".m4v", ".mkv", ".avi", ".webm", ".wmv", ".3gp"}
IMAGE = {".jpg", ".jpeg", ".png", ".gif", ".webp", ".heic", ".bmp", ".tif", ".tiff", ".jfif", ".avif"}
AUDIO = {".mp3", ".wav", ".m4a", ".flac", ".ogg", ".opus", ".aac", ".wma"}
TEXT = {".txt", ".md", ".csv", ".tsv", ".json", ".log", ".py", ".ps1", ".js", ".ts", ".yaml", ".yml",
        ".ini", ".cfg", ".toml", ".html", ".htm", ".xml", ".env", ".bat", ".cmd", ".srt", ".vtt"}
ZIPDOCS = {".docx": ["word/document.xml"], ".odt": ["content.xml"], ".ods": ["content.xml"],
           ".odp": ["content.xml"], ".xlsx": ["xl/sharedStrings.xml"], ".pptx": []}
OTHER_DOCS = {".doc", ".rtf", ".xls", ".ppt", ".epub", ".mobi", ".pages"}
INSTALLERS = {".exe", ".msi", ".msix", ".msixbundle", ".appx", ".appxbundle"}
ARCHIVES = {".zip", ".7z", ".rar", ".gz", ".tgz", ".tar"}
PARTIAL = {".crdownload", ".part", ".partial", ".download", ".opdownload", ".tmp", ".temp", ".!ut", ".aria2"}
TEXT_LIMIT = 4000
FACTS_VERSION = 3  # bump whenever extraction changes, so every file gets re-read once

logging.getLogger("pypdf").setLevel(logging.CRITICAL)


def kind_of(path: Path, is_dir: bool) -> str:
    if is_dir:
        return "folder"
    ext = path.suffix.lower()
    for kind, exts in (("video", VIDEO), ("image", IMAGE), ("audio", AUDIO), ("text", TEXT),
                       ("installer", INSTALLERS), ("archive", ARCHIVES)):
        if ext in exts:
            return kind
    if ext == ".pdf":
        return "pdf"
    if ext in ZIPDOCS or ext in OTHER_DOCS:
        return "document"
    if ext in (".lnk", ".url"):
        return "shortcut"
    return "other"


@dataclass
class Facts:
    text: str = ""                                   # readable content (document text, text found by OCR, ...)
    meta: dict = field(default_factory=dict)         # short labelled facts, e.g. {"downloaded from": "sefaria.org"}


# ----------------------------------------------------------------------------- small readers

def _clean(s) -> str:
    s = re.sub(r"[\x00-\x1f\x7f]", "", html.unescape(str(s)).replace("\n", " "))
    return re.sub(r"\s+", " ", s).strip()[:120]


def _xml_to_text(data: bytes) -> str:
    s = data.decode("utf-8", "ignore")
    s = re.sub(r"(?is)<(script|style)\b.*?</\1>", " ", s)
    s = re.sub(r"(?i)</(?:w:p|text:p|text:h|a:p|p|div|h\d|li|tr|br)>", "\n", s)
    s = html.unescape(re.sub(r"<[^>]+>", " ", s))
    return re.sub(r"[ \t]+", " ", s)


_CLOUD_HOSTS = re.compile(r"(amazonaws\.com|cloudfront\.net|googleusercontent\.com|akamai\w*\.net|azureedge\.net|"
                          r"blob\.core\.windows\.net|storage\.googleapis\.com)$")


def download_source(path: Path) -> str:
    """The website a file was downloaded from, from the hidden 'Zone.Identifier' note Windows attaches."""
    try:
        with open(str(path) + ":Zone.Identifier", encoding="utf-8", errors="ignore") as f:
            zone = f.read()
    except OSError:
        return ""
    urls = dict(re.findall(r"^(ReferrerUrl|HostUrl)=(.+?)\s*$", zone, re.M))
    hosts = []
    for key in ("ReferrerUrl", "HostUrl"):
        url = urls.get(key, "")
        if url.startswith(("http://", "https://")):
            host = (urlparse(url).hostname or "").removeprefix("www.")
            if host and host not in hosts:
                hosts.append(host)
    useful = [h for h in hosts if not _CLOUD_HOSTS.search(h)] or hosts
    return ", ".join(useful[:2])


def _office(path: Path, ext: str) -> Facts:
    meta, out = {}, ""
    with zipfile.ZipFile(path) as z:
        names = z.namelist()
        for part, tags in (("docProps/core.xml", (("dc:title", "title"), ("dc:subject", "subject"),
                                                  ("cp:keywords", "keywords"), ("dc:description", "description"),
                                                  ("dc:creator", "author"))),
                           ("meta.xml", (("dc:title", "title"), ("dc:subject", "subject"), ("meta:keyword", "keywords"),
                                         ("dc:description", "description"), ("meta:initial-creator", "author")))):
            if part in names:
                x = z.read(part).decode("utf-8", "ignore")
                for tag, label in tags:
                    m = re.search(rf"<{tag}\b[^>]*>(.*?)</{tag}>", x, re.S)
                    if m and m.group(1).strip() and label not in meta:
                        meta[label] = _clean(m.group(1))
        parts = ZIPDOCS[ext] or sorted((n for n in names if re.match(r"ppt/slides/slide\d+\.xml$", n)),
                                       key=lambda n: int(re.search(r"\d+", n).group()))
        for n in parts:
            if n in names:
                out += _xml_to_text(z.read(n)) + "\n"
            if len(out) >= TEXT_LIMIT:
                break
    return Facts(out[:TEXT_LIMIT], meta)


def _epub(path: Path) -> Facts:
    meta, out = {}, ""
    with zipfile.ZipFile(path) as z:
        names = z.namelist()
        opf = next((n for n in names if n.lower().endswith(".opf")), None)
        if opf:
            x = z.read(opf).decode("utf-8", "ignore")
            for tag, label in (("dc:title", "title"), ("dc:creator", "author"), ("dc:subject", "subject")):
                m = re.search(rf"<{tag}\b[^>]*>(.*?)</{tag}>", x, re.S)
                if m:
                    meta[label] = _clean(m.group(1))
        pages = [n for n in names if n.lower().endswith((".xhtml", ".html", ".htm"))
                 and not re.search(r"(nav|toc|cover|copyright)", n, re.I)]
        for n in pages[:6]:
            out += _xml_to_text(z.read(n)) + "\n"
            if len(out) >= TEXT_LIMIT:
                break
    return Facts(out[:TEXT_LIMIT], meta)


def _legacy_doc(path: Path) -> Facts:
    """Old binary Word/Excel files: pull out runs of readable text (good enough for matching words)."""
    with path.open("rb") as f:
        data = f.read(3_000_000)
    runs = [r.decode("utf-16-le") for r in re.findall(rb"(?:[\x20-\x7e]\x00){8,}", data)]
    runs += [r.decode("cp1252") for r in re.findall(rb"[\x20-\x7e]{16,}", data)]
    words = [r for r in runs if r.count(" ") >= 2 and not re.search(r"(Times New Roman|Microsoft|Normal\.dot)", r)]
    return Facts("\n".join(words)[:TEXT_LIMIT])


def _rtf(path: Path) -> Facts:
    s = path.read_bytes()[:400_000].decode("latin-1", "ignore")
    s = re.sub(r"\{\\\*[^{}]*\}", " ", s)
    s = re.sub(r"\\'[0-9a-fA-F]{2}", "", s)
    s = re.sub(r"\\[a-zA-Z]+-?\d* ?", " ", s)
    return Facts(re.sub(r"[{}\s]+", " ", s)[:TEXT_LIMIT])


def _text(path: Path) -> Facts:
    with path.open("rb") as f:
        raw = f.read(64_000)
    enc = "utf-16" if raw[:2] in (b"\xff\xfe", b"\xfe\xff") else "utf-8"
    s = raw.decode(enc, "ignore")
    meta = {}
    if path.suffix.lower() in (".html", ".htm", ".xml"):
        m = re.search(r"(?is)<title[^>]*>(.*?)</title>", s)
        if m and m.group(1).strip():
            meta["title"] = _clean(m.group(1))
        s = _xml_to_text(s.encode("utf-8"))
    return Facts(s[:TEXT_LIMIT], meta)


def _pdf(path: Path, ocr_jobs: list, tmpdir: Path) -> Facts:
    import pypdf
    reader = pypdf.PdfReader(str(path))
    if reader.is_encrypted:
        try:
            reader.decrypt("")
        except Exception:
            return Facts("", {"pdf": "password-protected"})
    meta = {}
    info = reader.metadata or {}
    for key, label in (("/Title", "title"), ("/Subject", "subject"), ("/Author", "author"), ("/Keywords", "keywords")):
        value = info.get(key)
        if value and str(value).strip() and not re.fullmatch(r"(untitled|microsoft word - .*)", str(value).strip(), re.I):
            meta[label] = _clean(value)
    meta["pages"] = str(len(reader.pages))
    text = ""
    for page in reader.pages[:3]:
        try:
            text += (page.extract_text() or "") + "\n"
        except Exception:
            pass
        if len(text) >= TEXT_LIMIT:
            break
    if len(text.strip()) < 40:
        # Scanned PDF: the pages are pictures. Pull out the page picture and read it with Windows OCR.
        meta["scanned"] = "yes"
        for i, page in enumerate(reader.pages[:2]):
            try:
                pics = list(page.images)
                if not pics:
                    continue
                img = max(pics, key=lambda p: len(p.data)).image
                img.thumbnail((2600, 2600))
                out = tmpdir / f"pdfpage-{abs(hash(str(path)))}-{i}.png"
                img.convert("RGB").save(out)
                ocr_jobs.append((out, path))
            except Exception:
                continue
    return Facts(text[:TEXT_LIMIT], meta)


def _zip(path: Path) -> Facts:
    with zipfile.ZipFile(path) as z:
        files = [n for n in z.namelist() if not n.endswith("/")]
    return Facts("\n".join(files[:80]), {"contains": f"{len(files)} files"})


def lnk_target(path: Path) -> str:
    """Where a Windows .lnk shortcut points (parsed from the file, per Microsoft's MS-SHLLINK format)."""
    data = path.read_bytes()
    if len(data) < 76 or data[:4] != b"L\x00\x00\x00":
        return ""
    flags = int.from_bytes(data[20:24], "little")
    pos = 76
    if flags & 0x01:  # has a target ID list: skip it
        pos += 2 + int.from_bytes(data[pos:pos + 2], "little")
    if flags & 0x02:  # has LinkInfo: holds the local path
        li = data[pos:]
        header = int.from_bytes(li[4:8], "little")
        if int.from_bytes(li[8:12], "little") & 1:
            if header >= 0x24:
                off = int.from_bytes(li[28:32], "little")
                if off:
                    end = li.find(b"\x00\x00", off)
                    while end != -1 and (end - off) % 2:
                        end = li.find(b"\x00\x00", end + 1)
                    return li[off:end].decode("utf-16-le", "ignore")
            off = int.from_bytes(li[16:20], "little")
            return li[off:li.find(b"\x00", off)].decode("cp1252", "ignore")
    m = re.search(r"[A-Za-z]:\\[^\x00]{3,200}", data.decode("utf-16-le", "ignore"))
    return m.group(0) if m else ""


def _shortcut(path: Path) -> Facts:
    if path.suffix.lower() == ".url":
        m = re.search(r"^URL=(.+)$", path.read_text(encoding="utf-8", errors="ignore"), re.M)
        url = m.group(1).strip() if m else ""
        return Facts(url, {"link": url[:120]} if url else {})
    target = lnk_target(path)
    return Facts(target, {"points to": target[:160]} if target else {})


def exe_info(path: Path) -> tuple[str | None, str | None, str | None]:
    """ProductName, ProductVersion, CompanyName from a Windows .exe (the 'Details' tab in Properties)."""
    try:
        ver = ctypes.WinDLL("version")
        ver.GetFileVersionInfoSizeW.argtypes = [wintypes.LPCWSTR, ctypes.POINTER(wintypes.DWORD)]
        ver.GetFileVersionInfoW.argtypes = [wintypes.LPCWSTR, wintypes.DWORD, wintypes.DWORD, ctypes.c_void_p]
        ver.VerQueryValueW.argtypes = [ctypes.c_void_p, wintypes.LPCWSTR,
                                       ctypes.POINTER(ctypes.c_void_p), ctypes.POINTER(wintypes.UINT)]
        size = ver.GetFileVersionInfoSizeW(str(path), None)
        if not size:
            return None, None, None
        buf = ctypes.create_string_buffer(size)
        if not ver.GetFileVersionInfoW(str(path), 0, size, buf):
            return None, None, None
        ptr, length = ctypes.c_void_p(), wintypes.UINT()
        codes = []
        if ver.VerQueryValueW(buf, r"\VarFileInfo\Translation", ctypes.byref(ptr), ctypes.byref(length)) \
                and length.value >= 4:
            lang, cp = ctypes.cast(ptr, ctypes.POINTER(ctypes.c_ushort * 2)).contents
            codes.append(f"{lang:04x}{cp:04x}")
        codes += ["040904b0", "040904e4", "000004b0"]

        def query(field_name):
            for code in codes:
                if ver.VerQueryValueW(buf, f"\\StringFileInfo\\{code}\\{field_name}",
                                      ctypes.byref(ptr), ctypes.byref(length)) and length.value:
                    value = ctypes.wstring_at(ptr, length.value).rstrip("\x00").strip()
                    if value:
                        return value
            return None

        return query("ProductName"), query("ProductVersion") or query("FileVersion"), query("CompanyName")
    except Exception:
        return None, None, None


def _installer(path: Path) -> Facts:
    if path.suffix.lower() != ".exe":
        return Facts()
    product, version, company = exe_info(path)
    meta = {k: _clean(v) for k, v in (("program", product), ("version", version), ("maker", company)) if v}
    return Facts(" ".join(filter(None, (product, company))), meta)


def _media(path: Path, ffprobe: str | None) -> Facts:
    if not ffprobe:
        return Facts()
    r = subprocess.run([ffprobe, "-v", "error", "-print_format", "json", "-show_format", "-show_streams", str(path)],
                       capture_output=True, timeout=30, creationflags=subprocess.CREATE_NO_WINDOW)
    data = json.loads(r.stdout or b"{}")
    fmt = data.get("format", {})
    tags = {k.lower(): str(v) for k, v in (fmt.get("tags") or {}).items()}
    meta = {}
    seconds = float(fmt.get("duration") or 0)
    if seconds:
        meta["length"] = f"{int(seconds // 60)} min {int(seconds % 60)} s" if seconds >= 60 else f"{seconds:.0f} s"
    created = tags.get("com.apple.quicktime.creationdate") or tags.get("creation_time") or tags.get("date")
    if created:
        meta["recorded"] = created[:10]
    make = tags.get("com.apple.quicktime.make") or tags.get("make") or ""
    model = tags.get("com.apple.quicktime.model") or tags.get("model") or ""
    if make or model:
        meta["device"] = _clean(f"{make} {model}")
    video = next((s for s in data.get("streams", []) if s.get("codec_type") == "video"
                  and s.get("disposition", {}).get("attached_pic") != 1), None)
    if video and video.get("width"):
        meta["resolution"] = f"{video['width']}x{video['height']}"
    for key in ("title", "artist", "album", "album_artist", "genre", "comment"):
        if tags.get(key):
            meta[key.replace("_", " ")] = _clean(tags[key])
    words = " ".join(meta.get(k, "") for k in ("title", "artist", "album", "genre", "comment"))
    return Facts(words, meta)


# ----------------------------------------------------------------------------- the store

class FactStore:
    """Gets Facts for many files at once (OCR runs as one batch), cached so each file is read only once."""

    def __init__(self, cache_path: Path, ocr_helper: Path, ffprobe: str | None):
        self.cache_path, self.ocr_helper, self.ffprobe = cache_path, ocr_helper, ffprobe
        try:
            data = json.loads(cache_path.read_text(encoding="utf-8"))
            self.files = data["files"] if data.get("version") == FACTS_VERSION else {}
        except Exception:
            self.files = {}
        self.dirty = False

    def get_many(self, paths: list[Path], say=None) -> dict[str, Facts]:
        result, todo = {}, []
        for p in paths:
            try:
                st = p.stat()
            except OSError:
                continue
            key, stamp = str(p), [st.st_size, st.st_mtime_ns]
            hit = self.files.get(key)
            if hit and hit[0] == stamp:
                result[key] = Facts(**hit[1])
            else:
                todo.append((p, stamp))
        if not todo:
            return result
        if say and len(todo) > 15:
            say(f"Reading {len(todo)} file(s) for the first time (text, PDFs, pictures, videos) - "
                f"only happens once per file...")
        tmpdir = Path(tempfile.mkdtemp(prefix="inbox-sorter-"))
        ocr_jobs: list[tuple[Path, Path]] = []
        try:
            for p, _ in todo:
                result[str(p)] = self.extract(p, ocr_jobs, tmpdir)
            if ocr_jobs:
                self.run_ocr(ocr_jobs, result, tmpdir)
        finally:
            shutil.rmtree(tmpdir, ignore_errors=True)
        for p, stamp in todo:
            f = result[str(p)]
            self.files[str(p)] = [stamp, {"text": f.text, "meta": f.meta}]
        self.dirty = True
        return result

    def extract(self, path: Path, ocr_jobs: list, tmpdir: Path) -> Facts:
        kind, ext = kind_of(path, False), path.suffix.lower()
        facts = Facts()
        try:
            if path.stat().st_size == 0:
                facts = Facts()
            elif kind == "text":
                facts = _text(path)
            elif ext in ZIPDOCS:
                facts = _office(path, ext)
            elif ext == ".epub":
                facts = _epub(path)
            elif ext in (".doc", ".xls", ".ppt"):
                facts = _legacy_doc(path)
            elif ext == ".rtf":
                facts = _rtf(path)
            elif kind == "pdf":
                facts = _pdf(path, ocr_jobs, tmpdir)
            elif kind == "image":
                if path.stat().st_size < 60_000_000:
                    ocr_jobs.append((path, path))
            elif kind in ("video", "audio"):
                facts = _media(path, self.ffprobe)
            elif ext == ".zip":
                facts = _zip(path)
            elif kind == "installer":
                facts = _installer(path)
            elif kind == "shortcut":
                facts = _shortcut(path)
        except Exception:
            pass  # unreadable / damaged file: the sorter still has the name, size, type and dates
        source = download_source(path)
        if source:
            facts.meta["downloaded from"] = source
        return facts

    def run_ocr(self, jobs: list, result: dict, tmpdir: Path):
        """One PowerShell call reads every picture with Windows' built-in OCR (see ocr-helper.ps1)."""
        listing = tmpdir / "ocr-list.json"
        listing.write_text(json.dumps([str(p) for p, _ in jobs]), encoding="utf-8")
        owner = {str(p).lower(): o for p, o in jobs}
        shell = os.path.join(os.environ.get("SystemRoot", r"C:\Windows"), "System32", "WindowsPowerShell", "v1.0",
                             "powershell.exe")
        try:
            r = subprocess.run([shell, "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", str(self.ocr_helper),
                                "-ListFile", str(listing)], capture_output=True, timeout=60 + 10 * len(jobs),
                               creationflags=subprocess.CREATE_NO_WINDOW)
        except Exception:
            return
        for line in r.stdout.decode("utf-8", "replace").splitlines():
            try:
                o = json.loads(line.lstrip("\ufeff"))
            except ValueError:
                continue
            src = owner.get(str(o.get("path", "")).lower())
            if src is None:
                continue
            f = result[str(src)]
            text = (o.get("text") or "").strip()
            if src.suffix.lower() == ".pdf":  # a scanned PDF page
                if text:
                    f.text = (f.text + "\n" + text)[:TEXT_LIMIT]
                    f.meta["text read by OCR"] = "yes"
                continue
            f.text = text[:TEXT_LIMIT]
            camera = _clean(f"{o.get('make') or ''} {o.get('model') or ''}")
            if camera:
                f.meta["camera"] = camera
            if o.get("taken"):
                f.meta["taken"] = o["taken"]
            w, h = int(o.get("width") or 0), int(o.get("height") or 0)
            if w and h:
                f.meta["size"] = f"{w}x{h}"
                if not camera and src.suffix.lower() == ".png" and w >= 1280 and 1.2 <= w / h <= 3.6:
                    f.meta["looks like"] = "screenshot"

    def save(self):
        if self.dirty:
            live = {k: v for k, v in self.files.items() if os.path.exists(k)}
            self.cache_path.write_text(json.dumps({"version": FACTS_VERSION, "files": live}, ensure_ascii=False),
                                       encoding="utf-8")
