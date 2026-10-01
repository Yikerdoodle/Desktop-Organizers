#!/usr/bin/env python3
r"""
SORT MY INBOX
=============
Suggests where each thing in "1 - Inbox" should go, and moves it only when you say yes.

How it decides - plain code first, AI only for the leftovers:
  1. FACTS  everything plain code can learn about each file (see facts.py): text, PDF pages,
            text inside pictures (Windows OCR), camera/date, video length/device, zip contents,
            the website it was downloaded from, what a shortcut points to, installer details
  2. RULES  certain cases -> "junk?": half-finished downloads, lock/temp files, empty files,
            exact duplicates, installers for programs you already have, zips already unzipped
  3. MATCH  math, no AI: compare the item to what's already in each folder
            (near-identical names, numbered sequences like IMG_0506 -> IMG_0507, shared words)
  4. TYPE   pictures -> Pictures, videos -> Videos, music -> Music (when nothing better matched)
  5. AI     only when 2-4 aren't sure: the local model (Ollama, via Jarvis) picks one of the
            few folders it's offered, or says UNSURE. It can't invent folders or touch files.
  6. YOU    approve. The code moves things and logs every move (sort-log.csv).
            "Junk" only ever goes to the Recycle Bin, and only after you say yes.

Run:  python inbox_sorter.py             normal, interactive
      python inbox_sorter.py --plan      just print the suggestions, change nothing (add --ai to include AI)
      python inbox_sorter.py --selftest  measure how well the matching works on your own folders
      python inbox_sorter.py --undo      undo the last sort
"""
from __future__ import annotations

import argparse
import csv
import ctypes
import difflib
import hashlib
import json
import math
import os
import re
import shutil
import stat
import subprocess
import sys
import time
import urllib.request
import winreg
from collections import Counter, defaultdict
from ctypes import wintypes
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path

from facts import PARTIAL, FactStore, Facts, exe_info, kind_of

HERE = Path(__file__).resolve().parent
CONFIG_PATH = HERE / "config.json"
CACHE_PATH = HERE / "facts-cache.json"
OCR_HELPER = HERE / "ocr-helper.ps1"
LOG_PATH = HERE.parent / "sort-log.csv"

DEFAULT_CONFIG = {
    "desktop": r"C:\Users\Public\Desktop",
    "inbox": "1 - Inbox",
    "libraries": {},
    "duplicate_search": [],
    "ignore_words": [],
    "skip_subfolders": ["older drafts", "old drafts", "old", "old versions", "backup", "backups"],
    "descriptions": {},
    "ffprobe": r"C:\ffmpeg\bin\ffprobe.exe",
    "thresholds": {"confident": 0.6, "margin": 0.15, "likely": 0.25, "likely_margin": 0.15},
    "ai": {"enabled": True, "model": "qwen3.5:4b", "ollama_url": "http://127.0.0.1:11434",
           "docker_container": "jarvis-ollama", "timeout_seconds": 240, "max_items_per_run": 8},
}


def load_config() -> dict:
    cfg = json.loads(json.dumps(DEFAULT_CONFIG))
    if CONFIG_PATH.exists():
        user = json.loads(CONFIG_PATH.read_text(encoding="utf-8-sig"))
        for key, value in user.items():
            if isinstance(value, dict) and isinstance(cfg.get(key), dict):
                cfg[key].update(value)
            else:
                cfg[key] = value
    return cfg


def is_hidden(path: Path) -> bool:
    try:
        attrs = path.stat().st_file_attributes
    except (OSError, AttributeError):
        return False
    return bool(attrs & (stat.FILE_ATTRIBUTE_HIDDEN | stat.FILE_ATTRIBUTE_SYSTEM))


def make_store(cfg: dict) -> FactStore:
    ffprobe = cfg.get("ffprobe")
    if not (ffprobe and Path(ffprobe).exists()):
        ffprobe = shutil.which("ffprobe")
    return FactStore(CACHE_PATH, OCR_HELPER, ffprobe)


# ============================================================================ words, names, sequences

STOP = set("""
of to in is on at by an or as be we he me my us it do go no so up if am
the and for with from this that into your you are was were have has had not but all any can our
out one two new old copy final draft drafts version maybe file files document doc docs untitled
image img photo screenshot screen shot page pages part its what when where which will would about
there their they them then than also just like get got make made more most some such very via per
etc www http https com org net io ca uk html pdf txt docx odt png jpg jpeg mov mp4 exe zip lnk url md
capture snip snapshot
""".split())
IGNORE: set[str] = set()  # words on everything (like your own name) that say nothing about where a file goes

# facts worth matching on (dates, sizes and page counts aren't)
MATCH_LABELS = ("title", "subject", "keywords", "description", "author", "downloaded from", "camera", "device",
                "program", "maker", "points to", "link", "artist", "album", "album artist", "genre", "comment",
                "looks like")


def tokens(s: str) -> list[str]:
    s = re.sub(r"([a-z])([A-Z])", r"\1 \2", s)
    out = []
    for w in re.findall(r"[A-Za-z\u00C0-\u024F]{2,}", s):
        w = w.lower()
        if w in STOP or w in IGNORE:
            continue
        if len(w) > 4 and w.endswith("ies"):
            w = w[:-3] + "y"
        elif len(w) > 4 and w.endswith("s") and not w.endswith("ss"):
            w = w[:-1]
        out.append(w)
    return out


GENERIC = re.compile(
    r"^(img|dsc|dscn|dcim|vid|pxl|mvimg|photo|image|picture|screenshot|screen ?shot|capture|scan|"
    r"camscanner|whatsapp (image|video|audio)|document|new text document|new document|untitled|"
    r"download|copy of)(?=[\s_\-\d.(]|$)", re.I)
SEQ = re.compile(r"^([A-Za-z]+[_\- ]?)(\d{3,6})$")


def is_generic(stem: str) -> bool:
    s = stem.strip()
    return bool(GENERIC.match(s)) or not re.search(r"[A-Za-z]{3,}", s) \
        or bool(re.fullmatch(r"[0-9a-fA-F_\-]{12,}", s))


def norm_name(stem: str) -> str:
    """Name with version noise removed, so 'Case 6' and 'Case maybe 7' look the same."""
    if is_generic(stem):
        return ""
    s = re.sub(r"([a-z])([A-Z])", r"\1 \2", stem).lower().replace("_", " ")
    if IGNORE:
        s = re.sub(r"\b(" + "|".join(map(re.escape, IGNORE)) + r")\b", " ", s)
    s = re.sub(r"\b(v|ver|version)\s*\d+\b", " ", s)
    s = re.sub(r"\b(maybe|final|copy|draft|new|old|updated|edited|edit)\b", " ", s)
    s = re.sub(r"\(\d+\)", " ", s)
    s = re.sub(r"\d+", "#", s)
    s = re.sub(r"[^a-z#]+", " ", s)
    return re.sub(r"\s+", " ", s).strip()


def seq_of(stem: str):
    m = SEQ.match(stem.strip())
    return (m.group(1).lower(), int(m.group(2))) if m else None


def bag(name_stem: str, facts: Facts) -> Counter:
    """Word bag: name counts 3x (strongest hint), useful facts 2x, plus words from the content."""
    info = " ".join(str(facts.meta[k]) for k in MATCH_LABELS if k in facts.meta)
    return Counter(tokens(name_stem) * 3 + tokens(info) * 2 + tokens(facts.text)[:400])


def facts_line(facts: Facts, limit: int = 300) -> str:
    return "; ".join(f"{k}: {v}" for k, v in facts.meta.items())[:limit]


@dataclass
class Item:
    path: Path
    name: str
    stem: str
    is_dir: bool
    kind: str
    size: int
    mtime: float
    facts: Facts
    raw: Counter
    norm: str
    seq: tuple | None
    scores: list = field(default_factory=list)


def make_item(path: Path, facts: Facts | None) -> Item:
    is_dir = path.is_dir()
    if is_dir:
        children = [p for p in path.rglob("*") if p.is_file()]
        size = sum(p.stat().st_size for p in children)
        facts = Facts("\n".join(p.name for p in children[:80]), {"contains": f"{len(children)} files"})
        stem = path.name
    else:
        size = path.stat().st_size
        facts = facts or Facts()
        stem = path.stem
    return Item(path, path.name, stem, is_dir, kind_of(path, is_dir), size, path.stat().st_mtime, facts,
                bag(stem, facts), norm_name(stem), seq_of(stem))


# ============================================================================ installed programs

WEAK_PROGRAM_WORDS = {"desktop", "studio", "player", "manager", "tool", "tools", "driver", "drivers",
                      "graphics", "client", "service", "runtime", "redistributable", "plus", "pro",
                      "free", "edition", "home", "media", "center", "control", "panel", "software",
                      "suite", "helper", "browser", "application", "launcher"}
GENERIC_PROGRAM_WORDS = {"setup", "installer", "install", "win", "windows", "bit", "full", "offline",
                         "online", "web", "update", "release", "latest", "international", "dch", "whql",
                         "unofficial", "uninstall", "app", "version", "the", "for", "and", "amd", "arm",
                         "x64", "x86", "en", "us"}
COMPANY_WORDS = {"inc", "llc", "ltd", "corporation", "corp", "co", "gmbh", "company", "limited", "the",
                 "sa", "ag", "bv", "srl", "team", "platforms", "technologies", "software", "studios", "studio"}


def program_tokens(s: str) -> set[str]:
    s = re.sub(r"https?://\S+|www\.\S+", " ", s)
    s = re.sub(r"([a-z])([A-Z])", r"\1 \2", s)
    return {w.lower() for w in re.findall(r"[A-Za-z]{2,}", s)} - GENERIC_PROGRAM_WORDS


def company_tokens(s: str | None) -> set[str]:
    return {w.lower() for w in re.findall(r"[A-Za-z]{2,}", s or "")} - COMPANY_WORDS


def version_tuple(v: str | None):
    nums = [int(n) for n in re.findall(r"\d+", v or "")[:4]]
    while nums and nums[-1] == 0 and len(nums) > 1:  # 3.7.9.0 == 3.7.9
        nums.pop()
    return tuple(nums) if nums else None


def installed_programs() -> list[tuple[str, str, str]]:
    """(name, version, publisher) for every program in Apps & Features."""
    out = []
    roots = [(winreg.HKEY_LOCAL_MACHINE, r"Software\Microsoft\Windows\CurrentVersion\Uninstall"),
             (winreg.HKEY_LOCAL_MACHINE, r"Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"),
             (winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Windows\CurrentVersion\Uninstall")]
    for hive, sub in roots:
        try:
            root = winreg.OpenKey(hive, sub)
        except OSError:
            continue
        with root:
            for i in range(winreg.QueryInfoKey(root)[0]):
                try:
                    with winreg.OpenKey(root, winreg.EnumKey(root, i)) as k:
                        name = str(winreg.QueryValueEx(k, "DisplayName")[0])
                        values = []
                        for field_name in ("DisplayVersion", "Publisher"):
                            try:
                                values.append(str(winreg.QueryValueEx(k, field_name)[0]))
                            except OSError:
                                values.append("")
                        out.append((name, *values))
                except OSError:
                    continue
    return out


def match_installed(item: Item, programs) -> tuple[str, str, str | None] | None:
    """Returns (installed name, installed version, installer version) if this installer's program is installed."""
    product, inst_ver, company = exe_info(item.path) if item.path.suffix.lower() == ".exe" else (None, None, None)
    maker = company_tokens(company)
    wanted = program_tokens(product) if product else set()
    if not wanted - WEAK_PROGRAM_WORDS:
        wanted = program_tokens(item.stem)
    if not wanted - WEAK_PROGRAM_WORDS:
        return None
    if inst_ver is None:
        m = re.search(r"\d+(?:\.\d+)+", item.stem)
        inst_ver = m.group(0) if m else None
    best, best_score = None, 0.0
    for name, version, publisher in programs:
        have = program_tokens(name)
        overlap = have & wanted
        if not overlap or not (overlap - WEAK_PROGRAM_WORDS) or not (have <= wanted or wanted <= have):
            continue
        pub = company_tokens(publisher)
        if maker and pub and not (maker & pub):
            continue  # same name, different maker (e.g. Facebook "Messenger" vs the game "The Messenger")
        score = len(overlap) / len(have | wanted)
        if score > best_score:
            best, best_score = (name, version, inst_ver), score
    return best if best_score >= 0.5 else None


# ============================================================================ exact duplicates

class DupIndex:
    """Finds byte-for-byte identical copies (same size first, then SHA-256)."""

    def __init__(self, roots: list[str], inbox: Path):
        self.inbox = inbox
        self.by_size: dict[int, list[Path]] = defaultdict(list)
        self.hashes: dict[Path, str] = {}
        seen, count = set(), 0
        for root in [str(inbox)] + list(roots):
            for dirpath, dirnames, filenames in os.walk(root):
                dirnames[:] = [d for d in dirnames if d not in ("node_modules", ".git", "__pycache__")]
                if dirpath in seen:
                    dirnames[:] = []
                    continue
                seen.add(dirpath)
                for fn in filenames:
                    p = Path(dirpath) / fn
                    try:
                        size = p.stat().st_size
                    except OSError:
                        continue
                    if size > 0:
                        self.by_size[size].append(p)
                    count += 1
                if count > 200_000:
                    return

    def _hash(self, p: Path) -> str:
        if p not in self.hashes:
            h = hashlib.sha256()
            with p.open("rb") as f:
                for chunk in iter(lambda: f.read(1 << 20), b""):
                    h.update(chunk)
            self.hashes[p] = h.hexdigest()
        return self.hashes[p]

    def copy_of(self, item: Item) -> Path | None:
        if item.is_dir or item.size == 0:
            return None
        for other in self.by_size.get(item.size, []):
            if other == item.path:
                continue
            both_in_inbox = self.inbox in other.parents
            # two identical files in the Inbox: keep the older one, flag the newer one
            if both_in_inbox and (other.stat().st_mtime, other.name) > (item.mtime, item.name):
                continue
            try:
                if self._hash(other) == self._hash(item.path):
                    return other
            except OSError:
                continue
        return None


# ============================================================================ 3. MATCH

NUMBERED = re.compile(r"^\d+ - ")


@dataclass
class Doc:
    name: str
    kind: str
    raw: Counter
    norm: str
    seq: tuple | None
    mtime: float = 0.0
    vec: dict = field(default_factory=dict)
    n2: float = 0.0


@dataclass
class Candidate:
    key: str
    path: Path
    desc: str = ""
    library: bool = False
    docs: list = field(default_factory=list)
    desc_raw: Counter = field(default_factory=Counter)
    total: dict = field(default_factory=dict)
    norm2: float = 0.0


def related(a: Candidate, b: Candidate) -> bool:
    return a.path == b.path or a.path in b.path.parents or b.path in a.path.parents


def build_candidates(cfg: dict, store: FactStore, say=None) -> list[Candidate]:
    desk = Path(cfg["desktop"])
    inbox = desk / cfg["inbox"]
    skip = {s.lower() for s in cfg["skip_subfolders"]}
    cands: list[Candidate] = []
    tops = sorted(p for p in desk.iterdir() if p.is_dir() and NUMBERED.match(p.name) and p != inbox)
    for top in tops:
        cands.append(Candidate(top.name, top))
        for sub in sorted(p for p in top.iterdir() if p.is_dir() and not is_hidden(p)):
            if sub.name.lower() not in skip:
                cands.append(Candidate(f"{top.name}\\{sub.name}", sub))
    by_path = {c.path: c for c in cands}
    owned: list[tuple[Candidate, Path]] = []
    for top in tops:
        for dirpath, dirnames, filenames in os.walk(top):
            here = Path(dirpath)
            owner = next((by_path[p] for p in [here, *here.parents] if p in by_path), None)
            for fn in filenames:
                p = here / fn
                if owner is None or fn.lower() == "desktop.ini" or is_hidden(p):
                    continue
                owned.append((owner, p))
    facts = store.get_many([p for _, p in owned], say)
    for owner, p in owned:
        if len(owner.docs) >= 800:
            continue
        f = facts.get(str(p), Facts())
        owner.docs.append(Doc(p.name, kind_of(p, False), bag(p.stem, f), norm_name(p.stem), seq_of(p.stem),
                              p.stat().st_mtime))
    for key, path in cfg["libraries"].items():
        cands.append(Candidate(key, Path(path), library=True))
    for c in cands:
        c.desc = cfg["descriptions"].get(c.key, "")
        c.desc_raw = Counter(tokens(c.key.split("\\")[-1]) * 2 + tokens(c.desc))
    return cands


def clamp(x: float) -> float:
    return max(0.0, min(1.0, x))


class Matcher:
    """TF-IDF word matching + near-identical names + numbered sequences. Pure math, no AI."""

    def __init__(self, cands: list[Candidate]):
        self.cands = cands
        self.df: Counter = Counter()
        n = 0
        for c in cands:
            for d in c.docs:
                self.df.update(d.raw.keys())
                n += 1
            if c.desc_raw:
                self.df.update(c.desc_raw.keys())
                n += 1
        self.n = n
        for c in cands:
            total: Counter = Counter()
            for d in c.docs:
                d.vec = self.weigh(d.raw)
                d.n2 = sum(w * w for w in d.vec.values())
                total.update(d.vec)
            for t, w in self.weigh(c.desc_raw).items():
                total[t] += 2 * w
            c.total = dict(total)
            c.norm2 = sum(w * w for w in total.values())

    def idf(self, t: str) -> float:
        return math.log((self.n + 1) / (self.df.get(t, 0) + 1)) + 1

    def weigh(self, raw: Counter) -> dict:
        return {t: (1 + math.log(k)) * self.idf(t) for t, k in raw.items()}

    def score(self, raw: Counter, norm: str, seq, kind: str, exclude: Doc | None = None):
        """Returns [(score 0-1, candidate, reason)] best first."""
        vec = self.weigh(raw)
        vn = math.sqrt(sum(w * w for w in vec.values()))
        results = []
        for c in self.cands:
            if c.library:
                continue
            # word overlap with the folder as a whole (centroid), minus the excluded doc in self-tests
            dot = sum(w * c.total.get(t, 0.0) for t, w in vec.items())
            n2 = c.norm2
            if exclude is not None and exclude in c.docs:
                dot -= sum(w * exclude.vec.get(t, 0.0) for t, w in vec.items())
                cross = sum(w * c.total.get(t, 0.0) for t, w in exclude.vec.items())
                n2 = n2 - 2 * cross + exclude.n2
            centroid = dot / (vn * math.sqrt(n2)) if dot > 0 and n2 > 1e-9 and vn else 0.0
            nearest, nearest_doc, name_r, name_doc, seq_doc = 0.0, None, 0.0, None, None
            for d in c.docs:
                if d is exclude:
                    continue
                if vn and d.n2:
                    ddot = sum(w * d.vec.get(t, 0.0) for t, w in vec.items())
                    if ddot > 0:
                        cs = ddot / (vn * math.sqrt(d.n2))
                        if cs > nearest:
                            nearest, nearest_doc = cs, d
                if norm and d.norm:
                    sm = difflib.SequenceMatcher(None, norm, d.norm)
                    if sm.real_quick_ratio() > name_r and sm.quick_ratio() > name_r:
                        r = sm.ratio()
                        if r > name_r:
                            name_r, name_doc = r, d
                if seq and d.seq and d.kind == kind and d.seq[0] == seq[0] and 0 < abs(d.seq[1] - seq[1]) <= 5:
                    seq_doc = d
            name_c = clamp((name_r - 0.55) / 0.4)
            text_c = 0.85 * clamp(max(centroid, nearest) / 0.45)
            seq_c = 0.92 if seq_doc else 0.0
            total = 1 - (1 - name_c) * (1 - text_c) * (1 - seq_c)
            strongest = max(name_c, text_c, seq_c)
            if strongest == 0:
                reason = "nothing in common"
            elif strongest == seq_c:
                reason = f'next in sequence after "{seq_doc.name}"'
            elif strongest == name_c:
                reason = f'name almost matches "{name_doc.name}"'
            else:
                ref = nearest_doc.vec if (nearest >= centroid and nearest_doc) else c.total
                shared = sorted(vec, key=lambda t: -vec[t] * ref.get(t, 0.0))[:3]
                shared = [t for t in shared if ref.get(t, 0.0) > 0]
                where = f'like "{nearest_doc.name}"' if nearest >= centroid and nearest_doc else "fits this folder"
                reason = f"{where} (words: {', '.join(shared)})"
            results.append((total, c, reason))
        results.sort(key=lambda r: -r[0])
        return results


def best_and_margin(scores):
    if not scores:
        return None, 0.0
    best = scores[0]
    second = next((s for s in scores[1:] if not related(s[1], best[1])), None)
    return best, best[0] - (second[0] if second else 0.0)


# ============================================================================ 2. RULES + 4. TYPE

@dataclass
class Suggestion:
    action: str                 # "move" | "junk" | "unsure"
    dest: Candidate | None
    confidence: float
    reason: str
    source: str                 # "rule" | "match" | "type" | "AI" | "-"
    fallback: "Suggestion | None" = None   # used if the AI can't decide (e.g. a picture -> Pictures)


SWITCH_CAPTURE = re.compile(r"^\d{16}(_[sc]|-[0-9A-Fa-f]{32})$")
SCREENSHOT = re.compile(r"^(screenshot|screen shot|capture|snip|snapshot)", re.I)
GAME_URL = re.compile(r"^URL=(steam://|com\.epicgames|roblox|battlenet://|origin2?://|uplay://|"
                      r"com\.ea\.|heroic://|itch://|xbox)", re.I | re.M)
GAME_PATH = re.compile(r"\\(steamapps|Epic Games|Riot Games|Roblox|GOG Galaxy|GOG Games|Ubisoft Game Launcher|"
                       r"EA Games|Electronic Arts|Battle\.net|XboxGames|itch|Minecraft Launcher|Rockstar Games)\\", re.I)


class Context:
    def __init__(self, cfg, cands, matcher, dups, programs):
        self.cfg, self.cands, self.matcher, self.dups, self.programs = cfg, cands, matcher, dups, programs

    def cand(self, key: str) -> Candidate | None:
        return next((c for c in self.cands if c.key.lower() == key.lower()), None)

    def games(self) -> Candidate | None:
        return next((c for c in self.cands if not c.library and "\\" not in c.key and "game" in c.key.lower()), None)


def rel(p: Path | str) -> str:
    s = str(p)
    for base, label in ((r"C:\Users\Public\Desktop" + "\\", "Desktop\\"), (r"C:\Users\Public" + "\\", "")):
        if s.lower().startswith(base.lower()):
            return label + s[len(base):]
    return s


def rules(item: Item, ctx: Context) -> Suggestion | None:
    name, ext = item.name, item.path.suffix.lower()
    if ext in PARTIAL:
        return Suggestion("junk", None, 0.99, "half-finished download / temp file", "rule")
    if name.startswith(("~$", ".~lock.")):
        return Suggestion("junk", None, 0.99, "leftover lock file from an office app", "rule")
    if item.is_dir and not any(item.path.iterdir()):
        return Suggestion("junk", None, 0.99, "empty folder", "rule")
    if not item.is_dir and item.size == 0:
        return Suggestion("junk", None, 0.95, "empty file (0 bytes)", "rule")
    copy = ctx.dups.copy_of(item)
    if copy:
        return Suggestion("junk", None, 0.99, f"exact copy of {rel(copy)}", "rule")
    if item.kind == "installer":
        found = match_installed(item, ctx.programs)
        setups = ctx.cand("Program installers")
        if found:
            prog, have, new = found
            hv, nv = version_tuple(have), version_tuple(new)
            if hv and nv and nv > hv:
                if setups:
                    return Suggestion("move", setups, 0.8, f"newer than your {prog} {have} - run it to update", "rule")
            elif hv and nv:
                return Suggestion("junk", None, 0.9, f"installer for {prog} {new}, you already have {have}", "rule")
            else:
                return Suggestion("junk", None, 0.8, f"installer for {prog}, which looks already installed", "rule")
        elif setups:
            return Suggestion("move", setups, 0.75, "installer for a program that isn't installed", "rule")
    if item.kind == "archive":
        roots = [item.path.parent] + [Path(r) for r in ctx.cfg["duplicate_search"]]
        for root in roots:
            unzipped = root / item.stem
            if unzipped.is_dir():
                return Suggestion("junk", None, 0.85, f"already unzipped: {rel(unzipped)}", "rule")
    games = ctx.games()
    if games and ext == ".url" and GAME_URL.search("URL=" + item.facts.meta.get("link", "")):
        return Suggestion("move", games, 0.95, "game launcher shortcut", "rule")
    if games and ext == ".lnk" and GAME_PATH.search(item.facts.meta.get("points to", "")):
        return Suggestion("move", games, 0.95, f"shortcut to a game ({short(item.facts.meta['points to'], 50)})", "rule")
    return None


def type_default(item: Item, ctx: Context) -> Suggestion | None:
    meta = item.facts.meta
    switch = ctx.cand("Pictures\\Nintendo Switch")
    if switch and item.kind in ("image", "video") and (SWITCH_CAPTURE.match(item.stem) or "nintendo" in
                                                        (meta.get("camera", "") + meta.get("device", "")).lower()):
        return Suggestion("move", switch, 0.9, f"Nintendo Switch {'capture' if item.kind == 'image' else 'clip'}", "type")
    shots = ctx.cand("Pictures\\Screenshots")
    if shots and item.kind == "image" and (SCREENSHOT.match(item.stem) or meta.get("looks like") == "screenshot"):
        return Suggestion("move", shots, 0.85, "screenshot", "type")
    for kind, key, what in (("image", "Pictures", "a picture"), ("video", "Videos", "a video"),
                            ("audio", "Music", "an audio file")):
        if item.kind == kind and ctx.cand(key):
            extra = f" ({meta.get('camera') or meta.get('artist')})" if meta.get("camera") or meta.get("artist") else ""
            return Suggestion("move", ctx.cand(key), 0.75, f"it's {what}{extra} and nothing else matched better", "type")
    return None


def decide(item: Item, ctx: Context) -> Suggestion:
    rule = rules(item, ctx)
    if rule:
        return rule
    item.scores = ctx.matcher.score(item.raw, item.norm, item.seq, item.kind)
    best, margin = best_and_margin(item.scores)
    th = ctx.cfg["thresholds"]
    if best and best[0] >= th["confident"] and margin >= th["margin"]:
        return Suggestion("move", best[1], best[0], best[2], "match")
    typed = type_default(item, ctx)
    # A picture with real text in it (a screenshot of something) deserves a closer look than "it's a picture"
    wordy_picture = item.kind == "image" and len(item.facts.text.strip()) >= 60
    if typed and not wordy_picture:
        return typed
    # "likely": weaker but clearly ahead - right ~9 times in 10 in the self-test. Used if the AI can't decide.
    likely = None
    if best and best[0] >= th.get("likely", 0.25) and margin >= th.get("likely_margin", 0.15):
        likely = Suggestion("move", best[1], best[0], best[2], "likely")
    top = f" (closest: {best[1].key}, {best[0]:.0%})" if best and best[0] > 0.2 else ""
    return Suggestion("unsure", None, best[0] if best else 0.0, "no strong match" + top, "-",
                      fallback=likely or typed)


# ---- groups: many similar files (e.g. Hours34(Monday).png ... Hours61(Thursday).png) get ONE AI question

_DAYS_MONTHS = ("monday|tuesday|wednesday|thursday|friday|saturday|sunday|january|february|march|april|may|june|"
                "july|august|september|october|november|december")


def group_key(item: Item):
    s = re.sub(rf"\b({_DAYS_MONTHS})\b", " ", item.stem.lower())
    s = re.sub(r"[^a-z]+", " ", s).strip()
    return (s, item.kind) if len(s) >= 3 else None


def ai_queue(unsure: list) -> list:
    """[(representative (item, sug), members [(item, sug)...])] - similar files are asked about once."""
    by_key, grouped, queue = defaultdict(list), set(), []
    for pair in unsure:
        key = group_key(pair[0])
        if key:
            by_key[key].append(pair)
    for members in by_key.values():
        if len(members) < 3:
            continue
        rep = max(members, key=lambda m: len(m[0].facts.text))
        rep_words = set(tokens(rep[0].facts.text))
        similar = [m for m in members if m is rep or not rep_words or
                   len(rep_words & set(tokens(m[0].facts.text))) / max(len(rep_words | set(tokens(m[0].facts.text))), 1) >= 0.3]
        if len(similar) >= 3:
            queue.append((rep, similar))
            grouped.update(id(m[1]) for m in similar)
    queue += [(p, [p]) for p in unsure if id(p[1]) not in grouped]
    return queue


# ============================================================================ 5. AI (only for the leftovers)

# Ollama's port isn't opened to Windows by the Jarvis setup, and "ollama run" mistakes folder paths
# in a prompt for image files. So we reach Ollama's normal web API from *inside* its container,
# with a tiny raw-socket HTTP client in Perl (the only scripting language that image ships).
PERL_POST = (
    'use IO::Socket::INET; local $/; my $b = <STDIN>;'
    'my $s = IO::Socket::INET->new(PeerAddr => "127.0.0.1", PeerPort => 11434, Proto => "tcp", Timeout => 10)'
    '  or die "connect failed: $!";'
    'binmode $s; print $s "POST /api/chat HTTP/1.0\\r\\nHost: localhost\\r\\nContent-Type: application/json\\r\\n"'
    '  . "Content-Length: " . length($b) . "\\r\\n\\r\\n" . $b;'
    'my $r = <$s>; $r =~ s/^.*?\\r\\n\\r\\n//s; print $r;'
)
LETTERS = "ABCDEFGHI"


class LocalAI:
    def __init__(self, cfg: dict):
        self.cfg = cfg["ai"]
        self.model = self.cfg["model"]
        self.mode: str | None = None
        self.note = ""
        self.last_error = ""

    def available(self) -> bool:
        if not self.cfg.get("enabled", True):
            self.note = "AI step is turned off in config.json"
            return False
        try:
            with urllib.request.urlopen(self.cfg["ollama_url"].rstrip("/") + "/api/tags", timeout=2) as r:
                names = [m["name"] for m in json.load(r).get("models", [])]
            if self.model in names:
                self.mode = "http"
                return True
        except Exception:
            pass
        ctr = self.cfg.get("docker_container")
        if ctr and shutil.which("docker"):
            try:
                r = subprocess.run(["docker", "exec", ctr, "ollama", "list"], capture_output=True, text=True,
                                   timeout=30, encoding="utf-8", errors="replace")
                if r.returncode == 0:
                    names = [ln.split()[0] for ln in r.stdout.splitlines()[1:] if ln.strip()]
                    if self.model in names:
                        self.mode = "docker"
                        return True
                    self.note = f'model "{self.model}" isn\'t installed in {ctr}'
                    return False
            except Exception:
                pass
        self.note = "local AI isn't running (start Docker Desktop / Jarvis for smarter suggestions)"
        return False

    def ask(self, prompt: str, allowed: list[str]) -> tuple[str, str] | None:
        """Ollama must answer with one of `allowed` - enforced by a JSON schema, then checked again here."""
        timeout = self.cfg.get("timeout_seconds", 240)
        # "reason" comes first on purpose: the model writes its thinking before it commits to a choice
        schema = {"type": "object", "required": ["reason", "choice"],
                  "properties": {"reason": {"type": "string"}, "choice": {"type": "string", "enum": allowed}}}
        body = json.dumps({"model": self.model, "stream": False, "think": False, "format": schema,
                           "messages": [{"role": "user", "content": prompt}],
                           "options": {"temperature": 0}, "keep_alive": "10m"})
        self.last_error = ""
        try:
            if self.mode == "http":
                req = urllib.request.Request(self.cfg["ollama_url"].rstrip("/") + "/api/chat",
                                             data=body.encode("utf-8"), headers={"Content-Type": "application/json"})
                with urllib.request.urlopen(req, timeout=timeout) as r:
                    reply = r.read().decode("utf-8", "replace")
            else:
                r = subprocess.run(["docker", "exec", "-i", self.cfg["docker_container"], "perl", "-e", PERL_POST],
                                   input=body.encode("utf-8"), capture_output=True, timeout=timeout)
                reply = r.stdout.decode("utf-8", "replace")
                if r.returncode:
                    self.last_error = r.stderr.decode("utf-8", "replace").strip()[:120]
            data = json.loads(reply)
            if "error" in data:
                self.last_error = str(data["error"])[:120]
                return None
            answer = json.loads(data["message"]["content"])
        except Exception as e:
            self.last_error = self.last_error or f"{type(e).__name__}: {e}"[:120]
            return None
        choice = str(answer.get("choice", "")).strip().upper()
        reason = re.sub(r"\s+", " ", str(answer.get("reason", ""))).strip()[:90]
        return (choice, reason) if choice in allowed else None


def ai_options(item: Item, ctx: Context) -> list[Candidate]:
    desk = [c for c in ctx.cands if not c.library]
    best = item.scores[0][0] if item.scores else 0.0
    if best < 0.3:
        options = desk
    else:
        options = [s[1] for s in item.scores[:3]]
        archive = next((c for c in desk if "\\" not in c.key and "archive" in c.key.lower()), None)
        if archive and archive not in options:
            options.append(archive)
    return options[:9]


def ai_prompt(item: Item, options: list[Candidate], siblings: list[Item] = ()) -> str:
    size = human_size(item.size)
    when = datetime.fromtimestamp(item.mtime).strftime("%Y-%m-%d")
    excerpt = re.sub(r"\s+", " ", item.facts.text).strip()[:700]
    label = "text found in the picture" if item.kind == "image" else "start of its content"
    facts = facts_line(item.facts)
    lines = [
        "You help sort ONE item from a desktop Inbox into the best folder.",
        "First think about what the item is and what it's for, then pick exactly one option letter.",
        "If none clearly fits, answer UNSURE.",
        "Answer JUNK only if the item is clearly useless (temporary, broken, leftover).",
        "",
        "ITEM",
        f"name: {item.name}",
        f"type: {item.kind}, {size}, last changed {when}",
    ]
    if facts:
        lines.append(f"facts: {facts}")
    lines.append(f"{label}: {excerpt}" if excerpt else "(no readable text inside)")
    if len(siblings) > 1:
        names = ", ".join(s.name for s in siblings[:6])
        lines.append(f"This is one of {len(siblings)} similar files that belong together: {names}...")
    lines += ["", "OPTIONS"]
    for letter, c in zip(LETTERS, options):
        recent = sorted((d for d in c.docs if d.norm), key=lambda d: -d.mtime)[:4]
        examples = ", ".join(d.name for d in recent) or "(nothing yet)"
        desc = f" - {c.desc}" if c.desc else ""
        lines.append(f"{letter}. {c.key}{desc}. Examples inside: {examples}")
    lines += ["JUNK - safe to throw away", "UNSURE - leave it for the human to decide", "",
              'Reply with JSON only: {"reason": "<max 15 words>", "choice": "<option letter, JUNK or UNSURE>"}']
    return "\n".join(lines)


# ============================================================================ 6. YOU (screen + moving)

class Col:
    R, DIM, B, G, Y, CY, RED = "\x1b[0m", "\x1b[90m", "\x1b[1m", "\x1b[32m", "\x1b[33m", "\x1b[36m", "\x1b[31m"


def human_size(n: float) -> str:
    for unit in ("bytes", "KB", "MB", "GB"):
        if n < 1024 or unit == "GB":
            return f"{n:.0f} {unit}" if unit == "bytes" else f"{n:.1f} {unit}"
        n /= 1024
    return f"{n} bytes"


def short(s: str, width: int) -> str:
    return s if len(s) <= width else s[: width - 3] + "..."


def getkey(valid: str) -> str:
    valid = valid.upper()
    if sys.stdin.isatty():
        import msvcrt
        while True:
            ch = msvcrt.getwch()
            if ch in ("\x00", "\xe0"):
                msvcrt.getwch()
                continue
            if ch == "\x03":
                raise KeyboardInterrupt
            ch = ch.upper()
            if ch in valid:
                print(ch)
                return ch
    for line in sys.stdin:  # piped input (tests/scripts): one key per line, anything invalid is skipped
        ch = line.replace("\ufeff", "").strip()[:1].upper()
        if ch and ch in valid:
            print(ch)
            return ch
    return "Q"  # ran out of input: stop safely


def pause(msg="  Press any key to close."):
    if sys.stdin.isatty():
        import msvcrt
        print(Col.DIM + msg + Col.R)
        msvcrt.getwch()


def tag(s: Suggestion) -> str:
    if s.source in ("match", "likely"):
        return f"{s.source} {s.confidence:.0%}"
    return s.source


def print_plan(pairs):
    groups = [("junk", f"{Col.RED}JUNK?{Col.R} {Col.DIM}(goes to the Recycle Bin - you can restore it from there){Col.R}"),
              ("move", f"{Col.G}MOVE{Col.R}"),
              ("unsure", f"{Col.Y}NOT SURE{Col.R} {Col.DIM}(stays in the Inbox unless you pick a folder){Col.R}")]
    n = 0
    for action, title in groups:
        rows = [(i, s) for i, s in pairs if s.action == action]
        if not rows:
            continue
        print(f"\n  {title}")
        for item, sug in rows:
            n += 1
            label = short(item.name, 40).ljust(40)
            if action == "move":
                print(f"  {n:>2}  {label} -> {Col.B}{short(sug.dest.key, 38)}{Col.R}")
                print(f"      {Col.DIM}{tag(sug)}: {short(sug.reason, 80)}{Col.R}")
            else:
                print(f"  {n:>2}  {label} {Col.DIM}{tag(sug)}: {short(sug.reason, 60)}{Col.R}")


def ordered(pairs):
    order = {"junk": 0, "move": 1, "unsure": 2}
    return sorted(pairs, key=lambda p: order[p[1].action])


def other_options(item: Item, sug: Suggestion, ctx: Context) -> list[Candidate]:
    out = []
    for _, c, _ in item.scores:
        if c is not sug.dest and c not in out:
            out.append(c)
        if len(out) == 3:
            break
    if not out:  # rules / type decisions have no scores - offer the main desktop folders
        out = [c for c in ctx.cands if not c.library and "\\" not in c.key and c is not sug.dest][:3]
    return out


def review(pairs, ctx: Context):
    decisions = []
    total = len(pairs)
    for n, (item, sug) in enumerate(pairs, 1):
        opts = other_options(item, sug, ctx)
        info = f"{item.kind}, {human_size(item.size)}, changed {datetime.fromtimestamp(item.mtime):%Y-%m-%d}"
        print(f"\n  {Col.B}({n}/{total})  {item.name}{Col.R}   {Col.DIM}{info}{Col.R}")
        facts = facts_line(item.facts, 110)
        if facts:
            print(f"         {Col.DIM}{facts}{Col.R}")
        if sug.action == "move":
            print(f"         suggestion: {Col.G}move to {sug.dest.key}{Col.R}   {Col.DIM}{tag(sug)}: {sug.reason}{Col.R}")
        elif sug.action == "junk":
            print(f"         suggestion: {Col.RED}Recycle Bin{Col.R}   {Col.DIM}{tag(sug)}: {sug.reason}{Col.R}")
        else:
            print(f"         {Col.Y}not sure{Col.R}   {Col.DIM}{sug.reason}{Col.R}")
        print("         other folders:  " + "   ".join(f"[{i}] {short(c.key, 30)}" for i, c in enumerate(opts, 1)))
        keys = "".join(str(i) for i in range(1, len(opts) + 1))
        yes = "[Y] yes   " if sug.action != "unsure" else ""
        while True:
            print(f"  {yes}[N] leave it   [{keys[:1]}-{keys[-1:]}] other folder   [J] junk   [O] open it   [Q] stop")
            k = getkey(("Y" if yes else "") + "NJOQ" + keys)
            if k == "O":
                os.startfile(item.path)
                continue
            break
        if k == "Q":
            break
        if k == "Y":
            decisions.append((item, sug.action, sug.dest, sug.reason))
        elif k == "J":
            decisions.append((item, "junk", None, "you chose junk"))
        elif k in keys:
            decisions.append((item, "move", opts[int(k) - 1], "you picked it"))
    return decisions


def recycle(path: Path) -> bool:
    class SHFILEOPSTRUCTW(ctypes.Structure):
        _fields_ = [("hwnd", wintypes.HWND), ("wFunc", wintypes.UINT), ("pFrom", wintypes.LPCWSTR),
                    ("pTo", wintypes.LPCWSTR), ("fFlags", ctypes.c_uint16), ("fAnyOperationsAborted", wintypes.BOOL),
                    ("hNameMappings", ctypes.c_void_p), ("lpszProgressTitle", wintypes.LPCWSTR)]
    FO_DELETE, FOF_SILENT, FOF_NOCONFIRMATION, FOF_ALLOWUNDO, FOF_NOERRORUI = 3, 0x4, 0x10, 0x40, 0x400
    op = SHFILEOPSTRUCTW(None, FO_DELETE, str(path) + "\0", None,
                         FOF_SILENT | FOF_NOCONFIRMATION | FOF_ALLOWUNDO | FOF_NOERRORUI, False, None, None)
    rc = ctypes.windll.shell32.SHFileOperationW(ctypes.byref(op))
    return rc == 0 and not op.fAnyOperationsAborted and not path.exists()


def unique_path(folder: Path, name: str, is_dir: bool) -> Path:
    target = folder / name
    stem, ext = (name, "") if is_dir else (Path(name).stem, Path(name).suffix)
    i = 2
    while target.exists():
        target = folder / f"{stem} ({i}){ext}"
        i += 1
    return target


def log_rows(rows):
    new = not LOG_PATH.exists()
    with LOG_PATH.open("a", newline="", encoding="utf-8-sig" if new else "utf-8") as f:
        w = csv.writer(f)
        if new:
            w.writerow(["run", "time", "action", "from", "to", "why"])
        w.writerows(rows)


def apply(decisions) -> Counter:
    run = datetime.now().strftime("%Y%m%d-%H%M%S")
    now = datetime.now().strftime("%Y-%m-%d %H:%M")
    done, rows = Counter(), []
    for item, action, dest, why in decisions:
        try:
            if action == "move":
                dest.path.mkdir(parents=True, exist_ok=True)
                target = unique_path(dest.path, item.name, item.is_dir)
                shutil.move(str(item.path), str(target))
                rows.append([run, now, "moved", str(item.path), str(target), why])
                done["moved"] += 1
            elif action == "junk":
                if recycle(item.path):
                    rows.append([run, now, "recycled", str(item.path), "Recycle Bin", why])
                    done["recycled"] += 1
                else:
                    done["failed"] += 1
        except OSError:
            done["failed"] += 1
            print(f"  {Col.Y}Couldn't move {item.name} (open in another app?){Col.R}")
    log_rows(rows)
    return done


def undo_last():
    if not LOG_PATH.exists():
        print("  Nothing to undo yet.")
        return
    with LOG_PATH.open(encoding="utf-8-sig") as f:
        rows = list(csv.DictReader(f))
    undone = {r["run"] for r in rows if r["action"] == "undo"}
    run = next((r["run"] for r in reversed(rows) if r["action"] in ("moved", "recycled") and r["run"] not in undone), None)
    if not run:
        print("  Nothing to undo.")
        return
    back, problems = 0, 0
    for r in reversed([r for r in rows if r["run"] == run and r["action"] == "moved"]):
        src, dst = Path(r["to"]), Path(r["from"])
        if src.exists() and not dst.exists():
            dst.parent.mkdir(parents=True, exist_ok=True)
            shutil.move(str(src), str(dst))
            back += 1
        else:
            problems += 1
    recycled = [r for r in rows if r["run"] == run and r["action"] == "recycled"]
    log_rows([[run, datetime.now().strftime("%Y-%m-%d %H:%M"), "undo", "", "", f"{back} moved back"]])
    when = datetime.strptime(run, "%Y%m%d-%H%M%S").strftime("%Y-%m-%d %H:%M")
    print(f"\n  {Col.G}Undid the sort from {when}: {back} item(s) moved back to the Inbox.{Col.R}")
    if problems:
        print(f"  {Col.Y}{problems} item(s) had been moved again since, so they were left where they are.{Col.R}")
    if recycled:
        print(f"  {len(recycled)} item(s) went to the Recycle Bin in that sort - open the Recycle Bin,")
        print("  right-click them > Restore, if you want them back:")
        for r in recycled:
            print(f"    - {Path(r['from']).name}")


# ============================================================================ self-test

def selftest(cands: list[Candidate], matcher: Matcher, cfg: dict):
    th = cfg["thresholds"]
    per: dict[str, Counter] = defaultdict(Counter)
    mistakes = []
    for c in cands:
        if c.library:
            continue
        top = c.key.split("\\")[0]
        for d in c.docs:
            best, margin = best_and_margin(matcher.score(d.raw, d.norm, d.seq, d.kind, exclude=d))
            if not best:
                continue
            s = per[top]
            s["files"] += 1
            s["best_ok"] += related(best[1], c)
            if best[0] >= th["confident"] and margin >= th["margin"]:
                s["confident"] += 1
                s["confident_ok"] += related(best[1], c)
                if not related(best[1], c):
                    mistakes.append((d.name, c.key, best[1].key, best[0]))
    total = sum(per.values(), Counter())
    print(f"\n  SELF-TEST: each of your {total['files']} existing files is hidden from the matcher in turn,")
    print("  then treated as if it were new. Where would it go?  (right = its folder, or that folder's parent/subfolder)\n")
    print(f"  {'folder':36} {'files':>5}  {'decided by math':>15}  {'...right':>9}  {'best guess right':>16}")
    for top in sorted(per):
        s = per[top]
        print(f"  {short(top, 36):36} {s['files']:>5}  {s['confident']:>15}  {s['confident_ok']:>4} of {s['confident']:<3}"
              f"  {s['best_ok'] / s['files']:>15.0%}")
    t = total
    print(f"  {'ALL':36} {t['files']:>5}  {t['confident']:>15}  {t['confident_ok']:>4} of {t['confident']:<3}"
          f"  {t['best_ok'] / max(t['files'], 1):>15.0%}")
    print(f"\n  {t['files'] - t['confident']} file(s) weren't clear-cut -> those would go to the AI step / you.")
    if mistakes:
        print("\n  Decided by math but WRONG:")
        for name, truth, guess, score in mistakes[:15]:
            print(f"    {short(name, 45):45}  is in {short(truth, 28):28} guessed {short(guess, 28)} ({score:.0%})")


# ============================================================================ main

def visible(folder: Path) -> list[Path]:
    return sorted(p for p in folder.iterdir() if not is_hidden(p) and p.name.lower() != "desktop.ini")


def open_up_folders(top: list[Path], ask: bool, open_all: bool) -> list[Path]:
    """A folder in the Inbox is one item - unless you choose to sort what's inside it (one level deep)."""
    paths = []
    for p in top:
        inside = visible(p) if p.is_dir() else []
        if inside:
            if open_all:
                paths.extend(inside)
                continue
            if ask:
                print(f'\n  The folder "{p.name}" has {len(inside)} thing(s) in it.')
                print(f"  {Col.CY}[Y]{Col.R} sort what's inside it one by one   {Col.CY}[N]{Col.R} treat it as one folder")
                if getkey("YN") == "Y":
                    paths.extend(inside)
                    continue
        paths.append(p)
    return paths


def main():
    ap = argparse.ArgumentParser(description="Suggests where Inbox items go; moves them when you say yes.")
    ap.add_argument("--plan", action="store_true", help="print suggestions only, change nothing")
    ap.add_argument("--ai", action="store_true", help="with --plan: also ask the local AI")
    ap.add_argument("--no-ai", action="store_true", help="skip the AI step")
    ap.add_argument("--open-folders", action="store_true", help="sort what's inside Inbox folders without asking")
    ap.add_argument("--selftest", action="store_true", help="measure matching quality on your folders")
    ap.add_argument("--undo", action="store_true", help="undo the last sort")
    args = ap.parse_args()

    os.system("")  # turns on colors in the Windows console
    if not sys.stdout.isatty():
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    os.system("title Sort my Inbox")

    if args.undo:
        undo_last()
        pause()
        return

    cfg = load_config()
    IGNORE.update(w.lower() for w in cfg["ignore_words"])
    inbox = Path(cfg["desktop"]) / cfg["inbox"]
    inbox.mkdir(exist_ok=True)
    store = make_store(cfg)
    say = lambda msg: print(f"  {Col.DIM}{msg}{Col.R}", flush=True)
    print(f"\n  {Col.CY}{Col.B}SORT MY INBOX{Col.R}")

    if args.selftest:
        cands = build_candidates(cfg, store, say)
        store.save()
        selftest(cands, Matcher(cands), cfg)
        return

    paths = open_up_folders(visible(inbox), ask=not args.plan, open_all=args.open_folders)
    if not paths:
        print(f"\n  {Col.G}The Inbox is empty. Nothing to sort.{Col.R}")
        if sys.stdin.isatty():
            print(f"\n  {Col.DIM}[U] undo the last sort   any other key: close{Col.R}")
            import msvcrt
            if msvcrt.getwch().upper() == "U":
                undo_last()
                pause()
        return

    say(f"Looking at {len(paths)} thing(s)... (reading folders, checking duplicates)")
    cands = build_candidates(cfg, store, say)
    matcher = Matcher(cands)
    dups = DupIndex(cfg["duplicate_search"], inbox)
    ctx = Context(cfg, cands, matcher, dups, installed_programs())
    inbox_facts = store.get_many([p for p in paths if p.is_file()], say)
    items = [make_item(p, inbox_facts.get(str(p))) for p in paths]
    store.save()
    pairs = [(it, decide(it, ctx)) for it in items]

    # ---- AI step, only for what the code couldn't settle
    need_ai = [(i, s) for i, s in pairs if s.action == "unsure"]
    use_ai = need_ai and not args.no_ai and (args.ai or not args.plan)
    if use_ai:
        ai = LocalAI(cfg)
        if ai.available():
            queue = ai_queue(need_ai)
            limit = cfg["ai"].get("max_items_per_run", 8)
            todo = queue[:limit]
            say(f"Asking the local AI about {len(todo)} thing(s) the rules weren't sure about"
                f" (similar files are asked about together; ~30 s each on this laptop)...")
            failures = 0
            for n, ((item, sug), members) in enumerate(todo, 1):
                options = ai_options(item, ctx)
                allowed = list(LETTERS[:len(options)]) + ["JUNK", "UNSURE"]
                start = time.time()
                group = f" (+{len(members) - 1} similar)" if len(members) > 1 else ""
                print(f"  {Col.DIM}  ({n}/{len(todo)}) {short(item.name, 45)}{group} ...{Col.R}", end="", flush=True)
                answer = ai.ask(ai_prompt(item, options, [m[0] for m in members]), allowed)
                print(f"{Col.DIM} {time.time() - start:.0f}s{Col.R}")
                if not answer:
                    failures += 1
                    for _, s in members:
                        s.reason += " - AI gave no usable answer" + (f" ({ai.last_error})" if ai.last_error else "")
                    if failures >= 2:  # the AI is struggling (e.g. Jarvis's Ollama restarting): stop asking
                        say("The local AI stopped answering properly, so the rest stay 'not sure' this time.")
                        break
                    continue
                failures = 0
                choice, why = answer
                for m_item, s in members:
                    also = f' (same kind of file as "{item.name}")' if m_item is not item else ""
                    if choice == "JUNK":
                        s.action, s.dest, s.source, s.reason = "junk", None, "AI", (why or "AI thinks it's junk") + also
                    elif choice == "UNSURE":
                        s.source, s.reason = "AI", f"AI isn't sure either{': ' + why if why else ''}"
                    else:
                        s.action, s.dest, s.source = "move", options[LETTERS.index(choice)], "AI"
                        s.reason = (why or "AI's pick") + also
            if len(queue) > limit:
                say(f"{len(queue) - limit} more weren't asked this time (limit {limit} per run) - they stay 'not sure'.")
        else:
            for _, sug in need_ai:
                sug.reason += f" - {ai.note}"
    # anything still unsure that has a safe default (a picture -> Pictures) gets that default
    for n, (item, sug) in enumerate(pairs):
        if sug.action == "unsure" and sug.fallback:
            sug.fallback.reason += " (AI didn't decide)" if use_ai else ""
            pairs[n] = (item, sug.fallback)

    pairs = ordered(pairs)
    print_plan(pairs)
    if args.plan:
        return

    print(f"\n  {Col.CY}[A]{Col.R} do all of the above   {Col.CY}[R]{Col.R} review one by one   "
          f"{Col.CY}[U]{Col.R} undo last sort   {Col.CY}[Q]{Col.R} quit, change nothing")
    k = getkey("ARUQ")
    if k == "Q":
        print("  Nothing changed.")
        pause()
        return
    if k == "U":
        undo_last()
        pause()
        return
    if k == "A":
        decisions = [(i, s.action, s.dest, s.reason) for i, s in pairs if s.action in ("move", "junk")]
    elif k == "R":
        decisions = review(pairs, ctx)
    else:
        return

    done = apply(decisions)
    left = len([p for p in inbox.iterdir() if not is_hidden(p)])
    parts = [f"{done['moved']} moved"] + ([f"{done['recycled']} sent to the Recycle Bin"] if done["recycled"] else []) \
        + ([f"{done['failed']} couldn't be moved"] if done["failed"] else [])
    print(f"\n  {Col.G}Done: {', '.join(parts)}. {left} left in the Inbox.{Col.R}")
    print(f"  {Col.DIM}Changed your mind? Open Sort my Inbox again and press U.{Col.R}")
    pause()


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n  Stopped. Nothing else was changed.")
