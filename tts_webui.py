#!/usr/bin/env python3
"""
tts_webui.py: DOCX/TXT hochladen, MP3 bekommen. Start in der Conda-Umgebung: python tts_webui.py
Startet Fish Audio Drama 3 (SGLang-Omni) bei Bedarf selbst, Merge und Encoding über ffmpeg.

Einstellungen stehen in config.toml, jede lässt sich per TTS_*-Umgebungsvariable überschreiben.
"""
from __future__ import annotations

import atexit
import http.client
import json
import math
import os
import queue
import re
import secrets
import shlex
import shutil
import signal
import site
import socket
import subprocess
import sys
import threading
import time
import tomllib
import unicodedata
import urllib.error
import urllib.parse
import urllib.request
import wave
import zipfile
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path
from xml.etree import ElementTree as ET

from flask import Flask, abort, jsonify, render_template, request, send_file


# ---------------------------------------------------------------- CONFIG
BASE = Path(__file__).resolve().parent


def _load_config():
    path = BASE / os.environ.get("TTS_CONFIG", "config.toml")
    if not path.is_file():
        return {}
    try:
        with path.open("rb") as handle:
            return tomllib.load(handle)
    except tomllib.TOMLDecodeError as exc:
        sys.exit(f"{path}: {exc}")


_CONFIG = _load_config()
_KNOWN_KEYS = set()


def _env(name, default, cast=str):
    """Reihenfolge: Umgebungsvariable TTS_X, dann Schlüssel x in config.toml, dann Default."""
    key = name.removeprefix("TTS_").lower()
    _KNOWN_KEYS.add(key)
    value = os.environ.get(name)
    if value in (None, ""):
        value = _CONFIG.get(key)
    return cast(value) if value not in (None, "") else default


HOST = _env("TTS_HOST", "0.0.0.0")
PORT = _env("TTS_PORT", 8090, int)
TITLE = _env("TTS_TITLE", "Audilekt")
DATA = BASE / _env("TTS_DATA_DIR", "data")
VOICES = BASE / _env("TTS_VOICES_DIR", "voices")
MAX_UPLOAD_MB = _env("TTS_MAX_UPLOAD_MB", 50, int)
KEEP_DAYS = _env("TTS_KEEP_DAYS", 30, int)

BACKEND_URL = _env("TTS_BACKEND_URL", "http://127.0.0.1:8010").rstrip("/")
AUTOSTART = _env("TTS_AUTOSTART", 1, int)                # Backend bei Bedarf selbst starten
MODEL_PATH = _env("TTS_MODEL", "./drama-3")              # lokaler Ordner oder Hugging-Face-Repo-ID
if (BASE / MODEL_PATH).is_dir():
    MODEL_PATH = str((BASE / MODEL_PATH).resolve())
ENGINE_MEMORY = _env("TTS_ENGINE_MEM", "0.75")           # Anteil des VRAM für die Engine
MAX_RUNNING = _env("TTS_MAX_RUNNING", "64")              # parallele Requests im Backend
BACKEND_ARGS = shlex.split(_env("TTS_BACKEND_ARGS", "")) # weitere Flags für sgl-omni serve
IDLE_STOP_SEC = _env("TTS_IDLE_STOP_SEC", 600, int)      # selbst gestartetes Backend danach beenden, 0 = nie
START_TIMEOUT_SEC = _env("TTS_START_TIMEOUT_SEC", 1200, int)
MIN_FREE_GIB = _env("TTS_MIN_FREE_GIB", 85.0, float)     # 0.75 Engine + 0.10 Vocoder + 6 GiB Reserve auf 96 GB

CONCURRENCY = _env("TTS_CONCURRENCY", 32, int)           # parallele Requests, höher = mehr GPU-Last
CHUNK_BYTES = _env("TTS_CHUNK_BYTES", 280, int)          # unter der internen Stückelung (~300 Byte)
MAX_DRAWS = _env("TTS_MAX_DRAWS", 3, int)                # Neuversuche je Abschnitt bei Stille/Loop
CPS = _env("TTS_CPS", 14.5, float)                       # Zeichen/s für die Längenschätzung
REQUEST_TIMEOUT_SEC = _env("TTS_REQUEST_TIMEOUT_SEC", 900, int)
# Leer = Server-Defaults (temperature 1.0, top_p 0.9, kein top_k). Laut SGLang-Omni-Quellcode
# lässt ein top_k von 30 einen Stream nach der ersten Phrase verstummen.
SAMPLING = {name: value for name, value in (("temperature", _env("TTS_TEMPERATURE", None, float)),
                                            ("top_p", _env("TTS_TOP_P", None, float)),
                                            ("top_k", _env("TTS_TOP_K", None, int))) if value is not None}

VOICE_URL = _env("TTS_VOICE_URL", "https://huggingface.co/datasets/zhaochenyang20/seed-tts-eval-mini/"
                                   "resolve/main/en/prompt-wavs/common_voice_en_10119832.wav")
VOICE_TEXT = "We asked over twenty different people, and they all said it was his."

READ_TABLES = _env("TTS_READ_TABLES", 1, int)            # DOCX-Tabellen zeilenweise vorlesen
ALLOW_TAGS = _env("TTS_ALLOW_TAGS", 0, int)              # 1 = [whisper] usw. an das Modell durchreichen
GAP_SENTENCE_MS = _env("TTS_GAP_SENTENCE_MS", 250, int)
GAP_PARAGRAPH_MS = _env("TTS_GAP_PARAGRAPH_MS", 650, int)
GAP_HEADING_MS = _env("TTS_GAP_HEADING_MS", 900, int)

FFMPEG = _env("TTS_FFMPEG", "/usr/bin/ffmpeg" if Path("/usr/bin/ffmpeg").exists() else "ffmpeg")
MP3_BITRATE = _env("TTS_MP3_BITRATE", "128k")
LOUDNORM = _env("TTS_LOUDNORM", 1, int)

for _key in sorted(_CONFIG.keys() - _KNOWN_KEYS):
    print(f"config.toml: unbekannter Schlüssel {_key!r} wird ignoriert", flush=True)

SAMPLE_RATE = 44100
FRAMES_PER_SEC = 21.5        # Codec-Frames je Sekunde, Einheit von max_new_tokens
MAX_FRAMES = 2048
JOBS_DIR = DATA / "jobs"
BACKEND_LOG = DATA / "backend.log"
ACTIVE = {"queued", "preparing", "waiting_gpu", "starting", "running", "encoding"}


class BackendError(RuntimeError):
    pass


class Cancelled(Exception):
    pass


# ---------------------------------------------------------------- Text aus DOCX/TXT
W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
_INLINE = {W + t for t in ("hyperlink", "ins", "moveTo", "smartTag", "sdt", "sdtContent",
                           "fldSimple", "customXml", "dir", "bdo")}
_HEADING = re.compile(r"(?i)^(heading|berschrift|title|titel|subtitle|untertitel)")
_FALSE = {"0", "false", "off"}


def _run_text(run, out):
    props = run.find(W + "rPr")
    if props is not None:
        hidden = props.find(W + "vanish")
        if hidden is not None and hidden.get(W + "val", "1") not in _FALSE:
            return
    for node in run:
        if node.tag == W + "t":
            out.append(node.text or "")
        elif node.tag in (W + "tab", W + "br", W + "cr"):
            out.append(" ")
        elif node.tag == W + "noBreakHyphen":
            out.append("-")


def _inline_text(parent, out):
    for node in parent:
        if node.tag == W + "r":
            _run_text(node, out)
        elif node.tag in _INLINE:
            _inline_text(node, out)


def _paragraph(par):
    out = []
    _inline_text(par, out)
    kind = "p"
    props = par.find(W + "pPr")
    if props is not None:
        style = props.find(W + "pStyle")
        if style is not None and _HEADING.match(style.get(W + "val", "")):
            kind = "h"
        if props.find(W + "outlineLvl") is not None:
            kind = "h"
    return "".join(out), kind


def _cell_text(cell):
    parts = []
    for node in cell:
        if node.tag == W + "p":
            parts.append(_paragraph(node)[0])
        elif node.tag == W + "tbl":
            parts.extend(_cell_text(c) for row in node.findall(W + "tr") for c in row.findall(W + "tc"))
    return " ".join(p.strip() for p in parts if p.strip())


def _block_items(parent, out):
    for node in parent:
        if node.tag == W + "p":
            out.append(_paragraph(node))
        elif node.tag == W + "tbl" and READ_TABLES:
            for row in node.findall(W + "tr"):
                cells = [t for t in (_cell_text(c) for c in row.findall(W + "tc")) if t]
                if cells:
                    out.append((", ".join(cells), "p"))
        elif node.tag in (W + "sdt", W + "sdtContent", W + "customXml"):
            _block_items(node, out)


def docx_paragraphs(path):
    with zipfile.ZipFile(path) as archive:
        info = archive.getinfo("word/document.xml")
        if info.file_size > 300 * 1024 * 1024:
            raise ValueError("document.xml ist zu groß")
        xml = archive.read(info)
    if b"<!DOCTYPE" in xml[:4096]:
        raise ValueError("DOCTYPE in document.xml")
    body = ET.fromstring(xml).find(W + "body")
    if body is None:
        raise ValueError("kein Dokumentkörper")
    out = []
    _block_items(body, out)
    return out


def _decode(raw):
    if raw[:3] == b"\xef\xbb\xbf":
        return raw[3:].decode("utf-8", "replace")
    if raw[:2] in (b"\xff\xfe", b"\xfe\xff"):
        return raw.decode("utf-16", "replace")
    try:
        return raw.decode("utf-8")
    except UnicodeDecodeError:
        return raw.decode("cp1252", "replace")


def txt_paragraphs(path):
    text = _decode(path.read_bytes()).replace("\r\n", "\n").replace("\r", "\n")
    text = re.sub(r"(\w)-\n(?=[a-zäöü])", r"\1", text)
    blocks = [b for b in re.split(r"\n[ \t]*\n+", text) if b.strip()]
    if len(blocks) == 1:
        lines = [l for l in blocks[0].split("\n") if l.strip()]
        longest = max((len(l) for l in lines), default=0)
        wrapped = (len(lines) > 5 and longest <= 120
                   and sum(len(l) >= 0.7 * longest for l in lines) / len(lines) > 0.6)
        if not wrapped:
            return [(l, "p") for l in lines]
    return [(b.replace("\n", " "), "p") for b in blocks]


_CHAR_MAP = str.maketrans({"\ufb00": "ff", "\ufb01": "fi", "\ufb02": "fl", "\ufb03": "ffi", "\ufb04": "ffl",
                           "\u00ad": "", "\u200b": "", "\u200c": "", "\u200d": "", "\ufeff": "",
                           "\u00a0": " ", "\u2009": " ", "\u202f": " "})


def clean(text):
    text = unicodedata.normalize("NFC", text).translate(_CHAR_MAP)
    text = re.sub(r"[\x00-\x08\x0e-\x1f]", " ", text)
    text = re.sub(r"<\|.*?\|>", " ", text)
    text = re.sub(r"https?://(?:www\.)?([^/\s]+)\S*", r"\1", text)
    if not ALLOW_TAGS:
        text = re.sub(r"\s*\[\d+(?:\s*[,;–-]\s*\d+)*\]", "", text)
        text = text.replace("[", "(").replace("]", ")")
    return re.sub(r"\s+", " ", text).strip()


_NOT_A_STOP = re.compile(
    r"(?:\b(?:Mr|Mrs|Ms|Dr|Prof|Sr|Jr|St|vs|etc|cf|al|Fig|Figs|Eq|Eqs|No|Nos|Vol|pp|ca|approx|"
    r"Inc|Ltd|Co|Ref|Refs|Sec|Ch|Dept|Univ|Jan|Feb|Mar|Apr|Jun|Jul|Aug|Sep|Sept|Oct|Nov|Dec)"
    r"|e\.g|i\.e|\b[A-Z])\.[\"”’»)]*\s*$")


def sentences(text):
    out, start = [], 0
    for match in re.finditer(r'[.!?…]+["”’»)]*\s+', text):
        piece = text[start:match.end()].strip()
        following = text[match.end():match.end() + 1]
        if _NOT_A_STOP.search(piece) or following.islower():
            continue
        out.append(piece)
        start = match.end()
    if start < len(text) and text[start:].strip():
        out.append(text[start:].strip())
    return out


def _size(text):
    return len(text.encode("utf-8"))


def _split_long(text, limit):
    if _size(text) <= limit:
        return [text]
    for pattern in (r"(?<=[;:])\s+", r"(?<=,)\s+", r"\s+(?=[–—-]\s)", r"\s+"):
        parts = re.split(pattern, text)
        if len(parts) < 2:
            continue
        packed, current = [], ""
        for part in parts:
            candidate = f"{current} {part}".strip()
            if current and _size(candidate) > limit:
                packed.append(current)
                current = part
            else:
                current = candidate
        packed.append(current)
        return [piece for item in packed for piece in _split_long(item, limit)]
    raw = text.encode("utf-8")
    return [raw[i:i + limit].decode("utf-8", "ignore") for i in range(0, len(raw), limit)]


def make_chunks(paragraphs):
    chunks = []
    for text, kind in paragraphs:
        pieces, current = [], ""
        for sentence in sentences(text):
            for fragment in _split_long(sentence, CHUNK_BYTES):
                candidate = f"{current} {fragment}".strip()
                if current and _size(candidate) > CHUNK_BYTES:
                    pieces.append(current)
                    current = fragment
                else:
                    current = candidate
        if current:
            pieces.append(current)
        for index, piece in enumerate(pieces):
            last = index == len(pieces) - 1
            if last and not re.search(r'[.!?…:;]["”’»)]*$', piece):
                piece += "."
            gap = GAP_SENTENCE_MS if not last else GAP_HEADING_MS if kind == "h" else GAP_PARAGRAPH_MS
            chunks.append({"t": piece, "g": gap})
    return chunks


def parse_document(path):
    if path.suffix == ".docx":
        try:
            paragraphs = docx_paragraphs(path)
        except (zipfile.BadZipFile, KeyError, ET.ParseError, ValueError) as exc:
            raise ValueError("Die Datei lässt sich nicht als DOCX lesen. "
                             "Speichere sie in Word neu als .docx und lade sie erneut hoch.") from exc
    else:
        paragraphs = txt_paragraphs(path)
    cleaned = [(c, kind) for text, kind in paragraphs if re.search(r"\w", c := clean(text))]
    return make_chunks(cleaned)


# ---------------------------------------------------------------- Audio
_TRIM = "silenceremove=start_periods=1:start_threshold=-50dB:start_silence={}"
_CHUNK_FILTER = ",".join((_TRIM.format("0.02"), "areverse", _TRIM.format("0.08"), "areverse", "volumedetect"))


def wav_seconds(path):
    with wave.open(str(path), "rb") as handle:
        return handle.getnframes() / handle.getframerate()


def write_silence(path, milliseconds):
    with wave.open(str(path), "wb") as handle:
        handle.setnchannels(1)
        handle.setsampwidth(2)
        handle.setframerate(SAMPLE_RATE)
        handle.writeframes(b"\0\0" * (SAMPLE_RATE * milliseconds // 1000))


def normalize_take(raw, dest):
    """Antwort des Modells nach 44,1 kHz/16 bit mono, Ränder trimmen. Liefert (Sekunden, mean dB)."""
    result = subprocess.run(
        [FFMPEG, "-hide_banner", "-nostats", "-y", "-i", "pipe:0", "-ac", "1", "-ar", str(SAMPLE_RATE),
         "-af", _CHUNK_FILTER, "-c:a", "pcm_s16le", "-f", "wav", str(dest)],
        input=raw, capture_output=True)
    if result.returncode != 0:
        return None
    match = re.search(r"mean_volume:\s*(-?[\d.]+|-inf)\s*dB", result.stderr.decode("utf-8", "replace"))
    return wav_seconds(dest), float(match.group(1)) if match else -math.inf


def estimate_seconds(text):
    return max(0.8, len(text) / CPS)


def judge_take(seconds, mean_db, estimate):
    if seconds < 0.2 or mean_db < -45:
        return "still"
    if seconds < 0.35 * estimate:
        return "zu kurz"
    if seconds > 2.3 * estimate + 1.5:
        return "zu lang"
    return None


def encode_mp3(job, chunks):
    lines = []
    for index, chunk in enumerate(chunks):
        lines.append(f"file 'chunks/{index:06d}.wav'")
        if index < len(chunks) - 1:
            gap = job.dir / f"gap_{chunk['g']}.wav"
            if not gap.exists():
                write_silence(gap, chunk["g"])
            lines.append(f"file '{gap.name}'")
    (job.dir / "concat.txt").write_text("\n".join(lines) + "\n", encoding="utf-8")
    command = [FFMPEG, "-hide_banner", "-nostats", "-y", "-f", "concat", "-safe", "0", "-i", "concat.txt"]
    if LOUDNORM:
        command += ["-af", "loudnorm=I=-18:TP=-1.5:LRA=11"]
    command += ["-ac", "1", "-ar", str(SAMPLE_RATE), "-c:a", "libmp3lame", "-b:a", MP3_BITRATE,
                "-id3v2_version", "3", "-metadata", f"title={Path(job.name).stem}",
                "-f", "mp3", "output.part"]
    result = subprocess.run(command, cwd=job.dir, capture_output=True)
    if result.returncode != 0:
        raise RuntimeError("ffmpeg konnte die MP3 nicht schreiben:\n"
                           + result.stderr.decode("utf-8", "replace")[-600:])
    (job.dir / "output.part").replace(job.dir / "output.mp3")


# ---------------------------------------------------------------- Backend (Drama 3)
_no_proxy = urllib.request.build_opener(urllib.request.ProxyHandler({}))
_address = urllib.parse.urlsplit(BACKEND_URL)
BACKEND_ADDRESS = (_address.hostname or "127.0.0.1", _address.port or 80)
_NET_ERRORS = (urllib.error.URLError, OSError, http.client.HTTPException)


def port_open(timeout=1.0):
    try:
        with socket.create_connection(BACKEND_ADDRESS, timeout):
            return True
    except OSError:
        return False


def speech_request(payload, timeout):
    req = urllib.request.Request(f"{BACKEND_URL}/v1/audio/speech", data=json.dumps(payload).encode("utf-8"),
                                 headers={"Content-Type": "application/json"}, method="POST")
    with _no_proxy.open(req, timeout=timeout) as response:
        return response.read()


def backend_answers():
    if not port_open():
        return False
    try:
        return len(speech_request({"input": "Ready.", "max_new_tokens": 60}, 300)) > 1000
    except _NET_ERRORS:
        return False


_gpu_cache = (0.0, None)


def gpu_info(max_age=3.0):
    global _gpu_cache
    now = time.time()
    if now - _gpu_cache[0] < max_age:
        return _gpu_cache[1]
    info = None
    try:
        result = subprocess.run(
            ["nvidia-smi", "--query-gpu=utilization.gpu,memory.used,memory.total,memory.free",
             "--format=csv,noheader,nounits", "--id=0"], capture_output=True, text=True, timeout=8)
        if result.returncode == 0:
            util, used, total, free = (float(v) for v in result.stdout.strip().split(","))
            info = {"util": round(util), "used_gib": used / 1024, "total_gib": total / 1024, "free_gib": free / 1024}
    except (OSError, ValueError, subprocess.TimeoutExpired):
        pass
    _gpu_cache = (now, info)
    return info


def _log_tail(chars=700):
    try:
        return BACKEND_LOG.read_text("utf-8", "replace")[-chars:].strip()
    except OSError:
        return ""


def backend_command():
    binary = Path(sys.executable).with_name("sgl-omni")
    codec = Path(MODEL_PATH) / "codec.pth"
    return [str(binary) if binary.exists() else "sgl-omni", "serve", "--model-path", MODEL_PATH,
            "--port", str(BACKEND_ADDRESS[1]), "--tts_engine.gpu_memory_fraction", ENGINE_MEMORY,
            "--tts_engine.factory.max_running_requests", MAX_RUNNING,
            *(["--vocoder.factory.codec_path", str(codec)] if codec.is_file() else []), *BACKEND_ARGS]


def backend_environment():
    env = dict(os.environ, PATH=f"{Path(sys.executable).parent}:{os.environ.get('PATH', '')}")
    # libcudart.so.13 & Co. aus den CUDA-Wheels der Umgebung (braucht sgl_kernel)
    libraries = [str(Path(p) / "nvidia/cu13/lib") for p in site.getsitepackages()]
    env["LD_LIBRARY_PATH"] = ":".join(filter(None, [*libraries, env.get("LD_LIBRARY_PATH")]))
    return env


class Backend:
    def __init__(self):
        self.lock = threading.RLock()
        self.proc = None
        self.busy = 0
        self.last_used = time.time()
        self.starting = False

    def state(self):
        if self.starting:
            return "starting"
        return "ready" if port_open(0.3) else "off"

    def acquire(self):
        with self.lock:
            self.busy += 1

    def release(self):
        with self.lock:
            self.busy -= 1
            self.last_used = time.time()

    def _running(self):
        return self.proc is not None and self.proc.poll() is None

    def _wait_for_gpu(self, job):
        while True:
            info = gpu_info(0)
            if info is None or info["free_gib"] >= MIN_FREE_GIB:
                return
            job.set("waiting_gpu", f"{info['free_gib']:.0f} GB frei, {MIN_FREE_GIB:.0f} GB nötig")
            if job.stop.wait(15):
                raise Cancelled

    def ensure_ready(self, job):
        if backend_answers():
            return
        if not AUTOSTART:
            raise BackendError(f"Das Sprachmodell antwortet nicht unter {BACKEND_URL}. "
                               "Starte sgl-omni von Hand oder setze autostart = true.")
        if MODEL_PATH.startswith((".", "/")) and not Path(MODEL_PATH).is_dir():
            raise BackendError(f"Modell nicht gefunden: {MODEL_PATH}. "
                               "Lade es mit: hf download fishaudio/drama-3 --local-dir ./drama-3")
        if not self._running():
            self._wait_for_gpu(job)
        job.set("starting")
        self.starting = True
        try:
            with self.lock:
                if not self._running():
                    try:
                        with BACKEND_LOG.open("ab") as log:
                            self.proc = subprocess.Popen(backend_command(), env=backend_environment(), stdout=log,
                                                         stderr=subprocess.STDOUT, stdin=subprocess.DEVNULL,
                                                         start_new_session=True, cwd=BASE)
                    except OSError as exc:
                        raise BackendError("sgl-omni lässt sich nicht starten. Läuft das Skript in der "
                                           f"Conda-Umgebung? ({exc})") from exc
                proc = self.proc
            deadline = time.time() + START_TIMEOUT_SEC
            while time.time() < deadline:
                if job.stop.is_set():
                    raise Cancelled
                if proc.poll() is not None:
                    raise BackendError("Das Sprachmodell konnte nicht starten. Ende von data/backend.log:\n"
                                       + _log_tail())
                if backend_answers():
                    return
                job.stop.wait(3)
            raise BackendError(f"Das Sprachmodell war nach {START_TIMEOUT_SEC} s nicht bereit. "
                               "Ende von data/backend.log:\n" + _log_tail())
        finally:
            self.starting = False

    def stop(self, grace=30):
        with self.lock:
            proc, self.proc = self.proc, None
            if proc is None:
                return
            group = proc.pid
            try:
                os.killpg(group, signal.SIGTERM)
            except ProcessLookupError:
                proc.poll()
                return
            deadline = time.time() + grace
            while time.time() < deadline:
                proc.poll()
                try:
                    os.killpg(group, 0)
                except ProcessLookupError:
                    return
                time.sleep(0.5)
            try:
                os.killpg(group, signal.SIGKILL)
            except ProcessLookupError:
                pass
            proc.wait(10)


backend = Backend()


# ---------------------------------------------------------------- Aufträge
class Job:
    FIELDS = ("id", "name", "source", "voice", "created", "finished", "status", "detail", "error",
              "total", "done", "audio_sec", "factor", "warnings")

    def __init__(self, **values):
        self.id = self.name = self.source = self.voice = ""
        self.created = self.finished = 0.0
        self.status, self.detail, self.error = "queued", "", ""
        self.total = self.done = 0
        self.audio_sec = 0.0
        self.factor = None
        self.warnings = []
        for key in self.FIELDS:
            if key in values:
                setattr(self, key, values[key])
        self.stop = threading.Event()
        self.deleted = False
        self.lock = threading.Lock()
        self.states = bytearray()
        self.heights = ""
        self.run_started = 0.0
        self.run_done = 0
        self.run_audio = 0.0

    @property
    def dir(self):
        return JOBS_DIR / self.id

    @property
    def mp3_name(self):
        stem = re.sub(r'[\x00-\x1f/\\:*?"<>|]+', " ", Path(self.name).stem).strip() or "audio"
        return f"{stem}.mp3"

    def save(self):
        if not self.dir.exists():
            return
        temp = self.dir / "job.json.tmp"
        temp.write_text(json.dumps({k: getattr(self, k) for k in self.FIELDS}, ensure_ascii=False), encoding="utf-8")
        temp.replace(self.dir / "job.json")

    def set(self, status, detail=""):
        self.status, self.detail = status, detail
        self.save()

    def public(self):
        data = {k: getattr(self, k) for k in ("id", "name", "status", "detail", "error", "total", "done",
                                              "audio_sec", "factor", "warnings")}
        data["download_name"] = self.mp3_name
        if self.status in ACTIVE and self.states:
            with self.lock:
                data["states"] = self.states.decode("ascii")
            data["heights"] = self.heights
        if self.status == "running" and self.run_done:
            data["eta"] = (self.total - self.done) * (time.time() - self.run_started) / self.run_done
        return data


JOBS: dict[str, Job] = {}
QUEUE: "queue.Queue[Job]" = queue.Queue()
PENDING, ACTIVE_CHUNK, DONE_CHUNK, WARN_CHUNK = b"0123"


def load_voice(name):
    wav_path, text_path = VOICES / f"{name}.wav", VOICES / f"{name}.txt"
    if not (wav_path.is_file() and text_path.is_file()):
        raise BackendError(f"Die Stimme {name} fehlt. Lege {name}.wav und {name}.txt in voices/ ab.")
    return {"audio_path": str(wav_path.resolve()), "text": text_path.read_text("utf-8").strip()}


def list_voices():
    return sorted(p.stem for p in VOICES.glob("*.wav") if p.with_suffix(".txt").is_file())


def ensure_default_voice():
    """Legt beim ersten Start die Standardstimme en0 an (Referenzclip aus dem Drama-3-README)."""
    if list_voices():
        return
    try:
        with urllib.request.urlopen(VOICE_URL, timeout=60) as response:
            audio = response.read()
    except (OSError, http.client.HTTPException) as exc:
        print(f"Standardstimme nicht geladen ({exc}). Lege name.wav und name.txt in {VOICES} ab.", flush=True)
        return
    (VOICES / "en0.txt").write_text(VOICE_TEXT + "\n", encoding="utf-8")
    (VOICES / "en0.wav").write_bytes(audio)


def call_backend(payload, job):
    delay, last = 2, ""
    for _ in range(5):
        if job.stop.is_set():
            raise Cancelled
        try:
            return speech_request(payload, REQUEST_TIMEOUT_SEC)
        except urllib.error.HTTPError as exc:
            body = exc.read()[:400].decode("utf-8", "replace")
            if 400 <= exc.code < 500 and exc.code not in (408, 429):
                raise BackendError(f"Das Sprachmodell lehnt die Anfrage ab (HTTP {exc.code}): {body}") from exc
            last = f"HTTP {exc.code}: {body}"
        except _NET_ERRORS as exc:
            last = str(exc)
        job.stop.wait(delay)
        delay = min(delay * 2, 20)
    raise BackendError(f"Das Sprachmodell antwortet nicht mehr: {last}")


def synthesize_chunk(job, index, chunk, voice):
    if job.stop.is_set():
        return
    final = job.dir / "chunks" / f"{index:06d}.wav"
    with job.lock:
        job.states[index] = ACTIVE_CHUNK
    estimate = estimate_seconds(chunk["t"])
    payload = {"input": chunk["t"], "references": [voice],
               "max_new_tokens": min(MAX_FRAMES, math.ceil(estimate * 2.5 * FRAMES_PER_SEC) + 40),
               **SAMPLING}
    best, best_score, reason = None, math.inf, "unlesbar"
    accepted = False
    for draw in range(MAX_DRAWS):
        raw = call_backend(payload, job)
        temp = final.with_name(f"{index:06d}.take{draw}.wav")
        take = normalize_take(raw, temp)
        if take is None:
            reason = "unlesbar"
            continue
        seconds, mean_db = take
        reason = judge_take(seconds, mean_db, estimate)
        if reason is None:
            temp.replace(final)
            accepted = True
            break
        if reason != "still":
            score = abs(math.log(seconds / estimate))
            if score < best_score:
                best, best_score = temp, score
    if not accepted:
        if best is not None:
            best.replace(final)
        else:
            write_silence(final, 500)
    for leftover in final.parent.glob(f"{index:06d}.take*.wav"):
        leftover.unlink(missing_ok=True)
    seconds = wav_seconds(final)
    with job.lock:
        job.done += 1
        job.run_done += 1
        job.audio_sec += seconds
        job.run_audio += seconds
        job.states[index] = DONE_CHUNK if accepted else WARN_CHUNK
        if not accepted:
            job.warnings.append({"i": index, "text": chunk["t"][:240], "why": reason})
    if not accepted:
        job.save()


def load_chunks(job):
    path = job.dir / "chunks.json"
    if path.exists():
        return json.loads(path.read_text("utf-8"))
    chunks = parse_document(job.dir / job.source)
    if not chunks:
        raise ValueError("Das Dokument enthält keinen vorlesbaren Text.")
    path.write_text(json.dumps(chunks, ensure_ascii=False), encoding="utf-8")
    return chunks


def prepare_progress(job, chunks):
    (job.dir / "chunks").mkdir(exist_ok=True)
    flagged = {w["i"] for w in job.warnings}
    states = bytearray([PENDING]) * len(chunks)
    done, audio = 0, 0.0
    for index in range(len(chunks)):
        path = job.dir / "chunks" / f"{index:06d}.wav"
        if path.exists():
            states[index] = WARN_CHUNK if index in flagged else DONE_CHUNK
            done += 1
            audio += wav_seconds(path)
    sizes = [_size(c["t"]) for c in chunks]
    longest = max(sizes)
    job.heights = "".join(str(max(1, round(9 * s / longest))) for s in sizes)
    job.total, job.done, job.audio_sec, job.states = len(chunks), done, audio, states
    job.run_done, job.run_audio = 0, 0.0
    return [i for i in range(len(chunks)) if states[i] == PENDING]


def run_job(job):
    job.stop.clear()
    job.error = ""
    job.set("preparing")
    backend.acquire()
    try:
        chunks = load_chunks(job)
        todo = prepare_progress(job, chunks)
        voice = load_voice(job.voice)
        if todo:
            backend.ensure_ready(job)
            job.run_started = time.time()
            job.set("running")
            failure = None
            with ThreadPoolExecutor(max_workers=CONCURRENCY) as pool:
                futures = [pool.submit(synthesize_chunk, job, i, chunks[i], voice) for i in todo]
                for future in as_completed(futures):
                    try:
                        future.result()
                    except Cancelled:
                        pass
                    except Exception as exc:  # erster harter Fehler stoppt den Auftrag
                        failure = failure or exc
                        job.stop.set()
            elapsed = time.time() - job.run_started
            if failure:
                raise failure
            if job.stop.is_set():
                raise Cancelled
        job.set("encoding")
        encode_mp3(job, chunks)
        shutil.rmtree(job.dir / "chunks", ignore_errors=True)
        for leftover in [*job.dir.glob("gap_*.wav"), job.dir / "concat.txt"]:
            leftover.unlink(missing_ok=True)
        if todo and job.run_audio:
            job.factor = job.run_audio / elapsed
        job.finished = time.time()
        job.set("done")
    except Cancelled:
        job.set("cancelled")
    except Exception as exc:  # Meldung landet in der Oberfläche
        job.error = str(exc) or exc.__class__.__name__
        job.finished = time.time()
        job.set("failed")
    finally:
        backend.release()
        if job.deleted:
            shutil.rmtree(job.dir, ignore_errors=True)


def worker():
    while True:
        job = QUEUE.get()
        if job.deleted or job.status != "queued":
            continue
        run_job(job)


def expire_jobs():
    limit = time.time() - KEEP_DAYS * 86400
    for job in list(JOBS.values()):
        if job.status not in ACTIVE and (job.finished or job.created) < limit:
            JOBS.pop(job.id, None)
            shutil.rmtree(job.dir, ignore_errors=True)


def housekeeping():
    last_expiry = 0.0
    while True:
        time.sleep(5)
        if IDLE_STOP_SEC > 0 and backend.busy == 0 and backend._running() \
                and time.time() - backend.last_used > IDLE_STOP_SEC:
            backend.stop()
        if time.time() - last_expiry > 3600:
            last_expiry = time.time()
            expire_jobs()


def load_jobs():
    for path in JOBS_DIR.glob("*/job.json"):
        try:
            job = Job(**json.loads(path.read_text("utf-8")))
        except (ValueError, TypeError, OSError):
            continue
        JOBS[job.id] = job
    for job in sorted(JOBS.values(), key=lambda j: j.created):
        if job.status in ACTIVE:
            job.status = "queued"
            QUEUE.put(job)


# ---------------------------------------------------------------- Web
app = Flask(__name__, template_folder=str(BASE / "templates"))
app.config["MAX_CONTENT_LENGTH"] = MAX_UPLOAD_MB * 1024 * 1024
app.config["TEMPLATES_AUTO_RELOAD"] = True               # Änderungen am Template ohne Neustart


def _problem(message, code):
    return jsonify(error=message), code


@app.errorhandler(413)
def _too_large(_):
    return _problem(f"Die Datei ist größer als {MAX_UPLOAD_MB} MB.", 413)


@app.get("/")
def index():
    page = render_template("index.html", title=TITLE, max_upload_mb=MAX_UPLOAD_MB, keep_days=KEEP_DAYS)
    return page, 200, {"Cache-Control": "no-store"}


@app.get("/api/state")
def state():
    jobs = sorted(JOBS.values(), key=lambda j: j.created, reverse=True)
    info = gpu_info() or {}
    return jsonify(jobs=[j.public() for j in jobs], voices=list_voices(),
                   backend={"state": backend.state(), **info})


@app.post("/api/jobs")
def create_job():
    upload = request.files.get("file")
    if upload is None or not upload.filename:
        return _problem("Es kam keine Datei an.", 400)
    name = Path(upload.filename.replace("\\", "/")).name
    extension = Path(name).suffix.lower()
    if extension not in (".docx", ".txt"):
        return _problem(f"{name}: Nur .docx und .txt lassen sich vorlesen.", 415)
    voices = list_voices()
    if not voices:
        return _problem("Keine Stimme vorhanden. Lege name.wav und name.txt in voices/ ab.", 409)
    voice = request.form.get("voice") or voices[0]
    if voice not in voices:
        return _problem(f"Die Stimme {voice} gibt es nicht.", 400)
    job = Job(id=f"{int(time.time() * 1000):x}-{secrets.token_hex(3)}", name=name,
              source=f"source{extension}", voice=voice, created=time.time())
    job.dir.mkdir(parents=True)
    upload.save(job.dir / job.source)
    job.save()
    JOBS[job.id] = job
    QUEUE.put(job)
    return jsonify(job.public()), 201


def _job_or_404(job_id):
    job = JOBS.get(job_id)
    if job is None:
        abort(404)
    return job


@app.post("/api/jobs/<job_id>/cancel")
def cancel_job(job_id):
    job = _job_or_404(job_id)
    if job.status in ACTIVE:
        job.stop.set()
        if job.status == "queued":
            job.set("cancelled")
    return jsonify(job.public())


@app.post("/api/jobs/<job_id>/retry")
def retry_job(job_id):
    job = _job_or_404(job_id)
    if job.status in ("failed", "cancelled"):
        job.stop.clear()
        job.error = ""
        job.set("queued")
        QUEUE.put(job)
    return jsonify(job.public())


@app.delete("/api/jobs/<job_id>")
def delete_job(job_id):
    job = _job_or_404(job_id)
    JOBS.pop(job_id, None)
    if job.status == "queued":
        job.status = "cancelled"
    if job.status in ACTIVE:
        job.deleted = True
        job.stop.set()
    else:
        shutil.rmtree(job.dir, ignore_errors=True)
    return "", 204


@app.get("/files/<job_id>")
def download(job_id):
    job = _job_or_404(job_id)
    path = job.dir / "output.mp3"
    if job.status != "done" or not path.is_file():
        abort(404)
    return send_file(path, mimetype="audio/mpeg", as_attachment="dl" in request.args,
                     download_name=job.mp3_name, conditional=True, max_age=0)


def shutdown(*_):
    """Strg+C und SIGTERM: Backend beenden und ohne Traceback raus. Offene Aufträge laufen beim nächsten Start weiter."""
    signal.signal(signal.SIGINT, signal.SIG_IGN)
    signal.signal(signal.SIGTERM, signal.SIG_IGN)
    print(f"\n{TITLE} wird beendet.", flush=True)
    backend.stop(grace=10)
    os._exit(0)


def main():
    signal.signal(signal.SIGINT, shutdown)
    signal.signal(signal.SIGTERM, shutdown)
    JOBS_DIR.mkdir(parents=True, exist_ok=True)
    VOICES.mkdir(parents=True, exist_ok=True)
    ensure_default_voice()
    load_jobs()
    threading.Thread(target=worker, daemon=True).start()
    threading.Thread(target=housekeeping, daemon=True).start()
    atexit.register(backend.stop)
    print(f"{TITLE}: http://{HOST}:{PORT}  Modell: {MODEL_PATH}", flush=True)
    try:
        from waitress import serve
    except ImportError:
        app.run(HOST, PORT, threaded=True)
        return
    serve(app, host=HOST, port=PORT, threads=12)


if __name__ == "__main__":
    main()
