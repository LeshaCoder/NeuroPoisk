import sys
import os
import io
import subprocess
import threading
import time
import re
import shutil
import tempfile
import zipfile
import html
from pathlib import Path
from io import BytesIO


# PyInstaller --windowed: stdout/stderr могут быть None, заглушаем до импортов
class _NullStream(io.TextIOBase):
    def write(self, s): return len(s) if s else 0
    def isatty(self):   return False
    def flush(self):    pass
    def readable(self): return False
    def writable(self): return True
    def seekable(self): return False

if sys.stdout is None: sys.stdout = _NullStream()
if sys.stderr is None: sys.stderr = _NullStream()


if getattr(sys, "frozen", False):
    BASE_DIR = Path(sys.executable).parent
else:
    BASE_DIR = Path(__file__).parent

os.environ["PATH"] = str(BASE_DIR) + os.pathsep + os.environ.get("PATH", "")

EMBED_DEVICE = "cpu"

SUPPORTED_EXTS = {
    ".txt", ".md", ".pdf", ".docx", ".csv",
    ".fb2", ".epub", ".mobi", ".azw3", ".azw",
    ".djvu", ".djv", ".rtf", ".odt", ".html", ".htm",
}


# --- Чтение файлов ---

def _decode_bytes(raw):
    if not raw:
        return ""
    if raw.startswith(b'\xef\xbb\xbf'):
        return raw[3:].decode("utf-8", errors="replace")
    if raw.startswith((b'\xff\xfe\x00\x00', b'\x00\x00\xfe\xff')):
        return raw.decode("utf-32", errors="replace")
    if raw.startswith((b'\xff\xfe', b'\xfe\xff')):
        return raw.decode("utf-16", errors="replace")
    try:
        import chardet
        r = chardet.detect(raw[:50_000])
        enc = r.get("encoding") or "utf-8"
        if r.get("confidence", 0) >= 0.5:
            t = raw.decode(enc, errors="replace")
            if t.count("\ufffd") / max(len(t), 1) < 0.05:
                return t
    except ImportError:
        pass
    for enc in ("utf-8", "cp1251", "cp866", "koi8-r", "latin-1"):
        try:
            return raw.decode(enc)
        except (UnicodeDecodeError, LookupError):
            pass
    return raw.decode("utf-8", errors="replace")


def _read_fb2(fp):
    try:
        raw = fp.read_bytes()
        if raw[:2] == b'PK':  # fb2 в zip
            with zipfile.ZipFile(BytesIO(raw)) as z:
                for name in z.namelist():
                    if name.endswith('.fb2'):
                        raw = z.read(name)
                        break
        text = _decode_bytes(raw)
        text = re.sub(r'<(title|p|v|stanza|poem|subtitle|epigraph|annotation)[^>]*>', '\n', text)
        text = re.sub(r'<[^>]+>', '', text)
        return re.sub(r'\n{3,}', '\n\n', html.unescape(text)).strip()
    except Exception:
        return ""

def _read_epub(fp):
    try:
        import ebooklib
        from ebooklib import epub
        from bs4 import BeautifulSoup
        book  = epub.read_epub(str(fp))
        parts = [BeautifulSoup(item.get_content(), "lxml").get_text("\n")
                 for item in book.get_items_of_type(ebooklib.ITEM_DOCUMENT)]
        return re.sub(r'\n{3,}', '\n\n', "\n".join(parts)).strip()
    except ImportError:
        pass
    try:
        parts = []
        with zipfile.ZipFile(str(fp)) as z:
            for name in sorted(z.namelist()):
                if name.endswith(('.html', '.xhtml', '.htm')):
                    text = _decode_bytes(z.read(name))
                    parts.append(html.unescape(re.sub(r'<[^>]+>', ' ', text)).strip())
        return re.sub(r'\n{3,}', '\n\n', "\n\n".join(parts)).strip()
    except Exception:
        return ""

def _read_mobi_azw(fp):
    try:
        import mobi
        tempdir, _ = mobi.extract(str(fp))
        result = ""
        for root, _, files in os.walk(tempdir):
            for fname in files:
                fpath = Path(root) / fname
                if fpath.suffix.lower() in ('.html', '.xhtml', '.htm'):
                    text = _decode_bytes(fpath.read_bytes())
                    result += html.unescape(re.sub(r'<[^>]+>', ' ', text)) + "\n"
        shutil.rmtree(tempdir, ignore_errors=True)
        return re.sub(r'\n{3,}', '\n\n', result).strip()
    except Exception:
        pass
    try:
        text = _decode_bytes(fp.read_bytes())
        text = re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f\x7f-\x9f]', '', text)
        return re.sub(r'\n{3,}', '\n\n', text).strip()
    except Exception:
        return ""

def _read_djvu(fp):
    try:
        r = subprocess.run(["djvutxt", str(fp)], capture_output=True, text=True, timeout=30)
        if r.returncode == 0:
            return re.sub(r'\n{3,}', '\n\n', r.stdout).strip()
    except Exception:
        pass
    return ""

def _read_rtf(fp):
    try:
        text = _decode_bytes(fp.read_bytes())
        text = re.sub(r'\\[a-z]+[-\d]*[ ]?', ' ', text)
        text = re.sub(r'[{}\\]', '', text)
        return re.sub(r'\s+', ' ', text).strip()
    except Exception:
        return ""

def _read_odt(fp):
    try:
        with zipfile.ZipFile(str(fp)) as z:
            if 'content.xml' in z.namelist():
                text = _decode_bytes(z.read('content.xml'))
                text = re.sub(r'<[^>]+>', ' ', text)
                return re.sub(r'\n{3,}', '\n\n', html.unescape(text)).strip()
    except Exception:
        pass
    return ""

def _read_html_file(fp):
    try:
        text = _decode_bytes(fp.read_bytes())
        try:
            from bs4 import BeautifulSoup
            return BeautifulSoup(text, "lxml").get_text("\n")
        except ImportError:
            pass
        text = re.sub(r'<script[^>]*>.*?</script>', '', text, flags=re.S | re.I)
        text = re.sub(r'<style[^>]*>.*?</style>',  '', text, flags=re.S | re.I)
        text = re.sub(r'<[^>]+>', ' ', text)
        return html.unescape(re.sub(r'\n{3,}', '\n\n', text)).strip()
    except Exception:
        return ""

def _read_pdf(fp):
    try:
        import fitz
        return "\n\n".join(page.get_text() for page in fitz.open(str(fp)))
    except ImportError:
        pass
    try:
        from pdfminer.high_level import extract_text
        return extract_text(str(fp))
    except ImportError:
        pass
    return ""

def _read_docx(fp):
    try:
        from docx import Document
        return "\n\n".join(p.text for p in Document(str(fp)).paragraphs if p.text.strip())
    except Exception:
        return ""

def _read_file_text(fp):
    ext = fp.suffix.lower()
    if ext == ".fb2":                         return _read_fb2(fp)
    if ext == ".epub":                        return _read_epub(fp)
    if ext in (".mobi", ".azw3", ".azw"):     return _read_mobi_azw(fp)
    if ext in (".djvu", ".djv"):              return _read_djvu(fp)
    if ext == ".rtf":                         return _read_rtf(fp)
    if ext == ".odt":                         return _read_odt(fp)
    if ext in (".html", ".htm"):              return _read_html_file(fp)
    if ext == ".pdf":                         return _read_pdf(fp)
    if ext == ".docx":                        return _read_docx(fp)
    return _decode_bytes(fp.read_bytes())


# --- Очистка текста от мусора книжных сайтов ---

_JUNK_LINE = [
    "bookscafe", "litres.ru", "flibusta", "lib.ru", "royallib", "aldebaran",
    "coollib", "электронной библиотек", "скачали книгу", "оставить отзыв",
    "все книги автора", "другие книги серии", "скачать книгу", "читать онлайн",
]
_JUNK_END = ["конец ознакомительного фрагмента", "купить полную версию"]

def _clean_text(txt):
    lower = txt.lower()
    cut   = len(txt)
    for m in _JUNK_END:
        i = lower.find(m)
        if i != -1 and i < cut:
            cut = i
    txt   = txt[:cut]
    lines = []
    for line in txt.splitlines():
        ll = line.lower().strip()
        if not ll:
            lines.append(line)
            continue
        if "http://" in ll or "https://" in ll or ll.startswith("www."):
            continue
        if any(m in ll for m in _JUNK_LINE):
            continue
        lines.append(line)
    return re.sub(r"\n{3,}", "\n\n", "\n".join(lines)).strip()


# --- Retrieval ---

TOP_K = 30

_LIT_QUERIES = [
    "главный герой персонаж", "начало первая глава", "сюжет события",
    "место действия", "отношения персонажей", "конец финал развязка",
    "описание природы", "диалог разговор", "кульминация", "семья дом быт",
    "деньги богатство", "мысли размышления", "работа профессия", "детство прошлое",
]

def _raw_retrieve(kb, q):
    for method in ["retrieve", "search", "query"]:
        if not hasattr(kb, method):
            continue
        try:
            r = getattr(kb, method)(q)
            if r:
                return r
        except Exception:
            pass
    return []

def retrieve_chunks(kb, question):
    seen, chunks = set(), []
    def add(cs):
        for c in cs:
            t = c.text if hasattr(c, "text") else str(c)
            k = t[:120].strip()
            if k not in seen:
                seen.add(k)
                chunks.append(c)
    add(_raw_retrieve(kb, question))
    for q in _LIT_QUERIES:
        if len(chunks) >= TOP_K:
            break
        add(_raw_retrieve(kb, q))
    return chunks[:TOP_K]

def _read_docs_direct(kb):
    docs_path = None
    for attr in ["docs_folder", "_docs_folder", "folder", "path"]:
        v = getattr(kb, attr, None)
        if v:
            p = Path(str(v))
            if p.exists():
                docs_path = p
                break
    if not docs_path and hasattr(kb, "_np_docs"):
        docs_path = Path(kb._np_docs)
    if not docs_path or not docs_path.exists():
        return ""
    files = [docs_path] if docs_path.is_file() else [
        f for f in sorted(docs_path.glob("**/*"))
        if f.is_file() and f.suffix.lower() in SUPPORTED_EXTS
    ]
    context = ""
    for fp in files[:5]:
        try:
            txt = _clean_text(_read_file_text(fp))
            if not txt:
                continue
            n    = len(txt)
            step = max(n // 8, 1)
            parts = [txt[i * step:i * step + 4000] for i in range(8)]
            context += "\n---\n" + "\n---\n".join(p for p in parts if p.strip())
        except Exception:
            pass
    return context

_BROAD = [
    "о чём", "про что", "краткое", "перескажи", "расскажи о книге",
    "что за книга", "содержание", "сюжет", "тема", "главная мысль",
    "о чем эта", "что происходит",
]

def _is_broad(q):
    return any(k in q.lower() for k in _BROAD)


# --- OpenRouter ---

_SYSTEM = """Ты — умный ассистент для анализа текстов книг.

Тебе даётся КОНТЕКСТ — несколько фрагментов из книг.
Ты должен ответить на вопрос пользователя, опираясь ТОЛЬКО на этот контекст.

ПРАВИЛА:
1. Отвечай только на русском языке.
2. Используй исключительно информацию из контекста.
3. НЕ придумывай факты.
4. Если в контексте нет ответа — напиши:
   "В предоставленных фрагментах нет ответа на этот вопрос".
5. Дай связный, понятный ответ своими словами.
6. ВСТАВЛЯЙ ссылки на цитаты внутри ответа в формате [1], [2], [3].
7. В конце обязательно приведи список цитат, на которые ты ссылался.

КАК РАБОТАТЬ:
- Сначала проанализируй все фрагменты
- Выбери самые релевантные части
- Объедини информацию
- Сформируй ответ
- Расставь ссылки [1], [2] в нужных местах

ФОРМАТ ОТВЕТА:

Ответ:
<твой ответ с ссылками [1], [2]>

Цитаты:
[1] "<точная цитата из текста>"
[2] "<точная цитата из текста>"
[3] "<если есть>"

ВАЖНО:
- Цитаты должны быть ДОСЛОВНЫМИ.
- Не добавляй ничего вне контекста.
- Не дублируй одинаковые цитаты.
- Используй только действительно нужные фрагменты.
"""

OPENROUTER_API_KEYS = [
    "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",
    "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",
    "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",
    "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",
    "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",
    "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",
]

_models_cache      = []
_models_cache_time = 0.0
_MODELS_TTL        = 600

def fetch_free_models():
    import urllib.request, json
    global _models_cache, _models_cache_time
    if _models_cache and (time.time() - _models_cache_time) < _MODELS_TTL:
        return _models_cache
    try:
        req = urllib.request.Request(
            "https://openrouter.ai/api/v1/models",
            headers={"Authorization": f"Bearer {OPENROUTER_API_KEYS[0]}"},
        )
        with urllib.request.urlopen(req, timeout=10) as r:
            data = json.loads(r.read())
        models = [m["id"] for m in data.get("data", []) if m.get("id", "").endswith(":free")]
        if models:
            _models_cache      = models
            _models_cache_time = time.time()
            return models
    except Exception:
        pass
    return _models_cache

def _or_post(model_id, context, question, api_key):
    import urllib.request, json
    user = f"ТЕКСТ:\n\"\"\"\n{context}\n\"\"\"\n\nВОПРОС: {question}\n\nОтвет строго по тексту:"
    body = json.dumps({
        "model": model_id,
        "messages": [
            {"role": "system", "content": _SYSTEM},
            {"role": "user",   "content": user},
        ],
        "max_tokens": 3000,
        "temperature": 0.1,
    }, ensure_ascii=False).encode("utf-8")
    req = urllib.request.Request(
        "https://openrouter.ai/api/v1/chat/completions",
        data=body,
        headers={
            "Content-Type":  "application/json",
            "Authorization": f"Bearer {api_key}",
            "HTTP-Referer":  "https://neuropoisk.local",
            "X-Title":       "Neuropoisk",
        },
    )
    with urllib.request.urlopen(req, timeout=60) as r:
        data = json.loads(r.read().decode("utf-8"))
    if "error" in data:
        raise RuntimeError(data["error"].get("message", "?"))
    result = data["choices"][0]["message"]["content"].strip()
    print(f"[OpenRouter] raw answer repr: {repr(result[:120])}")
    return result

def ask_openrouter(context, question):
    import urllib.error
    models = fetch_free_models()
    for model_id in models:
        for api_key in OPENROUTER_API_KEYS:
            try:
                return _or_post(model_id, context, question, api_key)
            except urllib.error.HTTPError as e:
                if e.code in (429, 401, 403):
                    continue   # лимит / невалидный ключ — следующий ключ
                break          # модель недоступна — следующая модель
            except Exception:
                continue
    return "⚠  Все модели и ключи OpenRouter временно недоступны."

def ask_with_kb(kb, question):
    context = ""
    if _is_broad(question):
        context = _read_docs_direct(kb)
    if not context:
        chunks = retrieve_chunks(kb, question)
        for c in chunks:
            t = _clean_text(c.text if hasattr(c, "text") else str(c))
            if t:
                context += f"\n---\n{t}\n"
    if not context:
        context = _read_docs_direct(kb)
    if not context:
        return "⚠  Не удалось получить текст из документов."
    return ask_openrouter(context, question)


# --- Индексация ---

_tmp_dirs = []

# Скорость обработки all-mpnet-base-v2 на CPU: ~3000 символов/сек (эмпирически)
_EMBED_CHARS_PER_SEC = 1000

def _collect_files(fp):
    """Вернуть список файлов для индексации (один файл или папка)."""
    if fp.is_file():
        return [fp]
    return [f for f in sorted(fp.rglob("*"))
            if f.is_file() and f.suffix.lower() in SUPPORTED_EXTS]


def make_kb(docs_path, progress):
    from piragi import Ragi
    fp  = Path(docs_path)
    NATIVE_EXTS = {".txt", ".md", ".pdf", ".docx", ".csv"}

    # --- Фаза 1 (5-40%): чтение и конвертация файлов ---
    if fp.is_file():
        ext = fp.suffix.lower()
        if ext not in NATIVE_EXTS:
            progress["pct"] = 5
            text = _read_file_text(fp)
            if not text:
                raise RuntimeError(
                    f"Не удалось прочитать {fp.name}.\n"
                    f"Формат {ext} может требовать доп. библиотек:\n"
                    f"  EPUB/MOBI → pip install ebooklib beautifulsoup4\n"
                    f"  DJVU      → установите DjVuLibre"
                )
            progress["pct"] = 35
            tmp_dir  = Path(tempfile.mkdtemp(prefix="neuropoisk_"))
            tmp_file = tmp_dir / (fp.stem + ".txt")
            tmp_file.write_text(_clean_text(text), encoding="utf-8")
            _tmp_dirs.append(tmp_dir)
            fp = tmp_file
        progress["pct"] = 40
    else:
        # Папка: конвертируем нестандартные файлы с реальным прогрессом
        files = _collect_files(fp)
        total = max(len(files), 1)
        needs_tmp = any(f.suffix.lower() not in NATIVE_EXTS for f in files)
        if needs_tmp:
            tmp_dir = Path(tempfile.mkdtemp(prefix="neuropoisk_"))
            _tmp_dirs.append(tmp_dir)
            for i, f in enumerate(files):
                pct = 5 + int((i / total) * 33)
                progress["pct"] = pct
                if f.suffix.lower() not in NATIVE_EXTS:
                    try:
                        text = _read_file_text(f)
                        if text:
                            out = tmp_dir / (f.stem + ".txt")
                            out.write_text(_clean_text(text), encoding="utf-8")
                    except Exception:
                        pass
                else:
                    try:
                        shutil.copy2(str(f), str(tmp_dir / f.name))
                    except Exception:
                        pass
            fp = tmp_dir
        progress["pct"] = 40

    # --- Фаза 2 (40-95%): индексация (embed в Ragi) ---
    try:
        char_count = fp.stat().st_size if fp.is_file() else sum(
            f.stat().st_size for f in fp.rglob("*") if f.is_file()
        )
    except Exception:
        char_count = 100_000

    est_sec = max(char_count / _EMBED_CHARS_PER_SEC, 3)
    print(f"[make_kb] char_count={char_count} bytes, est_sec={est_sec:.1f}s")
    progress["pct"]     = 40
    progress["est_sec"] = est_sec
    progress["t_index"] = time.time()

    index_dir = Path(tempfile.gettempdir()) / "NeuroPoiskov"
    index_dir.mkdir(parents=True, exist_ok=True)

    t_start = time.time()
    kb = Ragi(str(fp), persist_dir=str(index_dir), config={
        "embedding": {"model": "all-mpnet-base-v2", "device": EMBED_DEVICE},
    })
    real_sec = time.time() - t_start
    print(f"[make_kb] Ragi finished in {real_sec:.1f}s (est was {est_sec:.1f}s, char_count={char_count})")
    progress["pct"] = 100
    kb._np_docs = str(fp)
    return kb

def cleanup():
    try:
        shutil.rmtree(str(Path(tempfile.gettempdir()) / "NeuroPoiskov"), ignore_errors=True)
    except Exception:
        pass
    for d in _tmp_dirs:
        try:
            shutil.rmtree(str(d), ignore_errors=True)
        except Exception:
            pass


# --- GUI ---

try:
    import customtkinter as ctk
    import tkinter as _tk
except ImportError:
    sys.exit(1)

ctk.set_appearance_mode("light")
ctk.set_default_color_theme("dark-blue")

BG     = "#fff"
SURF   = "#f7f7f7"
SURF2  = "#00a2ff"
GOLD   = "#00a2ff"
GOLDD  = "#00a2ff"
CREAM  = "black"
MUTED  = "#6B7280"
GREEN  = "#4ADE80"
RED    = "#F87171"
BORDER = "#00a2ff"
BLUE = "#116cf5"

F_TITLE = ("Segoe UI", 22, "bold")
F_LABEL = ("Segoe UI", 11, "bold")
F_BTN   = ("Segoe UI", 15, "bold")
F_MONO  = ("Consolas", 15 , "bold")
F_SMALL = ("Segoe UI", 13)
F_ANS   = ("Segoe UI", 13)
F_BOLD  = ("Segoe UI", 13, "bold")

_CTRL_KEYS = {67: "<<Copy>>", 86: "<<Paste>>", 88: "<<Cut>>", 65: "<<SelectAll>>", 90: "<<Undo>>"}

def _fix_hotkeys(w):
    def _on(ev):
        if ev.state & 0x4:
            a = _CTRL_KEYS.get(ev.keycode)
            if a:
                w.event_generate(a)
                return "break"
    w.bind("<KeyPress>", _on, add=True)

def _textbox(parent, height, readonly=False, scroll_parent=None):
    frame = ctk.CTkFrame(parent, fg_color=SURF, corner_radius=8, border_width=1, border_color=BORDER)
    t = _tk.Text(
        frame, font=F_ANS, bg=SURF, fg=CREAM, insertbackground=CREAM,
        relief="flat", bd=0, highlightthickness=0, wrap="word",
        height=height, padx=10, pady=8, undo=True,
    )
    t.pack(fill="both", expand=True)
    _fix_hotkeys(t)

    # Text перехватывает MouseWheel сам — останавливаем всплытие к родителю
    def _on_scroll(ev):
        delta = int(-ev.delta / 120)
        top, bottom = t.yview()
        # Если листаем вниз и уже внизу — передаём в scroll_parent
        if delta > 0 and bottom >= 1.0 and scroll_parent is not None:
            scroll_parent._parent_canvas.yview_scroll(delta * 16, "units")
            return "break"
        # Если листаем вверх и уже вверху — передаём в scroll_parent
        if delta < 0 and top <= 0.0 and scroll_parent is not None:
            scroll_parent._parent_canvas.yview_scroll(delta * 16, "units")
            return "break"
        t.yview_scroll(delta * 2, "units")
        return "break"
    t.bind("<MouseWheel>", _on_scroll)

    if readonly:
        def _bl(ev):
            if ev.state & 0x4:
                return
            if ev.char and ev.char.isprintable():
                return "break"
        t.bind("<KeyPress>", _bl, add=True)

    frame.get    = lambda a, b: t.get(a, b)
    frame.insert = lambda a, s: t.insert(a, s)
    frame.delete = lambda a, b: t.delete(a, b)
    frame.config = t.config
    return frame, t

_FILETYPES = [
    ("Все книги и документы",
     "*.txt *.fb2 *.epub *.mobi *.azw3 *.azw *.djvu *.djv "
     "*.pdf *.docx *.rtf *.odt *.html *.htm *.md *.csv"),
    ("Электронные книги", "*.fb2 *.epub *.mobi *.azw3 *.azw *.djvu *.djv"),
    ("Документы",         "*.pdf *.docx *.rtf *.odt"),
    ("Текст",             "*.txt *.md *.csv *.html *.htm"),
    ("Все файлы",         "*.*"),
]


class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        # Говорим Windows что это отдельное приложение, а не python.exe
        try:
            import ctypes
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("neuropoisk.app.1")
        except Exception:
            pass
        self.title("НейроПоиск")
        self.configure(fg_color=BG)
        self.minsize(560, 460)
        try:
            _meipass = Path(getattr(sys, "_MEIPASS", ""))
            icon_path = (_meipass / "icon.ico") if (_meipass / "icon.ico").exists() else (BASE_DIR / "icon.ico")
            if icon_path.exists():
                self.iconbitmap(str(icon_path))
        except Exception:
            pass
        self._kb        = None
        self._busy      = False
        self._indexing  = False
        self._pending_q = None
        self._build()
        self._center()
        self.protocol("WM_DELETE_WINDOW", self._on_close)
        threading.Thread(target=fetch_free_models, daemon=True).start()

    def _center(self):
        self.update_idletasks()
        sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
        w = min(max(int(sw * 0.45), 580), 800)
        h = min(max(int(sh * 0.75), 520), 860)
        self.geometry(f"{w}x{h}+{(sw - w) // 2}+{(sh - h) // 2}")

    def _build(self):
        hdr = ctk.CTkFrame(self, fg_color="transparent")
        hdr.pack(fill="x", padx=32, pady=(20, 0))
        ctk.CTkLabel(hdr, text="НейроПоиск", font=F_TITLE, text_color=BLUE).pack(anchor="w")
        ctk.CTkFrame(self, height=1, fg_color=GOLDD).pack(fill="x", padx=32, pady=(10, 0))

        scroll = ctk.CTkScrollableFrame(
            self, fg_color="transparent",
            scrollbar_button_color=GOLDD,
            scrollbar_button_hover_color=GOLD,
        )
        scroll.pack(fill="both", expand=True, padx=24, pady=14)
        # Скролл основного фрейма только когда курсор не над текстовыми полями
        scroll._parent_canvas.bind(
            "<MouseWheel>",
            lambda e: scroll._parent_canvas.yview_scroll(int(-e.delta / 120) * 16, "units")
        )

        ctk.CTkLabel(scroll, text="ДОКУМЕНТ", font=F_LABEL, text_color="white", anchor="w").pack(fill="x")
        btn_row = ctk.CTkFrame(scroll, fg_color="transparent")
        btn_row.pack(fill="x", pady=(6, 0))
        ctk.CTkButton(
            btn_row, text="Папка", font=F_MONO, height=44,
            fg_color=SURF2, hover_color="#4dbeff", text_color="white",
            border_color=BORDER, border_width=1, corner_radius=8,
            command=self._pick_folder,
        ).pack(side="left", padx=(0, 8))
        ctk.CTkButton(
            btn_row, text="Файл", font=F_MONO, height=44,
            fg_color=SURF2, hover_color="#4dbeff", text_color="white",
            border_color=BORDER, border_width=1, corner_radius=8,
            command=self._pick_file,
        ).pack(side="left")

        self.doc_lbl    = ctk.CTkLabel(scroll, text="Файл не выбран", font=F_SMALL, text_color=MUTED, anchor="w")
        self.doc_lbl.pack(fill="x", pady=(6, 0))
        self.status_lbl = ctk.CTkLabel(scroll, text="", font=F_SMALL, text_color=MUTED, anchor="w")
        self.status_lbl.pack(fill="x", pady=(2, 0))
        self.progress   = ctk.CTkProgressBar(scroll, height=4, fg_color=SURF2, progress_color=GOLD)
        self.progress.set(0)

        ctk.CTkLabel(scroll, text="ВОПРОС", font=F_LABEL, text_color=GOLD, anchor="w").pack(fill="x", pady=(16, 0))
        self._q_frame, self.q_box = _textbox(scroll, height=4, scroll_parent=scroll)
        self._q_frame.pack(fill="x", pady=(6, 0))

        placeholder = "О чём этот документ?"
        self._placeholder_active = True
        self.q_box.insert("1.0", placeholder)
        self.q_box.config(fg=MUTED)

        def q_in(e):
            if self._placeholder_active:
                self.q_box.delete("1.0", "end")
                self.q_box.config(fg=CREAM)
                self._placeholder_active = False

        def q_out(e):
            if not self.q_box.get("1.0", "end").strip():
                self.q_box.insert("1.0", placeholder)
                self.q_box.config(fg=MUTED)
                self._placeholder_active = True

        self.q_box.bind("<FocusIn>",  q_in)
        self.q_box.bind("<FocusOut>", q_out)
        self.q_box.bind("<Control-Return>", lambda e: self._ask())

        ctk.CTkLabel(scroll, text="Ctrl+Enter — отправить", font=F_SMALL,
                     text_color="black", anchor="e").pack(fill="x", pady=(2, 0))

        self.ask_btn = ctk.CTkButton(
            scroll, text="Спросить  →", font=F_BTN, height=50,
            fg_color=GOLD, hover_color=GOLD, text_color="#fff",
            corner_radius=8, command=self._ask,
        )
        self.ask_btn.pack(fill="x", pady=(8, 0))

        ans_row = ctk.CTkFrame(scroll, fg_color="transparent")
        ans_row.pack(fill="x", pady=(16, 0))
        ctk.CTkLabel(ans_row, text="ОТВЕТ", font=F_LABEL, text_color=GOLD, anchor="w").pack(side="left")
        ctk.CTkButton(
            ans_row, text="Копировать", font=F_SMALL,
            fg_color=GOLD, hover_color=SURF2,
            text_color="white", border_width=0, height=24, width=120,
            command=self._copy,
        ).pack(side="right")
        self._ans_frame, self.ans_box = _textbox(scroll, height=14, readonly=True, scroll_parent=scroll)
        self._ans_frame.pack(fill="both", expand=True, pady=(10, 10))
        self.ans_box.tag_configure("bold", font=F_BOLD, foreground=CREAM)

    def _pick_folder(self):
        from tkinter import filedialog
        p = filedialog.askdirectory(title="Выберите папку с документами")
        if p:
            self._start_load(p)

    def _pick_file(self):
        from tkinter import filedialog
        p = filedialog.askopenfilename(title="Выберите файл", filetypes=_FILETYPES)
        if p:
            self._start_load(p)

    def _start_load(self, path):
        if self._indexing:
            return
        self.doc_lbl.configure(text=f"📄 {Path(path).name}", text_color=CREAM)
        self._kb            = None
        self._indexing      = True
        self._load_progress = {"pct": 0, "est_sec": None, "t_index": None, "t_start": time.time()}
        self.progress.pack(fill="x", pady=(4, 0))
        self.progress.configure(mode="determinate")
        self.progress.set(0)
        self._set_status("⏳ Индексирую...", MUTED)
        self._animate_loading()
        threading.Thread(target=self._load_w, args=(path,), daemon=True).start()

    _LOADING_DOTS = ["", ".", "..", "..."]

    def _animate_loading(self):
        if not self._indexing:
            return
        p    = self._load_progress
        dots = self._LOADING_DOTS[int(time.time() * 2) % 4]

        t_start  = p.get("t_start", time.time())
        elapsed  = time.time() - t_start
        elapsed_str = (f"{int(elapsed // 60)} мин {int(elapsed % 60)} сек"
                       if elapsed >= 60 else f"{int(elapsed)} сек")

        if p.get("t_index") and p.get("est_sec"):
            # Фаза 2 (embed, 40-99%): плавно интерполируем по времени индексации
            if p.get("pct", 0) < 100:
                phase_elapsed = time.time() - p["t_index"]
                est           = p["est_sec"]
                ratio         = min(phase_elapsed / est, 1.0)
                pct           = 40 + min(int(ratio * 58), 58)    # 40 → макс 98, линейно
                p["pct"]      = pct
            else:
                pct = 100

            # Оставшееся время считаем относительно общего старта
            total_est  = (p["t_index"] - t_start) + p["est_sec"]
            remaining  = max(total_est - elapsed, 0)
            if remaining > 1:
                eta  = (f"~{int(remaining // 60)} мин {int(remaining % 60)} сек"
                        if remaining >= 60 else f"~{int(remaining) + 1} сек")
                line2 = f"⏱  {elapsed_str}  |  осталось {eta}"
            else:
                line2 = f"⏱  {elapsed_str}"
        else:
            # Фаза 1 (чтение/конвертация, 0-40%): реальный pct из make_kb
            pct   = p.get("pct", 0)
            line2 = f"⏱  {elapsed_str}"

        self.progress.configure(mode="determinate")
        self.progress.set(pct / 100)

        line1 = f"📖  Читаю файл{dots}  {pct}%"
        text  = line1 + f"\n{line2}"
        if self._pending_q:
            text += "\n\n✉  Вопрос получен — отправится автоматически."
        self._set_ans(text, MUTED)
        self._anim_id = self.after(200, self._animate_loading)

    def _load_w(self, docs):
        try:
            kb = make_kb(docs, self._load_progress)
            self.after(0, lambda: self._load_ok(kb))
        except Exception as e:
            self.after(0, lambda err=str(e): self._load_err(err))

    def _load_ok(self, kb):
        self._kb       = kb
        self._indexing = False
        if hasattr(self, "_anim_id"):
            self.after_cancel(self._anim_id)
        self.progress.stop()
        self.progress.configure(mode="determinate")
        self.progress.set(1)
        self._set_status("✓ Готово — задайте вопрос!", GREEN)
        if self._pending_q:
            q, self._pending_q = self._pending_q, None
            self._run_ask(q)
        else:
            self._set_ans("✅  Файл прочитан и готов к работе.\n\nВведите вопрос и нажмите «Спросить».", GREEN)

    def _load_err(self, err):
        self._indexing = False
        if hasattr(self, "_anim_id"):
            self.after_cancel(self._anim_id)
        self.progress.stop()
        self.progress.configure(mode="determinate")
        self.progress.set(0)
        el = err.lower()
        if "lance" in el or "unable to copy" in el:
            hint = "Lance не работает на exFAT/FAT32/сетевых дисках.\nИндекс пишется в системный %TEMP%\\NeuroPoiskov."
        else:
            hint = err[:300]
        self._set_status("✕ Ошибка индексации — попробуйте другой файл", RED)
        self._set_ans(f"⚠  Ошибка индексации:\n\n{hint}\n\nВыберите другой файл.", RED)

    def _ask(self, _=None):
        if self._busy:
            return
        q = self.q_box.get("1.0", "end").strip()
        if not q or self._placeholder_active:
            self._set_ans("⚠  Введите вопрос.", RED)
            return
        if self._indexing:
            self._pending_q = q
            return
        if not self._kb:
            self._set_ans("⚠  Сначала выберите документ.", RED)
            return
        self._run_ask(q)

    def _run_ask(self, q):
        self._busy = True
        self.ask_btn.configure(text="Думаю...")
        self._set_ans("Думаю...", MUTED)
        threading.Thread(target=self._ask_w, args=(q,), daemon=True).start()

    def _ask_w(self, q):
        try:
            a = ask_with_kb(self._kb, q)
            self.after(0, lambda ans=a: self._ask_done(ans))
        except Exception as e:
            self.after(0, lambda err=str(e): self._ask_done(f"⚠  Ошибка: {err}"))

    def _ask_done(self, answer):
        self._busy = False
        self.ask_btn.configure(text="Спросить  →")
        print(f"[_ask_done] answer repr start: {repr(answer[:200])}")
        self._set_ans(answer, CREAM)

    def _set_status(self, t, c):
        self.status_lbl.configure(text=t, text_color=c)

    def _set_ans(self, text, color):
        text = text.replace("\\n", "\n").replace("\\r\\n", "\n")
        t = self.ans_box
        t.config(state="normal")
        t.delete("1.0", "end")
        t.config(fg=color)
        parts = re.split(r'\*\*(.+?)\*\*', text, flags=re.DOTALL)
        for i, part in enumerate(parts):
            t.insert("end", part, "bold" if i % 2 == 1 else "")
        t.config(state="disabled")

    def _copy(self):
        raw   = self.ans_box.get("1.0", "end").strip()
        clean = re.sub(r'\*\*(.+?)\*\*', r'\1', raw, flags=re.DOTALL)
        if clean:
            self.clipboard_clear()
            self.clipboard_append(clean)

    def _on_close(self):
        cleanup()
        self.destroy()


if __name__ == "__main__":
    App().mainloop()
