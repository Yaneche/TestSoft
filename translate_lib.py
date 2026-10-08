"""Translation cache with Google (free endpoint) + MyMemory fallback."""
import json, os, re, time, hashlib, urllib.request, urllib.parse

CACHE_FILE = "/workspace/translation_cache.json"
_lock_cache = {}

def load_cache():
    global _lock_cache
    if not _lock_cache:
        if os.path.exists(CACHE_FILE):
            with open(CACHE_FILE, "r", encoding="utf-8") as f:
                _lock_cache = json.load(f)
        else:
            _lock_cache = {}
    return _lock_cache

def save_cache():
    tmp = CACHE_FILE + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(_lock_cache, f, ensure_ascii=False)
    os.replace(tmp, CACHE_FILE)

def norm(s):
    s = re.sub(r"\s+", " ", s).strip()
    return s

def _google_batch(texts):
    # join with newline sentinel; google keeps \n between segments
    q = "\n".join(texts)
    url = ("https://translate.googleapis.com/translate_a/single?client=gtx"
           "&sl=en&tl=ru&dt=t&q=" + urllib.parse.quote(q))
    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
    with urllib.request.urlopen(req, timeout=20) as r:
        data = json.loads(r.read().decode("utf-8"))
    out = "".join(seg[0] or "" for seg in data[0])
    parts = out.split("\n")
    # google sometimes merges empty lines; pad
    while len(parts) < len(texts):
        parts.append("")
    return parts[: len(texts)]

def _mymemory(text):
    url = ("https://api.mymemory.translated.net/get?q=" + urllib.parse.quote(text) +
           "&langpair=en|ru")
    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
    with urllib.request.urlopen(req, timeout=20) as r:
        data = json.loads(r.read().decode("utf-8"))
    return data["responseData"]["translatedText"]

def translate_many(texts):
    """Return list of Russian strings for the given English texts (batched, cached)."""
    cache = load_cache()
    results = [None] * len(texts)
    todo_idx = []
    seen = {}
    for i, t in enumerate(texts):
        key = norm(t)
        if not key:
            results[i] = t
            continue
        if key in cache:
            results[i] = cache[key]
        elif key in seen:
            todo_idx.append(i)
            seen[key].append(i)
        else:
            seen[key] = [i]
            todo_idx.append(i)
    uniq = [k for k in seen]
    B = 40
    for start in range(0, len(uniq), B):
        batch = uniq[start:start + B]
        got = None
        for attempt in range(3):
            try:
                got = _google_batch(batch)
                break
            except Exception:
                time.sleep(1 + attempt)
        if got is None:
            got = []
            for btxt in batch:
                g = None
                for attempt in range(2):
                    try:
                        g = _mymemory(btxt)
                        break
                    except Exception:
                        time.sleep(1)
                got.append(g if g else btxt)
                time.sleep(0.15)
        for btxt, tr in zip(batch, got):
            tr = tr.strip() or btxt
            cache[btxt] = tr
            for i in seen[btxt]:
                results[i] = tr
        save_cache()
    return [r if r is not None else t for r, t in zip(results, texts)]
