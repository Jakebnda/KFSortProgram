#env python3
# Kwik Fill — Sorted Stores + Blackouts + Wobbler Kit (>=10 stores) Appendix

import os
import re
import sys
import json
import time
import traceback
from math import ceil
from datetime import datetime
from pathlib import Path
from collections import defaultdict, OrderedDict

import fitz  # PyMuPDF
import tkinter as tk
from tkinter import filedialog, messagebox

# ----------------------------- DEBUG -----------------------------

DEBUG = False  # set env KWIK_DEBUG=0 to disable
if os.environ.get("KWIK_DEBUG", "").strip().lower() in ("0", "false", "no"):
    DEBUG = False
DEBUG_LOG = "kwik_debug.log"

def dbg(msg: str):
    if not DEBUG:
        return
    ts = datetime.now().strftime('%Y-%m-%d %H:%M:%S.%f')[:-3]
    line = f"[{ts}] {msg}"
    print(line, flush=True)
    try:
        with open(DEBUG_LOG, "a", encoding="utf-8") as f:
            f.write(line + "\n")
    except Exception:
        pass

def dbg_ex(prefix: str = "EXCEPTION"):
    dbg(prefix + ":\n" + traceback.format_exc())

# ----------------------------- Windows foreground helpers -----------------------------

_WIN = sys.platform.startswith("win")
if _WIN:
    import ctypes
    user32 = ctypes.windll.user32
    SW_RESTORE = 9
    SW_SHOWNORMAL = 1

    def _force_foreground(win: tk.Toplevel):
        try:
            hwnd = int(win.winfo_id())
            user32.ShowWindow(hwnd, SW_RESTORE)
            user32.ShowWindow(hwnd, SW_SHOWNORMAL)
            user32.SetForegroundWindow(hwnd)
        except Exception:
            dbg_ex("_force_foreground")

# ----------------------------- Canon + constants -----------------------------

def canon(s: str) -> str:
    if not s:
        return ""
    s = re.sub(r'\s+', ' ', s.strip())
    s = re.sub(r'^\*+|\*+$', '', s)  # trim surrounding asterisks
    return s.lower()

BLACKOUT_JSON = "blackout_config.json"
ENV_FIT_JSON  = "envelope_fit.json"

_PREDETERMINED_WOBBLERS_CANON = {
    "shelf wobbler kit; alcohol version",
    "candy; counter kit",
    "shelf wobbler kit; non-alcohol version",
    "candy; shipper kit",
}

KIT_COUNTER = "*CANDY; COUNTER KIT*"
KIT_SHIPPER = "*CANDY; SHIPPER KIT*"
KIT_ALC     = "*Shelf Wobbler Kit; Alcohol Version*"
KIT_NONALC  = "*Shelf Wobbler Kit; Non-Alcohol Version*"

PROMO_WOBBLER_ALC_CANON    = canon("Shelf Wobbler Kit; Alcohol Version")
PROMO_WOBBLER_NONALC_CANON = canon("Shelf Wobbler Kit; Non-Alcohol Version")
TYPE_SHELF_WOBBLER_CANON   = canon("Shelf Wobbler")

HEADER_STORE_RE = re.compile(r'Store:\s?[A-Z]\d{4}')

STORE_TYPE_ORDER = [
    "Alcohol Counter + Shipper",
    "Alcohol Counter",
    "Alcohol Shipper",
    "Alcohol No Counter/Shipper",
    "Non-Alcohol Counter + Shipper",
    "Non-Alcohol Counter",
    "Non-Alcohol Shipper",
    "Non-Alcohol No Counter/Shipper",
    "Counter + Shipper", "Counter", "Shipper", "No Counter/Shipper",
    ""
]
STORE_TYPE_RANK = {t: i for i, t in enumerate(STORE_TYPE_ORDER)}

SPECIAL_SIGNAGE_LABELS = [
    ("Banner Sign", ("banner sign",)),
    ("Yard Sign", ("yard sign",)),
    ("A Frame", ("a frame", "a-frame")),
    ("Bollard Cover", ("bollard cover",)),
    ("Pole Sign", ("pole sign",)),
    ("Easel Sign", ("easel sign",)),
]

SPECIAL_SIGNAGE_LABEL_ORDER = [label for label, _ in SPECIAL_SIGNAGE_LABELS]

# Page / layout constants
MARGIN_L = 72
MARGIN_R = 72
MARGIN_T = 72
MARGIN_B = 72
LEADING  = 1.2  # line-height multiplier for wrapped text

# ----------------------------- JSON init / load / save -----------------------------

def ensure_json(path: str, default_obj):
    try:
        if not os.path.exists(path):
            dbg(f"ensure_json: creating {path}")
            with open(path, "w", encoding="utf-8") as f:
                json.dump(default_obj, f, indent=4, ensure_ascii=False)
            return json.loads(json.dumps(default_obj))
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
            dbg(f"ensure_json: loaded {path} (ok)")
            return data
    except json.JSONDecodeError:
        dbg(f"ensure_json: {path} corrupted; resetting")
        try:
            with open(path, "w", encoding="utf-8") as f2:
                json.dump(default_obj, f2, indent=4, ensure_ascii=False)
        except Exception:
            dbg_ex("ensure_json reset write failed")
        return json.loads(json.dumps(default_obj))
    except Exception:
        dbg_ex("ensure_json failure")
        return json.loads(json.dumps(default_obj))

def load_blackout_config() -> dict:
    return ensure_json(BLACKOUT_JSON, {})

def save_blackout_config(cfg: dict):
    try:
        with open(BLACKOUT_JSON, "w", encoding="utf-8") as f:
            json.dump(cfg, f, indent=4, ensure_ascii=False)
        dbg(f"save_blackout_config: saved {sum(len(v) for v in cfg.values())} rules across {len(cfg)} sign types")
    except Exception:
        dbg_ex("save_blackout_config failure")

def load_envelope_fit() -> dict:
    return ensure_json(ENV_FIT_JSON, {"will_fit": [], "wont_fit": []})

def save_envelope_fit(data: dict):
    try:
        with open(ENV_FIT_JSON, "w", encoding="utf-8") as f:
            json.dump({"will_fit": data.get("will_fit", []),
                       "wont_fit": data.get("wont_fit", [])}, f, indent=4, ensure_ascii=False)
        dbg(f"save_envelope_fit: will_fit={len(data.get('will_fit', []))}, wont_fit={len(data.get('wont_fit', []))}")
    except Exception:
        dbg_ex("save_envelope_fit failure")

# ----------------------------- Tk helpers: reliable modal -----------------------------

def _center_on_screen(win: tk.Toplevel, w: int = 760, h: int = 520):
    try:
        win.update_idletasks()
        sw = win.winfo_screenwidth()
        sh = win.winfo_screenheight()
        x = max(0, (sw - w) // 2)
        y = max(0, (sh - h) // 3)
        win.geometry(f"{w}x{h}+{x}+{y}")
    except Exception:
        pass

def _show_modal(win: tk.Toplevel, parent: tk.Tk, name: str = "modal"):
    try:
        def _on_close():
            dbg(f"_show_modal[{name}]: WM_DELETE_WINDOW")
            try:
                win.grab_release()
            except Exception:
                pass
            win.destroy()
        win.protocol("WM_DELETE_WINDOW", _on_close)

        win.withdraw()
        win.update_idletasks()
        _center_on_screen(win)
        try:
            if parent.state() != "withdrawn":
                win.transient(parent)
        except Exception:
            pass

        need_hide_parent = False
        try:
            if parent.state() in ("iconic", "withdrawn"):
                if _WIN:
                    parent.attributes("-alpha", 0.0)
                parent.deiconify()
                need_hide_parent = True
        except Exception:
            pass

        win.deiconify()
        win.state("normal")
        win.lift()
        win.attributes("-topmost", True)
        win.focus_force()
        win.update_idletasks()

        if _WIN:
            _force_foreground(win)
            win.after(250, lambda: (_force_foreground(win)))

        win.wait_visibility()
        win.grab_set()

        win.after(400, lambda: win.attributes("-topmost", False))

        dbg(f"_show_modal[{name}]: mapped={win.winfo_ismapped()} viewable={win.winfo_viewable()} state={win.state()}")
        win.wait_window()
        dbg(f"_show_modal[{name}]: closed")
    finally:
        if _WIN and 'need_hide_parent' in locals() and need_hide_parent:
            try:
                parent.withdraw()
                parent.attributes("-alpha", 1.0)
            except Exception:
                pass

# ----------------------------- GUI: blackout config -----------------------------

def gui_blackout_edit(root: tk.Tk):
    dbg("gui_blackout_edit: start")
    cfg = load_blackout_config()

    win = tk.Toplevel(root)
    win.title("Blackout Configuration (Add / Edit)")

    frame = tk.Frame(win)
    frame.pack(padx=10, pady=10, fill='both', expand=True)

    tk.Label(frame, text="Sign Type").grid(row=0, column=0, padx=5, pady=5, sticky='w')
    tk.Label(frame, text="Sign Version").grid(row=0, column=1, padx=5, pady=5, sticky='w')

    entries = []

    def add_entry(preset_type="", preset_version=""):
        r = len(entries) + 1
        e1 = tk.Entry(frame, width=28); e1.insert(0, preset_type)
        e2 = tk.Entry(frame, width=64); e2.insert(0, preset_version)
        e1.grid(row=r, column=0, padx=5, pady=3, sticky="ew")
        e2.grid(row=r, column=1, padx=5, pady=3, sticky="ew")
        entries.append((e1, e2))
        dbg(f"blackout_edit: add_entry r={r} type='{preset_type}' ver_len={len(preset_version)}")

    rows = 0
    for st, vers in cfg.items():
        for v in vers:
            add_entry(st, v); rows += 1
    if rows == 0:
        add_entry()
    dbg(f"blackout_edit: prepopulated rows={rows}")

    def save_config():
        try:
            new_cfg = {}
            for e1, e2 in entries:
                st = e1.get().strip()
                sv = e2.get().strip()
                if st and sv:
                    new_cfg.setdefault(st, []).append(sv)
            save_blackout_config(new_cfg)
            messagebox.showinfo("Blackout", "Saved to blackout_config.json", parent=win)
            win.destroy()
        except Exception:
            dbg_ex("blackout_edit save_config")

    btns = tk.Frame(win); btns.pack(pady=6)
    tk.Button(btns, text="Add Row", command=add_entry).pack(side='left', padx=5)
    tk.Button(btns, text="Save", command=save_config).pack(side='left', padx=5)

    _show_modal(win, root, name="blackout_edit")

def gui_blackout_delete(root: tk.Tk):
    dbg("gui_blackout_delete: start")
    cfg = load_blackout_config()

    win = tk.Toplevel(root)
    win.title("Blackout Configuration (Delete)")

    frame = tk.Frame(win); frame.pack(padx=10, pady=10, fill='both', expand=True)
    tk.Label(frame, text="Click 'Delete' to remove a rule").grid(row=0, column=0, columnspan=3, sticky='w', pady=(0,6))
    tk.Label(frame, text="Sign Type", font=("TkDefaultFont", 9, "bold")).grid(row=1, column=0, sticky='w')
    tk.Label(frame, text="Sign Version", font=("TkDefaultFont", 9, "bold")).grid(row=1, column=1, sticky='w')

    rows = [(st, v) for st, vs in cfg.items() for v in vs]
    dbg(f"blackout_delete: rows={len(rows)}")

    r = 2
    if not rows:
        tk.Label(frame, text="(No blackout rules)").grid(row=r, column=0, columnspan=3, sticky='w')
    else:
        for st, v in rows:
            tk.Label(frame, text=st).grid(row=r, column=0, sticky='w', padx=3, pady=2)
            tk.Label(frame, text=v).grid(row=r, column=1, sticky='w', padx=3, pady=2)
            def make_del(s=st, ver=v):
                return lambda: _delete_and_refresh(cfg, s, ver, win)
            tk.Button(frame, text="Delete", command=make_del()).grid(row=r, column=2, padx=3, pady=2)
            r += 1

    tk.Button(win, text="Close", command=win.destroy).pack(pady=6)
    _show_modal(win, root, name="blackout_delete")

def _delete_and_refresh(cfg, st, ver, win):
    try:
        dbg(f"_delete_and_refresh: '{st}' -> '{ver}'")
        lst = cfg.get(st, [])
        if ver in lst:
            lst.remove(ver)
        if not lst and st in cfg:
            del cfg[st]
        save_blackout_config(cfg)
        messagebox.showinfo("Blackout", f"Deleted: {st} — {ver}", parent=win)
        win.destroy()
    except Exception:
        dbg_ex("_delete_and_refresh failure")

# ----------------------------- Envelope-Fit config (checkboxes) -----------------------------

def gui_envelope_fit(root: tk.Tk, sign_types: list) -> dict:
    dbg(f"gui_envelope_fit: start with {len(sign_types)} sign types")
    prev = load_envelope_fit()

    try:
        if prev.get("will_fit") or prev.get("wont_fit"):
            dbg("gui_envelope_fit: existing config found")
            if messagebox.askyesno("Envelope Fit", "Reuse saved envelope_fit.json?", parent=root):
                dbg("gui_envelope_fit: reuse saved config")
                return prev
    except Exception:
        dbg_ex("gui_envelope_fit reuse prompt")

    win = tk.Toplevel(root)
    win.title("Envelope Fit — Select SIGN TYPES that FIT")

    frm = tk.Frame(win); frm.pack(padx=12, pady=12, fill="both", expand=True)
    tk.Label(frm, text="Check SIGN TYPES that FIT in envelopes. Unchecked = won't fit.").grid(row=0, column=0, sticky="w")

    vars_map = {}
    will_prev = {canon(s) for s in prev.get("will_fit", [])}
    sign_types_sorted = sorted(set(sign_types), key=lambda x: x.lower())

    for i, st in enumerate(sign_types_sorted, start=1):
        v = tk.IntVar(value=1 if canon(st) in will_prev else 0)
        cb = tk.Checkbutton(frm, text=st, variable=v, onvalue=1, offvalue=0, anchor="w", width=56)
        cb.grid(row=i, column=0, sticky="w")
        vars_map[st] = v

    btns = tk.Frame(win); btns.pack(fill="x", padx=12, pady=8)
    tk.Button(btns, text="Select All",  command=lambda: [v.set(1) for v in vars_map.values()]).pack(side="left", padx=4)
    tk.Button(btns, text="Select None", command=lambda: [v.set(0) for v in vars_map.values()]).pack(side="left", padx=4)

    result = {}
    def save_and_close():
        try:
            will, wont = [], []
            for st, v in vars_map.items():
                (will if v.get()==1 else wont).append(st)
            save_envelope_fit({"will_fit": will, "wont_fit": wont})
            result["will_fit"] = will
            result["wont_fit"] = wont
            dbg(f"gui_envelope_fit: saved will={len(will)} wont={len(wont)}")
            win.destroy()
        except Exception:
            dbg_ex("gui_envelope_fit save_and_close")

    tk.Button(btns, text="Save", command=save_and_close).pack(side="right")

    _show_modal(win, root, name="envelope_fit")
    return result if result else prev

# ----------------------------- Header / classification -----------------------------

def is_header_page(text: str) -> bool:
    return bool(HEADER_STORE_RE.search(text))

def extract_store_info(text: str) -> dict:
    out = {}
    try:
        for line in text.splitlines():
            if 'Sign Type' in line:
                break
            if 'Store:' in line:
                out['store'] = line.split('Store:')[-1].strip()
            elif 'Area:' in line:
                out['area'] = line.split('Area:')[-1].strip()
            elif 'Class:' in line:
                out['class'] = line.split('Class:')[-1].strip()
            else:
                m = re.search(r'\b(NY|PA|OH|New York|Pennsylvania|Ohio)\b', line, re.I)
                if m and 'location' not in out:
                    nm = m.group().upper()
                    map_ = {'NEW YORK': 'NY', 'PENNSYLVANIA': 'PA', 'OHIO': 'OH'}
                    out['location'] = map_.get(nm, nm)
    except Exception:
        dbg_ex("extract_store_info")
    return out

def clean_text_for_kits(text: str) -> str:
    t = re.sub(r'\s+', ' ', text)
    t = re.sub(r'\s*\*\s*C\s*A\s*N\s*D\s*Y\s*;\s*C\s*O\s*U\s*N\s*T\s*E\s*R\s*K\s*I\s*T\s*\*', KIT_COUNTER, t, flags=re.I)
    t = re.sub(r'\s*\*\s*C\s*A\s*N\s*D\s*Y\s*;\s*S\s*H\s*I\s*P\s*P\s*E\s*R\s*K\s*I\s*T\s*\*',   KIT_SHIPPER, t, flags=re.I)
    t = re.sub(r'\s*\*\s*S\s*h\s*e\s*l\s*f\s*\s*W\s*o\s*b\s*b\s*l\s*e\s*r\s*\s*K\s*i\s*t\s*;\s*A\s*l\s*c\s*o\s*h\s*o\s*l\s*\s*V\s*e\s*r\s*s\s*i\s*o\s*n\s*\*', KIT_ALC, t, flags=re.I)
    t = re.sub(r'\s*\*\s*S\s*h\s*e\s*l\s*f\s*\s*W\s*o\s*b\s*b\s*l\s*e\s*r\s*\s*K\s*i\s*t\s*;\s*N\s*o\s*n\s*-\s*A\s*l\s*c\s*o\s*h\s*o\s*l\s*\s*V\s*e\s*r\s*s\s*i\s*o\s*n\s*\*', KIT_NONALC, t, flags=re.I)
    return t

def classify_store(accum_text: str) -> dict:
    t = clean_text_for_kits(accum_text)
    is_counter    = KIT_COUNTER in t
    is_shipper    = KIT_SHIPPER in t
    is_alcohol    = KIT_ALC in t
    is_nonalcohol = (KIT_NONALC in t) and not is_alcohol

    alc = 'Alcohol' if is_alcohol else ('Non-Alcohol' if is_nonalcohol else '')
    if is_counter and is_shipper:
        store_type = f'{alc} Counter + Shipper'
    elif is_counter:
        store_type = f'{alc} Counter'
    elif is_shipper:
        store_type = f'{alc} Shipper'
    else:
        store_type = f'{alc} No Counter/Shipper'

    return {
        'store_type': store_type.strip(),
        'is_counter': is_counter,
        'is_shipper': is_shipper,
        'is_alcohol': is_alcohol,
        'is_non_alcohol': is_nonalcohol
    }

# ----------------------------- Geometry-aware rows -----------------------------

def detect_columns(page):
    try:
        rs = page.search_for('Sign Type')
        rp = page.search_for('Promotion Name')
        rq = page.search_for('Qty Ordered')
        rect_sign  = rs[0] if rs else None
        rect_promo = rp[0] if rp else None
        rect_qty   = rq[0] if rq else None
        w = page.rect.width
        if not rect_promo:
            rect_promo = fitz.Rect(w*0.25, 0, w*0.55, 20)
        if not rect_qty:
            rect_qty   = fitz.Rect(w*0.80, 0, w*0.95, 20)
        header_bottom = max([r.y1 for r in [rect_sign, rect_promo, rect_qty] if r] or [0])
        return {
            'x_promo_left': rect_promo.x0,
            'x_qty_left': rect_qty.x0,
            'header_bottom': header_bottom
        }
    except Exception:
        dbg_ex("detect_columns")
        return {'x_promo_left': page.rect.width*0.25, 'x_qty_left': page.rect.width*0.80, 'header_bottom': 0}

def iter_rows(page, y_min, y_max):
    try:
        words = page.get_text('words')  # x0,y0,x1,y1,txt,block,line,wordno
    except Exception:
        dbg_ex("iter_rows get_text words")
        words = []

    rows = defaultdict(list)
    for x0,y0,x1,y1,txt,blk,ln,wn in words:
        if y0 < y_min or y1 > y_max:
            continue
        rows[(blk, ln, round(y0, 1))].append((x0,y0,x1,y1,txt,wn))
    cols = detect_columns(page)
    x_prom = cols['x_promo_left']; x_qty = cols['x_qty_left']
    for key in sorted(rows.keys(), key=lambda k: k[2]):
        parts = sorted(rows[key], key=lambda t: t[-1])
        left, mid, right = [], [], []
        x0s, y0s, x1s, y1s = [], [], [], []
        for x0,y0,x1,y1,txt,wn in parts:
            xc = 0.5*(x0+x1)
            if xc < x_prom:
                left.append(txt)
            elif xc < x_qty:
                mid.append(txt)
            else:
                right.append(txt)
            x0s.append(x0); y0s.append(y0); x1s.append(x1); y1s.append(y1)
        rect = fitz.Rect(min(x0s), min(y0s), max(x1s), max(y1s))
        yield {
            'type_text': ' '.join(left).strip(),
            'promo_text': ' '.join(mid).strip(),
            'qty_text': ' '.join(right).strip(),
            'rect': rect
        }

# ----------------------------- Blackout / Highlight / Annotation -----------------------------

def is_predetermined_wobbler(promo: str) -> bool:
    return canon(promo) in _PREDETERMINED_WOBBLERS_CANON

def blackout_rows_on_page(page, blackout_cfg):
    try:
        if not blackout_cfg:
            return
        cols = detect_columns(page)
        y_min = cols['header_bottom'] + 2
        y_max = page.rect.y1 - 36
        canon_map = {canon(st): {canon(v) for v in vs} for st,vs in blackout_cfg.items()}
        last_type = None
        count = 0
        for row in iter_rows(page, y_min, y_max):
            st_this = canon(row['type_text'])
            st = st_this or last_type
            pv = canon(row['promo_text'])
            if st_this:
                last_type = st_this
            if not st or not pv:
                continue
            if st in canon_map and pv in canon_map[st]:
                page.draw_rect(row['rect'], color=(0,0,0), fill=(0,0,0), width=0)
                count += 1
        if count and DEBUG:
            dbg(f"blackout_rows_on_page: blacked_out_rows={count}")
    except Exception:
        dbg_ex("blackout_rows_on_page")

def highlight_keyword(page, needle, color):
    """
    Draws semi-transparent fill overlays (no stroke) over all matches of `needle`.
    Uses 60% opacity for clearer highlighting across all PDF viewers.
    """
    try:
        quads = page.search_for(needle, quads=True)
    except TypeError:
        # Fallback for environments without quads=True support
        quads = [fitz.Quad(r) for r in page.search_for(needle)]

    for q in quads:
        page.draw_quad(
            q,
            color=None,          # <-- no stroke
            width=0,             # <-- ensure no stroke width
            fill=color,          # RGB tuple e.g. (0.68, 0.85, 0.90)
            fill_opacity=0.60,   # <-- stronger highlight
            overlay=True
        )

def annotate_wobbler_kit(page, kit_name: str):
    try:
        if not kit_name:
            return
        cols = detect_columns(page)
        y_min = cols['header_bottom'] + 2
        y_max = page.rect.y1 - 36
        x_left_type = 12
        added = 0
        for row in iter_rows(page, y_min, y_max):
            if canon(row['type_text']) == TYPE_SHELF_WOBBLER_CANON:
                x = max(x_left_type, row['rect'].x0 + 4)
                y = row['rect'].y1 + 7
                if y < (page.rect.y1 - 8):
                    page.insert_text((x, y), f"Kit: {kit_name}", fontsize=8, color=(0.2,0.2,0.2))
                    added += 1
        if added and DEBUG:
            dbg(f"annotate_wobbler_kit: wrote '{kit_name}' x{added}")
    except Exception:
        dbg_ex("annotate_wobbler_kit")

def blackout_nonalc_wobbler_row_on_page(page):
    try:
        cols = detect_columns(page)
        y_min = cols['header_bottom'] + 2
        y_max = page.rect.y1 - 36
        n = 0
        last_type = None
        for row in iter_rows(page, y_min, y_max):
            st_this = canon(row['type_text'])
            if st_this:
                last_type = st_this
            st = last_type
            pv = canon(row['promo_text'])
            if not st or not pv:
                continue
            is_nonalc = (
                pv == PROMO_WOBBLER_NONALC_CANON
                or ('non' in pv and 'alcohol' in pv and 'wobbler' in pv)
            )
            if st == TYPE_SHELF_WOBBLER_CANON and is_nonalc:
                page.draw_rect(row['rect'], color=(0,0,0), fill=(0,0,0), width=0)
                n += 1
        if n and DEBUG:
            dbg(f"blackout_nonalc_wobbler_row_on_page: blacked_out={n}")
    except Exception:
        dbg_ex("blackout_nonalc_wobbler_row_on_page")

# ----------------------------- Store indexing + item extraction -----------------------------

def index_stores(doc: fitz.Document):
    dbg("index_stores: start")
    stores = []
    current = None
    accum = ""
    meta = {}
    try:
        for i, page in enumerate(doc):
            if i % 25 == 0:
                dbg(f"index_stores: at page {i}")
            try:
                text = page.get_text('text')
            except Exception:
                dbg_ex(f"index_stores get_text page {i}"); continue
            if not text.strip():
                continue
            if is_header_page(text):
                if current:
                    cls = classify_store(accum)
                    store_name = meta.get('store', f'UNKNOWN_{i}')
                    store_cls  = meta.get('class', '')
                    stores.append({
                        'store_id':   f"{store_name}|{store_cls}",
                        'store_name': store_name,
                        'store_type': cls['store_type'],
                        'location':   meta.get('location', ''),
                        'class':      store_cls,
                        'pages':      current['pages'],
                        'meta':       meta.copy()
                    })
                meta = extract_store_info(text)
                current = {'pages':[i]}
                accum = text + " "
            else:
                accum += text + " "
                if current:
                    current['pages'].append(i)

        if current:
            cls = classify_store(accum)
            store_name = meta.get('store', 'UNKNOWN_END')
            store_cls  = meta.get('class', '')
            stores.append({
                'store_id':   f"{store_name}|{store_cls}",
                'store_name': store_name,
                'store_type': cls['store_type'],
                'location':   meta.get('location', ''),
                'class':      store_cls,
                'pages':      current['pages'],
                'meta':       meta.copy()
            })
    except Exception:
        dbg_ex("index_stores")
    dbg(f"index_stores: found stores={len(stores)}")
    return stores

def extract_items_from_pages(doc: fitz.Document, pages):
    items = []
    last_type = None
    promo_buf = []
    try:
        for p in pages:
            page = doc[p]
            cols = detect_columns(page)
            y_min = cols['header_bottom'] + 2
            y_max = page.rect.y1 - 36
            for row in iter_rows(page, y_min, y_max):
                t = row['type_text'].strip()
                pr = row['promo_text'].strip()
                qt = row['qty_text'].strip()
                if t:
                    last_type = t
                if 'Sign Type Total' in (t + ' ' + pr):
                    last_type = None
                    promo_buf = []
                    continue
                if qt.isdigit() and last_type:
                    full_promo = ' '.join([p for p in (promo_buf + [pr]) if p]).strip()
                    if full_promo:
                        items.append({'type': last_type, 'promo': full_promo, 'qty': int(qt)})
                    promo_buf = []
                else:
                    if pr and not re.search(r'[a-z]+://|www\.', pr, re.I):
                        promo_buf.append(pr)
    except Exception:
        dbg_ex("extract_items_from_pages")
    return items

# ----------------------------- Wobbler kit grouping (post-determined) -----------------------------

def group_wobbler_kits(stores, min_stores=10):
    dbg("group_wobbler_kits: start")
    rep_text = {}
    combos = OrderedDict()
    try:
        for s in stores:
            items = s.get('items', [])
            wob = []
            for it in items:
                if canon(it['type']) != TYPE_SHELF_WOBBLER_CANON:
                    continue
                if is_predetermined_wobbler(it['promo']):
                    continue
                cp = canon(it['promo'])
                rep_text.setdefault(cp, it['promo'])
                wob.append((cp, it['qty']))
            if len(wob) <= 1:
                continue
            key = tuple(sorted(wob))
            combos.setdefault(key, []).append((s['store_id'], s['store_name']))

        kits = []
        idx = 1
        for key, store_list in combos.items():
            if len(store_list) < min_stores:
                continue
            items_disp = [{'promo': rep_text[p], 'qty': q} for (p,q) in key]
            store_ids  = [sid for sid,_ in store_list]
            store_names= sorted([nm for _,nm in store_list])
            kits.append({
                'kit_name': f'{idx}',
                'items': items_disp,
                'stores': store_names,
                'store_ids': store_ids,
                'store_count': len(store_list),
            })
            idx += 1
        kits.sort(key=lambda k: k['store_count'], reverse=True)
        kit_by_store_id = {}
        for kit in kits:
            for sid in kit['store_ids']:
                kit_by_store_id[sid] = kit['kit_name']
        dbg(f"group_wobbler_kits: kits={len(kits)}")
        return kits, kit_by_store_id
    except Exception:
        dbg_ex("group_wobbler_kits")
        return [], {}

# ----------------------------- Envelope-fit helpers -----------------------------

def unique_sign_types(stores) -> list:
    seen = OrderedDict()
    for s in stores:
        for it in s.get('items', []):
            st = it.get('type', '').strip()
            if st and canon(st) not in seen:
                seen[canon(st)] = st
    out = list(seen.values())
    dbg(f"unique_sign_types: {len(out)} types")
    return out

def compute_envelope_fit(stores, fit_cfg: dict):
    will = {canon(x) for x in fit_cfg.get("will_fit", [])}
    dbg(f"compute_envelope_fit: will_fit={len(will)}")
    for s in stores:
        items = s.get('items', [])
        if not items:
            s['fits_envelope'] = False
            continue
        s['fits_envelope'] = all(canon(it.get('type','')) in will for it in items)
    fits = [s for s in stores if s['fits_envelope']]
    not_fits = [s for s in stores if not s['fits_envelope']]
    dbg(f"compute_envelope_fit: fits={len(fits)} not_fits={len(not_fits)}")
    return fits, not_fits

# ----------------------------- Wrapped text helpers (no overflow) -----------------------------

def _line_height(fontsize: float, leading: float = LEADING) -> float:
    return fontsize * leading

def draw_wrapped_text(out_doc, page, x, y, text, max_width, fontsize=12, leading=LEADING, color=(0,0,0)):
    x_start = x
    paragraphs = str(text).splitlines() if text else [""]
    for para in paragraphs:
        words = para.split(" ")
        line = ""
        while words:
            peek = (line + (" " if line else "") + words[0]).strip()
            w = fitz.get_text_length(peek, fontname="helv", fontsize=fontsize)
            if w <= max_width:
                line = peek
                words.pop(0)
                if words:
                    continue
            else:
                if not line:
                    forced = words.pop(0)
                    trimmed = _ellipsize_to_width(forced, max_width, fontsize)
                    if trimmed:
                        forced = trimmed
                    line = forced
                else:
                    # will flush the current line without consuming the next word
                    pass
            if y > page.rect.y1 - MARGIN_B - _line_height(fontsize, leading):
                page = out_doc.new_page()
                x = x_start
                y = MARGIN_T
            page.insert_text((x, y), line, fontsize=fontsize, color=color)
            y += _line_height(fontsize, leading)
            line = ""
        if line:
            if y > page.rect.y1 - MARGIN_B - _line_height(fontsize, leading):
                page = out_doc.new_page()
                x = x_start
                y = MARGIN_T
            page.insert_text((x, y), line, fontsize=fontsize, color=color)
            y += _line_height(fontsize, leading)
    return page, y

def draw_heading(out_doc, page, text, y, fontsize=18):
    max_width = page.rect.x1 - MARGIN_R - MARGIN_L
    page, y = draw_wrapped_text(out_doc, page, MARGIN_L, y, text, max_width, fontsize=fontsize, leading=1.15)
    return page, y

def draw_label_value(out_doc, page, label, value, y, fontsize=12):
    max_width = page.rect.x1 - MARGIN_R - MARGIN_L
    page, y = draw_wrapped_text(out_doc, page, MARGIN_L, y, f"{label}: {value}", max_width, fontsize=fontsize)
    return page, y

def draw_bullets(out_doc, page, items, y, indent=16, fontsize=11):
    max_width = page.rect.x1 - MARGIN_R - MARGIN_L - indent
    for it in items:
        page, y = draw_wrapped_text(out_doc, page, MARGIN_L + indent, y, f"- {it}", max_width, fontsize=fontsize)
    return page, y

def _ellipsize_to_width(text, max_width, fontsize):
    """Trim with ellipsis so it fits the width (for short store codes this is usually a no-op)."""
    if fitz.get_text_length(text, fontname="helv", fontsize=fontsize) <= max_width:
        return text
    if fitz.get_text_length("…", fontname="helv", fontsize=fontsize) > max_width:
        return ""
    t = text
    while t and fitz.get_text_length(t + "…", fontname="helv", fontsize=fontsize) > max_width:
        t = t[:-1]
    return t + "…"

def draw_multicolumn_list(out_doc, page, items, y, columns=4, fontsize=10, col_gap=20, leading=1.15, header_on_new_pages=None, bullet="- "):
    """
    Render items in a strict grid across columns and rows without mid-column pagination.
    Prevents orphans like a single item on a new page by laying out per-page chunks.
    Repeats a small header (e.g., 'Stores (cont.)') when a new page is started.
    Returns (page, y_end).
    """
    width_total = page.rect.x1 - MARGIN_R - MARGIN_L
    col_width = (width_total - (columns - 1) * col_gap) / columns
    line_h = _line_height(fontsize, leading)

    i = 0
    while i < len(items):
        # rows available on this page starting at y
        rows_fit = int((page.rect.y1 - MARGIN_B - y) // line_h)
        if rows_fit <= 0:
            page = out_doc.new_page()
            y = MARGIN_T
            if header_on_new_pages:
                page, y = draw_wrapped_text(out_doc, page, MARGIN_L, y, header_on_new_pages, width_total, fontsize=12)
            continue

        per_page = rows_fit * columns
        chunk = items[i:i+per_page]
        # lay out row-major across columns (keeps all columns aligned)
        for r in range(rows_fit):
            for c in range(columns):
                idx = r + c*rows_fit
                if idx >= len(chunk):
                    continue
                name = chunk[idx]
                txt = (bullet + name) if bullet else name
                txt = _ellipsize_to_width(txt, col_width, fontsize)
                x = MARGIN_L + c * (col_width + col_gap)
                y_line = y + r * line_h
                page.insert_text((x, y_line), txt, fontsize=fontsize, color=(0,0,0))

        y += rows_fit * line_h
        i += len(chunk)
        if i < len(items):
            page = out_doc.new_page()
            y = MARGIN_T
            if header_on_new_pages:
                page, y = draw_wrapped_text(out_doc, page, MARGIN_L, y, header_on_new_pages, width_total, fontsize=12)
    return page, y

# ----------------------------- Store helpers -----------------------------

_STORE_NUM_RE = re.compile(r'[A-Z]\d{4}')

def extract_store_number(store: dict) -> str:
    for key in ("store_name", "store_id"):
        value = store.get(key)
        if not value:
            continue
        m = _STORE_NUM_RE.search(str(value))
        if m:
            return m.group(0)
    meta = store.get('meta') or {}
    value = meta.get('store', '')
    if value:
        m = _STORE_NUM_RE.search(str(value))
        if m:
            return m.group(0)
    return store.get('store_name', '') or ''

def detect_special_box_label(store: dict):
    items = store.get('items') or []
    if not items:
        return None
    types = {canon(it.get('type', '')) for it in items if it.get('type')}
    for label, needles in SPECIAL_SIGNAGE_LABELS:
        for needle in needles:
            for st in types:
                if needle in st:
                    return label
    return None

def render_store_group(out_doc, src_doc, stores, blackout_cfg, kit_by_store_id, section_title):
    if not stores:
        return
    page = out_doc.new_page()
    page, _ = draw_heading(out_doc, page, section_title, MARGIN_T, fontsize=18)
    current_type = None
    for store in stores:
        kit_name = kit_by_store_id.get(store['store_id'])
        store_type = store.get('store_type', '')
        if store_type != current_type:
            current_type = store_type
            heading = out_doc.new_page()
            heading_text = f"Store Type: {current_type}" if current_type else "Store Type: (Unspecified)"
            heading, _ = draw_heading(out_doc, heading, heading_text, MARGIN_T, fontsize=16)
        for p in store['pages']:
            out_doc.insert_pdf(src_doc, from_page=p, to_page=p)
            pg = out_doc[-1]
            blackout_rows_on_page(pg, blackout_cfg)
            if store.get('drop_nonalc_wobbler'):
                blackout_nonalc_wobbler_row_on_page(pg)
            highlight_keyword(pg, KIT_COUNTER, (0.68, 0.85, 0.90))
            highlight_keyword(pg, KIT_SHIPPER, (1.00, 0.71, 0.76))
            if kit_name:
                annotate_wobbler_kit(pg, kit_name)

# ----------------------------- Safe save -----------------------------

def ensure_unique_path(path: Path) -> Path:
    if not path.exists():
        return path
    n = 1
    while True:
        cand = path.with_name(f"{path.stem} ({n}){path.suffix}")
        if not cand.exists():
            return cand
        n += 1

# ----------------------------- NEW: per-store detection of both wobblers -----------------------------

def store_should_drop_nonalc(items) -> bool:
    has_alc = False
    has_non = False
    for it in items or []:
        if canon(it.get('type','')) != TYPE_SHELF_WOBBLER_CANON:
            continue
        p = canon(it.get('promo',''))
        if p == PROMO_WOBBLER_ALC_CANON:
            has_alc = True
        elif p == PROMO_WOBBLER_NONALC_CANON:
            has_non = True
        if has_alc and has_non:
            return True
    return False

# ----------------------------- Main processing -----------------------------

def process_pdf_sorted_with_kits_and_envelopes(input_file, output_file, root: tk.Tk):
    dbg(f"process: start input='{input_file}' output='{output_file}'")

    _ = load_blackout_config()
    _ = load_envelope_fit()

    try:
        with open(BLACKOUT_JSON, 'r', encoding='utf-8') as f:
            blackout_cfg = json.load(f)
        dbg(f"process: blackout rules types={len(blackout_cfg)}")
    except Exception:
        dbg_ex("process: load blackout config")
        blackout_cfg = {}

    t0 = time.time()

    try:
        with fitz.open(input_file) as doc:
            dbg(f"process: pdf pages={len(doc)}")
            # 1) Index stores and extract items
            stores = index_stores(doc)
            for idx, s in enumerate(stores):
                if idx % 20 == 0:
                    dbg(f"process: extracting items for store {idx+1}/{len(stores)}")
                s['items'] = extract_items_from_pages(doc, s['pages'])
                s['drop_nonalc_wobbler'] = store_should_drop_nonalc(s['items'])
                if DEBUG and s['drop_nonalc_wobbler']:
                    dbg(f"store '{s.get('store_name','?')}' -> drop Non-Alc Wobbler row")

            no_order_stores = [s for s in stores if not s.get('items')]
            stores_with_items = [s for s in stores if s.get('items')]
            dbg(f"process: no_order_stores={len(no_order_stores)} with_items={len(stores_with_items)}")

            # 2) Envelope-Fit GUI
            sign_types = unique_sign_types(stores_with_items)
            if not sign_types:
                messagebox.showerror("Envelope Fit", "No sign types found in this PDF.", parent=root)
                dbg("process: abort no sign types")
                return
            fit_cfg = gui_envelope_fit(root, sign_types)

            # 3) Compute per-store envelope fit
            fits, not_fits = compute_envelope_fit(stores_with_items, fit_cfg)

            # 4) Sort each bucket
            def store_sort_key(s):
                rank = STORE_TYPE_RANK.get(s['store_type'], 999)
                return (rank, s['location'], s['store_name'])
            fits_sorted = sorted(fits, key=store_sort_key)
            not_fits_sorted = sorted(not_fits, key=store_sort_key)
            box_special = {label: [] for label in SPECIAL_SIGNAGE_LABEL_ORDER}
            box_general = []
            for store in not_fits_sorted:
                label = detect_special_box_label(store)
                if label:
                    box_special[label].append(store)
                else:
                    box_general.append(store)
            box_total = len(not_fits_sorted)
            dbg_special = {label: len(box_special[label]) for label in SPECIAL_SIGNAGE_LABEL_ORDER if box_special[label]}
            dbg(f"process: fits_sorted={len(fits_sorted)} box_total={box_total} special={dbg_special}")

            # 5) Wobbler kits
            kits, kit_by_store_id = group_wobbler_kits(stores_with_items, min_stores=10)

            # 6) Build output
            out = fitz.open()

            no_order_display = []
            seen_no_order = set()
            for store in no_order_stores:
                candidate = extract_store_number(store).strip()
                if not candidate:
                    candidate = (store.get('store_name') or '').strip()
                if candidate and candidate not in seen_no_order:
                    seen_no_order.add(candidate)
                    no_order_display.append(candidate)
            no_order_line = ", ".join(no_order_display) if no_order_display else "None"

            # Summary page
            cover = out.new_page()
            y = MARGIN_T
            cover, y = draw_heading(out, cover, "Order Packaging Summary", y, fontsize=18)
            cover, y = draw_wrapped_text(out, cover, MARGIN_L, y, f"No Order Stores: {no_order_line}", cover.rect.x1 - MARGIN_R - MARGIN_L, fontsize=12)
            cover, y = draw_label_value(out, cover, "Envelope-Fit Stores", len(fits_sorted), y, fontsize=12)
            cover, y = draw_label_value(out, cover, "Box Stores", box_total, y, fontsize=12)
            for label in SPECIAL_SIGNAGE_LABEL_ORDER:
                stores_for_label = box_special[label]
                if stores_for_label:
                    cover, y = draw_label_value(out, cover, f"{label} Box Stores", len(stores_for_label), y, fontsize=11)
            cover, y = draw_label_value(out, cover, "Wobbler Kits (10+ stores)", len(kits), y, fontsize=12)
            cover, y = draw_wrapped_text(out, cover, MARGIN_L, y, "Envelope-Fit rule: a store fits only if ALL item Sign Types are marked as 'will fit'.", cover.rect.x1 - MARGIN_R - MARGIN_L, fontsize=10)
            if DEBUG:
                cover, y = draw_wrapped_text(out, cover, MARGIN_L, y, f"DEBUG log: {os.path.abspath(DEBUG_LOG)}", cover.rect.x1 - MARGIN_R - MARGIN_L, fontsize=8)

            # Envelope-friendly orders
            render_store_group(out, doc, fits_sorted, blackout_cfg, kit_by_store_id, "ENVELOPE-FRIENDLY ORDERS")

            # Box orders (special signage categories first)
            for label in SPECIAL_SIGNAGE_LABEL_ORDER:
                stores_for_label = box_special[label]
                if stores_for_label:
                    render_store_group(out, doc, stores_for_label, blackout_cfg, kit_by_store_id, f"BOX STORES — {label}")

            # Remaining box orders
            render_store_group(out, doc, box_general, blackout_cfg, kit_by_store_id, "BOX STORES")

            # Wobbler kits appendix (summary + details)
            cover2 = out.new_page()
            y = MARGIN_T
            cover2, y = draw_heading(out, cover2, "Wobbler Kits (Post-Determined, 10+ stores)", y, fontsize=18)
            excluded_list = ", ".join(sorted(_PREDETERMINED_WOBBLERS_CANON))
            cover2, y = draw_wrapped_text(out, cover2, MARGIN_L, y, f"Excluded (predetermined): {excluded_list}", cover2.rect.x1 - MARGIN_R - MARGIN_L, fontsize=10)
            if kits:
                for kit in kits:
                    cover2, y = draw_wrapped_text(out, cover2, MARGIN_L, y, f"{kit['kit_name']}: {kit['store_count']} stores", cover2.rect.x1 - MARGIN_R - MARGIN_L, fontsize=12)
            else:
                cover2, y = draw_wrapped_text(out, cover2, MARGIN_L, y, "No kits reached the 10+ store threshold.", cover2.rect.x1 - MARGIN_R - MARGIN_L, fontsize=12)

            for kit in kits:
                page = out.new_page()
                y = MARGIN_T
                page, y = draw_heading(out, page, f"{kit['kit_name']}  —  {kit['store_count']} stores", y, fontsize=16)

                # Items
                page, y = draw_wrapped_text(out, page, MARGIN_L, y, "Items:", page.rect.x1 - MARGIN_R - MARGIN_L, fontsize=12)
                item_lines = [f"{it['promo']}  (qty {it['qty']})" for it in kit['items']]
                page, y = draw_bullets(out, page, item_lines, y, indent=16, fontsize=11)

                if y > page.rect.y1 - MARGIN_B - _line_height(12):
                    page = out.new_page(); y = MARGIN_T

                # Stores (4 columns, strict grid, repeat header on new pages)
                page, y = draw_wrapped_text(out, page, MARGIN_L, y, "Stores:", page.rect.x1 - MARGIN_R - MARGIN_L, fontsize=12)
                page, y = draw_multicolumn_list(out_doc=out, page=page, items=kit['stores'], y=y,
                                                columns=4, fontsize=10, col_gap=20, leading=1.1,
                                                header_on_new_pages="Stores (cont.)", bullet="- ")

            # Save
            out_path = ensure_unique_path(Path(output_file))
            out.save(out_path.as_posix()); out.close()
    except Exception:
        dbg_ex("process: outer")
        messagebox.showerror("Error", "A fatal error occurred. See kwik_debug.log for details.", parent=root)
        return

    dt = time.time() - t0
    dbg(f"process: done in {dt:.2f}s")
    messagebox.showinfo("Complete",
                        f"Saved: {out_path}\nEnvelope-Fit Stores: {len(fits_sorted)}\nBox Stores: {box_total}\nNo Order Stores: {len(no_order_stores)}\nWobbler Kits (10+): {len(kits)}",
                        parent=root)

# ----------------------------- App wiring -----------------------------

def main():
    if DEBUG:
        try:
            with open(DEBUG_LOG, "w", encoding="utf-8") as f:
                f.write(f"KWIK DEBUG LOG — {datetime.now().isoformat()}\n")
        except Exception:
            pass
        dbg("main: debug enabled")

    root = tk.Tk(className="KwikFill")
    try:
        root.withdraw()
        root.update_idletasks()
    except Exception:
        dbg_ex("main root init")

    try:
        input_file = filedialog.askopenfilename(
            title="Select the Kwik-Fill PDF",
            filetypes=[("PDF files", "*.pdf")],
            parent=root
        )
        dbg(f"main: input_file='{input_file}'")
        if not input_file:
            messagebox.showerror("Error", "No input file selected.", parent=root)
            root.destroy(); return

        output_file = filedialog.asksaveasfilename(
            title="Save processed PDF as",
            defaultextension=".pdf",
            filetypes=[("PDF files", "*.pdf")],
            parent=root
        )
        dbg(f"main: output_file='{output_file}'")
        if not output_file:
            messagebox.showerror("Error", "No output file selected.", parent=root)
            root.destroy(); return

        try:
            if messagebox.askyesno("Blackout", "Delete existing blackout entries?", parent=root):
                gui_blackout_delete(root)
            if messagebox.askyesno("Blackout", "Add / edit blackout entries now?", parent=root):
                gui_blackout_edit(root)
        except Exception:
            dbg_ex("main blackout prompts")

        process_pdf_sorted_with_kits_and_envelopes(input_file, output_file, root)
    except Exception:
        dbg_ex("main outer")
        messagebox.showerror("Error", "Unexpected error. See kwik_debug.log.", parent=root)
    finally:
        try:
            root.destroy()
        except Exception:
            pass

if __name__ == "__main__":
    main()
