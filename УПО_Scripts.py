# -*- coding: utf-8 -*-
import os, sys, json, tkinter as tk
from tkinter import filedialog
import customtkinter as ctk
import tksvg
from PIL import Image
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from datetime import datetime

# ── paths ─────────────────────────────────────────────────────────────────────
BASE_DIR   = os.path.dirname(os.path.abspath(sys.argv[0]))
BUNDLE_DIR = getattr(sys, '_MEIPASS', BASE_DIR)
SVG_PATH   = os.path.join(BUNDLE_DIR, 'Security_Police_of_Ukraine_emblem.svg')
ICO_PATH   = os.path.join(BUNDLE_DIR, 'patrol-polycar.ico')
PNG_PATH   = os.path.join(BUNDLE_DIR, 'patrol-polycar.png')
_APP_DATA  = os.path.join(os.environ.get('APPDATA', BASE_DIR), 'УПО Scripts')
os.makedirs(_APP_DATA, exist_ok=True)
CFG_PATH   = os.path.join(_APP_DATA, 'upo_config.json')

UNITS = ['НР 10', 'НР 12', 'НР 13', 'НР 15', 'НР Умань 1', 'НР Умань 2']

# ── palette ───────────────────────────────────────────────────────────────────
BG       = '#151F2E'
CARD     = '#1B2A40'
BORDER   = '#263A58'
BTN      = '#1A4F96'
BTN_HOV  = '#133D7A'
GOLD     = '#C8A227'
WHITE    = '#FFFFFF'
TEXT     = '#D4DFF0'
MUTED    = '#6A8DB0'
ENTRY_BG = '#0E1724'
SUCCESS  = '#3C9E52'
WARN     = '#D89A00'
ERR      = '#C83030'
ALT_ROW  = 'B8C8E0'   # darker alternating row colour

ctk.set_appearance_mode('dark')

# ── config ────────────────────────────────────────────────────────────────────
def load_cfg():
    if os.path.exists(CFG_PATH):
        try:
            with open(CFG_PATH, 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception:
            pass
    return {'car_numbers': {u: '' for u in UNITS}}

def save_cfg(cfg):
    with open(CFG_PATH, 'w', encoding='utf-8') as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)

# ── helpers ───────────────────────────────────────────────────────────────────
def td_serial(td):
    return td.total_seconds() / 86400 if hasattr(td, 'total_seconds') else None

def get_car(unit, car_map):
    s = str(unit).strip()
    # direct match
    if s in car_map:
        return car_map[s]
    # case-insensitive fallback
    sl = s.lower()
    for k, v in car_map.items():
        if k.strip().lower() == sl:
            return v
    return ''

# ── conversion ────────────────────────────────────────────────────────────────
def read_df(path, sheet):
    try:
        return pd.read_excel(path, sheet_name=sheet, header=None)
    except Exception:
        return None

def convert(src1, src2, car_map, colorize=False):
    frames = []
    for p in [src1, src2]:
        if not p or not os.path.exists(p):
            continue
        for sh in [0, 1]:
            df = read_df(p, sh)
            if df is not None:
                frames.append(df)

    if not frames:
        raise ValueError('Не вдалось прочитати файли')

    df = pd.concat(frames, ignore_index=True)
    df = df.drop_duplicates(subset=[1, 2])          # унікальні за (час, назва)

    CUT = 6 * 3600
    def skey(r):
        s = r[1].total_seconds() if hasattr(r[1], 'total_seconds') else 0
        return s + 86400 if s < CUT else s

    df['_s'] = df.apply(skey, axis=1)
    df = df.sort_values('_s').drop(columns='_s').reset_index(drop=True)

    wb = Workbook()
    ws = wb.active
    try:
        ws.title = datetime.today().strftime('%d.%m.%y')
    except Exception:
        pass

    # ── output: 8 колонок ─────────────────────────────────────────────────────
    # 1: B (час дзвінка)         → H:MM:SS
    # 2: C (назва)               → text
    # 3: D (адреса)              → text
    # 4: E (час дзвінка HH:MM)   → HH:MM:SS  (той самий серіал, інший формат)
    # 5: F (наряд)               → text
    # 6: G (час відбуття)        → H:MM:SS
    # 7: ТРИВАЛІСТЬ = MINUTE(F6-D4)
    # 8: Номер авто НР

    HDR = ['ТП', 'Назва', 'Адреса', 'Час прийому виклику',
           'Наряд', 'Час відбуття', 'Тривалість (хв)', 'Номер авто НР']

    thin   = Side(style='thin', color='8899AA')
    brd    = Border(left=thin, right=thin, top=thin, bottom=thin)
    hdr_f  = Font(name='Arial', bold=True, size=10, color='FFFFFF')
    hdr_fi = PatternFill('solid', fgColor='1B3865')
    h_al   = Alignment(horizontal='center', vertical='center', wrap_text=True)

    for ci, h in enumerate(HDR, 1):
        c = ws.cell(1, ci, h)
        c.font = hdr_f; c.fill = hdr_fi; c.alignment = h_al; c.border = brd
    ws.row_dimensions[1].height = 30

    rf   = Font(name='Arial', size=10)
    cen  = Alignment(horizontal='center', vertical='center')
    left = Alignment(horizontal='left',   vertical='center')
    afil = PatternFill('solid', fgColor=ALT_ROW)

    for ri, (_, row) in enumerate(df.iterrows(), 2):
        call_s   = td_serial(row[1])
        depart_s = td_serial(row[6])
        unit     = row[5]
        car_num  = get_car(unit, car_map)

        def wr(col, val, fmt=None, al=cen):
            c = ws.cell(ri, col, val)
            c.font = rf; c.border = brd; c.alignment = al
            if fmt: c.number_format = fmt
            if colorize and ri % 2 == 0: c.fill = afil

        wr(1, call_s,                    'H:MM:SS')
        wr(2, row[2],                    al=left)
        wr(3, row[3],                    al=left)
        wr(4, call_s,                    'HH:MM:SS')
        wr(5, unit)
        wr(6, depart_s,                  'H:MM:SS')
        wr(7, f'=MINUTE(F{ri}-D{ri})')
        wr(8, car_num)

    for i, w in enumerate([10, 30, 34, 12, 12, 10, 12, 18], 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = 'A2'

    out_dir  = os.path.dirname(src1)
    date_str = datetime.today().strftime('%d.%m.%Y')
    out      = os.path.join(out_dir, f'УПО_{date_str}.xlsx')
    n = 1
    while os.path.exists(out):
        out = os.path.join(out_dir, f'УПО_{date_str}_{n}.xlsx'); n += 1
    wb.save(out)
    return out, len(df)

# ── settings dialog ───────────────────────────────────────────────────────────
class SettingsDialog(ctk.CTkToplevel):
    def __init__(self, parent, cfg, on_save):
        super().__init__(parent)
        self.title('Номери автомобілів')
        self.geometry('440x390')
        self.resizable(False, False)
        self.configure(fg_color=BG)
        self.grab_set()
        self.cfg = cfg
        self.on_save = on_save
        self.entries = {}
        try: self.iconbitmap(ICO_PATH)
        except Exception: pass

        ctk.CTkLabel(self, text='Номери автомобілів по підрозділах',
                     font=('Arial', 14, 'bold'), text_color=GOLD
                     ).pack(pady=(16, 8))

        fr = ctk.CTkFrame(self, fg_color=CARD, corner_radius=10,
                          border_width=1, border_color=BORDER)
        fr.pack(padx=20, fill='x')

        for u in UNITS:
            row = ctk.CTkFrame(fr, fg_color='transparent')
            row.pack(fill='x', padx=14, pady=6)
            ctk.CTkLabel(row, text=u, width=120, anchor='w',
                         font=('Arial', 12), text_color=TEXT
                         ).pack(side='left')
            e = ctk.CTkEntry(row, width=220, fg_color=ENTRY_BG,
                             border_color=BORDER, text_color=WHITE,
                             font=('Arial', 12), placeholder_text='держномер...')
            val = cfg.get('car_numbers', {}).get(u, '')
            if val: e.insert(0, val)
            e.pack(side='left', padx=(8, 0))
            self.entries[u] = e

        ctk.CTkButton(self, text='Зберегти',
                      fg_color=BTN, hover_color=BTN_HOV,
                      text_color=WHITE, font=('Arial', 13, 'bold'),
                      corner_radius=8, height=36, command=self._save
                      ).pack(pady=14)

    def _save(self):
        self.cfg['car_numbers'] = {u: self.entries[u].get().strip() for u in UNITS}
        save_cfg(self.cfg)
        self.on_save(self.cfg['car_numbers'])
        self.destroy()

# ── file row widget ───────────────────────────────────────────────────────────
class FileRow(ctk.CTkFrame):
    def __init__(self, parent, label, **kw):
        super().__init__(parent, fg_color='transparent', **kw)
        self.var = tk.StringVar()
        ctk.CTkLabel(self, text=label, width=70, anchor='w',
                     font=('Arial', 11), text_color=MUTED,
                     fg_color='transparent').pack(side='left')
        ctk.CTkEntry(self, textvariable=self.var, width=380, height=30,
                     fg_color=ENTRY_BG, border_color=BORDER,
                     text_color=WHITE, font=('Arial', 11),
                     placeholder_text='Оберіть .xlsx файл...'
                     ).pack(side='left', padx=(0, 6))
        ctk.CTkButton(self, text='📂', width=44, height=30,
                      fg_color=BTN, hover_color=BTN_HOV,
                      text_color=WHITE, font=('Arial', 12),
                      corner_radius=6, command=self._pick
                      ).pack(side='left')

    def _pick(self):
        p = filedialog.askopenfilename(
            title='Оберіть Excel файл',
            filetypes=[('Excel', '*.xlsx *.xls'), ('Всі файли', '*.*')])
        if p: self.var.set(p)

    def get(self): return self.var.get().strip()

# ── main window ───────────────────────────────────────────────────────────────
class UPOApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title('УПО Scripts')
        self.geometry('720x590')
        self.resizable(False, False)
        self.configure(fg_color=BG)
        try: self.iconbitmap(ICO_PATH)
        except Exception: pass

        self.cfg      = load_cfg()
        self.car_map  = self.cfg.get('car_numbers', {})
        self.colorize = tk.BooleanVar(value=False)

        self._build_bg()
        self._build_ui()

    def _build_bg(self):
        c = tk.Canvas(self, bg=BG, highlightthickness=0)
        c.place(x=0, y=0, width=720, height=590)
        # UPO emblem watermark
        if os.path.exists(SVG_PATH):
            try:
                self._emb = tksvg.SvgImage(file=SVG_PATH, scaletowidth=200)
                c.create_image(360, 295, image=self._emb)
                c.create_rectangle(0, 0, 720, 590, fill=BG, stipple='gray75')
            except Exception:
                pass
        c.create_rectangle(0, 0, 720, 3, fill=GOLD, outline='')
        c.create_rectangle(0, 587, 720, 590, fill=BTN, outline='')

    def _build_ui(self):
        # ── header: patrol car + title ────────────────────────────────────────
        hdr = ctk.CTkFrame(self, fg_color='transparent', width=720, height=120)
        hdr.place(x=0, y=8)
        hdr.pack_propagate(False)

        # patrol car image (left side)
        if os.path.exists(PNG_PATH):
            try:
                raw = Image.open(PNG_PATH).convert('RGBA')
                raw.thumbnail((240, 130), Image.LANCZOS)
                self._car_img = ctk.CTkImage(light_image=raw, dark_image=raw,
                                             size=(raw.width, raw.height))
                ctk.CTkLabel(hdr, image=self._car_img, text='',
                             fg_color='transparent').place(x=12, y=8)
            except Exception:
                pass

        # title (right side)
        ctk.CTkLabel(hdr, text='УПО  Scripts',
                     font=('Arial', 26, 'bold'), text_color=GOLD,
                     fg_color='transparent').place(x=270, y=18)
        ctk.CTkLabel(hdr, text='Обробка файлів виїздів',
                     font=('Arial', 11), text_color=MUTED,
                     fg_color='transparent').place(x=270, y=66)

        # ── files card ────────────────────────────────────────────────────────
        fc = ctk.CTkFrame(self, fg_color=CARD, corner_radius=10,
                          border_width=1, border_color=BORDER,
                          width=680, height=110)
        fc.place(relx=0.5, y=148, anchor='n')
        fc.pack_propagate(False)

        ctk.CTkLabel(fc, text='Вхідні файли', font=('Arial', 10),
                     text_color=MUTED, fg_color='transparent').place(x=16, y=8)

        self.file1 = FileRow(fc, 'Файл 1:')
        self.file1.place(x=14, y=32)
        self.file2 = FileRow(fc, 'Файл 2:')
        self.file2.place(x=14, y=68)

        # ── settings card ─────────────────────────────────────────────────────
        sc = ctk.CTkFrame(self, fg_color=CARD, corner_radius=10,
                          border_width=1, border_color=BORDER,
                          width=680, height=64)
        sc.place(relx=0.5, y=270, anchor='n')
        sc.pack_propagate(False)

        ctk.CTkLabel(sc, text='Номери автомобілів', font=('Arial', 11),
                     text_color=TEXT, fg_color='transparent').place(x=16, y=10)
        self.lbl_cars = ctk.CTkLabel(sc, text=self._summary(),
                                     font=('Arial', 10), text_color=MUTED,
                                     fg_color='transparent')
        self.lbl_cars.place(x=16, y=34)
        ctk.CTkButton(sc, text='Змінити', width=90, height=28,
                      fg_color=BTN, hover_color=BTN_HOV,
                      text_color=WHITE, font=('Arial', 11),
                      corner_radius=6, command=self._settings
                      ).place(x=576, y=18)

        # ── colorize ──────────────────────────────────────────────────────────
        ctk.CTkCheckBox(self, text='Розукрашення рядків',
                        variable=self.colorize,
                        font=('Arial', 11), text_color=MUTED,
                        fg_color=BTN, hover_color=BTN_HOV,
                        border_color=BORDER, checkmark_color=WHITE
                        ).place(relx=0.5, y=352, anchor='center')

        # ── process ───────────────────────────────────────────────────────────
        ctk.CTkButton(self, text='▶   Обробити',
                      font=('Arial', 14, 'bold'),
                      fg_color=BTN, hover_color=BTN_HOV,
                      text_color=WHITE, corner_radius=8,
                      height=44, width=210,
                      command=self._process
                      ).place(relx=0.5, y=410, anchor='center')

        self.lbl_status = ctk.CTkLabel(self, text='',
                                       font=('Arial', 11), text_color=MUTED,
                                       fg_color='transparent', wraplength=640)
        self.lbl_status.place(relx=0.5, y=470, anchor='center')

    def _summary(self):
        filled = sum(1 for v in self.car_map.values() if v)
        return f'Заповнено: {filled} / {len(UNITS)}'

    def _settings(self):
        SettingsDialog(self, self.cfg, self._on_saved)

    def _on_saved(self, car_numbers):
        self.car_map = car_numbers
        self.lbl_cars.configure(text=self._summary())
        self._st('Номери авто збережено', SUCCESS)

    def _process(self):
        s1 = self.file1.get()
        s2 = self.file2.get()
        if not s1 or not os.path.exists(s1):
            self._st('Оберіть хоча б перший файл', WARN)
            return
        self._st('Обробка...', MUTED); self.update()
        try:
            out, n = convert(s1, s2 or None, self.car_map, self.colorize.get())
            self._st(f'Готово — {n} рядків → {os.path.basename(out)}', SUCCESS)
        except Exception as e:
            self._st(f'Помилка: {e}', ERR)

    def _st(self, msg, col):
        self.lbl_status.configure(text=msg, text_color=col)


if __name__ == '__main__':
    app = UPOApp()
    app.mainloop()
