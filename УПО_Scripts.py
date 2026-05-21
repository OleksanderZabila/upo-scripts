# -*- coding: utf-8 -*-
import os
import sys
import json
import tkinter as tk
from tkinter import filedialog
import customtkinter as ctk
import tksvg
import pandas as pd
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from datetime import datetime

# ── paths ─────────────────────────────────────────────────────────────────────
BASE_DIR   = os.path.dirname(os.path.abspath(sys.argv[0]))
BUNDLE_DIR = getattr(sys, '_MEIPASS', BASE_DIR)
SVG_PATH   = os.path.join(BUNDLE_DIR, 'Security_Police_of_Ukraine_emblem.svg')
ICO_PATH   = os.path.join(BUNDLE_DIR, 'patrol-polycar.ico')
CFG_PATH   = os.path.join(BASE_DIR, 'upo_config.json')

UNITS = ['ТО 10', 'ТО 12', 'ТО 13', 'ТО 15', 'ТО Умань 1', 'ТО Умань 2']

# ── color palette (clean navy + police blue + gold) ───────────────────────────
BG       = '#1C2B40'   # window background
CARD     = '#243352'   # card / panel
BORDER   = '#2E4472'   # border
BTN      = '#1E5FAA'   # primary button
BTN_HOV  = '#174E90'   # button hover
GOLD     = '#D4AF37'   # accent / title
WHITE    = '#FFFFFF'
TEXT     = '#E8EEF8'   # body text
MUTED    = '#8AA4C4'   # secondary text
ENTRY_BG = '#141E2E'   # input background
SUCCESS  = '#4CAF50'
WARN     = '#FFC107'
ERR      = '#F44336'

ctk.set_appearance_mode('dark')


# ── config ────────────────────────────────────────────────────────────────────
def load_cfg():
    if os.path.exists(CFG_PATH):
        with open(CFG_PATH, 'r', encoding='utf-8') as f:
            return json.load(f)
    return {'car_numbers': {u: '' for u in UNITS}}


def save_cfg(cfg):
    with open(CFG_PATH, 'w', encoding='utf-8') as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)


# ── conversion ────────────────────────────────────────────────────────────────
def td_serial(td):
    return td.total_seconds() / 86400 if hasattr(td, 'total_seconds') else None


def convert_file(src_path, car_map, colorize=False):
    df19 = pd.read_excel(src_path, sheet_name=0, header=None)
    try:
        df20 = pd.read_excel(src_path, sheet_name=1, header=None)
        df_all = pd.concat([df19, df20], ignore_index=True)
    except Exception:
        df_all = df19.copy()

    df_all = df_all.drop_duplicates(subset=[1, 2])

    CUTOFF = 6 * 3600

    def sort_key(row):
        secs = row[1].total_seconds() if hasattr(row[1], 'total_seconds') else 0
        return secs + 86400 if secs < CUTOFF else secs

    df_all['_s'] = df_all.apply(sort_key, axis=1)
    df_all = df_all.sort_values('_s').drop(columns='_s').reset_index(drop=True)

    wb = Workbook()
    ws = wb.active
    try:
        ws.title = datetime.today().strftime('%d.%m.%y')
    except Exception:
        pass

    # ── output columns ────────────────────────────────────────────────────────
    # A: Адреса
    # B: Час передачі сигналу
    # C: Наряд
    # D: Час прибуття
    # E: Тривалість прибуття  (=MINUTE(D-B))
    # F: Номер автомобіля НР

    HDR = ['Адреса', 'Час передачі сигналу', 'Наряд',
           'Час прибуття', 'Тривалість прибуття', 'Номер автомобіля НР']

    thin   = Side(style='thin', color='AAAAAA')
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    hdr_font  = Font(name='Arial', bold=True, size=10)
    hdr_fill  = PatternFill('solid', fgColor='D0D8E8')
    hdr_align = Alignment(horizontal='center', vertical='center', wrap_text=True)

    for ci, h in enumerate(HDR, 1):
        c = ws.cell(1, ci, h)
        c.font      = hdr_font
        c.fill      = hdr_fill
        c.alignment = hdr_align
        c.border    = border
    ws.row_dimensions[1].height = 32

    row_font  = Font(name='Arial', size=10)
    c_center  = Alignment(horizontal='center', vertical='center')
    c_left    = Alignment(horizontal='left',   vertical='center')
    alt_fill  = PatternFill('solid', fgColor='EEF2FA')

    for ri, (_, row) in enumerate(df_all.iterrows(), 2):
        address    = row[3]
        call_td    = row[1]
        unit       = row[5]
        arrive_td  = row[6]
        car_num    = car_map.get(str(unit).strip(), '')

        call_s   = td_serial(call_td)
        arrive_s = td_serial(arrive_td)

        # A: Адреса
        c = ws.cell(ri, 1, address);   c.alignment = c_left
        # B: Час передачі сигналу
        c = ws.cell(ri, 2, call_s);    c.number_format = 'H:MM:SS'; c.alignment = c_center
        # C: Наряд
        c = ws.cell(ri, 3, unit);      c.alignment = c_center
        # D: Час прибуття
        c = ws.cell(ri, 4, arrive_s);  c.number_format = 'H:MM:SS'; c.alignment = c_center
        # E: Тривалість прибуття
        c = ws.cell(ri, 5, f'=MINUTE(D{ri}-B{ri})'); c.alignment = c_center
        # F: Номер автомобіля НР
        c = ws.cell(ri, 6, car_num);   c.alignment = c_center

        for ci in range(1, 7):
            cell = ws.cell(ri, ci)
            cell.font   = row_font
            cell.border = border
            if colorize and ri % 2 == 0:
                cell.fill = alt_fill

    # ── column widths ─────────────────────────────────────────────────────────
    widths = [38, 16, 12, 12, 14, 18]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    ws.freeze_panes = 'A2'

    # ── save ──────────────────────────────────────────────────────────────────
    out_dir  = os.path.dirname(src_path)
    date_str = datetime.today().strftime('%d.%m.%Y')
    out_path = os.path.join(out_dir, f'УПО_{date_str}.xlsx')
    n = 1
    while os.path.exists(out_path):
        out_path = os.path.join(out_dir, f'УПО_{date_str}_{n}.xlsx')
        n += 1
    wb.save(out_path)
    return out_path, len(df_all)


# ── settings dialog ───────────────────────────────────────────────────────────
class SettingsDialog(ctk.CTkToplevel):
    def __init__(self, parent, cfg, on_save):
        super().__init__(parent)
        self.title('Номери автомобілів')
        self.geometry('420x360')
        self.resizable(False, False)
        self.configure(fg_color=BG)
        self.grab_set()
        self.on_save = on_save
        self.cfg = cfg
        self.entries = {}

        try:
            self.iconbitmap(ICO_PATH)
        except Exception:
            pass

        ctk.CTkLabel(self, text='Номери автомобілів',
                     font=('Arial', 15, 'bold'), text_color=GOLD
                     ).pack(pady=(16, 8))

        frame = ctk.CTkFrame(self, fg_color=CARD, corner_radius=10,
                             border_width=1, border_color=BORDER)
        frame.pack(padx=20, fill='x')

        for unit in UNITS:
            row = ctk.CTkFrame(frame, fg_color='transparent')
            row.pack(fill='x', padx=14, pady=6)
            ctk.CTkLabel(row, text=unit, width=120, anchor='w',
                         font=('Arial', 12), text_color=TEXT
                         ).pack(side='left')
            e = ctk.CTkEntry(row, width=200, fg_color=ENTRY_BG,
                             border_color=BORDER, text_color=WHITE,
                             font=('Arial', 12),
                             placeholder_text='держномер...')
            val = cfg.get('car_numbers', {}).get(unit, '')
            if val:
                e.insert(0, val)
            e.pack(side='left', padx=(8, 0))
            self.entries[unit] = e

        ctk.CTkButton(self, text='Зберегти',
                      fg_color=BTN, hover_color=BTN_HOV,
                      text_color=WHITE, font=('Arial', 13, 'bold'),
                      corner_radius=8, height=36,
                      command=self._save
                      ).pack(pady=14)

    def _save(self):
        self.cfg['car_numbers'] = {u: self.entries[u].get().strip() for u in UNITS}
        save_cfg(self.cfg)
        self.on_save(self.cfg['car_numbers'])
        self.destroy()


# ── main app ──────────────────────────────────────────────────────────────────
class UPOApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title('УПО Scripts')
        self.geometry('680x500')
        self.resizable(False, False)
        self.configure(fg_color=BG)

        # window icon (top-left corner + taskbar)
        try:
            self.iconbitmap(ICO_PATH)
        except Exception:
            pass

        self.cfg     = load_cfg()
        self.car_map = self.cfg.get('car_numbers', {})
        self.file_var   = tk.StringVar()
        self.colorize   = tk.BooleanVar(value=False)

        self._build_bg()
        self._build_ui()

    # ── background: emblem watermark ─────────────────────────────────────────
    def _build_bg(self):
        canvas = tk.Canvas(self, bg=BG, highlightthickness=0)
        canvas.place(x=0, y=0, width=680, height=500)

        if os.path.exists(SVG_PATH):
            try:
                self._emblem = tksvg.SvgImage(file=SVG_PATH, scaletowidth=220)
                canvas.create_image(340, 250, image=self._emblem)
                canvas.create_rectangle(0, 0, 680, 500,
                                        fill=BG, stipple='gray75')
            except Exception:
                pass

        # top accent line
        canvas.create_rectangle(0, 0, 680, 3, fill=GOLD, outline='')
        canvas.create_rectangle(0, 497, 680, 500, fill=BTN, outline='')

    # ── UI ────────────────────────────────────────────────────────────────────
    def _build_ui(self):
        # title
        ctk.CTkLabel(self, text='УПО  Scripts',
                     font=('Arial', 26, 'bold'), text_color=GOLD,
                     fg_color='transparent'
                     ).place(relx=0.5, y=44, anchor='center')

        ctk.CTkLabel(self,
                     text='Обробка файлів пожежно-поліцейських виїздів',
                     font=('Arial', 11), text_color=MUTED,
                     fg_color='transparent'
                     ).place(relx=0.5, y=74, anchor='center')

        # ── file card ─────────────────────────────────────────────────────────
        fc = ctk.CTkFrame(self, fg_color=CARD, corner_radius=10,
                          border_width=1, border_color=BORDER,
                          width=620, height=90)
        fc.place(relx=0.5, y=158, anchor='center')
        fc.pack_propagate(False)

        ctk.CTkLabel(fc, text='Вхідний файл', font=('Arial', 10),
                     text_color=MUTED, fg_color='transparent'
                     ).place(x=18, y=10)

        ctk.CTkEntry(fc, textvariable=self.file_var,
                     width=460, height=32, fg_color=ENTRY_BG,
                     border_color=BORDER, text_color=WHITE,
                     font=('Arial', 11),
                     placeholder_text='Оберіть .xlsx файл...'
                     ).place(x=18, y=40)

        ctk.CTkButton(fc, text='📂', width=60, height=32,
                      fg_color=BTN, hover_color=BTN_HOV,
                      text_color=WHITE, font=('Arial', 13),
                      corner_radius=6, command=self._browse
                      ).place(x=488, y=40)

        # ── settings card ─────────────────────────────────────────────────────
        sc = ctk.CTkFrame(self, fg_color=CARD, corner_radius=10,
                          border_width=1, border_color=BORDER,
                          width=620, height=72)
        sc.place(relx=0.5, y=278, anchor='center')
        sc.pack_propagate(False)

        ctk.CTkLabel(sc, text='Номери автомобілів', font=('Arial', 11),
                     text_color=TEXT, fg_color='transparent'
                     ).place(x=18, y=10)

        self.lbl_cars = ctk.CTkLabel(sc, text=self._summary(),
                                     font=('Arial', 10), text_color=MUTED,
                                     fg_color='transparent')
        self.lbl_cars.place(x=18, y=34)

        ctk.CTkButton(sc, text='Змінити', width=88, height=28,
                      fg_color=BTN, hover_color=BTN_HOV,
                      text_color=WHITE, font=('Arial', 11),
                      corner_radius=6, command=self._open_settings
                      ).place(x=514, y=22)

        # ── colorize checkbox ─────────────────────────────────────────────────
        ctk.CTkCheckBox(self, text='Розукрашення рядків у файлі',
                        variable=self.colorize,
                        font=('Arial', 11), text_color=MUTED,
                        fg_color=BTN, hover_color=BTN_HOV,
                        border_color=BORDER, checkmark_color=WHITE
                        ).place(relx=0.5, y=340, anchor='center')

        # ── process button ─────────────────────────────────────────────────────
        ctk.CTkButton(self, text='▶   Обробити',
                      font=('Arial', 14, 'bold'),
                      fg_color=BTN, hover_color=BTN_HOV,
                      text_color=WHITE, corner_radius=8,
                      height=42, width=200,
                      command=self._process
                      ).place(relx=0.5, y=400, anchor='center')

        # ── status ────────────────────────────────────────────────────────────
        self.lbl_status = ctk.CTkLabel(self, text='',
                                       font=('Arial', 11), text_color=MUTED,
                                       fg_color='transparent', wraplength=580)
        self.lbl_status.place(relx=0.5, y=452, anchor='center')

    def _summary(self):
        filled = sum(1 for v in self.car_map.values() if v)
        return f'Заповнено: {filled} з {len(UNITS)}'

    def _browse(self):
        p = filedialog.askopenfilename(
            title='Оберіть файл виїздів',
            filetypes=[('Excel', '*.xlsx *.xls'), ('Всі файли', '*.*')])
        if p:
            self.file_var.set(p)
            self._status('', MUTED)

    def _open_settings(self):
        SettingsDialog(self, self.cfg, self._on_saved)

    def _on_saved(self, car_numbers):
        self.car_map = car_numbers
        self.lbl_cars.configure(text=self._summary())
        self._status('Налаштування збережено', SUCCESS)

    def _process(self):
        src = self.file_var.get().strip()
        if not src or not os.path.exists(src):
            self._status('Спочатку оберіть файл', WARN)
            return
        self._status('Обробка...', MUTED)
        self.update()
        try:
            out, n = convert_file(src, self.car_map, self.colorize.get())
            self._status(f'Готово — {n} рядків → {os.path.basename(out)}', SUCCESS)
        except Exception as e:
            self._status(f'Помилка: {e}', ERR)

    def _status(self, msg, color):
        self.lbl_status.configure(text=msg, text_color=color)


if __name__ == '__main__':
    app = UPOApp()
    app.mainloop()
