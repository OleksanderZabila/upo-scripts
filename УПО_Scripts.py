# -*- coding: utf-8 -*-
import os
import sys
import json
import tkinter as tk
from tkinter import filedialog, messagebox
import customtkinter as ctk
import tksvg
import pandas as pd
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from datetime import datetime

# ── paths ────────────────────────────────────────────────────────────────────
BASE_DIR   = os.path.dirname(os.path.abspath(sys.argv[0]))
BUNDLE_DIR = getattr(sys, '_MEIPASS', BASE_DIR)   # PyInstaller temp dir
SVG_PATH   = os.path.join(BUNDLE_DIR, 'Security_Police_of_Ukraine_emblem.svg')
CFG_PATH   = os.path.join(BASE_DIR, 'upo_config.json')

UNITS = ['ТО 10', 'ТО 12', 'ТО 13', 'ТО 15', 'ТО Умань 1', 'ТО Умань 2']

# ── colors / style ───────────────────────────────────────────────────────────
BG      = '#0A1A35'
ACCENT  = '#C8A800'
BTN_FG  = '#FFD700'
BTN_BG  = '#1B3260'
BTN_HOV = '#274A8A'
WHITE   = '#FFFFFF'
LIGHT   = '#D0D8E8'
CARD    = '#0F2347'
BORDER  = '#1E3A6A'

ctk.set_appearance_mode('dark')


# ── config ───────────────────────────────────────────────────────────────────
def load_cfg():
    if os.path.exists(CFG_PATH):
        with open(CFG_PATH, 'r', encoding='utf-8') as f:
            return json.load(f)
    return {'car_numbers': {u: '' for u in UNITS}}


def save_cfg(cfg):
    with open(CFG_PATH, 'w', encoding='utf-8') as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)


# ── conversion logic ──────────────────────────────────────────────────────────
def td_to_serial(td):
    return td.total_seconds() / 86400


def convert_file(src_path, car_map):
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

    # headers (matching the template)
    try:
        tmpl = load_workbook(src_path.replace(os.path.basename(src_path), 'Черкаси 19-20.05.26.xlsx'))
        tmpl_ws = tmpl.active
        headers_src = [c.value for c in tmpl_ws[1]]
    except Exception:
        headers_src = None

    headers = ['ТП', 'Назва', 'Адреса', 'Час прийому виклику',
               'Наряд', 'Номер авто', 'Час відбуття', 'Прибуття (хв)', '']

    wb = Workbook()
    ws = wb.active
    # sheet name = date from filename or today
    try:
        fname = os.path.splitext(os.path.basename(src_path))[0]
        sheet_name = fname[:31]
    except Exception:
        sheet_name = datetime.today().strftime('%d.%m.%y')
    ws.title = sheet_name

    # ── header row formatting ────────────────────────────────────────────────
    hdr_font   = Font(name='Arial', bold=True, size=10, color='FFFFFF')
    hdr_fill   = PatternFill('solid', fgColor='0F2347')
    hdr_align  = Alignment(horizontal='center', vertical='center', wrap_text=True)
    thin       = Side(style='thin', color='1E3A6A')
    border     = Border(left=thin, right=thin, top=thin, bottom=thin)

    for ci, h in enumerate(headers, 1):
        c = ws.cell(row=1, column=ci, value=h)
        c.font    = hdr_font
        c.fill    = hdr_fill
        c.alignment = hdr_align
        c.border  = border
    ws.row_dimensions[1].height = 30

    # ── data rows ────────────────────────────────────────────────────────────
    data_font  = Font(name='Arial', size=10)
    data_align = Alignment(vertical='center')
    center     = Alignment(horizontal='center', vertical='center')

    for ri, (_, row) in enumerate(df_all.iterrows(), 2):
        call_td    = row[1]
        name       = row[2]
        address    = row[3]
        unit       = row[5]
        depart_td  = row[6]

        call_s   = td_to_serial(call_td)   if hasattr(call_td,   'total_seconds') else None
        depart_s = td_to_serial(depart_td) if hasattr(depart_td, 'total_seconds') else None
        car_num  = car_map.get(str(unit).strip(), '')

        # col A: call time H:MM
        c = ws.cell(ri, 1, call_s);   c.number_format = 'H:MM:SS'; c.alignment = center
        # col B: name
        c = ws.cell(ri, 2, name);     c.alignment = data_align
        # col C: address
        c = ws.cell(ri, 3, address);  c.alignment = data_align
        # col D: call time HH:MM (text-style)
        c = ws.cell(ri, 4, call_s);   c.number_format = 'HH:MM:SS'; c.alignment = center
        # col E: unit
        c = ws.cell(ri, 5, unit);     c.alignment = center
        # col F: car number (auto-filled)
        c = ws.cell(ri, 6, car_num);  c.alignment = center
        # col G: departure time
        c = ws.cell(ri, 7, depart_s); c.number_format = 'H:MM:SS'; c.alignment = center
        # col H: minutes formula
        c = ws.cell(ri, 8, f'=MINUTE(G{ri}-D{ri})'); c.alignment = center
        # col I: empty
        ws.cell(ri, 9, None)

        for ci in range(1, 10):
            ws.cell(ri, ci).font   = data_font
            ws.cell(ri, ci).border = border

        # alternate row shading
        if ri % 2 == 0:
            fill = PatternFill('solid', fgColor='EDF1FA')
            for ci in range(1, 10):
                ws.cell(ri, ci).fill = fill

    # ── column widths ────────────────────────────────────────────────────────
    widths = [10, 34, 35, 13, 14, 16, 11, 11, 6]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    # freeze header
    ws.freeze_panes = 'A2'

    # ── save ─────────────────────────────────────────────────────────────────
    out_dir  = os.path.dirname(src_path)
    date_str = datetime.today().strftime('%d.%m.%Y')
    out_path = os.path.join(out_dir, f'УПО_результат_{date_str}.xlsx')
    # avoid overwrite
    n = 1
    base = out_path
    while os.path.exists(out_path):
        out_path = base.replace('.xlsx', f'_{n}.xlsx')
        n += 1
    wb.save(out_path)
    return out_path, len(df_all)


# ── settings dialog ──────────────────────────────────────────────────────────
class SettingsDialog(ctk.CTkToplevel):
    def __init__(self, parent, cfg, on_save):
        super().__init__(parent)
        self.title('Номери автомобілів')
        self.geometry('400x370')
        self.resizable(False, False)
        self.configure(fg_color=BG)
        self.grab_set()
        self.on_save = on_save
        self.cfg = cfg
        self.entries = {}

        ctk.CTkLabel(self, text='Номери автомобілів по підрозділах',
                     font=('Arial', 14, 'bold'), text_color=BTN_FG
                     ).pack(pady=(18, 10))

        frame = ctk.CTkFrame(self, fg_color=CARD, corner_radius=10)
        frame.pack(padx=24, pady=4, fill='x')

        for i, unit in enumerate(UNITS):
            row = ctk.CTkFrame(frame, fg_color='transparent')
            row.pack(fill='x', padx=14, pady=5)
            ctk.CTkLabel(row, text=unit, width=120, anchor='w',
                         font=('Arial', 12), text_color=LIGHT
                         ).pack(side='left')
            e = ctk.CTkEntry(row, width=180, fg_color='#0A1A35',
                             border_color=BORDER, text_color=WHITE,
                             font=('Arial', 12),
                             placeholder_text='напр. АА 1234 ВС')
            current = cfg.get('car_numbers', {}).get(unit, '')
            if current:
                e.insert(0, current)
            e.pack(side='left', padx=(8, 0))
            self.entries[unit] = e

        ctk.CTkButton(self, text='💾  Зберегти',
                      fg_color=ACCENT, hover_color='#A08800',
                      text_color='#000000', font=('Arial', 13, 'bold'),
                      corner_radius=8, height=38,
                      command=self._save
                      ).pack(pady=16)

    def _save(self):
        car_numbers = {u: self.entries[u].get().strip() for u in UNITS}
        self.cfg['car_numbers'] = car_numbers
        save_cfg(self.cfg)
        self.on_save(car_numbers)
        self.destroy()


# ── main app ──────────────────────────────────────────────────────────────────
class UPOApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title('УПО Scripts')
        self.geometry('700x530')
        self.resizable(False, False)
        self.configure(fg_color=BG)

        self.cfg = load_cfg()
        self.car_map = self.cfg.get('car_numbers', {})
        self.file_path = tk.StringVar()

        self._build_background()
        self._build_ui()

    # ── background canvas with emblem ─────────────────────────────────────
    def _build_background(self):
        canvas = tk.Canvas(self, bg=BG, highlightthickness=0)
        canvas.place(x=0, y=0, width=700, height=530)

        if os.path.exists(SVG_PATH):
            try:
                self._emblem = tksvg.SvgImage(file=SVG_PATH, scaletowidth=260)
                canvas.create_image(350, 265, image=self._emblem)
                # semi-transparent overlay to make emblem a watermark
                canvas.create_rectangle(0, 0, 700, 530,
                                        fill=BG, stipple='gray75')
            except Exception:
                pass

        # decorative top bar
        canvas.create_rectangle(0, 0, 700, 3, fill=ACCENT, outline='')
        canvas.create_rectangle(0, 527, 700, 530, fill=ACCENT, outline='')

    # ── main UI widgets ───────────────────────────────────────────────────
    def _build_ui(self):
        # ── title ────────────────────────────────────────────────────────
        ctk.CTkLabel(self,
                     text='УПО  Scripts',
                     font=('Arial', 28, 'bold'),
                     text_color=BTN_FG,
                     fg_color='transparent'
                     ).place(relx=0.5, y=48, anchor='center')

        ctk.CTkLabel(self,
                     text='Автоматична обробка файлів пожежних виїздів',
                     font=('Arial', 11),
                     text_color=LIGHT,
                     fg_color='transparent'
                     ).place(relx=0.5, y=82, anchor='center')

        # ── file selection card ───────────────────────────────────────────
        card = ctk.CTkFrame(self, fg_color=CARD, corner_radius=12,
                            border_width=1, border_color=BORDER,
                            width=600, height=100)
        card.place(relx=0.5, y=170, anchor='center')
        card.pack_propagate(False)

        ctk.CTkLabel(card, text='Вхідний файл (.xlsx)',
                     font=('Arial', 11), text_color=LIGHT,
                     fg_color='transparent'
                     ).place(x=20, y=12)

        self.entry_file = ctk.CTkEntry(card, textvariable=self.file_path,
                                       width=440, height=34,
                                       fg_color=BG, border_color=BORDER,
                                       text_color=WHITE, font=('Arial', 11),
                                       placeholder_text='Оберіть файл...')
        self.entry_file.place(x=20, y=44)

        ctk.CTkButton(card, text='📁', width=52, height=34,
                      fg_color=BTN_BG, hover_color=BTN_HOV,
                      text_color=WHITE, font=('Arial', 14),
                      corner_radius=6,
                      command=self._browse
                      ).place(x=470, y=44)

        # ── settings card ─────────────────────────────────────────────────
        card2 = ctk.CTkFrame(self, fg_color=CARD, corner_radius=12,
                             border_width=1, border_color=BORDER,
                             width=600, height=76)
        card2.place(relx=0.5, y=305, anchor='center')
        card2.pack_propagate(False)

        ctk.CTkLabel(card2,
                     text='⚙   Налаштування номерів авто',
                     font=('Arial', 11), text_color=LIGHT,
                     fg_color='transparent'
                     ).place(x=20, y=10)

        self.lbl_cars = ctk.CTkLabel(card2,
                                     text=self._car_summary(),
                                     font=('Arial', 10),
                                     text_color=ACCENT,
                                     fg_color='transparent')
        self.lbl_cars.place(x=20, y=34)

        ctk.CTkButton(card2, text='Змінити', width=90, height=28,
                      fg_color=BTN_BG, hover_color=BTN_HOV,
                      text_color=WHITE, font=('Arial', 11),
                      corner_radius=6,
                      command=self._open_settings
                      ).place(x=490, y=22)

        # ── process button ────────────────────────────────────────────────
        ctk.CTkButton(self,
                      text='▶   ОБРОБИТИ',
                      font=('Arial', 15, 'bold'),
                      fg_color=ACCENT,
                      hover_color='#A08800',
                      text_color='#000000',
                      corner_radius=10,
                      height=46,
                      width=220,
                      command=self._process
                      ).place(relx=0.5, y=408, anchor='center')

        # ── status label ──────────────────────────────────────────────────
        self.lbl_status = ctk.CTkLabel(self, text='',
                                       font=('Arial', 11),
                                       text_color=LIGHT,
                                       fg_color='transparent',
                                       wraplength=580)
        self.lbl_status.place(relx=0.5, y=470, anchor='center')

    def _car_summary(self):
        filled = sum(1 for v in self.car_map.values() if v)
        return f'Заповнено: {filled} з {len(UNITS)} підрозділів'

    def _browse(self):
        p = filedialog.askopenfilename(
            title='Оберіть файл виїздів',
            filetypes=[('Excel файли', '*.xlsx *.xls'), ('Всі файли', '*.*')]
        )
        if p:
            self.file_path.set(p)
            self._set_status('', 'info')

    def _open_settings(self):
        SettingsDialog(self, self.cfg, self._on_settings_saved)

    def _on_settings_saved(self, car_numbers):
        self.car_map = car_numbers
        self.lbl_cars.configure(text=self._car_summary())
        self._set_status('Налаштування збережено.', 'ok')

    def _process(self):
        src = self.file_path.get().strip()
        if not src or not os.path.exists(src):
            self._set_status('⚠  Спочатку оберіть файл!', 'warn')
            return
        self._set_status('⏳  Обробка...', 'info')
        self.update()
        try:
            out_path, n_rows = convert_file(src, self.car_map)
            fname = os.path.basename(out_path)
            self._set_status(f'✔  Готово! {n_rows} рядків → {fname}', 'ok')
        except Exception as e:
            self._set_status(f'✘  Помилка: {e}', 'err')

    def _set_status(self, msg, kind='info'):
        colors = {'ok': '#4CAF50', 'warn': '#FFC107', 'err': '#F44336', 'info': LIGHT}
        self.lbl_status.configure(text=msg, text_color=colors.get(kind, LIGHT))


if __name__ == '__main__':
    app = UPOApp()
    app.mainloop()
