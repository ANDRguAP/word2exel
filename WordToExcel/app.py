# -*- coding: utf-8 -*-
import os
import re
import threading
import traceback
from datetime import datetime, date, time
from difflib import SequenceMatcher

import tkinter as tk
from tkinter import filedialog, messagebox

from docx import Document
from docx.oxml.ns import qn
from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter


# ============================================================
# НАСТРОЙКИ
# ============================================================
FUZZY_THRESHOLD = 0.75      # порог схожести заголовков (режим сопоставления)
CLEAR_OLD_DATA = True       # очищать старые данные в Excel при сопоставлении по шапке
DATA_START_ROW = 2          # с какой строки писать данные при сопоставлении по шапке
CONVERT_NUMBERS = True      # превращать текстовые числа в настоящие числа при записи в Excel


# ============================================================
# ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ
# ============================================================
def normalize(text):
    if text is None:
        return ""
    text = str(text).strip().lower().replace("\n", " ")
    text = re.sub(r"[«»\"'`]", "", text)
    text = re.sub(r"[.:;,!?()\[\]{}]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def similarity(a, b):
    a, b = normalize(a), normalize(b)
    if not a or not b:
        return 0.0
    if a == b:
        return 1.0
    if a in b or b in a:
        return 0.95
    return SequenceMatcher(None, a, b).ratio()


def maybe_number(s):
    """Числа, записанные текстом, превращаем в настоящие числа Excel."""
    if not CONVERT_NUMBERS or not isinstance(s, str):
        return s
    t = s.strip().replace("\u00a0", " ").replace(" ", "")
    if re.fullmatch(r"-?\d+", t):
        digits = t.lstrip("-")
        if len(digits) > 1 and digits[0] == "0":   # коды вида "007" не трогаем
            return s
        return int(t)
    if re.fullmatch(r"-?\d+[.,]\d+", t):
        return float(t.replace(",", "."))
    return s


def fmt_val(v):
    """Значение ячейки Excel -> строка для Word."""
    if v is None:
        return ""
    if isinstance(v, datetime):
        return v.strftime("%d.%m.%Y %H:%M") if (v.hour or v.minute) else v.strftime("%d.%m.%Y")
    if isinstance(v, date):
        return v.strftime("%d.%m.%Y")
    if isinstance(v, time):
        return v.strftime("%H:%M:%S")
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v)


def unique_sheet_name(wb, base):
    base = re.sub(r"[\[\]:*?/\\]", " ", base)
    name = base[:31] or "Лист"
    i = 2
    while name in wb.sheetnames:
        suffix = f" ({i})"
        name = base[:31 - len(suffix)] + suffix
        i += 1
    return name


# ============================================================
# ЧТЕНИЕ ТАБЛИЦ WORD (с поддержкой объединённых ячеек)
# ============================================================
def table_to_matrix(table):
    """
    Преобразует таблицу Word в матрицу + список объединений.
    Возвращает (matrix, merges):
      matrix[i][j] — текст ячейки (или None, если ячейка накрыта объединением),
      merges — список (r1, c1, r2, c2) в 0-базных индексах.
    """
    from docx.table import _Cell

    try:
        grid_cols = len(table._tbl.tblGrid.findall(qn("w:gridCol")))
    except Exception:
        grid_cols = max((len(r.cells) for r in table.rows), default=1) or 1

    n_rows = len(table.rows)
    matrix = [[None] * grid_cols for _ in range(n_rows)]
    merges = []
    open_v = {}  # колонка -> информация об открытой вертикальной объединённой области

    for r_i, row in enumerate(table.rows):
        c_i = 0
        for tc in row._tr.tc_lst:
            while c_i < grid_cols and matrix[r_i][c_i] is not None:
                c_i += 1
            if c_i >= grid_cols:
                break

            tc_pr = tc.find(qn("w:tcPr"))
            span, vm = 1, None
            if tc_pr is not None:
                gs = tc_pr.find(qn("w:gridSpan"))
                if gs is not None:
                    try:
                        span = int(gs.get(qn("w:val")))
                    except (TypeError, ValueError):
                        span = 1
                vme = tc_pr.find(qn("w:vMerge"))
                if vme is not None:
                    vm = vme.get(qn("w:val")) or "continue"

            c_end = min(c_i + span, grid_cols) - 1

            if vm == "continue":
                # ячейка — продолжение вертикального объединения
                for c in range(c_i, c_end + 1):
                    if c in open_v:
                        open_v[c]["r_end"] = r_i
            else:
                try:
                    text = _Cell(tc, table).text.strip()
                except Exception:
                    try:
                        text = "\n".join(p.text for p in tc.xpath("./w:p")).strip()
                    except Exception:
                        text = ""
                matrix[r_i][c_i] = text
                for c in range(c_i + 1, c_end + 1):
                    matrix[r_i][c] = None  # накрыто горизонтальным объединением
                if span > 1 or vm == "restart":
                    entry = {"r_start": r_i, "r_end": r_i, "c_start": c_i, "c_end": c_end}
                    for c in range(c_i, c_end + 1):
                        open_v[c] = entry

            c_i = c_end + 1

    seen = set()
    for entry in list(open_v.values()):
        if id(entry) in seen:
            continue
        seen.add(id(entry))
        if entry["r_end"] > entry["r_start"] or entry["c_end"] > entry["c_start"]:
            merges.append((entry["r_start"], entry["c_start"], entry["r_end"], entry["c_end"]))

    # пустые строки/столбцы в конце таблицы убираем
    while matrix and all(v in (None, "") for v in matrix[-1]):
        matrix.pop()
    if matrix:
        width = max(1, max(len(r) for r in matrix))
        while width > 1 and all(row[width - 1] in (None, "") for row in matrix):
            width -= 1
        matrix = [row[:width] + [None] * (width - len(row)) for row in matrix]
        merges = [m for m in merges if m[0] < len(matrix) and m[1] < width]
    return matrix, merges


def read_word_tables(path):
    doc = Document(path)
    result = []
    for i, t in enumerate(doc.tables, 1):
        matrix, merges = table_to_matrix(t)
        if matrix:
            result.append({"name": f"Таблица {i}", "matrix": matrix, "merges": merges})
    return result


# ============================================================
# ЧТЕНИЕ ТАБЛИЦ EXCEL
# ============================================================
def sheet_to_matrix(ws):
    """Лист Excel -> (matrix, merges) с поддержкой объединённых ячеек."""
    n_rows, n_cols = ws.max_row, ws.max_column
    if not n_rows or not n_cols:
        return [], []
    vals = [[fmt_val(ws.cell(row=r, column=c).value) for c in range(1, n_cols + 1)]
            for r in range(1, n_rows + 1)]

    while vals and all(v == "" for v in vals[-1]):
        vals.pop()
    if not vals:
        return [], []
    width = max(len(r) for r in vals)
    vals = [r + [""] * (width - len(r)) for r in vals]
    while width > 1 and all(row[width - 1] == "" for row in vals):
        width -= 1
        vals = [row[:width] for row in vals]

    merges = []
    for mr in ws.merged_cells.ranges:
        r1, c1 = mr.min_row - 1, mr.min_col - 1
        r2, c2 = min(mr.max_row - 1, len(vals) - 1), min(mr.max_col - 1, width - 1)
        if r1 < len(vals) and c1 < width and (r2 > r1 or c2 > c1):
            merges.append((r1, c1, r2, c2))
    return vals, merges


# ============================================================
# НАПРАВЛЕНИЕ 1: WORD -> EXCEL
# ============================================================
def write_table_to_sheet(ws, matrix, merges, bold_header=True):
    for r, row in enumerate(matrix, 1):
        for c, v in enumerate(row, 1):
            if v not in (None, ""):
                ws.cell(row=r, column=c, value=maybe_number(v) if isinstance(v, str) else v)
    for (r1, c1, r2, c2) in merges:
        if r2 > r1 or c2 > c1:
            try:
                ws.merge_cells(start_row=r1 + 1, start_column=c1 + 1,
                               end_row=r2 + 1, end_column=c2 + 1)
            except Exception:
                pass

    # ширина столбцов по содержимому
    width = max(len(row) for row in matrix) if matrix else 1
    for c in range(1, width + 1):
        longest = 1
        for row in matrix:
            if c - 1 < len(row) and row[c - 1]:
                longest = max(longest, max(len(line) for line in str(row[c - 1]).split("\n")))
        ws.column_dimensions[get_column_letter(c)].width = min(max(longest + 2, 6), 60)

    # жирная шапка
    if bold_header and len(matrix) > 1:
        for c in range(1, max(len(row) for row in matrix) + 1):
            cell = ws.cell(row=1, column=c)
            if cell.value is not None:
                cell.font = cell.font.copy(bold=True)


def word_to_excel_as_is(word_path, excel_path, out_path, log):
    """Копирует ВСЕ таблицы из Word как есть (каждая — на свой лист)."""
    tables = read_word_tables(word_path)
    if not tables:
        raise ValueError("В файле Word не найдено ни одной таблицы.")
    log(f"Найдено таблиц в Word: {len(tables)}")

    if excel_path:
        wb = load_workbook(excel_path)
        log(f"Открываю Excel: {excel_path} (таблицы будут добавлены новыми листами)")
    else:
        wb = Workbook()
        wb.remove(wb.active)

    for t in tables:
        ws = wb.create_sheet(unique_sheet_name(wb, t["name"]))
        write_table_to_sheet(ws, t["matrix"], t["merges"])
        log(f"   ✓ '{ws.title}': {len(t['matrix'])} строк × {len(t['matrix'][0])} столбцов"
            + (f", объединений: {len(t['merges'])}" if t["merges"] else ""))

    wb.save(out_path)
    log(f"\n✓ Готово: {out_path}")


def word_to_excel_by_header(word_path, excel_path, out_path, log):
    """Старый режим: разложить данные Word по шапке существующего Excel."""
    raw_rows = []
    for t in read_word_tables(word_path):
        for row in t["matrix"]:
            cells = [v if v else "" for v in row]
            if any(cells):
                raw_rows.append(cells)
    if not raw_rows:
        raise ValueError("В Word не найдено ни одной таблицы с данными.")

    idx, w_head, w_data = detect_word_header(raw_rows)
    log(f"Шапка Word (строка {idx + 1}):")
    for i, h in enumerate(w_head):
        log(f"   [{i}] {h}")
    log(f"Строк данных: {len(w_data)}")

    log(f"\nЧитаю Excel: {excel_path}")
    wb = load_workbook(excel_path)
    ws = wb.active
    e_head = [ws.cell(row=1, column=c).value or "" for c in range(1, ws.max_column + 1)]
    log("Шапка Excel:")
    for i, h in enumerate(e_head):
        log(f"   [{i}] {h}")

    log("\nСопоставление столбцов:")
    mapping = build_mapping(w_head, e_head)
    for ex_idx, w_idx in mapping.items():
        if w_idx is not None:
            log(f"   ✓ '{e_head[ex_idx]}'  <-  '{w_head[w_idx]}'")
        else:
            log(f"   ✗ '{e_head[ex_idx]}' — нет соответствия в Word")

    if CLEAR_OLD_DATA:
        for row in range(DATA_START_ROW, ws.max_row + 1):
            for col in range(1, ws.max_column + 1):
                ws.cell(row=row, column=col).value = None

    for r_off, wrow in enumerate(w_data):
        excel_row = DATA_START_ROW + r_off
        for ex_idx, w_idx in mapping.items():
            if w_idx is None:
                continue
            v = wrow[w_idx] if w_idx < len(wrow) else None
            ws.cell(row=excel_row, column=ex_idx + 1, value=v if v not in ("", None) else None)

    wb.save(out_path)
    log(f"\n✓ Готово: {out_path}")


def detect_word_header(rows):
    markers = ["наимен", "код", "номер", "значен", "тип", "единица",
               "характер", "функц", "измер", "поз"]
    best_idx, best_score = 0, -1
    for i, row in enumerate(rows[:10]):
        joined = normalize(" ".join(str(c) for c in row))
        score = sum(1 for m in markers if m in joined)
        numeric = sum(1 for c in row if re.fullmatch(r"\d+", str(c).strip() or "x"))
        score -= numeric * 0.5
        if score > best_score:
            best_score, best_idx = score, i
    return best_idx, rows[best_idx], rows[best_idx + 1:]


def build_mapping(word_headers, excel_headers):
    mapping, used = {}, set()
    for ex_idx, ex_name in enumerate(excel_headers):
        if not normalize(ex_name):
            mapping[ex_idx] = None
            continue
        best_idx, best_score = None, 0.0
        for w_idx, w_name in enumerate(word_headers):
            if w_idx in used:
                continue
            s = similarity(ex_name, w_name)
            if s > best_score:
                best_score, best_idx = s, w_idx
        if best_idx is not None and best_score >= FUZZY_THRESHOLD:
            mapping[ex_idx] = best_idx
            used.add(best_idx)
        else:
            mapping[ex_idx] = None
    return mapping


# ============================================================
# НАПРАВЛЕНИЕ 2: EXCEL -> WORD
# ============================================================
def excel_to_word(excel_path, out_path, log, all_sheets=True):
    wb = load_workbook(excel_path, data_only=True)
    sheets = wb.sheetnames if all_sheets else [wb.active.title]

    if out_path and os.path.exists(out_path):
        doc = Document(out_path)
        log(f"Дописываю в существующий документ: {out_path}")
    else:
        doc = Document()
        log("Создаю новый документ Word")

    transferred = 0
    for name in sheets:
        ws = wb[name]
        matrix, merges = sheet_to_matrix(ws)
        if not matrix:
            log(f"   – Лист '{name}' пуст, пропускаю")
            continue

        n_cols = max(len(r) for r in matrix)
        table = doc.add_table(rows=len(matrix), cols=n_cols)
        try:
            table.style = "Table Grid"
        except Exception:
            pass

        for r, row in enumerate(matrix):
            for c, v in enumerate(row):
                table.cell(r, c).text = v

        for (r1, c1, r2, c2) in merges:
            if r2 > r1 or c2 > c1:
                try:
                    table.cell(r1, c1).merge(table.cell(r2, c2))
                except Exception:
                    pass

        # жирная первая строка (обычно это шапка)
        if len(matrix) > 1:
            for cell in table.rows[0].cells:
                for p in cell.paragraphs:
                    for run in p.runs:
                        run.bold = True

        doc.add_paragraph()
        log(f"   ✓ '{name}': {len(matrix)} строк × {n_cols} столбцов"
            + (f", объединений: {len(merges)}" if merges else ""))
        transferred += 1

    if transferred == 0:
        raise ValueError("В файле Excel не найдено данных для переноса.")

    doc.save(out_path)
    log(f"\n✓ Готово: {out_path}")


# ============================================================
# GUI
# ============================================================
class App:
    def __init__(self, root):
        self.root = root
        root.title("Word ↔ Excel: перенос таблиц в обе стороны")
        root.geometry("780x640")
        root.minsize(640, 480)

        self.direction = tk.StringVar(value="w2e")
        self.match_header = tk.BooleanVar(value=False)
        self.all_sheets = tk.BooleanVar(value=True)

        self.w2e_word = tk.StringVar()
        self.w2e_excel = tk.StringVar()
        self.w2e_out = tk.StringVar()
        self.e2w_excel = tk.StringVar()
        self.e2w_out = tk.StringVar()

        pad = {"padx": 10, "pady": 4}

        # --- Направление ---
        tk.Label(root, text="Направление переноса:", anchor="w",
                 font=("Arial", 11, "bold")).pack(fill="x", **pad)
        dir_frame = tk.Frame(root)
        dir_frame.pack(fill="x", padx=10)
        tk.Radiobutton(dir_frame, text="Word → Excel", variable=self.direction,
                       value="w2e", command=self.on_direction_change).pack(side="left", padx=5)
        tk.Radiobutton(dir_frame, text="Excel → Word", variable=self.direction,
                       value="e2w", command=self.on_direction_change).pack(side="left", padx=5)

        # --- Панель Word -> Excel ---
        self.frame_w2e = tk.Frame(root)
        self.frame_w2e.pack(fill="x")
        self._add_file_row(self.frame_w2e, "Файл Word (.docx):", self.w2e_word,
                           self.pick_w2e_word)
        self._add_file_row(self.frame_w2e, "Файл Excel (.xlsx), необязательно:",
                           self.w2e_excel, self.pick_w2e_excel)
        self._add_file_row(self.frame_w2e, "Куда сохранить (.xlsx):", self.w2e_out,
                           self.pick_w2e_out)
        tk.Checkbutton(self.frame_w2e,
                       text="Разложить по шапке существующего Excel (иначе копировать как есть)",
                       variable=self.match_header).pack(anchor="w", padx=10, pady=2)

        # --- Панель Excel -> Word ---
        self.frame_e2w = tk.Frame(root)
        self._add_file_row(self.frame_e2w, "Файл Excel (.xlsx):", self.e2w_excel,
                           self.pick_e2w_excel)
        self._add_file_row(self.frame_e2w, "Документ Word (.docx):", self.e2w_out,
                           self.pick_e2w_out)
        tk.Checkbutton(self.frame_e2w,
                       text="Все листы Excel → отдельные таблицы в Word",
                       variable=self.all_sheets).pack(anchor="w", padx=10, pady=2)

        # --- Кнопка ---
        self.btn_run = tk.Button(root, text="▶  ЗАПУСТИТЬ",
                                 command=self.run,
                                 bg="#4CAF50", fg="white",
                                 font=("Arial", 13, "bold"),
                                 height=2, cursor="hand2")
        self.btn_run.pack(fill="x", padx=10, pady=10)

        # --- Лог ---
        tk.Label(root, text="Журнал:", anchor="w").pack(fill="x", padx=10)
        log_frame = tk.Frame(root)
        log_frame.pack(fill="both", expand=True, padx=10, pady=(0, 10))

        self.log_box = tk.Text(log_frame, wrap="word", height=15)
        scroll = tk.Scrollbar(log_frame, command=self.log_box.yview)
        self.log_box.configure(yscrollcommand=scroll.set)
        self.log_box.pack(side="left", fill="both", expand=True)
        scroll.pack(side="right", fill="y")

        self.on_direction_change()

    # ---------- построение интерфейса ----------
    def _add_file_row(self, parent, label, var, command):
        tk.Label(parent, text=label, anchor="w").pack(fill="x", padx=10, pady=(6, 0))
        f = tk.Frame(parent)
        f.pack(fill="x", padx=10)
        tk.Entry(f, textvariable=var).pack(side="left", fill="x", expand=True)
        tk.Button(f, text="Выбрать…", width=12, command=command).pack(side="left", padx=5)

    def on_direction_change(self):
        if self.direction.get() == "w2e":
            self.frame_w2e.pack(fill="x", before=self.btn_run)
            self.frame_e2w.pack_forget()
        else:
            self.frame_e2w.pack(fill="x", before=self.btn_run)
            self.frame_w2e.pack_forget()

    # ---------- журналирование (потокобезопасно) ----------
    def log(self, msg):
        self.root.after(0, self._append_log, msg)

    def _append_log(self, msg):
        self.log_box.insert("end", msg + "\n")
        self.log_box.see("end")

    # ---------- выбор файлов ----------
    def pick_w2e_word(self):
        p = filedialog.askopenfilename(title="Выберите Word-файл",
                                       filetypes=[("Word", "*.docx"), ("Все файлы", "*.*")])
        if p:
            self.w2e_word.set(p)
            self.auto_out_w2e()

    def pick_w2e_excel(self):
        p = filedialog.askopenfilename(title="Выберите Excel-файл (шаблон)",
                                       filetypes=[("Excel", "*.xlsx"), ("Все файлы", "*.*")])
        if p:
            self.w2e_excel.set(p)
            self.auto_out_w2e()

    def pick_w2e_out(self):
        p = filedialog.asksaveasfilename(title="Сохранить результат как",
                                         defaultextension=".xlsx",
                                         filetypes=[("Excel", "*.xlsx")])
        if p:
            self.w2e_out.set(p)

    def pick_e2w_excel(self):
        p = filedialog.askopenfilename(title="Выберите Excel-файл",
                                       filetypes=[("Excel", "*.xlsx"), ("Все файлы", "*.*")])
        if p:
            self.e2w_excel.set(p)
            self.auto_out_e2w()

    def pick_e2w_out(self):
        p = filedialog.asksaveasfilename(title="Сохранить документ Word как",
                                         defaultextension=".docx",
                                         filetypes=[("Word", "*.docx")])
        if p:
            self.e2w_out.set(p)

    def auto_out_w2e(self):
        base = self.w2e_excel.get() or self.w2e_word.get()
        if base:
            b, _ = os.path.splitext(base)
            self.w2e_out.set(f"{b}_result.xlsx")

    def auto_out_e2w(self):
        base = self.e2w_excel.get()
        if base:
            b, _ = os.path.splitext(base)
            self.e2w_out.set(f"{b}_result.docx")

    # ---------- запуск ----------
    def run(self):
        if self.direction.get() == "w2e":
            w = self.w2e_word.get()
            e = self.w2e_excel.get()
            o = self.w2e_out.get()
            if not w:
                messagebox.showwarning("Не хватает данных", "Укажите файл Word.")
                return
            match = self.match_header.get()
            if match and not e:
                messagebox.showwarning("Не хватает данных",
                                       "Для сопоставления по шапке укажите файл Excel.")
                return
            if not o:
                self.auto_out_w2e()
                o = self.w2e_out.get()
            self._start(self._worker_w2e, (w, e, o, match))
        else:
            e = self.e2w_excel.get()
            o = self.e2w_out.get()
            if not e:
                messagebox.showwarning("Не хватает данных", "Укажите файл Excel.")
                return
            if not o:
                self.auto_out_e2w()
                o = self.e2w_out.get()
            self._start(self._worker_e2w, (e, o, self.all_sheets.get()))

    def _start(self, worker, args):
        self.log_box.delete("1.0", "end")
        self.btn_run.config(state="disabled", text="Обработка…")
        threading.Thread(target=self._run_worker, args=(worker, args), daemon=True).start()

    def _run_worker(self, worker, args):
        try:
            worker(*args)
        except Exception as ex:
            self.log("\nОШИБКА:\n" + traceback.format_exc())
            self.root.after(0, messagebox.showerror, "Ошибка", str(ex))
        finally:
            self.root.after(0, lambda: self.btn_run.config(state="normal", text="▶  ЗАПУСТИТЬ"))

    def _worker_w2e(self, w, e, o, match):
        self.log(f"Режим: Word → Excel\n")
        if match:
            word_to_excel_by_header(w, e, o, self.log)
        else:
            word_to_excel_as_is(w, e, o, self.log)
        self.root.after(0, messagebox.showinfo, "Готово", f"Файл сохранён:\n{o}")

    def _worker_e2w(self, e, o, all_sheets):
        self.log(f"Режим: Excel → Word\n")
        excel_to_word(e, o, self.log, all_sheets)
        self.root.after(0, messagebox.showinfo, "Готово", f"Документ сохранён:\n{o}")


if __name__ == "__main__":
    root = tk.Tk()
    App(root)
    root.mainloop()
