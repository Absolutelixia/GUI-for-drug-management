# path: app/medicine_gui.py
import tkinter as tk
from tkinter import ttk, messagebox
import pandas as pd
from pathlib import Path
import sys

# 尝试 Pillow（为加载 JPEG）；无 Pillow 时仅支持 PNG/GIF
try:
    from PIL import Image, ImageTk
    PIL_OK = True
except Exception:
    PIL_OK = False

def resource_path(relpath: str) -> Path:
    """返回打包/源码均可用的资源绝对路径。"""
    # 为什么：PyInstaller onefile 解压到 _MEIPASS；onedir 资源在 exe 同目录；源码在脚本目录
    if getattr(sys, "frozen", False):
        base = Path(getattr(sys, "_MEIPASS", Path(sys.executable).parent))
    else:
        base = Path(__file__).resolve().parent
    return (base / relpath).resolve()

EXCEL_PATH = resource_path("Drug list.xlsx")
LOGO_PATH = resource_path("logo.jpg")  # 如用 PNG，请改成 logo.png

# ---------- 核心修复：统一布尔解析 ----------
def parse_bool(val, default=False) -> bool:
    """将各种文本/数值安全解析为布尔。避免 bool('False') == True 的陷阱。"""
    if isinstance(val, bool):
        return val
    if val is None:
        return default
    # 数值
    if isinstance(val, (int, float)):
        if pd.isna(val):
            return default
        return val != 0
    s = str(val).strip().lower()
    if s in {"true", "t", "yes", "y", "1", "✓", "✔", "是", "對", "对"}:
        return True
    if s in {"false", "f", "no", "n", "0", "✗", "x", "否", "不"}:
        return False
    if s in {"", "none", "nan"}:
        return default
    # 其它未知文本，按 default
    return default

class MedicineGUI:
    def __init__(self, master):
        self.master = master
        self.master.title("Medicine Inventory · Team Astro. & Co.")
        self.master.geometry("1280x860")
        self.master.minsize(1300, 720)

        try:
            self.style = ttk.Style()
            for theme in ("vista", "clam", "default"):
                if theme in self.style.theme_names():
                    self.style.theme_use(theme); break
            base_font = ("Segoe UI", 10)
            title_font = ("Segoe UI", 16, "bold")
            self.master.option_add("*Font", base_font)
            self.style.configure("TFrame", padding=8)
            self.style.configure("TLabel", padding=(2, 2))
            self.style.configure("TButton", padding=(10, 6))
            self.style.configure("Header.TLabel", font=title_font)
            self.style.configure("Card.TLabelframe", padding=12)
            self.style.configure("Card.TLabelframe.Label", font=("Segoe UI", 11, "bold"))
            self.style.configure("Treeview", rowheight=26)
            self.style.configure("Treeview.Heading", font=("Segoe UI", 10, "bold"))
        except Exception:
            pass

        self.medicine_list = []

        # ---- Top Header ----
        header = ttk.Frame(self.master)
        header.grid(row=0, column=0, sticky="ew")
        header.columnconfigure(0, weight=1)
        header.columnconfigure(1, weight=0)  # 右侧放 logo

        title = ttk.Label(header, text="Medicine Inventory", style="Header.TLabel")
        subtitle = ttk.Label(header, text="Team Astro.& Co.")
        title.grid(row=0, column=0, sticky="w")
        subtitle.grid(row=1, column=0, sticky="w", pady=(0, 4))

        # ---- Logo（两行高度）----
        self.logo_img = None  # 防 GC
        logo_h = 50  # 大约两行文字高度
        logo_padx = (12, 0)
        logo_path = LOGO_PATH
        try:
            if logo_path.exists():
                if PIL_OK:
                    im = Image.open(logo_path)
                    # 等比缩放到指定高度
                    w, h = im.size
                    if h != 0:
                        new_w = max(1, int(w * (logo_h / h)))
                        im = im.resize((new_w, logo_h), Image.LANCZOS)
                    self.logo_img = ImageTk.PhotoImage(im)
                else:
                    # 无 Pillow 时：仅当为 PNG/GIF 可用
                    self.logo_img = tk.PhotoImage(file=str(logo_path))
                logo_label = ttk.Label(header, image=self.logo_img)
                # 跨两行，贴右侧
                logo_label.grid(row=0, column=1, rowspan=2, sticky="e", padx=logo_padx)
        except Exception:
            # 静默失败：不影响主功能
            pass

        # ---- Form Card ----
        form = ttk.Labelframe(self.master, text="Record")
        form.grid(row=1, column=0, sticky="ew", padx=8, pady=6)
        for c in range(8): form.columnconfigure(c, weight=1)

        self.where_label = ttk.Label(form, text="Where:");      self.where_entry = ttk.Entry(form)
        self.location_label = ttk.Label(form, text="Location:"); self.location_entry = ttk.Entry(form)
        self.name_label = ttk.Label(form, text="Name:");         self.name_entry = ttk.Entry(form)
        self.quantity_label = ttk.Label(form, text="Quantity:"); self.quantity_entry = ttk.Entry(form)
        self.function_label = ttk.Label(form, text="Function:"); self.function_entry = ttk.Entry(form)
        self.last_one_var = tk.BooleanVar()
        self.last_one_checkbutton = ttk.Checkbutton(form, text="Last One?", variable=self.last_one_var)
        self.to_order_var = tk.BooleanVar()
        self.to_order_checkbutton = ttk.Checkbutton(form, text="To Order?", variable=self.to_order_var)
        self.signature_label = ttk.Label(form, text="Note:");    self.signature_entry = ttk.Entry(form)
        self.reference_label = ttk.Label(form, text="Reference:"); self.reference_entry = ttk.Entry(form)

        # Grid positions
        self.where_label.grid(row=0, column=0, sticky="w", padx=(4, 2), pady=4)
        self.where_entry.grid(row=0, column=1, sticky="ew", padx=(2, 8), pady=4)
        self.location_label.grid(row=0, column=2, sticky="w", padx=(4, 2), pady=4)
        self.location_entry.grid(row=0, column=3, sticky="ew", padx=(2, 8), pady=4)
        self.name_label.grid(row=0, column=4, sticky="w", padx=(4, 2), pady=4)
        self.name_entry.grid(row=0, column=5, sticky="ew", padx=(2, 8), pady=4)
        self.quantity_label.grid(row=0, column=6, sticky="w", padx=(4, 2), pady=4)
        self.quantity_entry.grid(row=0, column=7, sticky="ew", padx=(2, 8), pady=4)

        self.function_label.grid(row=1, column=0, sticky="w", padx=(4, 2), pady=4)
        self.function_entry.grid(row=1, column=1, sticky="ew", padx=(2, 8), pady=4)
        self.last_one_checkbutton.grid(row=1, column=2, sticky="w", padx=(4, 2), pady=4)
        self.to_order_checkbutton.grid(row=1, column=3, sticky="w", padx=(4, 2), pady=4)
        self.signature_label.grid(row=1, column=4, sticky="w", padx=(4, 2), pady=4)
        self.signature_entry.grid(row=1, column=5, sticky="ew", padx=(2, 8), pady=4)
        self.reference_label.grid(row=1, column=6, sticky="w", padx=(4, 2), pady=4)
        self.reference_entry.grid(row=1, column=7, sticky="ew", padx=(2, 8), pady=4)

        # ---- Actions ----
        actions = ttk.Frame(self.master)
        actions.grid(row=2, column=0, sticky="ew", padx=8)
        for c in range(6): actions.columnconfigure(c, weight=0)
        actions.columnconfigure(4, weight=1)
        self.add_button = ttk.Button(actions, text="Add", command=self.add_medicine)
        self.copy_button = ttk.Button(actions, text="Copy Row(s)", command=self.copy_row)
        self.undo_button = ttk.Button(actions, text="Undo (Ctrl+Z)", command=self.undo_action)
        self.search_label = ttk.Label(actions, text="Search:")
        self.search_entry = ttk.Entry(actions)
        self.search_button = ttk.Button(actions, text="Search", command=self.search_medicine)
        self.add_button.grid(row=0, column=0, padx=(0, 8), pady=4)
        self.copy_button.grid(row=0, column=1, padx=(0, 8), pady=4)
        self.undo_button.grid(row=0, column=2, padx=(0, 8), pady=4)
        self.search_label.grid(row=0, column=3, sticky="e", padx=(8, 4))
        self.search_entry.grid(row=0, column=4, sticky="ew", padx=(0, 8))
        self.search_button.grid(row=0, column=5, padx=(0, 0))

        sep = ttk.Separator(self.master, orient="horizontal")
        sep.grid(row=3, column=0, sticky="ew", padx=8, pady=(4, 6))

        # ---- Table ----
        table_card = ttk.Frame(self.master)
        table_card.grid(row=4, column=0, sticky="nsew", padx=8, pady=(0, 8))
        self.master.rowconfigure(4, weight=1)
        table_card.rowconfigure(0, weight=1)
        table_card.columnconfigure(0, weight=1)

        columns = ("where","location","name","quantity","function","last_one","to_order","signature","reference")
        self.medicine_table = ttk.Treeview(table_card, columns=columns, show='headings', selectmode="extended")

        self.medicine_table.heading("where", text="Where")
        self.medicine_table.heading("location", text="Location")
        self.medicine_table.heading("name", text="Name")
        self.medicine_table.heading("quantity", text="Quantity")
        self.medicine_table.heading("function", text="Function")
        self.medicine_table.heading("last_one", text="Last One?")
        self.medicine_table.heading("to_order", text="To Order?")
        self.medicine_table.heading("signature", text="Note")
        self.medicine_table.heading("reference", text="Reference")

        self.medicine_table.column("where", width=130, anchor="center")
        self.medicine_table.column("location", width=130, anchor="center")
        self.medicine_table.column("name", width=200, anchor="w")
        self.medicine_table.column("quantity", width=100, anchor="center")
        self.medicine_table.column("function", width=160, anchor="w")
        self.medicine_table.column("last_one", width=100, anchor="center")
        self.medicine_table.column("to_order", width=100, anchor="center")
        self.medicine_table.column("signature", width=140, anchor="w")
        self.medicine_table.column("reference", width=200, anchor="w")

        for col in columns:
            self.medicine_table.heading(col, command=lambda c=col: self.sort_by_column(c))

        vsb = ttk.Scrollbar(table_card, orient="vertical", command=self.medicine_table.yview)
        hsb = ttk.Scrollbar(table_card, orient="horizontal", command=self.medicine_table.xview)
        self.medicine_table.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.medicine_table.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")

        self.medicine_table.tag_configure("oddrow", background="#FAFAFA")
        self.medicine_table.tag_configure("evenrow", background="#F0F0F5")
        self.medicine_table.bind("<Double-1>", self.edit_medicine)
        self.medicine_table.bind("<Button-3>", self.delete_medicine)

        self.status_var = tk.StringVar(value="Ready")
        status = ttk.Label(self.master, textvariable=self.status_var, anchor="w")
        status.grid(row=5, column=0, sticky="ew", padx=8, pady=(0, 6))

        self.undo_stack = []
        self.max_undo = 20
        self.master.bind_all('<Control-z>', lambda e: self.undo_action())
        self.master.protocol("WM_DELETE_WINDOW", self.on_close)

        self._enable_mousewheel(self.medicine_table)

        try:
            df = pd.read_excel(EXCEL_PATH)
            note_col = 'Note' if 'Note' in df.columns else ('Signature' if 'Signature' in df.columns else None)
            loc_col = 'Location' if 'Location' in df.columns else None
            ref_col = 'Reference' if 'Reference' in df.columns else None
            for _, row in df.iterrows():
                note_val = row[note_col] if note_col else ""
                loc_val = row[loc_col] if loc_col else ""
                ref_val = row[ref_col] if ref_col else ""
                last_one_val = parse_bool(row.get('Last one?', False))
                to_order_val = parse_bool(row.get('To order?', False))
                medicine = Medicine(
                    row.get('Where?', ''), loc_val, row.get('Name', ''), row.get('Quantity', ''),
                    row.get('Function', ''), last_one_val, to_order_val, note_val, ref_val
                )
                self.medicine_list.append(medicine)
                tag = "evenrow" if len(self.medicine_list) % 2 == 0 else "oddrow"
                self.medicine_table.insert("", "end", values=(
                    medicine.where, medicine.location, medicine.name, medicine.quantity, medicine.function,
                    medicine.last_one, medicine.to_order, medicine.signature, medicine.reference
                ), tags=(tag,))
            self.status_var.set(f"Loaded {len(self.medicine_list)} item(s) from Excel.")
        except Exception as e:
            self.status_var.set(f"Excel load failed: {e}")

    # ============== 辅助：滚轮/排序/撤销/保存 ==============
    def _enable_mousewheel(self, widget: ttk.Treeview):
        widget.bind("<MouseWheel>", lambda e: (widget.yview_scroll(-int(e.delta/120), "units"), "break"))
        widget.bind("<Shift-MouseWheel>", lambda e: (widget.xview_scroll(-int(e.delta/120), "units"), "break"))
        widget.bind("<Button-4>", lambda e: (widget.yview_scroll(-1, "units"), "break"))
        widget.bind("<Button-5>", lambda e: (widget.yview_scroll(1, "units"), "break"))

    def add_medicine(self):
        medicine = Medicine(
            self.where_entry.get(), self.location_entry.get(), self.name_entry.get(), self.quantity_entry.get(),
            self.function_entry.get(), self.last_one_var.get(), self.to_order_var.get(),
            self.signature_entry.get(), self.reference_entry.get()
        )
        self.medicine_list.append(medicine)
        tag = "evenrow" if len(self.medicine_list) % 2 == 0 else "oddrow"
        self.medicine_table.insert("", "end", values=(
            medicine.where, medicine.location, medicine.name, medicine.quantity, medicine.function,
            medicine.last_one, medicine.to_order, medicine.signature, medicine.reference
        ), tags=(tag,))
        self._push_undo(("add", self._snapshot(medicine)))

    def edit_medicine(self, event):
        if not self.medicine_table.selection():
            self.status_var.set("No row selected."); return
        item = self.medicine_table.selection()[0]
        medicine = self.medicine_list[self.medicine_table.index(item)]

        self.edit_window = tk.Toplevel(self.master)
        self.edit_window.title("Edit Medicine")
        self.edit_window.geometry("500x420")
        self.edit_window.resizable(False, False)

        frame = ttk.Labelframe(self.edit_window, text="Edit")
        frame.pack(fill="both", expand=True, padx=10, pady=10)
        for c in range(2): frame.columnconfigure(c, weight=1)

        ttk.Label(frame, text="Where:").grid(row=0, column=0, sticky="w", padx=4, pady=6)
        self.e_where_entry = ttk.Entry(frame); self.e_where_entry.grid(row=0, column=1, sticky="ew", padx=4, pady=6)
        self.e_where_entry.insert(0, medicine.where)

        ttk.Label(frame, text="Location:").grid(row=1, column=0, sticky="w", padx=4, pady=6)
        self.e_location_entry = ttk.Entry(frame); self.e_location_entry.grid(row=1, column=1, sticky="ew", padx=4, pady=6)
        self.e_location_entry.insert(0, medicine.location)

        ttk.Label(frame, text="Name:").grid(row=2, column=0, sticky="w", padx=4, pady=6)
        self.e_name_entry = ttk.Entry(frame); self.e_name_entry.grid(row=2, column=1, sticky="ew", padx=4, pady=6)
        self.e_name_entry.insert(0, medicine.name)

        ttk.Label(frame, text="Quantity:").grid(row=3, column=0, sticky="w", padx=4, pady=6)
        self.e_quantity_entry = ttk.Entry(frame); self.e_quantity_entry.grid(row=3, column=1, sticky="ew", padx=4, pady=6)
        self.e_quantity_entry.insert(0, medicine.quantity)

        ttk.Label(frame, text="Function:").grid(row=4, column=0, sticky="w", padx=4, pady=6)
        self.e_function_entry = ttk.Entry(frame); self.e_function_entry.grid(row=4, column=1, sticky="ew", padx=4, pady=6)
        self.e_function_entry.insert(0, medicine.function)

        ttk.Label(frame, text="Note:").grid(row=5, column=0, sticky="w", padx=4, pady=6)
        self.e_signature_entry = ttk.Entry(frame); self.e_signature_entry.grid(row=5, column=1, sticky="ew", padx=4, pady=6)
        self.e_signature_entry.insert(0, medicine.signature)

        ttk.Label(frame, text="Reference:").grid(row=6, column=0, sticky="w", padx=4, pady=6)
        self.e_reference_entry = ttk.Entry(frame); self.e_reference_entry.grid(row=6, column=1, sticky="ew", padx=4, pady=6)
        self.e_reference_entry.insert(0, medicine.reference)

        self.e_last_one_var = tk.BooleanVar(value=parse_bool(medicine.last_one))
        self.e_to_order_var = tk.BooleanVar(value=parse_bool(medicine.to_order))
        ttk.Checkbutton(frame, text="Last One?", variable=self.e_last_one_var).grid(row=7, column=0, sticky="w", padx=4, pady=6)
        ttk.Checkbutton(frame, text="To Order?", variable=self.e_to_order_var).grid(row=7, column=1, sticky="w", padx=4, pady=6)

        btns = ttk.Frame(self.edit_window); btns.pack(fill="x", padx=10, pady=(0, 10))
        ttk.Button(btns, text="Save", command=lambda: self.save_medicine(item)).pack(side="right")

    def save_medicine(self, item):
        medicine = self.medicine_list[self.medicine_table.index(item)]
        old_snap = self._snapshot(medicine)
        medicine.where = self.e_where_entry.get()
        medicine.location = self.e_location_entry.get()
        medicine.name = self.e_name_entry.get()
        medicine.quantity = self.e_quantity_entry.get()
        medicine.function = self.e_function_entry.get()
        medicine.last_one = parse_bool(self.e_last_one_var.get())
        medicine.to_order = parse_bool(self.e_to_order_var.get())
        medicine.signature = self.e_signature_entry.get()
        medicine.reference = self.e_reference_entry.get()
        self.medicine_table.item(item, values=(
            medicine.where, medicine.location, medicine.name, medicine.quantity, medicine.function,
            medicine.last_one, medicine.to_order, medicine.signature, medicine.reference
        ))
        self.edit_window.destroy()
        self._push_undo(("edit", old_snap, self._snapshot(medicine)))

    def delete_medicine(self, event):
        if not self.medicine_table.selection():
            self.status_var.set("No row selected."); return
        item = self.medicine_table.selection()[0]
        medicine = self.medicine_list[self.medicine_table.index(item)]
        if not messagebox.askyesno("Delete Medicine", f"Are you sure you want to delete {medicine.name}?"):
            return
        snap = self._snapshot(medicine)
        self.medicine_list.remove(medicine)
        self.medicine_table.delete(item)
        self._push_undo(("delete", snap))

    def search_medicine(self):
        kw = self.search_entry.get().strip().lower()
        for item in self.medicine_table.get_children():
            vals = self.medicine_table.item(item).get("values", [])
            name_val = str(vals[2]).lower() if len(vals) > 2 else ""
            func_val = str(vals[4]).lower() if len(vals) > 4 else ""
            if kw and (kw in name_val or kw in func_val):
                self.medicine_table.selection_add(item); self.medicine_table.see(item)
            else:
                self.medicine_table.selection_remove(item)

    def copy_row(self):
        selected_items = self.medicine_table.selection()
        if not selected_items:
            self.status_var.set("No row selected."); return
        lines = []
        for iid in selected_items:
            vals = self.medicine_table.item(iid).get("values", [])
            lines.append("\t".join(str(v) for v in vals))
        text = "\n".join(lines)
        self.master.clipboard_clear(); self.master.clipboard_append(text)
        self.status_var.set(f"Copied {len(selected_items)} row(s) to clipboard.")

    def copy_table(self):
        table = ""
        for column in self.medicine_table["columns"]: table += column + "\t"
        table += "\n"
        for item in self.medicine_table.get_children():
            for value in self.medicine_table.item(item)["values"]:
                table += str(value) + "\t"
            table += "\n"
        self.master.clipboard_clear(); self.master.clipboard_append(table)
        self.status_var.set("Table copied to clipboard.")

    def sort_by_column(self, col_key: str):
        col_index_map = {"where":0,"location":1,"name":2,"quantity":3,"function":4,"last_one":5,"to_order":6,"signature":7,"reference":8}
        idx = col_index_map.get(col_key, 0)
        def classify(value):
            s = "" if value is None else str(value).strip()
            low = s.lower()
            if low == "true": return (0, 1.0)
            if low == "false": return (0, 0.0)
            try:
                raw = s[1:] if s.startswith("-") else s
                if raw.replace(".", "", 1).isdigit():
                    return (0, float(s))
            except Exception:
                pass
            if s and s[0].isascii() and s[0].isalpha(): return (1, s.lower())
            return (2, s)
        children = list(self.medicine_table.get_children(""))
        items = []
        for iid in children:
            vals = self.medicine_table.item(iid).get("values", [])
            v = vals[idx] if idx < len(vals) else ""
            items.append((classify(v), iid))
        items.sort(key=lambda x: x[0])
        for new_pos, (_, iid) in enumerate(items): self.medicine_table.move(iid, "", new_pos)
        for i, iid in enumerate(self.medicine_table.get_children("")):
            self.medicine_table.item(iid, tags=("evenrow",) if i % 2 == 0 else ("oddrow",))

    def _snapshot(self, m: 'Medicine'):
        return {
            'Where?': m.where, 'Location': m.location, 'Name': m.name, 'Quantity': m.quantity,
            'Function': m.function, 'Last one?': parse_bool(m.last_one), 'To order?': parse_bool(m.to_order),
            'Note': m.signature, 'Reference': m.reference,
        }

    def _push_undo(self, action_tuple):
        self.undo_stack.append(action_tuple)
        if len(self.undo_stack) > self.max_undo: self.undo_stack.pop(0)
        self.status_var.set(f"Undo saved ({len(self.undo_stack)}/{self.max_undo}).")

    def undo_action(self, evt=None):
        if not self.undo_stack:
            self.status_var.set("Nothing to undo."); return
        action = self.undo_stack.pop()
        kind = action[0]
        if kind == "add":
            self._apply_delete_by_snapshot(action[1])
        elif kind == "delete":
            self._apply_add_from_snapshot(action[1])
        elif kind == "edit":
            self._apply_update(from_snap=action[2], to_snap=action[1])
        self.status_var.set("Undo done.")

    def _apply_add_from_snapshot(self, snap: dict):
        m = Medicine(
            snap.get('Where?',''), snap.get('Location',''), snap.get('Name',''), snap.get('Quantity',''),
            snap.get('Function',''), parse_bool(snap.get('Last one?', False)), parse_bool(snap.get('To order?', False)),
            snap.get('Note',''), snap.get('Reference','')
        )
        self.medicine_list.append(m)
        tag = "evenrow" if len(self.medicine_list) % 2 == 0 else "oddrow"
        self.medicine_table.insert("", "end", values=(
            m.where, m.location, m.name, m.quantity, m.function, m.last_one, m.to_order, m.signature, m.reference
        ), tags=(tag,))

    def _apply_delete_by_snapshot(self, snap: dict):
        idx_to_del = None
        for i, m in enumerate(self.medicine_list):
            if self._snapshot(m) == snap:
                idx_to_del = i; break
        if idx_to_del is None: return
        children = self.medicine_table.get_children()
        if idx_to_del < len(children): self.medicine_table.delete(children[idx_to_del])
        self.medicine_list.pop(idx_to_del)

    def _apply_update(self, from_snap: dict, to_snap: dict):
        idx_to_edit = None
        for i, m in enumerate(self.medicine_list):
            if self._snapshot(m) == from_snap:
                idx_to_edit = i; break
        if idx_to_edit is None: return
        m = self.micine_list[idx_to_edit] if False else self.medicine_list[idx_to_edit]
        m.where = to_snap.get('Where?',''); m.location = to_snap.get('Location',''); m.name = to_snap.get('Name','')
        m.quantity = to_snap.get('Quantity',''); m.function = to_snap.get('Function','')
        m.last_one = parse_bool(to_snap.get('Last one?', False))
        m.to_order = parse_bool(to_snap.get('To order?', False))
        m.signature = to_snap.get('Note',''); m.reference = to_snap.get('Reference','')
        children = self.medicine_table.get_children()
        if idx_to_edit < len(children):
            self.medicine_table.item(children[idx_to_edit], values=(
                m.where, m.location, m.name, m.quantity, m.function, m.last_one, m.to_order, m.signature, m.reference
            ))

    def on_close(self):
        try:
            cols = ['Where?','Location','Name','Quantity','Function','Last one?','To order?','Note','Reference']
            rows = []
            for item in self.medicine_table.get_children():
                vals = self.medicine_table.item(item).get('values', [])
                if len(vals) >= 9:
                    rows.append({
                        'Where?': vals[0], 'Location': vals[1], 'Name': vals[2], 'Quantity': vals[3],
                        'Function': vals[4], 'Last one?': parse_bool(vals[5]), 'To order?': parse_bool(vals[6]),
                        'Note': vals[7], 'Reference': vals[8],
                    })
            pd.DataFrame(rows, columns=cols).to_excel(EXCEL_PATH, index=False)
        except Exception as e:
            messagebox.showerror("Save Error", f"Failed to save Excel: {e}")
        finally:
            self.master.destroy()

class Medicine:
    def __init__(self, where, location, name, quantity, function, last_one, to_order, signature, reference):
        self.where = where; self.location = location; self.name = name; self.quantity = quantity
        self.function = function; self.last_one = last_one; self.to_order = to_order
        self.signature = signature; self.reference = reference

    def snapshot(self):
        return {
            'Where?': self.where, 'Location': self.location, 'Name': self.name, 'Quantity': self.quantity,
            'Function': self.function, 'Last one?': parse_bool(self.last_one), 'To order?': parse_bool(self.to_order),
            'Note': self.signature, 'Reference': self.reference
        }

    def __str__(self):
        return f"{self.where}, {self.location}, {self.name}, {self.quantity}, {self.function}, {self.last_one}, {self.to_order}, {self.signature}, {self.reference}"

if __name__ == "__main__":
    root = tk.Tk()
    app = MedicineGUI(root)
    root.mainloop()


