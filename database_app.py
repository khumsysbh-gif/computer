#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
ໂປຣແກຣມບັນທຶກຖານຂໍ້ມູນບຸກຄະລາກອນ (ເພີ່ມ / ແກ້ໄຂ / ລຶບ / ຄົ້ນຫາ / ສົ່ງອອກ Excel)
ຮອງຮັບ 3 ພາສາ: ລາວ, English, ไทย
"""
#python -m streamlit run database_app.py
import hashlib
import hmac
import os
import sqlite3
import tkinter as tk
from datetime import datetime
from io import BytesIO
from tkinter import filedialog, messagebox, simpledialog, ttk

try:
    from PIL import Image, ImageOps, ImageTk
    from openpyxl import Workbook
    from openpyxl.drawing.image import Image as XLImage
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter
    MISSING_LIB = None
except ImportError as _e:
    MISSING_LIB = str(_e)

DB_FILE = "data.db"
PBKDF2_ROUNDS = 200_000
MAX_LOGIN_ATTEMPTS = 3
MIN_PASSWORD_LEN = 6

DATE_FORMAT = "%d/%m/%Y"
DATE_HINT = "ວວ/ດດ/ປປປປ"

PHOTO_SIZE = (300, 400)
PREVIEW_SIZE = (150, 200)

# ລະບົບຄຳສັບ 3 ພາສາ (Localization Dictionary)
LANG = {
    "lo": {
        "title": "ໂປຣແກຣມບັນທຶກຖານຂໍ້ມູນບຸກຄະລາກອນ",
        "form": "ຂໍ້ມູນ",
        "photo": "ຮູບ 3×4",
        "pick_photo": "📷 ເລືອກຮູບ",
        "remove_photo": "ລຶບຮູບ",
        "add": "➕ ເພີ່ມ",
        "update": "✏️ ບັນທຶກການແກ້ໄຂ",
        "delete": "🗑 ລຶບ",
        "clear": "ລ້າງຟອມ",
        "export": "📥 ດາວໂຫຼດ Excel",
        "change_pw": "🔑 ປ່ຽນລະຫັດຜ່ານ",
        "logout": "🚪 ອອກຈາກລະບົບ",
        "search": "ຄົ້ນຫາ",
        "show_all": "ສະແດງທັງໝົດ",
        "no_photo": "ຮູບ 3×4\n(ຍັງບໍ່ມີຮູບ)",
        "required_hint": "* ຈຳເປັນຕ້ອງໃສ່     ຮູບແບບວັນທີ: ວວ/ດດ/ປປປປ  (ເຊັ່ນ 25/10/2026)",
        "select_row_err": "ກະລຸນາເລືອກແຖວທີ່ຕ້ອງການກ່ອນ",
        "success": "ສຳເລັດ",
        "error": "ຜິດພາດ",
    },
    "en": {
        "title": "Personnel Database Management System",
        "form": "Information",
        "photo": "Photo 3×4",
        "pick_photo": "📷 Select Photo",
        "remove_photo": "Remove Photo",
        "add": "➕ Add",
        "update": "✏️ Save Update",
        "delete": "🗑 Delete",
        "clear": "Clear Form",
        "export": "📥 Export Excel",
        "change_pw": "🔑 Change Password",
        "logout": "🚪 Logout",
        "search": "Search",
        "show_all": "Show All",
        "no_photo": "Photo 3×4\n(No Photo)",
        "required_hint": "* Required     Date format: DD/MM/YYYY  (e.g., 25/10/2026)",
        "select_row_err": "Please select a row first",
        "success": "Success",
        "error": "Error",
    },
    "th": {
        "title": "โปรแกรมบันทึกฐานข้อมูลบุคลากร",
        "form": "ข้อมูล",
        "photo": "รูป 3×4",
        "pick_photo": "📷 เลือกรูป",
        "remove_photo": "ลบรูป",
        "add": "➕ เพิ่ม",
        "update": "✏️ บันทึกการแก้ไข",
        "delete": "🗑 ลบ",
        "clear": "ล้างฟอร์ม",
        "export": "📥 ดาวน์โหลด Excel",
        "change_pw": "🔑 เปลี่ยนรหัสผ่าน",
        "logout": "🚪 ออกจากระบบ",
        "search": "ค้นหา",
        "show_all": "แสดงทั้งหมด",
        "no_photo": "รูป 3×4\n(ยังไม่มีรูป)",
        "required_hint": "* จำเป็นต้องใส่     รูปแบบวันที่: วว/ดด/ปปปป  (เช่น 25/10/2026)",
        "select_row_err": "กรุณาเลือกแถวที่ต้องการก่อน",
        "success": "สำเร็จ",
        "error": "ผิดพลาด",
    }
}

# ຊື່ຟີલ્ડ ແລະ ການແປພາສາໃນຟອມ
FIELDS = [
    ("military_id",     {"lo": "ເລກນາຍທະຫານ", "en": "Military ID", "th": "เลข 3 นายทหาร"}, "text"),
    ("id_card",         {"lo": "ເລກບັດປະຈຳຕົວ", "en": "ID Card", "th": "เลขบัตรประจำตัว"}, "text"),
    ("name",            {"lo": "ຊື່", "en": "Name", "th": "ชื่อ"}, "text"),
    ("surname",         {"lo": "ນາມສະກຸນ", "en": "Surname", "th": "นามสกุล"}, "text"),
    ("gender",          {"lo": "ເພດ", "en": "Gender", "th": "เพศ"}, "combo"),
    ("rank",            {"lo": "ຊັ້ນ", "en": "Rank", "th": "ชั้นยศ"}, "text"),
    ("birth_date",      {"lo": "ວັນເດືອນປີເກີດ", "en": "Birth Date", "th": "วันเดือนปีเกิด"}, "date"),
    ("revolution_date", {"lo": "ວັນເດືອນປີເຂົ້າການປະຕິວັດ", "en": "Revolution Date", "th": "วันเดือนปีเข้าปฏิวัติ"}, "date"),
    ("promotion_date",  {"lo": "ວັນເດືອນປີເລື່ອນຊັ້ນ", "en": "Promotion Date", "th": "วันเดือนปีเลื่อนชั้น"}, "date"),
    ("duty",            {"lo": "ໜ້າທີ່ຮັບຜິດຊອບ", "en": "Duty", "th": "หน้าที่รับผิดชอบ"}, "text"),
    ("party_position",  {"lo": "ຕຳແໜ່ງພັກ", "en": "Party Position", "th": "ตำแหน่งพรรค"}, "text"),
    ("state_position",  {"lo": "ຕຳແໜ່ງລັດ", "en": "State Position", "th": "ตำแหน่งรัฐ"}, "text"),
    ("unit",            {"lo": "ກົມກອງສັງກັດ", "en": "Unit", "th": "สังกัด/กรมกอง"}, "text"),
    ("reserve_party_date", {"lo": "ວັນເດືອນປີເຂົ້າພັກສຳຮອງ", "en": "Reserve Party Date", "th": "วันเข้าพรรคสำรอง"}, "date"),
    ("full_party_date",    {"lo": "ວັນເດືອນປີເຂົ້າພັກສົມບູນ", "en": "Full Party Date", "th": "วันเข้าพรรคสมบูรณ์"}, "date"),
    ("education",       {"lo": "ລະດັບການສຶກສາ", "en": "Education", "th": "ระดับการศึกษา"}, "combo"),
    ("phone",           {"lo": "ເບີໂທ", "en": "Phone", "th": "เบอร์โทร"}, "text"),
    ("address",         {"lo": "ທີ່ຢູ່", "en": "Address", "th": "ที่อยู่"}, "text"),
    ("note",            {"lo": "ໝາຍເຫດ", "en": "Note", "th": "หมายเหตุ"}, "text"),
]
FIELD_KEYS = [f[0] for f in FIELDS]

COMBO_OPTIONS = {
    "gender": {
        "lo": ["ຊາຍ", "ຍິງ"], 
        "en": ["Male", "Female"], 
        "th": ["ชาย", "หญิง"]
    },
    "education": {
        "lo": ["ປະຖົມ", "ມັດທະຍົມຕອນຕົ້ນ", "ມັດທະຍົມຕອນປາຍ", "ຊັ້ນກາງ / ວິຊາຊີບ", "ຊັ້ນສູງ", "ປະລິນຍາຕີ", "ປະລິນຍາໂທ", "ປະລິນຍາເອກ"],
        "en": ["Primary", "Lower Secondary", "Upper Secondary", "Vocational/Diploma", "Higher Diploma", "Bachelor's", "Master's", "Doctorate"],
        "th": ["ประถม", "มัธยมศึกษาตอนต้น", "มัธยมศึกษาตอนปลาย", "ปวส./อนุปริญญา", "ปริญญาตรี", "ปริญญาโท", "ปริญญาเอก"]
    },
}

TABLE_COLS = ["military_id", "name", "surname", "gender", "rank", "unit", "party_position", "phone"]
TABLE_WIDTHS = {"military_id": 120, "name": 130, "surname": 130, "gender": 60, "rank": 100, "unit": 180, "party_position": 140, "phone": 110}


class Database:
    def __init__(self, path=DB_FILE):
        self.conn = sqlite3.connect(path)
        self.conn.row_factory = sqlite3.Row
        self._create_tables()
        self._migrate()

    def _create_tables(self):
        with self.conn:
            self.conn.execute("""CREATE TABLE IF NOT EXISTS auth (id INTEGER PRIMARY KEY CHECK (id = 1), salt BLOB NOT NULL, pw_hash BLOB NOT NULL)""")
            self.conn.execute("""CREATE TABLE IF NOT EXISTS records (id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT NOT NULL)""")

    def _migrate(self):
        existing = {r["name"] for r in self.conn.execute("PRAGMA table_info(records)")}
        with self.conn:
            for key in FIELD_KEYS:
                if key not in existing:
                    self.conn.execute(f"ALTER TABLE records ADD COLUMN {key} TEXT")
            if "photo" not in existing:
                self.conn.execute("ALTER TABLE records ADD COLUMN photo BLOB")

    @staticmethod
    def _hash(password: str, salt: bytes) -> bytes:
        return hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), salt, PBKDF2_ROUNDS)

    def has_password(self) -> bool:
        return self.conn.execute("SELECT 1 FROM auth WHERE id = 1").fetchone() is not None

    def set_password(self, password: str):
        salt = os.urandom(16)
        pw_hash = self._hash(password, salt)
        with self.conn:
            self.conn.execute("INSERT OR REPLACE INTO auth (id, salt, pw_hash) VALUES (1, ?, ?)", (salt, pw_hash))

    def verify_password(self, password: str) -> bool:
        row = self.conn.execute("SELECT salt, pw_hash FROM auth WHERE id = 1").fetchone()
        return False if row is None else hmac.compare_digest(self._hash(password, row["salt"]), row["pw_hash"])

    def add(self, data: dict, photo: bytes = None):
        cols = FIELD_KEYS + ["photo"]
        vals = [data.get(k, "") for k in FIELD_KEYS] + [photo]
        sql = f"INSERT INTO records ({', '.join(cols)}) VALUES ({', '.join('?' * len(cols))})"
        with self.conn:
            self.conn.execute(sql, vals)

    def update(self, rec_id, data: dict, photo: bytes = None):
        cols = FIELD_KEYS + ["photo"]
        vals = [data.get(k, "") for k in FIELD_KEYS] + [photo, rec_id]
        sql = f"UPDATE records SET {', '.join(c + '=?' for c in cols)} WHERE id=?"
        with self.conn:
            self.conn.execute(sql, vals)

    def delete(self, rec_id):
        with self.conn:
            self.conn.execute("DELETE FROM records WHERE id=?", (rec_id,))

    def get(self, rec_id):
        return self.conn.execute("SELECT * FROM records WHERE id=?", (rec_id,)).fetchone()

    def all(self):
        return self.conn.execute("SELECT id, " + ", ".join(TABLE_COLS) + " FROM records ORDER BY id DESC").fetchall()

    def all_full(self):
        return self.conn.execute("SELECT * FROM records ORDER BY id ASC").fetchall()

    def search(self, keyword):
        like = f"%{keyword}%"
        where = " OR ".join(f"{k} LIKE ?" for k in FIELD_KEYS)
        return self.conn.execute(f"SELECT id, " + ", ".join(TABLE_COLS) + f" FROM records WHERE {where} ORDER BY id DESC", [like] * len(FIELD_KEYS)).fetchall()


def process_photo(path: str) -> bytes:
    img = Image.open(path)
    img = ImageOps.exif_transpose(img).convert("RGB")
    w, h = img.size
    target = PHOTO_SIZE[0] / PHOTO_SIZE[1]
    if w / h > target:
        new_w = int(h * target)
        left = (w - new_w) // 2
        img = img.crop((left, 0, left + new_w, h))
    else:
        new_h = int(w / target)
        top = (h - new_h) // 2
        img = img.crop((0, top, w, top + new_h))
    img = img.resize(PHOTO_SIZE, Image.LANCZOS)
    buf = BytesIO()
    img.save(buf, "JPEG", quality=90)
    return buf.getvalue()


def export_to_excel(rows, path: str, lang_code: str):
    wb = Workbook()
    ws = wb.active
    ws.title = "Personnel Data"

    labels = [f[1][lang_code] for f in FIELDS]
    headers = ["No.", "Photo 3×4"] + labels
    thin = Side(style="thin", color="999999")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    head_fill = PatternFill("solid", fgColor="D9E1F2")
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left = Alignment(horizontal="left", vertical="center", wrap_text=True)

    for c, text in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=c, value=text)
        cell.font = Font(name="Phetsarath OT", bold=True)
        cell.fill = head_fill
        cell.alignment = center
        cell.border = border
    ws.row_dimensions[1].height = 36

    for n, row in enumerate(rows, start=1):
        r = n + 1
        ws.cell(row=r, column=1, value=n)
        ws.cell(row=r, column=2, value="")
        for j, key in enumerate(FIELD_KEYS, start=3):
            cell_val = ws.cell(row=r, column=j, value=row[key] or "")
            cell_val.font = Font(name="Phetsarath OT")
        for c in range(1, len(headers) + 1):
            cell = ws.cell(row=r, column=c)
            cell.border = border
            cell.alignment = center if c <= 2 else left
        ws.row_dimensions[r].height = 80
        if row["photo"]:
            xl_img = XLImage(BytesIO(row["photo"]))
            xl_img.width, xl_img.height = 75, 100
            ws.add_image(xl_img, f"B{r}")

    ws.column_dimensions["A"].width = 8
    ws.column_dimensions["B"].width = 14
    wb.save(path)


def authenticate(db: Database) -> bool:
    root = tk.Tk()
    root.withdraw()
    try:
        if not db.has_password():
            messagebox.showinfo("Setup", "First time setup. Please set your password.")
            while True:
                pw1 = simpledialog.askstring("Setup", f"New Password (min {MIN_PASSWORD_LEN} chars):", show="*")
                if not pw1: return False
                if len(pw1) < MIN_PASSWORD_LEN: continue
                pw2 = simpledialog.askstring("Confirm", "Re-enter password:", show="*")
                if pw1 == pw2:
                    db.set_password(pw1)
                    return True
        for attempt in range(1, MAX_LOGIN_ATTEMPTS + 1):
            pw = simpledialog.askstring("Login", f"Password (Attempt {attempt}/{MAX_LOGIN_ATTEMPTS}):", show="*")
            if not pw: return False
            if db.verify_password(pw): return True
        return False
    finally:
        root.destroy()


class App(tk.Tk):
    def __init__(self, db: Database):
        super().__init__()
        self.db = db
        self.current_lang = "lo"  # ພາສາເລີ່ມຕົ້ນ (ລາວ)
        self.selected_id = None
        self.photo_bytes = None
        self._photo_tk = None
        self.is_logged_out = False

        self.geometry("1320x800")
        self.minsize(1100, 680)
        self.option_add("*Font", "Phetsarath 11")

        self.vars = {k: tk.StringVar() for k in FIELD_KEYS}
        
        self.ui_elements = {}
        self._build_ui()
        self.refresh()

    def _t(self, key):
        return LANG[self.current_lang].get(key, key)

    def _build_ui(self):
        # ລຶບ Widget ເກົ່າຖ້າມີການຣີເຟຣຊພາສາ
        for widget in self.winfo_children():
            widget.destroy()

        self.title(self._t("title"))

        # Toolbar & Language Switcher
        bar = ttk.Frame(self)
        bar.pack(fill="x", padx=10, pady=5)

        self.btn_add = ttk.Button(bar, text=self._t("add"), command=self.on_add)
        self.btn_add.pack(side="left", padx=3)
        self.btn_update = ttk.Button(bar, text=self._t("update"), command=self.on_update)
        self.btn_update.pack(side="left", padx=3)
        self.btn_delete = ttk.Button(bar, text=self._t("delete"), command=self.on_delete)
        self.btn_delete.pack(side="left", padx=3)
        self.btn_clear = ttk.Button(bar, text=self._t("clear"), command=self.clear_form)
        self.btn_clear.pack(side="left", padx=3)
        self.btn_export = ttk.Button(bar, text=self._t("export"), command=self.on_export)
        self.btn_export.pack(side="left", padx=(15, 3))
        self.btn_changepw = ttk.Button(bar, text=self._t("change_pw"), command=self.on_change_password)
        self.btn_changepw.pack(side="left", padx=3)
        self.btn_logout = ttk.Button(bar, text=self._t("logout"), command=self.on_logout)
        self.btn_logout.pack(side="left", padx=(15, 3))

        # 🌐 Combobox ເລືອກພາສາ
        self.lang_cb = ttk.Combobox(bar, values=["ລາວ", "English", "ไทย"], width=10, state="readonly")
        lang_map = {"lo": "ລາວ", "en": "English", "th": "ไทย"}
        self.lang_cb.set(lang_map[self.current_lang])
        self.lang_cb.pack(side="left", padx=15)
        self.lang_cb.bind("<<ComboboxSelected>>", self.change_language)

        # Search Bar
        ttk.Button(bar, text=self._t("search"), command=self.on_search).pack(side="right", padx=3)
        self.var_search = tk.StringVar()
        entry = ttk.Entry(bar, textvariable=self.var_search, width=20)
        entry.pack(side="right", padx=3)
        entry.bind("<Return>", lambda e: self.on_search())
        ttk.Button(bar, text=self._t("show_all"), command=self.refresh).pack(side="right", padx=3)

        # Form Section
        top = ttk.Frame(self)
        top.pack(fill="x", padx=10, pady=(5, 5))

        self.frm_label = ttk.LabelFrame(top, text=self._t("form"))
        self.frm_label.pack(side="left", fill="both", expand=True)

        per_row = 3
        for i, (key, labels, kind) in enumerate(FIELDS):
            r, c = divmod(i, per_row)
            txt = labels[self.current_lang] + (" *" if key == "name" else "")
            ttk.Label(self.frm_label, text=txt).grid(row=r, column=c * 2, padx=(8, 4), pady=5, sticky="e")
            
            if kind == "combo":
                opts = COMBO_OPTIONS[key][self.current_lang]
                w = ttk.Combobox(self.frm_label, textvariable=self.vars[key], values=opts, width=21, state="readonly")
            else:
                w = ttk.Entry(self.frm_label, textvariable=self.vars[key], width=24)
            w.grid(row=r, column=c * 2 + 1, padx=(0, 10), pady=5, sticky="w")

        last_row = (len(FIELDS) - 1) // per_row + 1
        ttk.Label(self.frm_label, text=self._t("required_hint"), foreground="#666666").grid(
            row=last_row, column=0, columnspan=per_row * 2, padx=8, pady=(2, 6), sticky="w"
        )

        # Photo Frame
        self.pf_label = ttk.LabelFrame(top, text=self._t("photo"))
        self.pf_label.pack(side="left", padx=(10, 0), fill="y")
        self.canvas = tk.Canvas(self.pf_label, width=PREVIEW_SIZE[0], height=PREVIEW_SIZE[1], bg="white", highlightthickness=1, highlightbackground="#888888")
        self.canvas.pack(padx=10, pady=(10, 6))
        
        btns = ttk.Frame(self.pf_label)
        btns.pack(pady=(0, 8))
        self.btn_pick = ttk.Button(btns, text=self._t("pick_photo"), command=self.on_pick_photo)
        self.btn_pick.pack(side="left", padx=3)
        self.btn_rem = ttk.Button(btns, text=self._t("remove_photo"), command=self.on_remove_photo)
        self.btn_rem.pack(side="left", padx=3)

        self._show_photo()

        # Table Section
        table_frame = ttk.Frame(self)
        table_frame.pack(fill="both", expand=True, padx=10, pady=(5, 10))

        cols = ["id"] + TABLE_COLS
        self.tree = ttk.Treeview(table_frame, columns=cols, show="headings", selectmode="browse")
        self.tree.heading("id", text="ID")
        self.tree.column("id", width=50, anchor="center")
        for col_name in TABLE_COLS:
            # ຊອກຫາປ້າຍຊື່ຕາມພາສາປັດຈຸບັນ
            label_text = next(f[1][self.current_lang] for f in FIELDS if f[0] == col_name)
            self.tree.heading(col_name, text=label_text)
            self.tree.column(col_name, width=TABLE_WIDTHS[col_name], anchor="w")

        sb = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        self.tree.bind("<<TreeviewSelect>>", self.on_select)

        self.status = tk.StringVar()
        ttk.Label(self, textvariable=self.status, anchor="w").pack(fill="x", padx=10, pady=(0, 6))

    def change_language(self, _event=None):
        val = self.lang_cb.get()
        if val == "ລາວ": self.current_lang = "lo"
        elif val == "English": self.current_lang = "en"
        elif val == "ไทย": self.current_lang = "th"
        self._build_ui()
        self.fill_table(self.db.all())

    def _show_photo(self):
        self.canvas.delete("all")
        if self.photo_bytes:
            img = Image.open(BytesIO(self.photo_bytes)).resize(PREVIEW_SIZE, Image.LANCZOS)
            self._photo_tk = ImageTk.PhotoImage(img)
            self.canvas.create_image(0, 0, anchor="nw", image=self._photo_tk)
        else:
            self._photo_tk = None
            self.canvas.create_text(PREVIEW_SIZE[0] // 2, PREVIEW_SIZE[1] // 2, text=self._t("no_photo"), fill="#999999", justify="center")

    def on_pick_photo(self):
        path = filedialog.askopenfilename(title="Select Photo", filetypes=[("Image files", "*.jpg *.jpeg *.png *.bmp"), ("All files", "*.*")])
        if not path: return
        try:
            self.photo_bytes = process_photo(path)
        except Exception as e:
            messagebox.showerror(self._t("error"), str(e))
            return
        self._show_photo()

    def on_remove_photo(self):
        self.photo_bytes = None
        self._show_photo()

    def _read_form(self):
        data = {k: self.vars[k].get().strip() for k in FIELD_KEYS}
        if not data["name"]:
            messagebox.showwarning("Warning", "Please input Name")
            return None
        return data

    def clear_form(self):
        self.selected_id = None
        for v in self.vars.values(): v.set("")
        self.photo_bytes = None
        self._show_photo()
        if self.tree.selection():
            self.tree.selection_remove(self.tree.selection())

    def fill_table(self, rows):
        self.tree.delete(*self.tree.get_children())
        for r in rows:
            vals = [r["id"]] + [r[c] or "" for c in TABLE_COLS]
            self.tree.insert("", "end", values=vals)
        self.status.set(f"Total: {len(rows)}")

    def refresh(self):
        self.fill_table(self.db.all())
        self.clear_form()

    def on_select(self, _event=None):
        sel = self.tree.selection()
        if not sel: return
        rec_id = int(self.tree.item(sel[0], "values")[0])
        rec = self.db.get(rec_id)
        if rec is None: return
        self.selected_id = rec_id
        for k in FIELD_KEYS:
            self.vars[k].set(rec[k] or "")
        self.photo_bytes = rec["photo"]
        self._show_photo()

    def on_add(self):
        data = self._read_form()
        if not data: return
        self.db.add(data, self.photo_bytes)
        self.refresh()
        self.status.set("Added successfully")

    def on_update(self):
        if self.selected_id is None:
            messagebox.showinfo("Info", self._t("select_row_err"))
            return
        data = self._read_form()
        if not data: return
        self.db.update(self.selected_id, data, self.photo_bytes)
        self.refresh()
        self.status.set("Updated successfully")

    def on_delete(self):
        if self.selected_id is None:
            messagebox.showinfo("Info", self._t("select_row_err"))
            return
        if messagebox.askyesno("Confirm", "Delete this record?"):
            self.db.delete(self.selected_id)
            self.refresh()
            self.status.set("Deleted successfully")

    def on_search(self):
        kw = self.var_search.get().strip()
        if not kw:
            self.refresh()
            return
        self.fill_table(self.db.search(kw))

    def on_export(self):
        rows = self.db.all_full()
        if not rows: return
        path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")])
        if not path: return
        export_to_excel(rows, path, self.current_lang)
        messagebox.showinfo(self._t("success"), f"Exported successfully to:\n{path}")

    def on_change_password(self):
        old = simpledialog.askstring("Change Password", "Current password:", show="*")
        if not old or not self.db.verify_password(old):
            messagebox.showerror("Error", "Incorrect current password")
            return
        new1 = simpledialog.askstring("Change Password", f"New password (min {MIN_PASSWORD_LEN}):", show="*")
        if not new1 or len(new1) < MIN_PASSWORD_LEN: return
        new2 = simpledialog.askstring("Change Password", "Confirm new password:", show="*")
        if new1 != new2:
            messagebox.showerror("Error", "Passwords do not match")
            return
        self.db.set_password(new1)
        messagebox.showinfo(self._t("success"), "Password updated")

    def on_logout(self):
        if messagebox.askyesno("Logout", "Are you sure you want to logout?"):
            self.is_logged_out = True
            self.destroy()

def main():
    if MISSING_LIB:
        return
    db = Database()
    while True:
        if not authenticate(db): break
        app = App(db)
        app.mainloop()
        if not getattr(app, "is_logged_out", False): break

if __name__ == "__main__":
    main()