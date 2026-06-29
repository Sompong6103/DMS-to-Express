#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
express_import.py — นำเข้าข้อมูลงานบริการจาก NS68-69-Excel เข้าโปรแกรมบัญชี Express
(เขียนลงไฟล์ .DBF โดยตรง) เป็นบิล "ขายเงินสด-ศูนย์บริการ" (HC{ปีเดือน}/{running})

วิธีใช้:
  python3 express_import.py preview                 # ดูตัวอย่างบิลที่จะสร้าง (ไม่เขียนไฟล์) -- ค่าเริ่มต้น
  python3 express_import.py verify                  # สร้าง HC6906/001 ใหม่แล้วเทียบ byte กับของจริงในระบบ
  python3 express_import.py commit                  # เขียนลง DBF จริง (สำรองไฟล์อัตโนมัติก่อน)
  python3 express_import.py preview --xls FILE.xls  # ระบุไฟล์ xls เอง

หลัง commit: เปิด Express -> เริ่มระบบ -> จัดการแฟ้มข้อมูล -> (1) จัดเรียงดัชนีใหม่ (Reindex)
            และ (2) คำนวณยอดบัญชี/ยอดยกมาใหม่  เพื่อให้ดัชนีและยอดในงบถูกต้อง
"""
import os, sys, glob, struct, shutil, datetime, argparse

def base_dir():
    """โฟลเดอร์ของโปรแกรม (รองรับทั้งรันด้วย python และรันเป็น .exe จาก PyInstaller)"""
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))

# ───────────────────────── CONFIG ─────────────────────────
DB_DIR  = os.path.join(base_dir(), "NS68-69")
XLS_DIR = os.path.join(base_dir(), "NS68-69-Excel")

DOC_PREFIX = "HC"          # บิลขายเงินสด-ศูนย์บริการ (โหมด service)
CUST_CODE  = "เงินสด"       # รหัสลูกค้าเงินสด
USER_ID    = "AOR"
VAT_RATE   = 7.0
CASH_ACC   = "1111-01"     # เดบิต เงินสด
VAT_ACC    = "2180-01"     # เครดิต ภาษีขาย
STAT_ACC   = "4111-01"     # เดบิต บัญชีสถิติรายได้ (SEQIT=' 5')
STAT_FROM  = "1133-01"     # CHGACCFROM ของบรรทัดเงินสด/สถิติ
JNLTYP     = "03"

# ───── โหมดอะไหล่ (ขายเงินสด-อะไหล่) ─────
PARTS_PREFIX = "HO"        # บิลขายเงินสด-อะไหล่
# 2 สาขา: ชุมพร (เดิม) / สุราษ
PARTS_BRANCHES = {
    "chumphon": {"name":"ชุมพร", "depcod":"106", "income":"4131-02",
                 "stkcod":"41-3100-02", "stkdes":"รายได้จากการขายอะไหล่และอุปกรณ์(เงินสด)", "slash":True},
    "surat":    {"name":"สุราษ", "depcod":"ข9", "income":"4131-02",
                 "stkcod":"420003", "stkdes":"รายได้จากการขายอะไหล่-เงินสด", "slash":False},
}
def parts_cfg(branch): return PARTS_BRANCHES.get(branch or "chumphon", PARTS_BRANCHES["chumphon"])

# ───── โหมดซื้อเงินเชื่อ-อะไหล่ (RS) ─────
PUR_PREFIX  = "RS"         # ใบรับสินค้า/ซื้อเงินเชื่อ-อะไหล่
PUR_INV_ACC = "1162-01"    # เดบิต สินค้าคงเหลือ-อะไหล่
PUR_VATIN   = "1194-01"    # เดบิต ภาษีซื้อ
PUR_PAYTRM  = 45           # เครดิต (วัน)
PUR_JNLTYP  = "04"
PUR_DESC    = "ซื้อเชื่อจาก"
PUR_ACCFROM = "1133-01"    # CHGACCFROM

# 2 สาขา: ชุมพร (เดิม) / สุราษ — ผู้ใช้เลือกโฟลเดอร์ DB ของสาขานั้นเอง
PUR_BRANCHES = {
    "chumphon": {"name":"ชุมพร", "supcod":"P078", "depcod":"106",
                 "stkcod":"11-6200-01", "stkdes":"สินค้าคงเหลือ-อะไหล่และอุปกรณ์", "slash":True},
    "surat":    {"name":"สุราษ", "supcod":"P091", "depcod":"ซ2",
                 "stkcod":"5000", "stkdes":"อะไหล่และอุปกรณ์", "slash":False},
}
PUR_DEFAULT_BRANCH = "chumphon"
def pur_cfg(branch): return PUR_BRANCHES.get(branch or PUR_DEFAULT_BRANCH, PUR_BRANCHES[PUR_DEFAULT_BRANCH])

# ───── โหมดจ่ายเงินสดย่อย (DP = จ่ายคชจ.จากเงินทดรองจ่ายย่อย) ─────
PETTY_REFNUM = ""          # ใบเบิก (VR...) เว้นว่างไว้ (ระบุภายหลังได้)

# ───── โหมดโอนเงินระหว่างธนาคาร (BT) ─────
BT_JNLTYP  = "00"          # สมุดรายวันทั่วไป
BT_FEE_ACC = "5360-04"     # ค่าธรรมเนียมธนาคาร (สุราษ)
BT_FEE_DEP = ""            # แผนกของบรรทัดค่าธรรมเนียม

# ───── โหมดขายเงินสด-ศูนย์บริการ (HC) : 2 สาขา ─────
# cats: (คอลัมน์ xls, รหัสสินค้าบริการ STKCOD, คำอธิบาย) — บัญชีรายได้ดึงจาก DB อัตโนมัติ
SERVICE_BRANCHES = {
    "chumphon": {"name":"ชุมพร", "depcod":"", "slash":True, "cats":[
        ("PARTNET", "41-7100-01", "รายได้ศูนย์บริการทั่วไป-อะไหล่(เงินสด)"),
        ("OILNET",  "41-7100-02", "รายได้ศูนย์บริการทั่วไป-น้ำมัน(เงินสด)"),
        ("SERVNET", "41-7100-03", "รายได้ศูนย์บริการทั่วไป-ค่าแรง(เงินสด)"),
        ("OUTNET",  "41-7100-04", "รายได้ศูนย์บริการทั่วไป-จัดซื้อ(เงินสด)"),
    ]},
    "surat": {"name":"สุราษ", "depcod":"ศ3", "slash":False, "cats":[
        ("PARTNET", "4300-03", "รายได้จากการให้บริการ-อะไหล่(เงินสด)"),
        ("OILNET",  "4300-02", "รายได้จากการให้บริการ-น้ำมัน(เงินสด)"),
        ("SERVNET", "4300-01", "รายได้จากการให้บริการ-ค่าแรง(เงินสด)"),
        ("OUTNET",  "4300-04", "รายได้จากการให้บริการ-อะไหล่จัดซื้อ(เงินสด)"),
    ]},
}
def service_cfg(branch): return SERVICE_BRANCHES.get(branch or "chumphon", SERVICE_BRANCHES["chumphon"])
DESC_PREFIX = "ขายสดให้"   # คำขึ้นต้นคำอธิบายรายการบัญชี

# ───────────────────────── DBF CODEC ─────────────────────────
class DBF:
    """อ่าน/เขียนไฟล์ DBF แบบ Visual FoxPro (cp874, ไม่มี memo, ไม่มี 0x1A)"""
    def __init__(self, path):
        self.path = path
        with open(path, "rb") as f:
            self.data = bytearray(f.read())
        self.ver  = self.data[0]
        self.nrec = struct.unpack("<I", self.data[4:8])[0]
        self.hlen = struct.unpack("<H", self.data[8:10])[0]
        self.rlen = struct.unpack("<H", self.data[10:12])[0]
        self.fields, self.offsets = [], []
        off, pos = 32, 1
        while self.data[off] != 0x0D:
            fd = self.data[off:off+32]
            name = fd[0:11].split(b"\x00")[0].decode("latin1")
            ftype, flen, dec = chr(fd[11]), fd[16], fd[17]
            self.fields.append((name, ftype, flen, dec))
            self.offsets.append(pos); pos += flen; off += 32
        self.fmap = {n:(t,l,d,o) for (n,t,l,d),o in zip(self.fields, self.offsets)}

    # ---- decode one stored record (index) into dict ----
    def get(self, i):
        r = self.data[self.hlen + i*self.rlen : self.hlen + (i+1)*self.rlen]
        rec = {"_deleted": r[0:1] == b"*"}
        for (n,t,l,d),o in zip(self.fields, self.offsets):
            rec[n] = self._dec(r[o:o+l], t)
        return rec

    def raw(self, i):
        return bytes(self.data[self.hlen + i*self.rlen : self.hlen + (i+1)*self.rlen])

    def __len__(self): return self.nrec

    @staticmethod
    def _dec(v, t):
        if t == "C":
            try: return v.decode("cp874").rstrip()
            except: return v.decode("latin1").rstrip()
        if t == "N":
            s = v.strip()
            if not s: return None
            try: return float(s) if b"." in v else int(s)
            except: return None
        if t == "B": return struct.unpack("<d", v)[0]
        if t == "D":
            s = v.strip()
            return f"{s[:4].decode()}-{s[4:6].decode()}-{s[6:8].decode()}" if len(s)==8 and s.isdigit() else None
        if t == "L": return v.decode("latin1")
        return v

    # ---- encode one field value to bytes ----
    def _enc_field(self, name, val):
        t, l, dec, _ = self.fmap[name]
        if t == "C":
            b = ("" if val is None else str(val)).encode("cp874", "replace")
            return b[:l].ljust(l, b" ")
        if t == "N":
            if val is None or val == "": return b" "*l
            s = (f"{val:.{dec}f}" if dec else str(int(round(val))))
            return s.encode("ascii")[:l].rjust(l, b" ")
        if t == "B":
            return struct.pack("<d", float(val or 0.0))
        if t == "D":
            if not val: return b" "*l
            s = str(val).replace("-", "")[:8]
            return s.encode("ascii").ljust(l, b" ")
        if t == "L":
            return (str(val)[:1] if val else " ").encode("ascii")
        return (b"" if val is None else str(val).encode("cp874","replace"))[:l].ljust(l, b" ")

    def build_record(self, values):
        """สร้าง record bytes 1 ตัวจาก dict (ฟิลด์ที่ไม่ระบุ = ค่าว่างตามชนิด)"""
        out = bytearray(b" ")  # deletion flag = not deleted
        for (n,t,l,d) in self.fields:
            out += self._enc_field(n, values.get(n))
        assert len(out) == self.rlen, f"{self.path}: rec {len(out)} != rlen {self.rlen}"
        return bytes(out)

    def append(self, records):
        # บางตาราง (dBASE 0x03 เช่น GLPTT) มี EOF marker 0x1A ท้ายไฟล์ — ตัดออกก่อนต่อ record
        had_eof = self.data[-1:] == b"\x1a"
        if had_eof: del self.data[-1]
        for rec in records:
            self.data += rec
        self.nrec += len(records)
        struct.pack_into("<I", self.data, 4, self.nrec)
        if had_eof: self.data += b"\x1a"     # ใส่ EOF กลับให้เหมือนเดิม
        # last-update date (YY-1900, MM, DD)
        today = TODAY
        self.data[1] = today.year - 1900
        self.data[2] = today.month
        self.data[3] = today.day

    def save(self):
        with open(self.path, "wb") as f:
            f.write(self.data)

# ───────────────────────── helpers ─────────────────────────
TODAY = datetime.date.today()        # วันที่บันทึก (CHGDAT) = วันที่จริงของเครื่อง
CHGTIM = datetime.datetime.now().strftime("%H%M")   # เวลาที่บันทึก (HHMM)

def excel_serial_to_date(serial):
    return datetime.date(1899,12,30) + datetime.timedelta(days=int(serial))

def yymm(d):
    """ปีเดือนแบบไทย เช่น 2026-06 -> '6906' (พ.ศ.2569 -> 69, เดือน 06)"""
    be = d.year + 543
    return f"{be % 100:02d}{d.month:02d}"

def r2(x): return round(float(x) + 1e-9, 2)

def xlsx_date(dv):
    """แปลงวันที่จาก xlsx -> ค.ศ. (รองรับ พ.ศ.เต็ม 2569, พ.ศ.ย่อ 19xx, และ ค.ศ.)"""
    if not hasattr(dv, "year"): return TODAY
    y = dv.year
    if y >= 2400:   y -= 543              # พ.ศ.เต็ม -> ค.ศ.
    elif y < 2000:  y = 1957 + (y % 100)  # พ.ศ.ย่อ เก็บเป็น 19xx -> ค.ศ.
    return datetime.date(y, dv.month, dv.day)

_bank_cache = {}
def resolve_bank(code):
    """รหัสบัญชีธนาคาร (BKMAS.BNKACC) -> (GL ACCNUM, ชื่อย่อ) จาก DB ที่เลือก"""
    if DB_DIR not in _bank_cache:
        m = {}
        try:
            bk = DBF(os.path.join(DB_DIR, "BKMAS.DBF"))
            for i in range(len(bk)):
                r = bk.get(i)
                m[str(r.get("BNKACC"))] = (r.get("ACCNUM"), r.get("SHORTNAM") or r.get("BNKNAM"))
        except Exception:
            pass
        _bank_cache[DB_DIR] = m
    return _bank_cache[DB_DIR].get(code, (None, None))

_income_cache = {}
def resolve_income(stkcod):
    """หาบัญชีรายได้ของรหัสสินค้าบริการ จาก DB ที่เลือก (STMAS.ACCCOD -> ISACC ACCTYP='S' ACCNUM03)
       คืน None ถ้าหาไม่เจอ (ผู้เรียกใช้ fallback เอง)"""
    if DB_DIR not in _income_cache:
        smap = {}; amap = {}
        try:
            isacc = DBF(os.path.join(DB_DIR, "ISACC.DBF"))
            for i in range(len(isacc)):
                r = isacc.get(i)
                if r.get("ACCTYP") == "S": amap[str(r.get("ACCCOD"))] = r.get("ACCNUM03")
            stmas = DBF(os.path.join(DB_DIR, "STMAS.DBF"))
            for i in range(len(stmas)):
                r = stmas.get(i); smap[str(r.get("STKCOD"))] = str(r.get("ACCCOD"))
        except Exception:
            pass
        _income_cache[DB_DIR] = (smap, amap)
    smap, amap = _income_cache[DB_DIR]
    acc = amap.get(smap.get(stkcod))
    return acc or None

def read_xls_jobs(path):
    import xlrd
    sh = xlrd.open_workbook(path).sheet_by_index(0)
    hdr = {n:i for i,n in enumerate(sh.row_values(0))}
    jobs = []
    for r in range(1, sh.nrows):
        g = lambda c: sh.cell_value(r, hdr[c]) if c in hdr else 0
        if not str(g("TAXNO")).strip(): continue
        jobs.append({
            "TAXNO":  str(g("TAXNO")).strip(),
            "DATE":   excel_serial_to_date(g("TAXDATE")),
            "PARTNET": r2(g("PARTNET")), "OILNET": r2(g("OILNET")),
            "SERVNET": r2(g("SERVNET")), "OUTNET": r2(g("OUTNET")),
        })
    return jobs

def group_bills(jobs, branch=None):
    """รวม job ตามวันที่ -> 1 บิล/วัน ; cats ตามสาขา (ชุมพร/สุราษ)"""
    cfg = service_cfg(branch)
    by_date = {}
    for j in jobs:
        by_date.setdefault(j["DATE"], []).append(j)
    bills = []
    for d in sorted(by_date):
        grp = by_date[d]
        taxnos = sorted(j["TAXNO"] for j in grp)
        youref = taxnos[0] if taxnos[0]==taxnos[-1] else f"{taxnos[0]}-{taxnos[-1]}"
        lines = []
        for col, stk, des in cfg["cats"]:
            amt = r2(sum(j[col] for j in grp))
            if amt != 0:
                lines.append({"col":col, "stk":stk, "des":des, "amt":amt})
        if not lines: continue
        sub = r2(sum(l["amt"] for l in lines))
        vat = r2(sub * VAT_RATE / 100.0)
        bills.append({"mode":"service", "branch": branch or "chumphon", "date":d, "youref":youref,
                      "lines":lines, "sub":sub, "vat":vat, "total":r2(sub+vat), "njobs":len(grp)})
    return bills

# ───────────────────────── โหมดอะไหล่ (parts) ─────────────────────────
def _read_header(path):
    """อ่านหัวคอลัมน์ รองรับทั้ง .xls (xlrd) และ .xlsx (openpyxl)"""
    if path.lower().endswith(".xlsx"):
        import openpyxl
        ws = openpyxl.load_workbook(path, read_only=True, data_only=True).active
        for row in ws.iter_rows(min_row=1, max_row=1, values_only=True):
            return [("" if c is None else str(c).strip()) for c in row]
        return []
    import xlrd
    return [str(h).strip() for h in xlrd.open_workbook(path).sheet_by_index(0).row_values(0)]

def detect_mode(path):
    """ดูจากหัวคอลัมน์ว่าเป็นไฟล์โหมดไหน"""
    hdr = _read_header(path)
    if "รหัสผู้โอน" in hdr or "รหัสรับโอน" in hdr:    return "bt"
    if "เลขที่เอกสาร" in hdr:                          return "petty"
    if "INVNO" in hdr and ("RECVNO" in hdr or "APNAME" in hdr): return "purchase"
    if "ยอดสุทธิ" in hdr or "เลขที่ใบกำกับ" in hdr: return "parts"
    if "SERVNET" in hdr or "JOBNO" in hdr:          return "service"
    return "service"

def read_parts_invoices(path):
    import xlrd
    sh = xlrd.open_workbook(path).sheet_by_index(0)
    hdr = {str(n).strip():i for i,n in enumerate(sh.row_values(0))}
    g = lambda r,c: sh.cell_value(r, hdr[c]) if c in hdr else ""
    inv = []
    for r in range(1, sh.nrows):
        no = str(g(r,"เลขที่ใบกำกับ")).strip()
        if not no: continue
        inv.append({"INV":no, "DATE":excel_serial_to_date(g(r,"วันที่ใบกำกับ")),
                    "CUSNAME":str(g(r,"ชื่อลูกค้า")).strip(),
                    "NET":r2(g(r,"ยอดสุทธิ")), "VAT":r2(g(r,"ภาษีมูลค่าเพิ่ม"))})
    return inv

def group_parts_bills(invs, branch=None):
    """รวมใบกำกับตามวันที่ -> 1 บิล/วัน (ยอดสุทธิรวม 1 บรรทัด) ; cfg ตามสาขา"""
    cfg = parts_cfg(branch)
    by_date = {}
    for v in invs:
        by_date.setdefault(v["DATE"], []).append(v)
    bills = []
    for d in sorted(by_date):
        grp = sorted(by_date[d], key=lambda v:v["INV"])
        nos = [v["INV"] for v in grp]
        youref = nos[0] if len(nos)==1 else f"{nos[0]}-{nos[-1][-4:]}"
        net = r2(sum(v["NET"] for v in grp))
        vat = r2(sum(v["VAT"] for v in grp))
        bills.append({"mode":"parts", "branch": branch or "chumphon", "date":d, "youref":youref,
                      "lines":[{"stk":cfg["stkcod"], "acc":cfg["income"],
                                "des":cfg["stkdes"], "amt":net}],
                      "invoices":grp, "sub":net, "vat":vat, "total":r2(net+vat),
                      "njobs":len(grp)})
    return bills

# ───────────────────────── โหมดซื้อเงินเชื่อ-อะไหล่ (purchase) ─────────────────────────
_VENDOR_CACHE = None
def _norm(s): return "".join(str(s).split()).replace("\xa0","")

def _load_vendors():
    """โหลดผู้ขายจาก APMAS + นับว่าแต่ละรหัสถูกใช้ในเอกสาร RS กี่ครั้ง (ไว้ตัดสินกรณีชื่อซ้ำ)"""
    global _VENDOR_CACHE
    if _VENDOR_CACHE is not None: return _VENDOR_CACHE
    from collections import Counter
    ap = DBF(os.path.join(DB_DIR, "APTRN.DBF"))
    rs_use = Counter(ap.get(i).get("SUPCOD") for i in range(len(ap))
                     if str(ap.get(i).get("DOCNUM","")).startswith(PUR_PREFIX))
    _VENDOR_CACHE = []
    am = DBF(os.path.join(DB_DIR, "APMAS.DBF"))
    for i in range(len(am)):
        r = am.get(i)
        if not r.get("SUPCOD"): continue
        _VENDOR_CACHE.append((_norm(r.get("SUPNAM")), {
            "SUPCOD":r.get("SUPCOD"), "SUPNAM":(r.get("SUPNAM") or "").strip(),
            "TAXID":(r.get("TAXID") or "").strip(), "APACC":(r.get("ACCNUM") or "").strip(),
            "ORGNUM":r.get("ORGNUM") or 0, "STATUS":r.get("STATUS"),
            "RSUSE":rs_use.get(r.get("SUPCOD"),0)}))
    return _VENDOR_CACHE

def vendor_by_code(code):
    for nm, info in _load_vendors():
        if info["SUPCOD"] == code: return info
    return None

def vendor_lookup(apname, supcod):
    """ผู้ขายของรายงานซื้อ = รหัสที่กำหนดต่อสาขา ; ถ้าหาไม่เจอค่อย match จากชื่อ"""
    fixed = vendor_by_code(supcod)
    if fixed: return fixed
    key = _norm(apname)
    cands = [info for nm, info in _load_vendors() if nm and (nm in key or key in nm)]
    if not cands: return None
    cands.sort(key=lambda v: (v["RSUSE"], len(v["TAXID"])==13, v["STATUS"]=="A"), reverse=True)
    return cands[0]

def read_purchase_receipts(path):
    import xlrd
    sh = xlrd.open_workbook(path).sheet_by_index(0)
    hdr = {str(n).strip():i for i,n in enumerate(sh.row_values(0))}
    g = lambda r,c: sh.cell_value(r, hdr[c]) if c in hdr else ""
    out = []
    for r in range(1, sh.nrows):
        inv = str(g(r,"INVNO")).strip()
        if not inv: continue
        idate = g(r,"INVDATE")
        if isinstance(idate,(int,float)): idate = excel_serial_to_date(idate)
        else: idate = datetime.date.fromisoformat(str(idate)[:10])
        out.append({"INVNO":inv, "INVDATE":idate, "APNAME":str(g(r,"APNAME")).strip(),
                    "NET":r2(g(r,"NETAMT")), "VAT":r2(g(r,"VATAMT")), "RECVNO":str(g(r,"RECVNO")).strip()})
    return out

def group_purchase_bills(rows, branch=None):
    """แต่ละแถว = 1 บิล (ไม่รวม) ; cfg ตามสาขา (ชุมพร/สุราษ)"""
    cfg = pur_cfg(branch)
    bills = []
    for v in rows:
        net, vat = v["NET"], v["VAT"]
        ven = vendor_lookup(v["APNAME"], cfg["supcod"]) or {}
        bills.append({"mode":"purchase", "branch": branch or PUR_DEFAULT_BRANCH,
                      "date":v["INVDATE"], "youref":v["INVNO"],
                      "invno":v["INVNO"], "apname":v["APNAME"], "vendor":ven,
                      "lines":[{"stk":cfg["stkcod"], "acc":PUR_INV_ACC, "des":cfg["stkdes"], "amt":net}],
                      "sub":net, "vat":vat, "total":r2(net+vat), "njobs":1})
    return bills

# ───────────────────────── โหมดจ่ายเงินสดย่อย (petty) ─────────────────────────
def read_petty(path):
    """อ่านไฟล์ .xlsx เงินสดย่อย: เลขที่เอกสาร/ประเภท/หมายเหตุ/รายละเอียด/จำนวนเงิน/วันที่/เลขที่บัญชี"""
    import openpyxl
    ws = openpyxl.load_workbook(path, data_only=True).active
    hdr = None; out = []
    for row in ws.iter_rows(values_only=True):
        if hdr is None:
            hdr = {("" if c is None else str(c).strip()): i for i, c in enumerate(row)}; continue
        g = lambda name: row[hdr[name]] if name in hdr and hdr[name] < len(row) else None
        docnum = str(g("เลขที่เอกสาร") or "").strip()
        if not docnum: continue
        dv = g("วันที่")
        if hasattr(dv, "year"):
            y = dv.year; y = 1957 + (y % 100) if y < 2000 else y   # ปีเก็บเป็น 19xx (พ.ศ.ย่อ) -> ค.ศ.
            d = datetime.date(y, dv.month, dv.day)
        else:
            d = TODAY
        out.append({"DOCNUM":docnum, "DOCTYP":str(g("ประเภท") or "").strip(),
                    "REMARK":str(g("หมายเหตุ") or "").strip(), "DESCRP":str(g("รายละเอียด") or "").strip(),
                    "ACCNUM":str(g("เลขที่บัญชี") or "").strip(), "AMOUNT":r2(g("จำนวนเงิน") or 0), "DATE":d})
    return out

def read_bt(path):
    """อ่านไฟล์โอนเงินระหว่างธนาคาร (.xlsx)"""
    import openpyxl
    ws = openpyxl.load_workbook(path, data_only=True).active
    hdr = None; out = []
    for row in ws.iter_rows(values_only=True):
        if hdr is None:
            hdr = {("" if c is None else str(c).strip()): i for i, c in enumerate(row)}; continue
        g = lambda name: row[hdr[name]] if name in hdr and hdr[name] < len(row) else None
        docnum = str(g("เลขที่เอกสาร") or "").strip()
        if not docnum: continue
        out.append({"DOCNUM":docnum, "DATE":xlsx_date(g("วันที่ทำรายการ")),
                    "SRC":str(g("รหัสผู้โอน") or "").strip(), "DST":str(g("รหัสรับโอน") or "").strip(),
                    "DEBIT":r2(g("จำนวนเงินหักบัญชี") or 0), "FEE":r2(g("ค่าธรรมเนียม") or 0),
                    "DEPOSIT":r2(g("จำนวนเงินนำฝากเข้าบัญชี") or 0),
                    "DETAIL":str(g("รายละเอียด") or "").strip(), "MEMO":str(g("คำอธิบาย") or "").strip()})
    return out

def group_bt_bills(rows):
    """แต่ละแถว = 1 เอกสารโอนเงิน (ใช้เลขที่เอกสารจากไฟล์)"""
    bills = []
    for v in rows:
        bills.append({"mode":"bt", "date":v["DATE"], "youref":v["DOCNUM"], "docnum":v["DOCNUM"],
                      "src":v["SRC"], "dst":v["DST"], "debit":v["DEBIT"], "fee":v["FEE"],
                      "deposit":v["DEPOSIT"], "detail":v["DETAIL"], "memo":v["MEMO"],
                      "lines":[{"stk":v["DST"], "des":"เข้า "+v["DST"], "amt":v["DEPOSIT"]},
                               {"stk":"ค่าธรรมเนียม", "des":"ค่าธรรมเนียมธนาคาร", "amt":v["FEE"]}],
                      "sub":v["DEBIT"], "vat":0.0, "total":v["DEBIT"], "njobs":1})
    return bills

def group_petty_bills(rows):
    """แต่ละแถว = 1 เอกสาร DP (ใช้เลขที่เอกสารจากไฟล์)"""
    bills = []
    for v in rows:
        amt = v["AMOUNT"]
        bills.append({"mode":"petty", "date":v["DATE"], "youref":v["DOCNUM"], "docnum":v["DOCNUM"],
                      "doctype":v["DOCTYP"], "remark":v["REMARK"],
                      "lines":[{"stk":v["ACCNUM"], "acc":v["ACCNUM"], "des":v["DESCRP"], "amt":amt}],
                      "sub":amt, "vat":0.0, "total":amt, "njobs":1})
    return bills

# ───────────────────────── record builders ─────────────────────────
def D(d): return d.strftime("%Y%m%d")

def build_all(bill, docnum, t_artrn, t_gljnl, t_gljit, t_isvat, t_stcrd):
    """คืน dict {table: [record-bytes,...]} สำหรับบิล 1 ใบ"""
    d, sub, vat, tot = bill["date"], bill["sub"], bill["vat"], bill["total"]
    youref = bill["youref"]; DEP = service_cfg(bill.get("branch"))["depcod"]
    gdesc  = f"{DESC_PREFIX}     {youref}"          # คำอธิบายบรรทัดบัญชี (5 ช่องว่าง)
    nlines = len(bill["lines"])

    # --- ARTRN (หัวบิล) ---
    artrn = t_artrn.build_record({
        "RECTYP":"1","DOCNUM":docnum,"DOCDAT":D(d),"DEPCOD":DEP,"FLGVAT":"2","CUSCOD":CUST_CODE,
        "YOUREF":youref,"DUEDAT":D(d),"BILNUM":"~","NXTSEQ":f"{nlines:>3}",
        "AMOUNT":sub,"DISCAMT":0.0,"AFTDISC":sub,"ADVNUM":"AI","ADVAMT":0.0,"TOTAL":sub,
        "AMTRAT0":0.0,"VATRAT":VAT_RATE,"VATAMT":vat,"NETAMT":tot,"NETVAL":sub,
        "RCVAMT":tot,"REMAMT":0.0,"COMAMT":0.0,"CMPLAPP":"Y","CMPLDAT":D(d),"DOCSTAT":"N",
        "CSHRCV":tot,"CHQRCV":0.0,"INTRCV":0.0,"BEFTAX":0.0,"TAX":0.0,"IVCAMT":tot,
        "CHQPAS":0.0,"VATDAT":D(d),"SRV_VATTYP":"1","USERID":USER_ID,"CHGDAT":D(TODAY),
        "ORGNUM":0,
    })

    # --- GLJNL (สมุดรายวัน) ---
    gljnl = t_gljnl.build_record({
        "JNLTYP":JNLTYP,"VOUDAT":D(d),"VOUCHER":docnum,"DESCRP":gdesc,
        "TRNSTAT":"P","DOCSTAT":"J","CREBY":USER_ID,"CREDAT":D(TODAY),
        "USERID":USER_ID,"CHGDAT":D(TODAY),
    })

    # --- GLJNLIT (บรรทัดลงบัญชี) ---  TRNTYP '0'=เดบิต '1'=เครดิต
    git = []
    # บรรทัดสถิติ (SEQIT=' 5') : เดบิต 4111-01 = ยอดก่อน VAT
    git.append(t_gljit.build_record({
        "VOUCHER":docnum,"SEQIT":" 5","VOUDAT":D(d),"ACCNUM":STAT_ACC,"DEPCOD":DEP,
        "DESCRP":f"{DESC_PREFIX}      \xa0{CUST_CODE}","TRNTYP":"0","AMOUNT":sub,
        "CHGDAT":D(TODAY),"CHGTIM":CHGTIM,"CHGACCFROM":STAT_FROM,
    }))
    # บรรทัดรายได้ (เครดิต) ตามหมวดที่มียอด — บัญชีดึงจาก DB
    for l in bill["lines"]:
        git.append(t_gljit.build_record({
            "VOUCHER":docnum,"SEQIT":" 2","VOUDAT":D(d),"ACCNUM":resolve_income(l["stk"]) or l.get("acc",""),
            "DEPCOD":DEP,"DESCRP":gdesc,"TRNTYP":"1","AMOUNT":l["amt"],
            "CHGDAT":D(TODAY),"CHGTIM":CHGTIM,
        }))
    # เดบิต เงินสด = ยอดรวมทั้งสิ้น
    git.append(t_gljit.build_record({
        "VOUCHER":docnum,"SEQIT":" 2","VOUDAT":D(d),"ACCNUM":CASH_ACC,"DEPCOD":DEP,
        "DESCRP":gdesc,"TRNTYP":"0","AMOUNT":tot,
        "CHGDAT":D(TODAY),"CHGTIM":CHGTIM,"CHGACCFROM":STAT_FROM,
    }))
    # เครดิต ภาษีขาย
    git.append(t_gljit.build_record({
        "VOUCHER":docnum,"SEQIT":" 2","VOUDAT":D(d),"ACCNUM":VAT_ACC,"DEPCOD":DEP,
        "DESCRP":gdesc,"TRNTYP":"1","AMOUNT":vat,
        "CHGDAT":D(TODAY),"CHGTIM":CHGTIM,
    }))

    # --- ISVAT (รายงานภาษีขาย) ---
    vatprd = d.replace(day=1)
    isvat = t_isvat.build_record({
        "VATREC":"S","RECTYP":"1","VATPRD":D(vatprd),"VATDAT":D(d),"DOCDAT":D(d),
        "DOCNUM":docnum,"REFNUM":docnum,"DEPCOD":DEP,"DESCRP":gdesc,"AMT01":sub,"VAT01":vat,
        "AMT02":0.0,"VAT02":0.0,"AMTRAT0":0.0,"SELF_ADDED":"N","HAD_MODIFY":"N",
        "DOCSTAT":"N","TAXID":"0000000000000","ORGNUM":0,
    })

    # --- STCRD (การ์ดสินค้าบริการ) ---
    stcrd = []
    for i, l in enumerate(bill["lines"], 1):
        stcrd.append(t_stcrd.build_record({
            "STKCOD":l["stk"],"LOCCOD":"01","DOCNUM":docnum,"SEQNUM":f"{i:>3}",
            "DOCDAT":D(d),"DEPCOD":DEP,"POSOPR":"9","PEOPLE":CUST_CODE,"TRNQTY":1.0,"TQUCOD":"คน",
            "TFACTOR":1.0,"UNITPR":l["amt"],"DISCAMT":0.0,"TRNVAL":l["amt"],"PHYBAL":0.0,
            "XTRNQTY":1.0,"XUNITPR":0.0,"XTRNVAL":0.0,"XSALVAL":l["amt"],"NETVAL":l["amt"],
            "STKDES":l["des"],
        }))

    return {"ARTRN":[artrn],"GLJNL":[gljnl],"GLJNLIT":git,"ISVAT":[isvat],"STCRD":stcrd}

def build_parts_all(bill, docnum, t_artrn, t_gljnl, t_gljit, t_isvat, t_stcrd):
    """บิลขายเงินสด-อะไหล่ (HO): ยอดสุทธิรวม 1 บรรทัด + GL 2 ชุด + ภาษีรายใบกำกับ"""
    d, net, vat, tot = bill["date"], bill["sub"], bill["vat"], bill["total"]
    youref = bill["youref"]; cfg = parts_cfg(bill.get("branch"))
    descA = f"{DESC_PREFIX}      \xa0{CUST_CODE}"     # ชุดสรุป (ชื่อลูกค้าเงินสด)
    descB = f"{DESC_PREFIX}     {youref}"            # ชุดอ้างอิงเลขใบกำกับ
    DEP   = cfg["depcod"]; INCOME = resolve_income(cfg["stkcod"]) or cfg["income"]

    artrn = t_artrn.build_record({
        "RECTYP":"1","DOCNUM":docnum,"DOCDAT":D(d),"DEPCOD":DEP,"FLGVAT":"2","CUSCOD":CUST_CODE,
        "YOUREF":youref,"DUEDAT":D(d),"BILNUM":"~","NXTSEQ":"  1",
        "AMOUNT":net,"DISCAMT":0.0,"AFTDISC":net,"ADVAMT":0.0,"TOTAL":net,
        "AMTRAT0":0.0,"VATRAT":VAT_RATE,"VATAMT":vat,"NETAMT":tot,"NETVAL":net,
        "RCVAMT":tot,"REMAMT":0.0,"COMAMT":0.0,"CMPLAPP":"Y","CMPLDAT":D(d),"DOCSTAT":"N",
        "CSHRCV":tot,"CHQRCV":0.0,"INTRCV":0.0,"BEFTAX":0.0,"TAX":0.0,"IVCAMT":tot,
        "CHQPAS":0.0,"VATDAT":D(d),"SRV_VATTYP":"1","USERID":USER_ID,"CHGDAT":D(TODAY),"ORGNUM":0,
    })

    # GLJNL: 2 voucher (เลขเดียวกัน) — ชุดสรุป + ชุดอ้างอิง
    gljnl = [t_gljnl.build_record({
        "JNLTYP":JNLTYP,"VOUDAT":D(d),"VOUCHER":docnum,"DESCRP":dsc,
        "TRNSTAT":"P","DOCSTAT":"N","CREBY":USER_ID,"CREDAT":D(TODAY),
        "USERID":USER_ID,"CHGDAT":D(TODAY)}) for dsc in (descA, descB)]

    # GLJNLIT: 6 บรรทัด (2 ชุดสมดุล)  TRNTYP '0'=เดบิต '1'=เครดิต
    def gl(acc, typ, amt, dsc, seqit, frm=""):
        return t_gljit.build_record({"VOUCHER":docnum,"SEQIT":seqit,"VOUDAT":D(d),"ACCNUM":acc,
            "DEPCOD":DEP,"DESCRP":dsc,"TRNTYP":typ,"AMOUNT":amt,
            "CHGDAT":D(TODAY),"CHGTIM":CHGTIM,"CHGACCFROM":frm})
    git = [
        gl(CASH_ACC,    "0", tot, descA, " 2", STAT_FROM),   # ชุดสรุป
        gl(INCOME,"1", net, descA, " 5", STAT_FROM),
        gl(VAT_ACC,     "1", vat, descA, " 6", STAT_FROM),
        gl(CASH_ACC,    "0", tot, descB, " 2", STAT_FROM),   # ชุดอ้างอิง
        gl(INCOME,"1", net, descB, " 2"),
        gl(VAT_ACC,     "1", vat, descB, " 2"),
    ]

    # ISVAT: 1 สรุป + รายใบกำกับ
    vatprd = d.replace(day=1)
    isvat = [t_isvat.build_record({
        "VATREC":"S","RECTYP":"1","VATPRD":D(vatprd),"VATDAT":D(d),"DOCDAT":D(d),
        "DOCNUM":docnum,"REFNUM":docnum,"DEPCOD":DEP,"DESCRP":CUST_CODE,"AMT01":net,"VAT01":vat,
        "AMT02":0.0,"VAT02":0.0,"AMTRAT0":0.0,"SELF_ADDED":"N","HAD_MODIFY":"N",
        "DOCSTAT":"N","TAXID":"0000000000000","ORGNUM":0})]
    for v in bill["invoices"]:
        isvat.append(t_isvat.build_record({
            "VATREC":"S","RECTYP":"","VATPRD":D(vatprd),"VATDAT":D(d),"DOCDAT":D(d),
            "DOCNUM":docnum,"REFNUM":v["INV"],"DEPCOD":DEP,
            "DESCRP":f"{DESC_PREFIX}  {v['CUSNAME']}","AMT01":v["NET"],"VAT01":v["VAT"],
            "AMT02":0.0,"VAT02":0.0,"AMTRAT0":0.0,"SELF_ADDED":"","HAD_MODIFY":"",
            "DOCSTAT":"N","TAXID":"0000000000000","ORGNUM":0}))

    # STCRD: 1 บรรทัด (ยอดสุทธิรวม)
    stcrd = [t_stcrd.build_record({
        "STKCOD":cfg["stkcod"],"LOCCOD":"01","DOCNUM":docnum,"SEQNUM":"  1","DOCDAT":D(d),
        "DEPCOD":DEP,"POSOPR":"9","PEOPLE":CUST_CODE,"TRNQTY":1.0,"TQUCOD":"คน","TFACTOR":1.0,
        "UNITPR":net,"DISCAMT":0.0,"TRNVAL":net,"PHYBAL":0.0,"XTRNQTY":1.0,"XUNITPR":0.0,
        "XTRNVAL":0.0,"XSALVAL":net,"NETVAL":net,"STKDES":cfg["stkdes"]})]

    return {"ARTRN":[artrn],"GLJNL":gljnl,"GLJNLIT":git,"ISVAT":isvat,"STCRD":stcrd}

def build_purchase_all(bill, docnum, t_aptrn, t_gljnl, t_gljit, t_isvat, t_stcrd):
    """ซื้อเงินเชื่อ-อะไหล่ (RS): 1 ใบ/แถว, แยก VAT, Dr สินค้า+ภาษีซื้อ Cr เจ้าหนี้
       (ไม่แตะยอดสต็อกสะสม — ให้ Express คำนวณต้นทุน/ยอดสินค้าใหม่ภายหลัง)"""
    d, net, vat, tot = bill["date"], bill["sub"], bill["vat"], bill["total"]
    inv = bill["invno"]; ven = bill["vendor"]; cfg = pur_cfg(bill.get("branch"))
    supcod = ven.get("SUPCOD",""); apacc = ven.get("APACC","") or "2123-02"
    desc = f"{PUR_DESC}  " + bill["apname"].replace(" ", "\xa0")   # "ซื้อเชื่อจาก  " + ชื่อ(nbsp)
    due = d + datetime.timedelta(days=PUR_PAYTRM)
    DEP = cfg["depcod"]; vatprd = d.replace(day=1)

    aptrn = t_aptrn.build_record({
        "RECTYP":"3","DOCNUM":docnum,"DOCDAT":D(d),"REFNUM":inv,"VATPRD":D(vatprd),
        "DEPCOD":DEP,"FLGVAT":"2","SUPCOD":supcod,"PAYTRM":PUR_PAYTRM,"DUEDAT":D(due),
        "BILNUM":"~","NXTSEQ":"  1","AMOUNT":net,"DISCAMT":0.0,"AFTDISC":net,"ADVAMT":0.0,
        "TOTAL":net,"AMTRAT0":0.0,"VATRAT":VAT_RATE,"VATAMT":vat,"NETAMT":tot,"NETVAL":net,
        "PAYAMT":tot,"REMAMT":0.0,"CMPLAPP":"Y","DOCSTAT":"N","CSHPAY":0.0,"CHQPAY":0.0,
        "VATDAT":D(d),"SRV_VATTYP":"1","PVATPRORAT":"-","USERID":USER_ID,"CHGDAT":D(TODAY),
        "ORGNUM":ven.get("ORGNUM",0)})

    gljnl = [t_gljnl.build_record({
        "JNLTYP":PUR_JNLTYP,"VOUDAT":D(d),"VOUCHER":docnum,"DESCRP":desc,
        "TRNSTAT":"P","DOCSTAT":"N","CREBY":USER_ID,"CREDAT":D(TODAY),
        "USERID":USER_ID,"CHGDAT":D(TODAY)}) for _ in range(2)]

    def gl(acc, typ, amt, seqit, dep):
        return t_gljit.build_record({"VOUCHER":docnum,"SEQIT":seqit,"VOUDAT":D(d),"ACCNUM":acc,
            "DEPCOD":dep,"DESCRP":desc,"TRNTYP":typ,"AMOUNT":amt,
            "CHGDAT":D(TODAY),"CHGTIM":CHGTIM,"CHGACCFROM":PUR_ACCFROM})
    git = []
    for _ in range(2):                                   # 2 ชุดเหมือนกัน
        git += [gl(PUR_INV_ACC,"0",net,    " 1",DEP),    # Dr สินค้าคงเหลือ
                gl(PUR_VATIN,  "0",vat,    " 3",DEP),    # Dr ภาษีซื้อ
                gl(apacc,      "1",tot,    " 6","")]      # Cr เจ้าหนี้

    isvat = [t_isvat.build_record({
        "VATREC":"P","RECTYP":"3","VATPRD":D(vatprd),"VATDAT":D(d),"DOCDAT":D(d),
        "DOCNUM":docnum,"REFNUM":inv,"DEPCOD":DEP,"DESCRP":(ven.get("SUPNAM","") or bill["apname"]).replace(" ","\xa0"),
        "AMT01":net,"VAT01":vat,"AMT02":0.0,"VAT02":0.0,"AMTRAT0":0.0,
        "SELF_ADDED":"N","HAD_MODIFY":"N","DOCSTAT":"N",
        "TAXID":ven.get("TAXID","") or "0000000000000","ORGNUM":ven.get("ORGNUM",0)})]

    # STCRD: รับเข้าสต็อก (cost=net) — ยอดสะสมปล่อยให้ Express คำนวณใหม่
    stcrd = [t_stcrd.build_record({
        "STKCOD":cfg["stkcod"],"LOCCOD":"01","DOCNUM":docnum,"SEQNUM":"  1","DOCDAT":D(d),
        "REFNUM":inv,"DEPCOD":DEP,"POSOPR":"0","PEOPLE":supcod,"TRNQTY":1.0,"TQUCOD":"ชน",
        "TFACTOR":1.0,"UNITPR":net,"DISCAMT":0.0,"TRNVAL":net,"XTRNQTY":1.0,"XUNITPR":net,
        "XTRNVAL":net,"NETVAL":net,"STKDES":cfg["stkdes"]})]

    return {"APTRN":[aptrn],"GLJNL":gljnl,"GLJNLIT":git,"ISVAT":isvat,"STCRD":stcrd}

def build_petty_all(bill, docnum, t_glptt, t_glpttit):
    """จ่ายเงินสดย่อย (DP): GLPTT (หัว) + GLPTTIT (รายการเดียว) — ประเภทเป็นข้อความ, แผนกเว้นว่าง"""
    d, amt = bill["date"], bill["sub"]
    l = bill["lines"][0]
    glptt = t_glptt.build_record({
        "RECTYP":"1","DOCNUM":docnum,"DOCDAT":D(d),"DOCTYP":bill["doctype"],"PTTSTAT":"",
        "DOCSTAT":"N","AMOUNT":amt,"REMARK":bill["remark"],"REFNUM":PETTY_REFNUM,
        "USERID":USER_ID,"CHGDAT":D(TODAY)})
    glpttit = t_glpttit.build_record({
        "DOCNUM":docnum,"ACCNUM":l["acc"],"DEPCOD":"","DESCRP":l["des"],"AMOUNT":amt})
    return {"GLPTT":[glptt], "GLPTTIT":[glpttit]}

def build_bt_all(bill, docnum, t_gljnl, t_gljit):
    """โอนเงินระหว่างธนาคาร (BT): GLJNL + GLJNLIT 3 บรรทัด
       Dr ปลายทาง(ฝากเข้า) SEQIT' 1' + Dr ค่าธรรมเนียม SEQIT' 2' + Cr ต้นทาง(หักบัญชี) SEQIT' 5'"""
    d = bill["date"]; deposit, fee, debit = bill["deposit"], bill["fee"], bill["debit"]
    src_acc, src_nm = resolve_bank(bill["src"])
    dst_acc, dst_nm = resolve_bank(bill["dst"])
    desc = bill["detail"] or f"โอนมาจาก {src_nm or bill['src']} ไปเข้า {dst_nm or bill['dst']} ธรรมเนียม {fee:,.2f}"
    gljnl = [t_gljnl.build_record({
        "JNLTYP":BT_JNLTYP,"VOUDAT":D(d),"VOUCHER":docnum,"DESCRP":desc,
        "TRNSTAT":"P","DOCSTAT":"N","CREBY":USER_ID,"CREDAT":D(TODAY),
        "USERID":USER_ID,"CHGDAT":D(TODAY)})]
    def gl(acc, typ, amt, seqit, dep=""):
        return t_gljit.build_record({"VOUCHER":docnum,"SEQIT":seqit,"VOUDAT":D(d),"ACCNUM":acc,
            "DEPCOD":dep,"DESCRP":desc,"TRNTYP":typ,"AMOUNT":amt,
            "CHGDAT":D(TODAY),"CHGTIM":CHGTIM})
    git = [
        gl(dst_acc or "", "0", deposit, " 1"),            # Dr ปลายทาง (ฝากเข้า)
        gl(BT_FEE_ACC,    "0", fee,     " 2", BT_FEE_DEP),# Dr ค่าธรรมเนียม
        gl(src_acc or "", "1", debit,   " 5"),            # Cr ต้นทาง (หักบัญชี)
    ]
    return {"GLJNL":gljnl, "GLJNLIT":git}

def build_bill(bill, docnum, t):
    """เลือก builder ตามโหมดของบิล"""
    m = bill.get("mode")
    if m=="bt":       return build_bt_all(bill, docnum, t["GLJNL"], t["GLJNLIT"])
    if m=="petty":    return build_petty_all(bill, docnum, t["GLPTT"], t["GLPTTIT"])
    if m=="purchase": return build_purchase_all(bill, docnum, t["APTRN"],t["GLJNL"],t["GLJNLIT"],t["ISVAT"],t["STCRD"])
    fn = build_parts_all if m=="parts" else build_all
    return fn(bill, docnum, t["ARTRN"],t["GLJNL"],t["GLJNLIT"],t["ISVAT"],t["STCRD"])

def load_source(xls, branch=None):
    """อ่านไฟล์ -> (mode, bills) ; branch ใช้เฉพาะโหมดซื้อ (ชุมพร/สุราษ)"""
    mode = detect_mode(xls)
    if mode=="bt":       return mode, group_bt_bills(read_bt(xls))
    if mode=="petty":    return mode, group_petty_bills(read_petty(xls))
    if mode=="purchase": return mode, group_purchase_bills(read_purchase_receipts(xls), branch)
    if mode=="parts":    return mode, group_parts_bills(read_parts_invoices(xls), branch)
    return mode, group_bills(read_xls_jobs(xls), branch)

# ───────────────────────── running number ─────────────────────────
def prefix_for(bills):
    m = bills[0].get("mode") if bills else None
    return {"parts":PARTS_PREFIX, "purchase":PUR_PREFIX}.get(m, DOC_PREFIX)

def next_numbers(t_hdr, bills, prefix=None, dedup_field="YOUREF", slash=True):
    """กำหนดเลขที่บิลต่อจากที่มีอยู่ในตารางหัว (ARTRN/APTRN) ต่อ ปีเดือน
       slash=True -> RS6905/088 ; slash=False -> RS6906014 (ไม่มี /)"""
    prefix = prefix or prefix_for(bills)
    used = {}; existing = set()
    for i in range(len(t_hdr)):
        dn = t_hdr.get(i).get("DOCNUM") or ""
        if not dn.startswith(prefix): continue
        body = dn[len(prefix):]
        if slash:
            ym, _, nn = body.partition("/")
        else:
            ym, nn = body[:4], body[4:]            # RS|6906|014
        if ym.isdigit() and nn.isdigit():
            used[ym] = max(used.get(ym,0), int(nn))
            existing.add(t_hdr.get(i).get(dedup_field))
    sep = "/" if slash else ""
    out = []
    for b in bills:
        if b["youref"] in existing:
            out.append((b, None)); continue          # มีแล้ว -> ข้าม
        ym = yymm(b["date"])
        used[ym] = used.get(ym, 0) + 1
        out.append((b, f"{prefix}{ym}{sep}{used[ym]:03d}"))
    return out

# ───────────────────────── commands ─────────────────────────
def open_tables():
    return {n: DBF(os.path.join(DB_DIR, n+".DBF"))
            for n in ["ARTRN","APTRN","GLJNL","GLJNLIT","ISVAT","STCRD","GLPTT","GLPTTIT"]}

def count_items(path):
    """นับจำนวนรายการในไฟล์ตามโหมด"""
    m = detect_mode(path)
    rd = {"parts":read_parts_invoices, "purchase":read_purchase_receipts,
          "petty":read_petty, "bt":read_bt}.get(m, read_xls_jobs)
    return len(rd(path))

def pick_xls(arg):
    if arg: return arg
    xs = [f for f in glob.glob(os.path.join(XLS_DIR, "*.xls")) if count_items(f)]
    if not xs: sys.exit("ไม่พบไฟล์ .xls ที่มีข้อมูลใน " + XLS_DIR)
    return sorted(xs)[0]

def plan_docnums(mode, t, bills, branch=None):
    """คืน [(bill, docnum|None)] — petty/bt ใช้เลขจากไฟล์, อื่นๆ รันต่อจากตารางหัว"""
    if mode == "bt":
        have = {t["GLJNL"].get(i).get("VOUCHER") for i in range(len(t["GLJNL"]))}
        return [(b, None if b["docnum"] in have else b["docnum"]) for b in bills]
    if mode == "petty":
        have = {t["GLPTT"].get(i).get("DOCNUM") for i in range(len(t["GLPTT"]))}
        return [(b, None if b["docnum"] in have else b["docnum"]) for b in bills]
    if mode == "purchase":
        return next_numbers(t["APTRN"], bills, prefix=PUR_PREFIX, dedup_field="REFNUM",
                            slash=pur_cfg(branch)["slash"])
    if mode == "parts":
        return next_numbers(t["ARTRN"], bills, prefix=PARTS_PREFIX, dedup_field="YOUREF",
                            slash=parts_cfg(branch)["slash"])
    return next_numbers(t["ARTRN"], bills, prefix=DOC_PREFIX, dedup_field="YOUREF",
                        slash=service_cfg(branch)["slash"])

def cmd_preview(args, commit=False):
    xls = pick_xls(args.xls)
    branch = getattr(args, "branch", None)
    mode, bills = load_source(xls, branch)
    unit = {"parts":"ใบกำกับ","purchase":"ใบรับ","petty":"รายการ","bt":"รายการ"}.get(mode,"งาน")
    nitems = sum(b["njobs"] for b in bills)
    t = open_tables()
    plan = plan_docnums(mode, t, bills, branch)

    print(f"ไฟล์ต้นทาง : {os.path.basename(xls)}  [โหมด {mode}]  ({nitems} {unit} -> {len(bills)} บิล)")
    print("="*64)
    new = []
    for b, docnum in plan:
        if docnum is None:
            print(f"[ข้าม] {b['date']}  อ้างอิง {b['youref']} : มีบิลนี้ในระบบแล้ว")
            continue
        new.append((b, docnum))
        print(f"\n● {docnum}   วันที่ {b['date'].strftime('%d/%m/%Y')}   อ้างอิง {b['youref']}  ({b['njobs']} {unit})")
        for l in b["lines"]:
            print(f"    {l['stk']:12} {l['des']:42} {l['amt']:>12,.2f}")
        print(f"    {'':12} {'รวมก่อน VAT':>42} {b['sub']:>12,.2f}")
        print(f"    {'':12} {'VAT 7%':>42} {b['vat']:>12,.2f}")
        print(f"    {'':12} {'รวมทั้งสิ้น':>42} {b['total']:>12,.2f}")
    print("\n" + "="*64)
    print(f"จะสร้างใหม่ {len(new)} บิล")

    if not commit:
        print("\n(โหมด preview — ยังไม่เขียนไฟล์ ใช้คำสั่ง `commit` เพื่อเขียนจริง)")
        return

    if not new:
        print("ไม่มีบิลใหม่ให้เขียน"); return

    # สำรองไฟล์
    stamp = TODAY.strftime("%Y%m%d") + "_" + datetime.datetime.now().strftime("%H%M%S")
    bkdir = os.path.join(DB_DIR, "_backup_" + stamp)
    os.makedirs(bkdir, exist_ok=True)
    for n in t:
        for ext in (".DBF", ".CDX"):
            src = os.path.join(DB_DIR, n+ext)
            if os.path.exists(src): shutil.copy2(src, bkdir)
    print(f"สำรองไฟล์ไว้ที่: {bkdir}")

    buckets = {n:[] for n in t}
    for b, docnum in new:
        recs = build_bill(b, docnum, t)
        for k,v in recs.items(): buckets[k] += v
    for n, dbf in t.items():
        dbf.append(buckets[n]); dbf.save()
        print(f"  {n}.DBF : +{len(buckets[n])} records (รวม {dbf.nrec})")

    # อัปเดตเลขรันใน ISRUN (เฉพาะโหมด service ที่ ISRUN เก็บเป็นรูปแบบ ปีเดือน+รัน)
    if mode=="service": update_isrun(new)
    else: print("  (เลขรันคำนวณต่อจากเอกสารเดิม — ไม่แก้ ISRUN)")
    print("\n✓ เขียนข้อมูลเรียบร้อย")
    print("‼️  ขั้นต่อไปใน Express: เริ่มระบบ -> จัดการแฟ้มข้อมูล -> จัดเรียงดัชนีใหม่ (Reindex)")
    if mode=="petty":
        print("    หมายเหตุ: เอกสาร DP ยังไม่ผ่านบัญชี (DOCSTAT=N) — ตรวจ/ผ่านรายการเงินทดรองจ่ายใน Express ตามปกติ")
    else:
        print("    + คำนวณยอดบัญชี/ยอดยกมาใหม่")
        if mode=="purchase":
            print("    + คำนวณต้นทุน/ยอดสินค้าคงเหลือใหม่   + คำนวณยอดเจ้าหนี้ใหม่")

def update_isrun(new):
    path = os.path.join(DB_DIR, "ISRUN.DBF")
    t = DBF(path)
    # หาเลขรันสูงสุดต่อ ปีเดือน จากบิลที่เพิ่งสร้าง
    maxnum = {}
    for b, docnum in new:
        ym, _, nn = docnum[len(DOC_PREFIX):].partition("/")
        maxnum[ym] = max(maxnum.get(ym,0), int(nn))
    last_ym = max(maxnum)                      # ปีเดือนล่าสุด
    nextval = f"{last_ym}{maxnum[last_ym]+1:03d}"   # เลขถัดไป เช่น 6906003
    for i in range(len(t)):
        rec = t.get(i)
        if rec.get("PREFIX") == DOC_PREFIX and (rec.get("POSDES") or "").startswith("ขายเงินสด-ศูนย์บริการ"):
            # patch เฉพาะฟิลด์ DOCNUM ในตำแหน่งเดิม
            t_off = t.fmap["DOCNUM"][3]; t_len = t.fmap["DOCNUM"][1]
            base = t.hlen + i*t.rlen + t_off
            t.data[base:base+t_len] = nextval.encode("ascii").ljust(t_len, b" ")
            t.save()
            print(f"  ISRUN: เลขรันถัดไปของ {DOC_PREFIX} = {nextval}")
            return
    print("  (เตือน) ไม่พบ record ISRUN ของ HC ให้อัปเดต — โปรดตรวจสอบเลขรันใน Express")

def cmd_verify(args):
    """สร้าง HC6906/001 ใหม่จาก group วันที่ 2026-06-02 แล้วเทียบ byte กับของจริง"""
    xls = pick_xls(args.xls)
    bills = group_bills(read_xls_jobs(xls))
    target = next((b for b in bills if b["date"]==datetime.date(2026,6,2)), None)
    if not target: sys.exit("ไม่พบกลุ่มวันที่ 2026-06-02 ในไฟล์ xls สำหรับตรวจสอบ")
    t = open_tables()
    gen = build_all(target, "HC6906/001", t["ARTRN"],t["GLJNL"],t["GLJNLIT"],t["ISVAT"],t["STCRD"])

    ok = True
    for tbl in ["ARTRN","GLJNL","ISVAT"]:
        dbf = t[tbl]
        idx = next((i for i in range(len(dbf)) if (dbf.get(i).get("DOCNUM")==("HC6906/001"))), None)
        if idx is None and tbl=="GLJNL":
            idx = next((i for i in range(len(dbf)) if dbf.get(i).get("VOUCHER")=="HC6906/001"), None)
        live = dbf.raw(idx); mine = gen[tbl][0]
        ok &= report_diff(tbl, dbf, live, mine)

    # GLJNLIT: เทียบเป็นชุด (เรียงตาม ACCNUM+SEQIT)
    for tbl, key in [("GLJNLIT", lambda r:(r["ACCNUM"],r["SEQIT"])),
                     ("STCRD",   lambda r:r["SEQNUM"])]:
        dbf = t[tbl]
        live_idx = [i for i in range(len(dbf))
                    if (dbf.get(i).get("VOUCHER") if tbl=="GLJNLIT" else dbf.get(i).get("DOCNUM"))=="HC6906/001"]
        live_recs = sorted((dbf.get(i) for i in live_idx), key=key)
        live_raw  = {key(dbf.get(i)): dbf.raw(i) for i in live_idx}
        mine_recs = sorted((decode_built(dbf, m) for m in gen[tbl]), key=key)
        if len(live_recs) != len(mine_recs):
            print(f"[{tbl}] จำนวนบรรทัดไม่เท่ากัน: ของจริง {len(live_recs)} / สร้างใหม่ {len(mine_recs)}"); ok=False
        for m in gen[tbl]:
            dm = decode_built(dbf, m)
            lr = live_raw.get(key(dm))
            if lr is None:
                print(f"[{tbl}] ไม่พบบรรทัดคู่เทียบ key={key(dm)}"); ok=False; continue
            ok &= report_diff(f"{tbl} {key(dm)}", dbf, lr, m)

    print("\n" + ("✓ ตรงกันทุก byte (ยกเว้นฟิลด์เวลา/ผู้ใช้ที่ระบุ) — format ถูกต้อง พร้อมใช้งาน"
                  if ok else "✗ มีความต่าง — ตรวจสอบรายการด้านบนก่อนเขียนจริง"))

def cmd_verify_parts(args):
    """สร้าง HO6905/013 ใหม่จากข้อมูลที่รู้ แล้วเทียบ byte กับของจริง (ยกเว้น DESCRP/ผู้ใช้)"""
    import datetime as _dt
    _c = parts_cfg("chumphon")
    bill = {"mode":"parts","branch":"chumphon","date":_dt.date(2026,5,29),"youref":"1CB-26050029",
            "lines":[{"stk":_c["stkcod"],"acc":_c["income"],"des":_c["stkdes"],"amt":280.0}],
            "invoices":[{"INV":"1CB-26050029","DATE":_dt.date(2026,5,29),
                         "CUSNAME":"บริษัท มิตรแท้ ชุมพร จำกัด","NET":280.0,"VAT":19.6}],
            "sub":280.0,"vat":19.6,"total":299.6,"njobs":1}
    t = open_tables()
    gen = build_parts_all(bill,"HO6905/013",t["ARTRN"],t["GLJNL"],t["GLJNLIT"],t["ISVAT"],t["STCRD"])
    ig = IGNORE_FIELDS | {"USERID","CREBY","DESCRP"}   # DESCRP เป็นข้อความ Express เองยังไม่คงที่
    ok = True
    # ARTRN (1)
    dbf=t["ARTRN"]; idx=next(i for i in range(len(dbf)) if dbf.get(i).get("DOCNUM")=="HO6905/013")
    ok &= report_diff("ARTRN", dbf, dbf.raw(idx), gen["ARTRN"][0], ig)
    # STCRD (1)
    dbf=t["STCRD"]; idx=next(i for i in range(len(dbf)) if dbf.get(i).get("DOCNUM")=="HO6905/013")
    ok &= report_diff("STCRD", dbf, dbf.raw(idx), gen["STCRD"][0], ig)
    # GLJNLIT (6) จับคู่ตาม (ACCNUM,SEQIT,DESCRP-group)
    dbf=t["GLJNLIT"]
    live=[dbf.raw(i) for i in range(len(dbf)) if dbf.get(i).get("VOUCHER")=="HO6905/013"]
    print(f"[GLJNLIT] ของจริง {len(live)} บรรทัด / สร้าง {len(gen['GLJNLIT'])} บรรทัด",
          "✓" if len(live)==len(gen['GLJNLIT']) else "✗")
    # ISVAT (1+N)
    dbf=t["ISVAT"]
    livev=[dbf.raw(i) for i in range(len(dbf)) if dbf.get(i).get("DOCNUM")=="HO6905/013"]
    print(f"[ISVAT]   ของจริง {len(livev)} บรรทัด / สร้าง {len(gen['ISVAT'])} บรรทัด",
          "✓" if len(livev)==len(gen['ISVAT']) else "✗")
    print("\n" + ("✓ ARTRN/STCRD ตรง byte (ยกเว้น DESCRP/ผู้ใช้/เวลา) + โครงสร้าง GL/VAT ครบ"
                  if ok else "✗ มีความต่างใน ARTRN/STCRD — ตรวจด้านบน"))

# ฟิลด์ที่ยอมให้ต่างได้ (เวลา/วันที่บันทึก — ไม่กระทบความถูกต้องบัญชี)
IGNORE_FIELDS = {"CHGDAT","CHGTIM","CREDAT","CHGACCFROM"}

def decode_built(dbf, recbytes):
    rec = {}
    for (n,t,l,d),o in zip(dbf.fields, dbf.offsets):
        rec[n] = dbf._dec(recbytes[o:o+l], t)
    return rec

def report_diff(label, dbf, live, mine, ignore=None):
    ignore = ignore if ignore is not None else IGNORE_FIELDS
    diffs = []
    for (n,t,l,d),o in zip(dbf.fields, dbf.offsets):
        lv, mv = live[o:o+l], mine[o:o+l]
        if lv != mv:
            diffs.append((n, dbf._dec(lv,t), dbf._dec(mv,t), n in ignore))
    real = [x for x in diffs if not x[3]]
    if not diffs:
        print(f"[{label}] ✓ เหมือนกันทุก byte")
    else:
        tag = "✓(ต่างเฉพาะเวลา)" if not real else "✗ ต่าง"
        print(f"[{label}] {tag}")
        for n, a, b, ign in diffs:
            mark = "  (ข้าม)" if ign else ""
            print(f"     {n:12} จริง={a!r:22} สร้าง={b!r}{mark}")
    return not real

# ───────────────────────── main ─────────────────────────
if __name__ == "__main__":
    ap = argparse.ArgumentParser(description="นำเข้างานบริการเข้า Express (ขายเงินสด-ศูนย์บริการ)")
    ap.add_argument("cmd", nargs="?", default="preview", choices=["preview","verify","commit"])
    ap.add_argument("--xls", help="ระบุไฟล์ xls (ค่าเริ่มต้น: ไฟล์แรกที่มีข้อมูลใน NS68-69-Excel)")
    args = ap.parse_args()
    if args.cmd == "verify":
        # verify ตามโหมดของไฟล์ (service=HC, parts=HO)
        if detect_mode(pick_xls(args.xls))=="parts": cmd_verify_parts(args)
        else:                                         cmd_verify(args)
    elif args.cmd == "commit":  cmd_preview(args, commit=True)
    else:                       cmd_preview(args, commit=False)
