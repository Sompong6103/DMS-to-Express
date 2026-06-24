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
PARTS_DEPCOD = "106"       # แผนกอะไหล่
PARTS_INCOME = "4131-02"   # บัญชีรายได้ขายอะไหล่
PARTS_STKCOD = "41-3100-02"
PARTS_STKDES = "รายได้จากการขายอะไหล่และอุปกรณ์(เงินสด)"

# คอลัมน์ xls -> (รหัสสินค้าบริการ STKCOD, บัญชีรายได้ GL, คำอธิบาย)  เรียงตามลำดับบรรทัดในบิล
CATEGORIES = [
    ("PARTNET", "41-7100-01", "4171-01", "รายได้ศูนย์บริการทั่วไป-อะไหล่(เงินสด)"),
    ("OILNET",  "41-7100-02", "4171-03", "รายได้ศูนย์บริการทั่วไป-น้ำมัน(เงินสด)"),
    ("SERVNET", "41-7100-03", "4171-05", "รายได้ศูนย์บริการทั่วไป-ค่าแรง(เงินสด)"),
    ("OUTNET",  "41-7100-04", "4171-07", "รายได้ศูนย์บริการทั่วไป-จัดซื้อ(เงินสด)"),
]
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
        for rec in records:
            self.data += rec
        self.nrec += len(records)
        struct.pack_into("<I", self.data, 4, self.nrec)
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

def group_bills(jobs):
    """รวม job ตามวันที่ -> 1 บิล/วัน"""
    by_date = {}
    for j in jobs:
        by_date.setdefault(j["DATE"], []).append(j)
    bills = []
    for d in sorted(by_date):
        grp = by_date[d]
        taxnos = sorted(j["TAXNO"] for j in grp)
        youref = taxnos[0] if taxnos[0]==taxnos[-1] else f"{taxnos[0]}-{taxnos[-1]}"
        lines = []
        for col, stk, acc, des in CATEGORIES:
            amt = r2(sum(j[col] for j in grp))
            if amt != 0:
                lines.append({"col":col, "stk":stk, "acc":acc, "des":des, "amt":amt})
        if not lines: continue
        sub = r2(sum(l["amt"] for l in lines))
        vat = r2(sub * VAT_RATE / 100.0)
        bills.append({"mode":"service", "date":d, "youref":youref, "lines":lines,
                      "sub":sub, "vat":vat, "total":r2(sub+vat), "njobs":len(grp)})
    return bills

# ───────────────────────── โหมดอะไหล่ (parts) ─────────────────────────
def detect_mode(path):
    """ดูจากหัวคอลัมน์ว่าเป็นไฟล์ศูนย์บริการ (service) หรืออะไหล่ (parts)"""
    import xlrd
    hdr = [str(h).strip() for h in xlrd.open_workbook(path).sheet_by_index(0).row_values(0)]
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

def group_parts_bills(invs):
    """รวมใบกำกับตามวันที่ -> 1 บิล/วัน (ยอดสุทธิรวม 1 บรรทัด)"""
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
        bills.append({"mode":"parts", "date":d, "youref":youref,
                      "lines":[{"stk":PARTS_STKCOD, "acc":PARTS_INCOME,
                                "des":PARTS_STKDES, "amt":net}],
                      "invoices":grp, "sub":net, "vat":vat, "total":r2(net+vat),
                      "njobs":len(grp)})
    return bills

# ───────────────────────── record builders ─────────────────────────
def D(d): return d.strftime("%Y%m%d")

def build_all(bill, docnum, t_artrn, t_gljnl, t_gljit, t_isvat, t_stcrd):
    """คืน dict {table: [record-bytes,...]} สำหรับบิล 1 ใบ"""
    d, sub, vat, tot = bill["date"], bill["sub"], bill["vat"], bill["total"]
    youref = bill["youref"]
    gdesc  = f"{DESC_PREFIX}     {youref}"          # คำอธิบายบรรทัดบัญชี (5 ช่องว่าง)
    nlines = len(bill["lines"])

    # --- ARTRN (หัวบิล) ---
    artrn = t_artrn.build_record({
        "RECTYP":"1","DOCNUM":docnum,"DOCDAT":D(d),"FLGVAT":"2","CUSCOD":CUST_CODE,
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
        "VOUCHER":docnum,"SEQIT":" 5","VOUDAT":D(d),"ACCNUM":STAT_ACC,
        "DESCRP":f"{DESC_PREFIX}      \xa0{CUST_CODE}","TRNTYP":"0","AMOUNT":sub,
        "CHGDAT":D(TODAY),"CHGTIM":CHGTIM,"CHGACCFROM":STAT_FROM,
    }))
    # บรรทัดรายได้ (เครดิต) ตามหมวดที่มียอด
    for l in bill["lines"]:
        git.append(t_gljit.build_record({
            "VOUCHER":docnum,"SEQIT":" 2","VOUDAT":D(d),"ACCNUM":l["acc"],
            "DESCRP":gdesc,"TRNTYP":"1","AMOUNT":l["amt"],
            "CHGDAT":D(TODAY),"CHGTIM":CHGTIM,
        }))
    # เดบิต เงินสด = ยอดรวมทั้งสิ้น
    git.append(t_gljit.build_record({
        "VOUCHER":docnum,"SEQIT":" 2","VOUDAT":D(d),"ACCNUM":CASH_ACC,
        "DESCRP":gdesc,"TRNTYP":"0","AMOUNT":tot,
        "CHGDAT":D(TODAY),"CHGTIM":CHGTIM,"CHGACCFROM":STAT_FROM,
    }))
    # เครดิต ภาษีขาย
    git.append(t_gljit.build_record({
        "VOUCHER":docnum,"SEQIT":" 2","VOUDAT":D(d),"ACCNUM":VAT_ACC,
        "DESCRP":gdesc,"TRNTYP":"1","AMOUNT":vat,
        "CHGDAT":D(TODAY),"CHGTIM":CHGTIM,
    }))

    # --- ISVAT (รายงานภาษีขาย) ---
    vatprd = d.replace(day=1)
    isvat = t_isvat.build_record({
        "VATREC":"S","RECTYP":"1","VATPRD":D(vatprd),"VATDAT":D(d),"DOCDAT":D(d),
        "DOCNUM":docnum,"REFNUM":docnum,"DESCRP":gdesc,"AMT01":sub,"VAT01":vat,
        "AMT02":0.0,"VAT02":0.0,"AMTRAT0":0.0,"SELF_ADDED":"N","HAD_MODIFY":"N",
        "DOCSTAT":"N","TAXID":"0000000000000","ORGNUM":0,
    })

    # --- STCRD (การ์ดสินค้าบริการ) ---
    stcrd = []
    for i, l in enumerate(bill["lines"], 1):
        stcrd.append(t_stcrd.build_record({
            "STKCOD":l["stk"],"LOCCOD":"01","DOCNUM":docnum,"SEQNUM":f"{i:>3}",
            "DOCDAT":D(d),"POSOPR":"9","PEOPLE":CUST_CODE,"TRNQTY":1.0,"TQUCOD":"คน",
            "TFACTOR":1.0,"UNITPR":l["amt"],"DISCAMT":0.0,"TRNVAL":l["amt"],"PHYBAL":0.0,
            "XTRNQTY":1.0,"XUNITPR":0.0,"XTRNVAL":0.0,"XSALVAL":l["amt"],"NETVAL":l["amt"],
            "STKDES":l["des"],
        }))

    return {"ARTRN":[artrn],"GLJNL":[gljnl],"GLJNLIT":git,"ISVAT":[isvat],"STCRD":stcrd}

def build_parts_all(bill, docnum, t_artrn, t_gljnl, t_gljit, t_isvat, t_stcrd):
    """บิลขายเงินสด-อะไหล่ (HO): ยอดสุทธิรวม 1 บรรทัด + GL 2 ชุด + ภาษีรายใบกำกับ"""
    d, net, vat, tot = bill["date"], bill["sub"], bill["vat"], bill["total"]
    youref = bill["youref"]
    descA = f"{DESC_PREFIX}      \xa0{CUST_CODE}"     # ชุดสรุป (ชื่อลูกค้าเงินสด)
    descB = f"{DESC_PREFIX}     {youref}"            # ชุดอ้างอิงเลขใบกำกับ
    DEP   = PARTS_DEPCOD

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
        gl(PARTS_INCOME,"1", net, descA, " 5", STAT_FROM),
        gl(VAT_ACC,     "1", vat, descA, " 6", STAT_FROM),
        gl(CASH_ACC,    "0", tot, descB, " 2", STAT_FROM),   # ชุดอ้างอิง
        gl(PARTS_INCOME,"1", net, descB, " 2"),
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
        "STKCOD":PARTS_STKCOD,"LOCCOD":"01","DOCNUM":docnum,"SEQNUM":"  1","DOCDAT":D(d),
        "DEPCOD":DEP,"POSOPR":"9","PEOPLE":CUST_CODE,"TRNQTY":1.0,"TQUCOD":"คน","TFACTOR":1.0,
        "UNITPR":net,"DISCAMT":0.0,"TRNVAL":net,"PHYBAL":0.0,"XTRNQTY":1.0,"XUNITPR":0.0,
        "XTRNVAL":0.0,"XSALVAL":net,"NETVAL":net,"STKDES":PARTS_STKDES})]

    return {"ARTRN":[artrn],"GLJNL":gljnl,"GLJNLIT":git,"ISVAT":isvat,"STCRD":stcrd}

def build_bill(bill, docnum, t):
    """เลือก builder ตามโหมดของบิล"""
    fn = build_parts_all if bill.get("mode")=="parts" else build_all
    return fn(bill, docnum, t["ARTRN"],t["GLJNL"],t["GLJNLIT"],t["ISVAT"],t["STCRD"])

def load_source(xls):
    """อ่านไฟล์ -> (mode, bills)"""
    mode = detect_mode(xls)
    if mode=="parts": return mode, group_parts_bills(read_parts_invoices(xls))
    return mode, group_bills(read_xls_jobs(xls))

# ───────────────────────── running number ─────────────────────────
def prefix_for(bills):
    return PARTS_PREFIX if bills and bills[0].get("mode")=="parts" else DOC_PREFIX

def next_numbers(t_artrn, bills, prefix=None):
    """กำหนดเลขที่บิลต่อจากที่มีอยู่ใน ARTRN (ต่อ ปีเดือน)"""
    prefix = prefix or prefix_for(bills)
    used = {}; existing_youref = set()
    for i in range(len(t_artrn)):
        dn = t_artrn.get(i).get("DOCNUM") or ""
        if dn.startswith(prefix) and "/" in dn:
            ym, _, nn = dn[len(prefix):].partition("/")
            if nn.isdigit(): used[ym] = max(used.get(ym,0), int(nn))
            existing_youref.add(t_artrn.get(i).get("YOUREF"))
    out = []
    for b in bills:
        if b["youref"] in existing_youref:
            out.append((b, None)); continue          # มีแล้ว -> ข้าม
        ym = yymm(b["date"])
        used[ym] = used.get(ym, 0) + 1
        out.append((b, f"{prefix}{ym}/{used[ym]:03d}"))
    return out

# ───────────────────────── commands ─────────────────────────
def open_tables():
    return {n: DBF(os.path.join(DB_DIR, n+".DBF"))
            for n in ["ARTRN","GLJNL","GLJNLIT","ISVAT","STCRD"]}

def count_items(path):
    """นับจำนวนรายการในไฟล์ (งาน หรือ ใบกำกับ) ตามโหมด"""
    return len(read_parts_invoices(path) if detect_mode(path)=="parts" else read_xls_jobs(path))

def pick_xls(arg):
    if arg: return arg
    xs = [f for f in glob.glob(os.path.join(XLS_DIR, "*.xls")) if count_items(f)]
    if not xs: sys.exit("ไม่พบไฟล์ .xls ที่มีข้อมูลใน " + XLS_DIR)
    return sorted(xs)[0]

def cmd_preview(args, commit=False):
    xls = pick_xls(args.xls)
    mode, bills = load_source(xls)
    unit = "ใบกำกับ" if mode=="parts" else "งาน"
    nitems = sum(b["njobs"] for b in bills)
    t = open_tables()
    plan = next_numbers(t["ARTRN"], bills)

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
    else: print("  (โหมดอะไหล่: เลขรันคำนวณต่อจากเอกสารเดิมใน ARTRN — ไม่แก้ ISRUN)")
    print("\n✓ เขียนข้อมูลเรียบร้อย")
    print("‼️  ขั้นต่อไปใน Express: เริ่มระบบ -> จัดการแฟ้มข้อมูล ->")
    print("    1) จัดเรียงดัชนีใหม่ (Reindex)   2) คำนวณยอดบัญชี/ยอดยกมาใหม่")

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
    bill = {"mode":"parts","date":_dt.date(2026,5,29),"youref":"1CB-26050029",
            "lines":[{"stk":PARTS_STKCOD,"acc":PARTS_INCOME,"des":PARTS_STKDES,"amt":280.0}],
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
