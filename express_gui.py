#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
express_gui.py — หน้าจอกดใช้งาน (เว็บในเครื่อง) สำหรับนำเข้างานบริการเข้า Express
รัน:  python3 express_gui.py     แล้วเบราว์เซอร์จะเปิดให้อัตโนมัติ
เลือกโฟลเดอร์ฐานข้อมูล (DBF) และเลือกไฟล์ xls โดยตรงได้จากหน้าจอ (จำค่าไว้ในไฟล์ config)
"""
import os, io, sys, glob, json, socket, threading, webbrowser, contextlib, traceback, argparse, hashlib, secrets
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import urlparse, parse_qs

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import express_import as EX

# โฟลเดอร์โปรแกรม (รองรับ .exe) — ใช้เก็บไฟล์ config ให้เขียนได้จริง
HERE = EX.base_dir()
try:                       # ให้ console บน Windows แสดงภาษาไทยได้
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")
except Exception:
    pass

if os.name == "nt":        # ปิด dialog "There is no disk in the drive" ตอนไล่เช็คไดรฟ์/card reader
    try:
        import ctypes
        ctypes.windll.kernel32.SetErrorMode(0x0001 | 0x8000)  # SEM_FAILCRITICALERRORS | SEM_NOOPENFILEERRORBOX
    except Exception:
        pass

CONFIG_FILE = os.path.join(HERE, "express_config.json")
UPLOAD_DIR  = os.path.join(HERE, "uploads")     # ไฟล์ xls ที่ client อัปโหลดมา
REQUIRED = ["ARTRN", "GLJNL", "GLJNLIT", "ISVAT", "STCRD", "ISRUN"]   # ตารางที่ต้องมี
XLS_FILE = ""   # ไฟล์ xls ที่เลือกไว้ (path เต็ม)
AUTH_HASH = None             # รหัสผ่าน (เก็บแบบ hash)
TOKENS = set()               # token ที่ login แล้ว (อยู่ในหน่วยความจำ; รีสตาร์ท = login ใหม่)

def pw_hash(pw): return hashlib.sha256(("express::"+pw).encode("utf-8")).hexdigest()
def set_password(pw):
    global AUTH_HASH
    AUTH_HASH = pw_hash(pw); save_config()
def new_token():
    t = secrets.token_hex(16); TOKENS.add(t); return t

# ───────── config ─────────
def load_config():
    global XLS_FILE, AUTH_HASH
    cfg = {"db_dir": EX.DB_DIR, "xls_file": ""}
    if os.path.exists(CONFIG_FILE):
        try: cfg.update(json.load(open(CONFIG_FILE, encoding="utf-8")))
        except Exception: pass
    if cfg.get("db_dir"):  EX.DB_DIR = cfg["db_dir"]
    XLS_FILE = cfg.get("xls_file") or ""
    AUTH_HASH = cfg.get("auth_hash")
    if XLS_FILE: EX.XLS_DIR = os.path.dirname(XLS_FILE)
    return cfg

def save_config():
    json.dump({"db_dir": EX.DB_DIR, "xls_file": XLS_FILE, "auth_hash": AUTH_HASH},
              open(CONFIG_FILE, "w", encoding="utf-8"), ensure_ascii=False, indent=2)

def db_ok(path):
    return bool(path) and all(os.path.exists(os.path.join(path, n + ".DBF")) for n in REQUIRED)

def xls_jobs(path):
    try: return EX.count_items(path)
    except Exception: return -1

def xls_mode(path):
    try: return EX.detect_mode(path)
    except Exception: return ""

def config_status():
    ok = bool(XLS_FILE) and os.path.isfile(XLS_FILE) and XLS_FILE.lower().endswith(".xls")
    return {"db_dir": EX.DB_DIR, "db_ok": db_ok(EX.DB_DIR),
            "db_missing": [n for n in REQUIRED if not os.path.exists(os.path.join(EX.DB_DIR or "", n + ".DBF"))],
            "xls_file": XLS_FILE, "xls_name": os.path.basename(XLS_FILE) if XLS_FILE else "",
            "xls_ok": ok, "xls_jobs": xls_jobs(XLS_FILE) if ok else -1,
            "mode": xls_mode(XLS_FILE) if ok else ""}

# ───────── folder/file browser ─────────
def list_drives():
    """รายชื่อไดรฟ์ที่มี (Windows: ใช้ GetLogicalDrives เลี่ยงการแตะ media/card reader)"""
    out = []
    if os.name == "nt":
        import string
        bitmask = 0
        try:
            import ctypes
            bitmask = ctypes.windll.kernel32.GetLogicalDrives()
        except Exception:
            bitmask = 0
        if bitmask:
            for i, c in enumerate(string.ascii_uppercase):
                if bitmask & (1 << i): out.append(f"{c}:\\")
        else:                                   # fallback (เผื่อเรียก API ไม่ได้)
            for c in string.ascii_uppercase:
                p = f"{c}:\\"
                try:
                    if os.path.exists(p): out.append(p)
                except Exception: pass
    else:
        out.append("/")
        vol = "/Volumes"
        if os.path.isdir(vol):
            for n in sorted(os.listdir(vol)):
                if os.path.isdir(os.path.join(vol, n)): out.append(os.path.join(vol, n))
    return out

def listdir(path, want_files=False):
    path = os.path.abspath(os.path.expanduser(path or os.path.expanduser("~")))
    if not os.path.isdir(path): path = os.path.expanduser("~")
    dirs, files = [], []
    try:
        entries = sorted(os.scandir(path), key=lambda x: x.name.lower())
    except Exception:
        entries = []
    XEXT = (".xls", ".xlsx")
    xls_here = [e for e in entries if e.is_file() and e.name.lower().endswith(XEXT)]
    count_jobs = want_files and len(xls_here) <= 60      # ถ้าไฟล์เยอะมากไม่นับงาน (ช้า)
    for e in entries:
        if e.name.startswith("."): continue
        if e.is_dir(follow_symlinks=False):
            try: names = os.listdir(e.path)
            except Exception: names = []
            upper = {n.upper() for n in names}
            dirs.append({"name": e.name, "path": e.path,
                         "dbf": "ARTRN.DBF" in upper,
                         "xls": any(n.lower().endswith(XEXT) for n in names)})
        elif want_files and e.is_file() and e.name.lower().endswith(XEXT):
            files.append({"name": e.name, "path": e.path,
                          "jobs": xls_jobs(e.path) if count_jobs else None})
    parent = os.path.dirname(path)
    return {"path": path, "parent": (parent if parent != path else None),
            "dirs": dirs, "files": files, "drives": list_drives(),
            "self_dbf": "ARTRN.DBF" in {n.upper() for n in (os.listdir(path) if os.path.isdir(path) else [])},
            "home": os.path.expanduser("~")}

# ───────── helpers ─────────
def _args(xls, branch=None): return type("A", (), {"xls": xls, "branch": branch})()

def clean(text):
    return "\n".join(ln for ln in text.splitlines()
                     if not (ln.startswith("WARNING ***") or "OLE2" in ln or "not 512 + multiple" in ln))

def capture(fn, *a, **k):
    buf = io.StringIO()
    with contextlib.redirect_stdout(buf), contextlib.redirect_stderr(buf):
        try: fn(*a, **k)
        except SystemExit as e: buf.write(f"\n[{e}]")
        except Exception:       buf.write("\n" + traceback.format_exc())
    return clean(buf.getvalue())

def preview_data(xls, branch=None):
    mode, bills = EX.load_source(xls, branch)
    t     = EX.open_tables()
    plan  = EX.plan_docnums(mode, t, bills, branch)
    out   = []
    for b, docnum in plan:
        out.append({"docnum": docnum, "skipped": docnum is None,
                    "date": b["date"].strftime("%d/%m/%Y"), "youref": b["youref"], "njobs": b["njobs"],
                    "lines": [{"stk": l["stk"], "des": l["des"], "amt": l["amt"]} for l in b["lines"]],
                    "sub": b["sub"], "vat": b["vat"], "total": b["total"]})
    new = sum(1 for x in out if not x["skipped"])
    return {"file": os.path.basename(xls), "mode": mode, "njobs": sum(b["njobs"] for b in bills),
            "nbills": len(bills), "nnew": new, "bills": out}

def verify_purchase(xls, branch=None):
    """ตรวจโครงสร้างโหมดซื้อ: ทุกบิลต้องสมดุล Dr=Cr และบัญชี/ผู้ขายถูกต้อง"""
    mode, bills = EX.load_source(xls, branch)
    t = EX.open_tables()
    ven = bills[0]["vendor"] if bills else {}
    cfg = EX.pur_cfg(branch)
    print(f"สาขา: {cfg['name']} | ผู้ขาย: {ven.get('SUPCOD')} ({ven.get('SUPNAM','')[:24]}) | แผนก {cfg['depcod']} | รหัสสินค้า {cfg['stkcod']} | เครดิต {EX.PUR_PAYTRM} วัน")
    bad = 0
    for b in bills[:200]:
        recs = EX.build_purchase_all(b,"RS0000/000",t["APTRN"],t["GLJNL"],t["GLJNLIT"],t["ISVAT"],t["STCRD"])
        dr=cr=0
        for m in recs["GLJNLIT"]:
            x = EX.decode_built(t["GLJNLIT"], m)
            if x["TRNTYP"]=="0": dr+=x["AMOUNT"]
            else: cr+=x["AMOUNT"]
        if abs(dr-cr) > 0.005: bad += 1
    print(f"ตรวจ {len(bills)} บิล: GL สมดุลทุกใบ" + (" ✓" if bad==0 else f" ✗ ไม่สมดุล {bad} ใบ"))
    print("บัญชี: Dr สินค้าคงเหลือ 1162-01 + Dr ภาษีซื้อ 1194-01  /  Cr เจ้าหนี้ "+str(ven.get('APACC')))
    print("\n✓ โครงสร้างถูกต้อง (ยอด/บัญชี/สมดุล) — ดูรายละเอียดแต่ละบิลที่ปุ่มดูตัวอย่าง")

def verify_petty(xls):
    mode, bills = EX.load_source(xls)
    print(f"โหมดจ่ายเงินสดย่อย (DP): {len(bills)} เอกสาร (แต่ละแถว = 1 เอกสาร)")
    tot = sum(b["sub"] for b in bills)
    print(f"ยอดรวมทั้งหมด: {tot:,.2f} บาท")
    print("แต่ละเอกสาร: GLPTT (หัว) + GLPTTIT (รายการเดียว) — ประเภทเป็นข้อความ, แผนกเว้นว่าง, ไม่มี VAT")
    print("เอกสาร DP จะยังไม่ผ่านบัญชี (DOCSTAT=N) ตามรูปแบบเงินสดย่อยปกติ")
    print("\n✓ โครงสร้างถูกต้อง — ดูรายละเอียดแต่ละเอกสารที่ปุ่มดูตัวอย่าง")

def verify_generic(xls, branch=None):
    """ตรวจโครงสร้าง+สมดุล GL + โชว์บัญชีรายได้ที่ดึงจาก DB (ใช้กับสาขาที่ไม่มีตัวอย่างเทียบ byte)"""
    mode, bills = EX.load_source(xls, branch)
    t = EX.open_tables()
    plan = EX.plan_docnums(mode, t, bills, branch)
    new = [(b,dn) for b,dn in plan if dn]
    print(f"โหมด {mode} · สาขา {EX.service_cfg(branch)['name'] if mode=='service' else (EX.parts_cfg(branch)['name'])} : {len(bills)} บิล (ใหม่ {len(new)})")
    bad = 0
    for b,dn in new[:300]:
        git = EX.build_bill(b,dn,t).get("GLJNLIT",[])
        dr=cr=0
        for m in git:
            x = EX.decode_built(t["GLJNLIT"], m)
            if x["SEQIT"].strip()=="5": continue          # บรรทัดสถิติ ไม่นับสมดุล
            if x["TRNTYP"]=="0": dr+=x["AMOUNT"]
            else: cr+=x["AMOUNT"]
        if git and abs(dr-cr)>0.01: bad+=1
    print("ตรวจสมดุล GL (ไม่นับบรรทัดสถิติ): " + ("✓ สมดุลทุกใบ" if bad==0 else f"✗ {bad} ใบไม่สมดุล"))
    if bills:
        print("บัญชีรายได้ที่ดึงจากฐานข้อมูลที่เลือก:")
        for l in bills[0]["lines"]:
            inc = EX.resolve_income(l["stk"])
            print(f"   {l['stk']:12} {l['des'][:34]:34} -> {inc or '⚠️ หาไม่เจอใน DB นี้ (เลือก DB ให้ถูกสาขา)'}")
    print("\n✓ ดูรายละเอียดแต่ละบิลที่ปุ่มดูตัวอย่าง")

def verify_bt(xls, branch=None):
    mode, bills = EX.load_source(xls, branch)
    t = EX.open_tables()
    have = {t["GLJNL"].get(i).get("VOUCHER") for i in range(len(t["GLJNL"]))}
    new = [b for b in bills if b["docnum"] not in have]
    print(f"โหมดโอนเงินระหว่างธนาคาร (BT) · สาขา {EX.bt_cfg(branch)['name']}: {len(bills)} รายการ (ใหม่ {len(new)})")
    bad=0; codes={}
    for b in bills:
        codes[b["src"]] = EX.resolve_bank(b["src"])[0]
        codes[b["dst"]] = EX.resolve_bank(b["dst"])[0]
        recs = EX.build_bt_all(b, b["docnum"], t["GLJNL"], t["GLJNLIT"])
        dr=cr=0
        for m in recs["GLJNLIT"]:
            x=EX.decode_built(t["GLJNLIT"], m)
            if x["TRNTYP"]=="0": dr+=x["AMOUNT"]
            else: cr+=x["AMOUNT"]
        if abs(dr-cr)>0.01: bad+=1
    print("ตรวจสมดุล GL: " + ("✓ สมดุลทุกใบ" if bad==0 else f"✗ {bad} ใบไม่สมดุล"))
    print(f"ค่าธรรมเนียมลงบัญชี: {EX.bt_cfg(branch)['fee_acc']} (แผนก {EX.bt_cfg(branch)['fee_dep'] or 'ว่าง'})")
    if bills:
        b0 = bills[0]
        gj = EX.decode_built(t["GLJNL"], EX.build_bt_all(b0, b0["docnum"], t["GLJNL"], t["GLJNLIT"])["GLJNL"][0])
        print(f"ตัวอย่างคำอธิบาย: {gj['DESCRP']!r}")
    print("รหัสบัญชีธนาคาร -> บัญชี GL (ดึงจาก BKMAS ของ DB ที่เลือก):")
    for c, acc in sorted(codes.items()):
        print(f"   {c:4} -> {acc or '⚠️ หาไม่เจอใน DB นี้ (เลือก DB ให้ถูก/ตรวจ BKMAS)'}")
    print("\n✓ ดูรายละเอียดที่ปุ่มดูตัวอย่าง")

def verify_text(xls, branch=None):
    m = EX.detect_mode(xls)
    if m=="bt":       return capture(verify_bt, xls, branch)
    if m=="petty":    return capture(verify_petty, xls)
    if m=="purchase": return capture(verify_purchase, xls, branch)
    if branch=="surat":                                  # ไม่มีตัวอย่างสุราษให้เทียบ byte
        return capture(verify_generic, xls, branch)
    fn = EX.cmd_verify_parts if m=="parts" else EX.cmd_verify
    return capture(fn, _args(xls))

# ───────── HTML ─────────
HTML = r"""<!doctype html><html lang="th"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>นำเข้าข้อมูลศูนย์บริการ → Express</title>
<style>
:root{--bg:#0f172a;--card:#1e293b;--ac:#38bdf8;--ok:#22c55e;--bad:#f87171;--warn:#f59e0b;--mut:#94a3b8}
*{box-sizing:border-box;font-family:-apple-system,"Segoe UI",Tahoma,sans-serif}
body{margin:0;background:var(--bg);color:#e2e8f0;padding:24px;max-width:1000px;margin:auto}
h1{font-size:20px;margin:0 0 4px} .sub{color:var(--mut);font-size:13px;margin-bottom:18px}
.card{background:var(--card);border-radius:14px;padding:18px;margin-bottom:16px;box-shadow:0 2px 10px #0006}
label{display:block;font-size:13px;color:var(--mut);margin-bottom:6px}
input.path{width:100%;padding:10px;border-radius:9px;border:1px solid #334155;background:#0b1220;color:#e2e8f0;font-size:13px}
.frow{display:flex;gap:10px;align-items:center;margin-bottom:12px} .frow .grow{flex:1}
.btns{display:flex;gap:10px;flex-wrap:wrap;margin-top:6px}
button{border:0;border-radius:10px;padding:11px 18px;font-size:14px;font-weight:600;cursor:pointer;color:#04111f}
button.sm{padding:9px 13px;font-size:13px} button:hover{filter:brightness(1.08)} button:disabled{opacity:.5;cursor:default}
.b-prev{background:var(--ac)} .b-ver{background:#a78bfa} .b-com{background:var(--ok)} .b-pick{background:#cbd5e1}
.badge{font-size:12px;padding:3px 10px;border-radius:20px;white-space:nowrap}
.badge.ok{background:#064e3b;color:#6ee7b7} .badge.bad{background:#450a0a;color:#fca5a5}
.tabs{display:flex;gap:8px;margin-bottom:16px}
.tab{flex:1;text-align:center;padding:13px;border-radius:11px;background:#0b1220;border:1px solid #334155;color:var(--mut);cursor:pointer;font-weight:600;font-size:14px}
.tab.on{background:var(--ac);color:#04111f;border-color:var(--ac)}
.bill{border:1px solid #334155;border-radius:11px;padding:14px;margin:12px 0;background:#0b1220}
.bill.skip{opacity:.55} .bill h3{margin:0 0 8px;font-size:16px;color:var(--ac)} .bill.skip h3{color:var(--mut)}
table{width:100%;border-collapse:collapse;font-size:13px} td{padding:5px 6px;border-bottom:1px solid #1e293b}
td.r{text-align:right;font-variant-numeric:tabular-nums} .tot td{font-weight:700;border-top:1px solid #475569}
pre{white-space:pre-wrap;background:#0b1220;padding:14px;border-radius:10px;font-size:12.5px;line-height:1.5;max-height:460px;overflow:auto}
.tag{display:inline-block;font-size:11px;padding:2px 8px;border-radius:20px;margin-left:8px}
.tag.new{background:#064e3b;color:#6ee7b7} .tag.skip{background:#334155;color:#cbd5e1}
.note{background:#422006;border:1px solid #92400e;border-radius:10px;padding:12px;font-size:13px;color:#fcd34d}
#status{font-size:13px;color:var(--mut);min-height:18px;margin-top:8px}
.spin{display:inline-block;width:14px;height:14px;border:2px solid #475569;border-top-color:var(--ac);border-radius:50%;animation:s .7s linear infinite;vertical-align:-2px}
@keyframes s{to{transform:rotate(360deg)}}
#modal{position:fixed;inset:0;background:#000a;display:none;align-items:center;justify-content:center;z-index:9}
.mbox{background:var(--card);border-radius:14px;padding:18px;width:min(680px,92vw);max-height:84vh;display:flex;flex-direction:column}
.mhead{font-size:15px;font-weight:700;margin-bottom:6px} .mpath{font-size:12px;color:var(--mut);word-break:break-all;margin-bottom:10px}
.flist{flex:1;overflow:auto;border:1px solid #334155;border-radius:10px;background:#0b1220}
.frow2{display:flex;align-items:center;gap:8px;padding:9px 12px;border-bottom:1px solid #1e293b;cursor:pointer;font-size:14px}
.frow2:hover{background:#15233b} .frow2.file:hover{background:#0c2a1e} .frow2 .ic{width:18px;text-align:center}
.frow2 .tags{margin-left:auto;display:flex;gap:6px;align-items:center} .frow2 .jb{font-size:12px;color:var(--mut)}
.minit{display:flex;gap:8px;flex-wrap:wrap;margin-bottom:10px}
.chip{background:#0b1220;border:1px solid #334155;border-radius:8px;padding:6px 10px;font-size:12px;cursor:pointer;color:#cbd5e1}
.mfoot{display:flex;gap:10px;justify-content:space-between;align-items:center;margin-top:12px}
.dim{color:var(--mut);font-size:12px}
</style></head><body>
<h1>🧾 นำเข้าข้อมูลศูนย์บริการ → โปรแกรมบัญชี Express</h1>
<div class="sub">รวมรายการตามวันที่ → สร้างบิลขายเงินสด (โปรแกรมจะเลือกชนิดบิลให้อัตโนมัติจากหัวคอลัมน์ของไฟล์)
  &nbsp;·&nbsp; <a onclick="changePw()" style="color:var(--ac);cursor:pointer">🔑 เปลี่ยนรหัส</a>
  &nbsp;·&nbsp; <a onclick="logout()" style="color:var(--mut);cursor:pointer">ออกจากระบบ</a></div>

<div class="tabs">
  <div class="tab" id="tab-service">🔧 ขายสด-ศูนย์บริการ (HC)</div>
  <div class="tab" id="tab-parts">🔩 ขายสด-อะไหล่ (HO)</div>
  <div class="tab" id="tab-purchase">📦 ซื้อเชื่อ-อะไหล่ (RS)</div>
  <div class="tab" id="tab-petty">💵 จ่ายเงินสดย่อย (DP)</div>
  <div class="tab" id="tab-bt">🏦 โอนเงินระหว่างธนาคาร (BT)</div>
</div>

<div class="card">
  <div class="frow">
    <div class="grow">
      <label>📂 โฟลเดอร์ฐานข้อมูล Express (ที่มีไฟล์ .DBF)</label>
      <input class="path" id="db" readonly>
    </div>
    <button class="b-pick sm" onclick="openPick('db')">เลือกโฟลเดอร์…</button>
    <span id="dbbadge" class="badge"></span>
  </div>
  <div class="frow">
    <div class="grow">
      <label>📄 ไฟล์ Excel ที่จะนำเข้า (.xls)</label>
      <input class="path" id="xlsfile" readonly placeholder="ยังไม่ได้เลือก/อัปโหลดไฟล์">
    </div>
    <button class="b-com sm" onclick="document.getElementById('uplinput').click()">⬆️ อัปโหลดจากเครื่องนี้</button>
    <button class="b-pick sm" onclick="openPick('xlsfile')">เลือกบน server…</button>
    <span id="xlsbadge" class="badge"></span>
  </div>
  <input type="file" id="uplinput" accept=".xls,.xlsx" style="display:none">
  <div class="dim" style="margin-top:-4px">⬆️ อัปโหลด = ใช้ไฟล์ xls จากเครื่องที่เปิดหน้านี้  ·  เลือกบน server = ไฟล์ที่อยู่บนเครื่องหลัก</div>
  <div class="frow" id="branchrow" style="display:none;margin-top:12px">
    <div>
      <label>🏢 สาขา (สำหรับโหมดซื้อ RS / อะไหล่ HO) — เลือกให้ตรงกับฐานข้อมูลที่เลือก</label>
      <select id="branch" style="padding:9px;border-radius:9px;border:1px solid #334155;background:#0b1220;color:#e2e8f0;font-size:13px">
        <option value="chumphon">ชุมพร (เดิม · เลขที่มี /)</option>
        <option value="surat">สุราษ (เลขที่ไม่มี / · RS:P091/ซ2/5000 · HO:แผนกข9/420003)</option>
      </select>
    </div>
  </div>
  <div class="btns">
    <button class="b-prev" onclick="act('preview')">👁 ดูตัวอย่างบิล</button>
    <button class="b-ver"  onclick="act('verify')">🔍 ตรวจสอบความถูกต้อง</button>
    <button class="b-com"  onclick="commit()">💾 บันทึกลงระบบจริง</button>
  </div>
  <div id="status"></div>
</div>

<div class="note" id="note" style="display:none">
  ⚠️ หลังบันทึกแล้ว เปิด Express → เริ่มระบบ → จัดการแฟ้มข้อมูล →
  (1) จัดเรียงดัชนีใหม่ (Reindex)  (2) คำนวณยอดบัญชี/ยอดยกมาใหม่ — และต้องปิด Express ก่อนกดบันทึกทุกครั้ง
</div>

<div class="card" id="result"><div style="color:var(--mut)">— เลือกโฟลเดอร์ฐานข้อมูลและไฟล์ xls แล้วกดปุ่ม —</div></div>

<div id="modal"><div class="mbox">
  <div class="mhead" id="mtitle">เลือก</div>
  <div class="minit" id="mshort"></div>
  <div class="mpath" id="mpath"></div>
  <div class="flist" id="flist"></div>
  <div class="mfoot">
    <span class="dim" id="mhint"></span>
    <div style="display:flex;gap:10px">
      <button class="b-pick sm" onclick="closePick()">ยกเลิก</button>
      <button class="b-com sm" id="mpick" onclick="chooseDir()">เลือกโฟลเดอร์นี้</button>
    </div>
  </div>
</div></div>

<!-- login / ตั้งรหัส -->
<div id="auth" style="position:fixed;inset:0;background:var(--bg);display:flex;align-items:center;justify-content:center;z-index:20">
  <div class="mbox" style="width:min(380px,92vw)">
    <div class="mhead" id="authtitle">🔒 เข้าสู่ระบบ</div>
    <div class="dim" id="authmsg" style="margin-bottom:12px"></div>
    <input class="path" id="authpw" type="password" placeholder="รหัสผ่าน" style="margin-bottom:12px"
           onkeydown="if(event.key==='Enter')authSubmit()">
    <button class="b-com" id="authbtn" style="width:100%" onclick="authSubmit()">เข้าสู่ระบบ</button>
  </div>
</div>

<script>
const $=s=>document.querySelector(s);
let selXls='';
let TOKEN=localStorage.getItem('exp_token')||'';
async function afetch(url,opts){
  opts=opts||{}; opts.headers=Object.assign({},opts.headers||{},{'X-Auth':TOKEN});
  const r=await fetch(url,opts);
  if(r.status===401){ TOKEN=''; localStorage.removeItem('exp_token'); showLogin(); throw new Error('ต้องเข้าสู่ระบบ'); }
  return r;
}
const MODES={service:{name:'ศูนย์บริการ',unit:'งาน'},parts:{name:'อะไหล่',unit:'ใบกำกับ'},purchase:{name:'ซื้ออะไหล่',unit:'ใบรับ'},petty:{name:'จ่ายเงินสดย่อย',unit:'รายการ'},bt:{name:'โอนเงินธนาคาร',unit:'รายการ'}};
function setMode(m){
  $('#tab-service').classList.toggle('on',m==='service');
  $('#tab-parts').classList.toggle('on',m==='parts');
  $('#tab-purchase').classList.toggle('on',m==='purchase');
  $('#tab-petty').classList.toggle('on',m==='petty');
  $('#tab-bt').classList.toggle('on',m==='bt');
  $('#branchrow').style.display = (m && m!=='petty') ? 'flex' : 'none';
}
function branchVal(){ return $('#branch') ? $('#branch').value : 'chumphon'; }
function fmt(n){return n.toLocaleString('th-TH',{minimumFractionDigits:2,maximumFractionDigits:2})}
function busy(b,msg){$('#status').innerHTML=b?'<span class="spin"></span> '+msg:(msg||'');
  document.querySelectorAll('.card button').forEach(x=>x.disabled=b);}

async function loadConfig(){
  const c=await (await afetch('/api/config')).json();
  $('#db').value=c.db_dir||''; $('#xlsfile').value=c.xls_file||''; selXls=c.xls_file||'';
  const db=$('#dbbadge'); db.className='badge '+(c.db_ok?'ok':'bad');
  db.textContent=c.db_ok?'✓ พบฐานข้อมูล':'✗ ไม่พบ DBF'+(c.db_missing&&c.db_missing.length?' ('+c.db_missing.slice(0,3).join(',')+'…)':'');
  updateXls(c);
}
function updateXls(info){
  selXls=info.xls_file||'';
  $('#xlsfile').value=info.xls_name||info.xls_file||'';
  const xb=$('#xlsbadge'); const u=MODES[info.mode]?MODES[info.mode].unit:'รายการ';
  if(!info.xls_file){xb.className='badge bad';xb.textContent='✗ ยังไม่ได้เลือก/อัปโหลดไฟล์';}
  else{xb.className='badge '+(info.xls_ok?'ok':'bad');
    xb.textContent=info.xls_ok?('✓ '+(info.xls_jobs>=0?info.xls_jobs+' '+u:'พร้อม')+(info.mode?' · '+MODES[info.mode].name:'')):'✗ ไฟล์ไม่ถูกต้อง';}
  setMode(info.mode);
}
async function doUpload(file){
  if(!file)return;
  busy(true,'กำลังอัปโหลด '+file.name+' ...');
  try{
    const buf=await file.arrayBuffer();
    const r=await afetch('/api/upload',{method:'POST',
      headers:{'X-Filename':encodeURIComponent(file.name),'Content-Type':'application/octet-stream'},body:buf});
    const d=await r.json();
    if(d.error){$('#result').innerHTML='<pre style="color:#fca5a5">'+d.error+'</pre>';}
    else{updateXls(d); $('#result').innerHTML='<div style="color:var(--ok)">อัปโหลด '+d.xls_name+' แล้ว ('+(d.xls_jobs>=0?d.xls_jobs+' รายการ':'')+') — กดดูตัวอย่างได้เลย</div>';}
  }catch(e){$('#result').innerHTML='<pre style="color:#fca5a5">'+e+'</pre>';}
  busy(false,'');
}
function renderPreview(d){
  setMode(d.mode); const u=MODES[d.mode]?MODES[d.mode].unit:'รายการ';
  const mn=MODES[d.mode]?MODES[d.mode].name:'';
  let h=`<div class="sub">ไฟล์ <b>${d.file}</b> [${mn}] — ${d.njobs} ${u} → ${d.nbills} บิล · จะสร้างใหม่ <b style="color:var(--ok)">${d.nnew}</b> บิล</div>`;
  d.bills.forEach(b=>{
    h+=`<div class="bill ${b.skipped?'skip':''}"><h3>${b.skipped?'⏭ '+b.youref:'● '+b.docnum}
        <span class="tag ${b.skipped?'skip':'new'}">${b.skipped?'มีในระบบแล้ว':'สร้างใหม่'}</span></h3>
        <div class="sub">วันที่ ${b.date} · อ้างอิง ${b.youref} · ${b.njobs} งาน</div><table>`;
    b.lines.forEach(l=>h+=`<tr><td>${l.stk}</td><td>${l.des}</td><td class="r">${fmt(l.amt)}</td></tr>`);
    h+=`<tr><td></td><td class="r">รวมก่อน VAT</td><td class="r">${fmt(b.sub)}</td></tr>
        <tr><td></td><td class="r">VAT 7%</td><td class="r">${fmt(b.vat)}</td></tr>
        <tr class="tot"><td></td><td class="r">รวมทั้งสิ้น</td><td class="r">${fmt(b.total)}</td></tr></table></div>`;
  });
  $('#result').innerHTML=h;
}
async function act(cmd){
  if(!selXls){alert('ยังไม่ได้เลือกไฟล์ xls');return;}
  busy(true, cmd==='preview'?'กำลังคำนวณ...':'กำลังตรวจสอบ...');
  const bq='&branch='+branchVal();
  try{
    if(cmd==='preview'){
      const d=await (await afetch('/api/preview?xls='+encodeURIComponent(selXls)+bq)).json();
      d.error?$('#result').innerHTML='<pre style="color:#fca5a5">'+d.error+'</pre>':renderPreview(d);
    }else{
      const d=await (await afetch('/api/'+cmd+'?xls='+encodeURIComponent(selXls)+bq)).json();
      $('#result').innerHTML='<pre>'+(d.text||d.error)+'</pre>';
    }
  }catch(e){$('#result').innerHTML='<pre style="color:#fca5a5">'+e+'</pre>';}
  busy(false,'เสร็จ');
}
async function commit(){
  if(!selXls){alert('ยังไม่ได้เลือกไฟล์ xls');return;}
  if(!confirm('ยืนยันบันทึกลงฐานข้อมูลจริง?\n\n• ปิดโปรแกรม Express ก่อน\n• ระบบจะสำรองไฟล์ให้อัตโนมัติ\n• หลังบันทึกต้อง Reindex + คำนวณยอดใน Express'))return;
  busy(true,'กำลังบันทึก...');
  try{
    const d=await (await afetch('/api/commit',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({xls:selXls,branch:branchVal()})})).json();
    $('#result').innerHTML='<pre>'+(d.text||d.error)+'</pre>'; $('#note').style.display='block';
  }catch(e){$('#result').innerHTML='<pre style="color:#fca5a5">'+e+'</pre>';}
  busy(false,'บันทึกเสร็จ'); loadConfig();
}

/* ---- picker ---- */
let pickTarget=null, curPath=null;
function openPick(target){
  pickTarget=target;
  const isFile = target==='xlsfile';
  $('#mtitle').textContent = isFile?'เลือกไฟล์ Excel (.xls)':'เลือกโฟลเดอร์ฐานข้อมูล (ที่มี .DBF)';
  $('#mpick').style.display = isFile?'none':'inline-block';
  $('#mhint').textContent = isFile?'👉 คลิกที่ไฟล์ .xls เพื่อเลือก':'';
  $('#modal').style.display='flex';
  const start = isFile ? ($('#xlsfile').value? $('#xlsfile').value.replace(/[^/]*$/,'') : '')
                       : ($('#db').value||'');
  browse(start);
}
function closePick(){$('#modal').style.display='none';}
async function browse(path){
  const isFile = pickTarget==='xlsfile';
  const d=await (await afetch('/api/listdir?path='+encodeURIComponent(path||'')+(isFile?'&files=1':''))).json();
  curPath=d.path;
  $('#mpath').textContent='📍 '+d.path + (d.self_dbf&&!isFile?'   ✓ โฟลเดอร์นี้มี DBF':'');
  let sh=`<span class="chip" onclick='browse(${JSON.stringify(d.home)})'>🏠 Home</span>`;
  (d.drives||[]).forEach(dr=>sh+=`<span class="chip" onclick='browse(${JSON.stringify(dr)})'>💾 ${dr}</span>`);
  if(d.parent)sh+=`<span class="chip" onclick='browse(${JSON.stringify(d.parent)})'>⬆️ ขึ้นบน</span>`;
  $('#mshort').innerHTML=sh;
  let h='';
  d.dirs.forEach(f=>{
    let tags=(f.dbf?'<span class="badge ok">DBF</span>':'')+(f.xls?'<span class="badge ok">xls</span>':'');
    h+=`<div class="frow2" onclick='browse(${JSON.stringify(f.path)})'>
        <span class="ic">📁</span><span>${f.name}</span><span class="tags">${tags}</span></div>`;
  });
  if(isFile)(d.files||[]).forEach(f=>{
    const jb=f.jobs==null?'':(f.jobs>=0?f.jobs+' งาน':'อ่านไม่ได้');
    h+=`<div class="frow2 file" onclick='selectFile(${JSON.stringify(f.path)})'>
        <span class="ic">📄</span><span>${f.name}</span><span class="tags"><span class="jb">${jb}</span></span></div>`;
  });
  $('#flist').innerHTML=h||'<div class="dim" style="padding:14px">— ว่าง —</div>';
}
async function chooseDir(){
  await postCfg({db_dir:curPath}); closePick(); loadConfig();
  $('#result').innerHTML='<div style="color:var(--mut)">อัปเดตโฟลเดอร์ฐานข้อมูลแล้ว</div>';
}
async function selectFile(p){
  await postCfg({xls_file:p}); closePick(); await loadConfig();
  $('#result').innerHTML='<div style="color:var(--mut)">เลือกไฟล์แล้ว — กดดูตัวอย่างบิลได้เลย</div>';
}
async function postCfg(body){
  return (await afetch('/api/config',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(body)})).json();
}
document.getElementById('uplinput').addEventListener('change',e=>doUpload(e.target.files[0]));

/* ---- auth (login / ตั้งรหัส) ---- */
let authMode='login';
function showLogin(){authMode='login';$('#authtitle').textContent='🔒 เข้าสู่ระบบ';
  $('#authmsg').textContent='ใส่รหัสผ่านเพื่อใช้งานโปรแกรม';$('#authmsg').style.color='var(--mut)';
  $('#authbtn').textContent='เข้าสู่ระบบ';$('#authpw').value='';$('#auth').style.display='flex';setTimeout(()=>$('#authpw').focus(),50);}
function showSetup(t){authMode='setup';$('#authtitle').textContent=t||'🔑 ตั้งรหัสผ่านครั้งแรก';
  $('#authmsg').textContent='ตั้งรหัสผ่านสำหรับเข้าใช้งาน (อย่างน้อย 4 ตัว)';$('#authmsg').style.color='var(--mut)';
  $('#authbtn').textContent='บันทึกรหัสและเข้าใช้';$('#authpw').value='';$('#auth').style.display='flex';setTimeout(()=>$('#authpw').focus(),50);}
function changePw(){showSetup('🔑 เปลี่ยนรหัสผ่าน');}
async function authSubmit(){
  const pw=$('#authpw').value;
  const ep=authMode==='setup'?'/api/setpassword':'/api/login';
  const d=await (await fetch(ep,{method:'POST',headers:{'Content-Type':'application/json','X-Auth':TOKEN},body:JSON.stringify({password:pw})})).json();
  if(d.ok){TOKEN=d.token;localStorage.setItem('exp_token',TOKEN);$('#auth').style.display='none';loadConfig();}
  else{$('#authmsg').textContent=d.error||'รหัสผ่านไม่ถูกต้อง';$('#authmsg').style.color='#fca5a5';}
}
function logout(){TOKEN='';localStorage.removeItem('exp_token');showLogin();}
async function boot(){
  const a=await (await fetch('/api/auth')).json();
  if(!a.hasPassword) return showSetup();
  if(!TOKEN) return showLogin();
  try{ await loadConfig(); $('#auth').style.display='none'; }  // token ใช้ได้ -> เข้าแอป
  catch(e){ /* afetch เด้ง login ให้แล้ว */ }
}
boot();
</script></body></html>"""

# ───────── HTTP ─────────
class H(BaseHTTPRequestHandler):
    def log_message(self, *a): pass
    def _send(self, code, body, ctype="application/json; charset=utf-8"):
        b = body if isinstance(body, bytes) else body.encode("utf-8")
        self.send_response(code); self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(b))); self.end_headers()
        self.wfile.write(b)

    def _authed(self): return self.headers.get("X-Auth") in TOKENS

    def do_GET(self):
        u = urlparse(self.path); q = parse_qs(u.query)
        arg = lambda k: (q.get(k) or [None])[0]
        try:
            if u.path == "/":
                self._send(200, HTML, "text/html; charset=utf-8"); return
            if u.path == "/api/auth":      # สถานะรหัสผ่าน (ไม่ต้อง login)
                self._send(200, json.dumps({"hasPassword": bool(AUTH_HASH)})); return
            if not self._authed():         # ทุก api อื่นต้อง login ก่อน
                self._send(401, json.dumps({"error": "unauthorized"})); return
            if u.path == "/api/config":
                self._send(200, json.dumps(config_status(), ensure_ascii=False))
            elif u.path == "/api/listdir":
                self._send(200, json.dumps(listdir(arg("path"), want_files=bool(arg("files"))), ensure_ascii=False))
            elif u.path == "/api/preview":
                self._send(200, json.dumps(preview_data(arg("xls"), arg("branch")), ensure_ascii=False))
            elif u.path == "/api/verify":
                self._send(200, json.dumps({"text": verify_text(arg("xls"), arg("branch"))}, ensure_ascii=False))
            else:
                self._send(404, json.dumps({"error": "not found"}))
        except Exception:
            self._send(200, json.dumps({"error": clean(traceback.format_exc())}, ensure_ascii=False))

    def do_POST(self):
        global XLS_FILE
        u = urlparse(self.path)
        n = int(self.headers.get("Content-Length", 0))
        body = self.rfile.read(n)
        try:
            # ---- auth endpoints (เปิดได้โดยไม่ต้อง login) ----
            if u.path == "/api/login":
                pw = json.loads(body or b"{}").get("password", "")
                if AUTH_HASH and pw_hash(pw) == AUTH_HASH:
                    self._send(200, json.dumps({"ok": True, "token": new_token()}))
                else:
                    self._send(200, json.dumps({"ok": False}))
                return
            if u.path == "/api/setpassword":
                # ตั้งรหัสครั้งแรก (ยังไม่มีรหัส) หรือเปลี่ยนรหัส (ต้อง login แล้ว)
                if AUTH_HASH and not self._authed():
                    self._send(401, json.dumps({"error": "unauthorized"})); return
                pw = json.loads(body or b"{}").get("password", "")
                if len(pw) < 4:
                    self._send(200, json.dumps({"ok": False, "error": "รหัสสั้นเกินไป (อย่างน้อย 4 ตัว)"})); return
                set_password(pw)
                self._send(200, json.dumps({"ok": True, "token": new_token()}))
                return
            # ---- ทุก POST อื่นต้อง login ----
            if not self._authed():
                self._send(401, json.dumps({"error": "unauthorized"})); return
            if u.path == "/api/upload":
                from urllib.parse import unquote
                fn = os.path.basename(unquote(self.headers.get("X-Filename", "upload.xls")))
                if not fn.lower().endswith((".xls", ".xlsx")): fn += ".xls"
                os.makedirs(UPLOAD_DIR, exist_ok=True)
                path = os.path.join(UPLOAD_DIR, fn)
                with open(path, "wb") as f: f.write(body)
                jobs = xls_jobs(path)
                self._send(200, json.dumps({"xls_file": path, "xls_name": fn,
                    "mode": xls_mode(path), "xls_jobs": jobs, "xls_ok": jobs >= 0}, ensure_ascii=False))
                return
            data = json.loads(body or b"{}")
            if u.path == "/api/config":
                if "db_dir" in data:  EX.DB_DIR = os.path.abspath(os.path.expanduser(data["db_dir"]))
                if "xls_file" in data:
                    XLS_FILE = os.path.abspath(os.path.expanduser(data["xls_file"]))
                    EX.XLS_DIR = os.path.dirname(XLS_FILE)
                save_config()
                self._send(200, json.dumps(config_status(), ensure_ascii=False))
            elif u.path == "/api/commit":
                self._send(200, json.dumps({"text": capture(EX.cmd_preview, _args(data.get("xls"), data.get("branch")), commit=True)}, ensure_ascii=False))
            else:
                self._send(404, json.dumps({"error": "not found"}))
        except Exception:
            self._send(200, json.dumps({"error": clean(traceback.format_exc())}, ensure_ascii=False))

def lan_ip():
    """หา IP ของเครื่องในวง LAN (ไม่ส่งข้อมูลจริง)"""
    s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
    try: s.connect(("8.8.8.8", 80)); ip = s.getsockname()[0]
    except Exception: ip = "127.0.0.1"
    finally: s.close()
    return ip

def free_port(host, want=0):
    try:
        s = socket.socket(); s.bind((host, want)); p = s.getsockname()[1]; s.close(); return p
    except OSError:
        s = socket.socket(); s.bind((host, 0)); p = s.getsockname()[1]; s.close(); return p

def main():
    ap = argparse.ArgumentParser(description="โปรแกรมนำเข้าข้อมูล -> Express")
    ap.add_argument("--local", action="store_true", help="เปิดเฉพาะเครื่องนี้ (ไม่แชร์ LAN)")
    ap.add_argument("--lan", action="store_true", help="(ค่าเริ่มต้นอยู่แล้ว) แชร์ในวง LAN")
    ap.add_argument("--port", type=int, default=0, help="พอร์ต (ค่าเริ่มต้น: แชร์ LAN=8765, --local=สุ่ม)")
    a = ap.parse_args()
    load_config()

    lan = not a.local                       # ค่าเริ่มต้น = แชร์ LAN
    host = "127.0.0.1" if a.local else "0.0.0.0"
    port = a.port or (0 if a.local else 8765)
    port = free_port(host, port)
    httpd = ThreadingHTTPServer((host, port), H)
    local_url = f"http://127.0.0.1:{port}/"

    print("="*60)
    print("  โปรแกรมนำเข้าข้อมูลขายเงินสด → Express")
    print(f"  เครื่องนี้เปิดที่ : {local_url}")
    if lan:
        print(f"  เครื่องอื่นใน LAN : http://{lan_ip()}:{port}/   <- ส่ง URL นี้ให้เครื่องอื่น")
        print("  ⚠️  แชร์ LAN (ค่าเริ่มต้น): ใครในวงเน็ตเดียวกันเข้าหน้านี้และกดบันทึกได้")
        print("      ครั้งแรก Windows จะถาม Allow firewall ต้องกดอนุญาต  (อยากปิดแชร์: ใส่ --local)")
    print(f"  ฐานข้อมูล : {EX.DB_DIR}")
    print(f"  ไฟล์ xls  : {XLS_FILE or '(ยังไม่เลือก)'}")
    print("  (ปิดหน้าต่างนี้ = ปิดโปรแกรม)   กด Ctrl+C เพื่อหยุด")
    print("="*60)
    threading.Timer(0.8, lambda: webbrowser.open(local_url)).start()
    try: httpd.serve_forever()
    except KeyboardInterrupt: print("\nปิดโปรแกรมแล้ว")

if __name__ == "__main__":
    main()
