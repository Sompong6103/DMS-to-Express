#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
express_gui.py — หน้าจอกดใช้งาน (เว็บในเครื่อง) สำหรับนำเข้างานบริการเข้า Express
รัน:  python3 express_gui.py     แล้วเบราว์เซอร์จะเปิดให้อัตโนมัติ
เลือกโฟลเดอร์ฐานข้อมูล (DBF) และเลือกไฟล์ xls โดยตรงได้จากหน้าจอ (จำค่าไว้ในไฟล์ config)
"""
import os, io, sys, glob, json, socket, threading, webbrowser, contextlib, traceback
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

CONFIG_FILE = os.path.join(HERE, "express_config.json")
REQUIRED = ["ARTRN", "GLJNL", "GLJNLIT", "ISVAT", "STCRD", "ISRUN"]   # ตารางที่ต้องมี
XLS_FILE = ""   # ไฟล์ xls ที่เลือกไว้ (path เต็ม)

# ───────── config ─────────
def load_config():
    global XLS_FILE
    cfg = {"db_dir": EX.DB_DIR, "xls_file": ""}
    if os.path.exists(CONFIG_FILE):
        try: cfg.update(json.load(open(CONFIG_FILE, encoding="utf-8")))
        except Exception: pass
    if cfg.get("db_dir"):  EX.DB_DIR = cfg["db_dir"]
    XLS_FILE = cfg.get("xls_file") or ""
    if XLS_FILE: EX.XLS_DIR = os.path.dirname(XLS_FILE)
    return cfg

def save_config():
    json.dump({"db_dir": EX.DB_DIR, "xls_file": XLS_FILE},
              open(CONFIG_FILE, "w", encoding="utf-8"), ensure_ascii=False, indent=2)

def db_ok(path):
    return bool(path) and all(os.path.exists(os.path.join(path, n + ".DBF")) for n in REQUIRED)

def xls_jobs(path):
    try: return len(EX.read_xls_jobs(path))
    except Exception: return -1

def config_status():
    ok = bool(XLS_FILE) and os.path.isfile(XLS_FILE) and XLS_FILE.lower().endswith(".xls")
    return {"db_dir": EX.DB_DIR, "db_ok": db_ok(EX.DB_DIR),
            "db_missing": [n for n in REQUIRED if not os.path.exists(os.path.join(EX.DB_DIR or "", n + ".DBF"))],
            "xls_file": XLS_FILE, "xls_name": os.path.basename(XLS_FILE) if XLS_FILE else "",
            "xls_ok": ok, "xls_jobs": xls_jobs(XLS_FILE) if ok else -1}

# ───────── folder/file browser ─────────
def listdir(path, want_files=False):
    path = os.path.abspath(os.path.expanduser(path or os.path.expanduser("~")))
    if not os.path.isdir(path): path = os.path.expanduser("~")
    dirs, files = [], []
    try:
        entries = sorted(os.scandir(path), key=lambda x: x.name.lower())
    except PermissionError:
        entries = []
    xls_here = [e for e in entries if e.is_file() and e.name.lower().endswith(".xls")]
    count_jobs = want_files and len(xls_here) <= 60      # ถ้าไฟล์เยอะมากไม่นับงาน (ช้า)
    for e in entries:
        if e.name.startswith("."): continue
        if e.is_dir(follow_symlinks=False):
            try: names = os.listdir(e.path)
            except Exception: names = []
            upper = {n.upper() for n in names}
            dirs.append({"name": e.name, "path": e.path,
                         "dbf": "ARTRN.DBF" in upper,
                         "xls": any(n.lower().endswith(".xls") for n in names)})
        elif want_files and e.is_file() and e.name.lower().endswith(".xls"):
            files.append({"name": e.name, "path": e.path,
                          "jobs": xls_jobs(e.path) if count_jobs else None})
    parent = os.path.dirname(path)
    return {"path": path, "parent": (parent if parent != path else None),
            "dirs": dirs, "files": files,
            "self_dbf": "ARTRN.DBF" in {n.upper() for n in (os.listdir(path) if os.path.isdir(path) else [])},
            "home": os.path.expanduser("~")}

# ───────── helpers ─────────
def _args(xls): return type("A", (), {"xls": xls})()

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

def preview_data(xls):
    jobs  = EX.read_xls_jobs(xls)
    bills = EX.group_bills(jobs)
    t     = EX.open_tables()
    plan  = EX.next_numbers(t["ARTRN"], bills)
    out   = []
    for b, docnum in plan:
        out.append({"docnum": docnum, "skipped": docnum is None,
                    "date": b["date"].strftime("%d/%m/%Y"), "youref": b["youref"], "njobs": b["njobs"],
                    "lines": [{"stk": l["stk"], "des": l["des"], "amt": l["amt"]} for l in b["lines"]],
                    "sub": b["sub"], "vat": b["vat"], "total": b["total"]})
    new = sum(1 for x in out if not x["skipped"])
    return {"file": os.path.basename(xls), "njobs": len(jobs),
            "nbills": len(bills), "nnew": new, "bills": out}

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
<div class="sub">รวมงานบริการตามวันที่ → สร้างบิลขายเงินสด-ศูนย์บริการ (HC ปีเดือน/เลขรัน)</div>

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
      <input class="path" id="xlsfile" readonly placeholder="ยังไม่ได้เลือกไฟล์">
    </div>
    <button class="b-pick sm" onclick="openPick('xlsfile')">เลือกไฟล์…</button>
    <span id="xlsbadge" class="badge"></span>
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

<script>
const $=s=>document.querySelector(s);
let selXls='';
function fmt(n){return n.toLocaleString('th-TH',{minimumFractionDigits:2,maximumFractionDigits:2})}
function busy(b,msg){$('#status').innerHTML=b?'<span class="spin"></span> '+msg:(msg||'');
  document.querySelectorAll('.card button').forEach(x=>x.disabled=b);}

async function loadConfig(){
  const c=await (await fetch('/api/config')).json();
  $('#db').value=c.db_dir||''; $('#xlsfile').value=c.xls_file||''; selXls=c.xls_file||'';
  const db=$('#dbbadge'); db.className='badge '+(c.db_ok?'ok':'bad');
  db.textContent=c.db_ok?'✓ พบฐานข้อมูล':'✗ ไม่พบ DBF'+(c.db_missing&&c.db_missing.length?' ('+c.db_missing.slice(0,3).join(',')+'…)':'');
  const xb=$('#xlsbadge');
  if(!c.xls_file){xb.className='badge bad';xb.textContent='✗ ยังไม่ได้เลือกไฟล์';}
  else{xb.className='badge '+(c.xls_ok?'ok':'bad');
    xb.textContent=c.xls_ok?('✓ '+(c.xls_jobs>=0?c.xls_jobs+' งาน':'พร้อม')):'✗ ไฟล์ไม่ถูกต้อง';}
}
function renderPreview(d){
  let h=`<div class="sub">ไฟล์ <b>${d.file}</b> — ${d.njobs} งาน → ${d.nbills} บิล · จะสร้างใหม่ <b style="color:var(--ok)">${d.nnew}</b> บิล</div>`;
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
  try{
    if(cmd==='preview'){
      const d=await (await fetch('/api/preview?xls='+encodeURIComponent(selXls))).json();
      d.error?$('#result').innerHTML='<pre style="color:#fca5a5">'+d.error+'</pre>':renderPreview(d);
    }else{
      const d=await (await fetch('/api/'+cmd+'?xls='+encodeURIComponent(selXls))).json();
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
    const d=await (await fetch('/api/commit',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({xls:selXls})})).json();
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
  const d=await (await fetch('/api/listdir?path='+encodeURIComponent(path||'')+(isFile?'&files=1':''))).json();
  curPath=d.path;
  $('#mpath').textContent='📍 '+d.path + (d.self_dbf&&!isFile?'   ✓ โฟลเดอร์นี้มี DBF':'');
  let sh=`<span class="chip" onclick='browse(${JSON.stringify(d.home)})'>🏠 Home</span>`;
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
  return (await fetch('/api/config',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(body)})).json();
}
loadConfig();
</script></body></html>"""

# ───────── HTTP ─────────
class H(BaseHTTPRequestHandler):
    def log_message(self, *a): pass
    def _send(self, code, body, ctype="application/json; charset=utf-8"):
        b = body if isinstance(body, bytes) else body.encode("utf-8")
        self.send_response(code); self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(b))); self.end_headers()
        self.wfile.write(b)

    def do_GET(self):
        u = urlparse(self.path); q = parse_qs(u.query)
        arg = lambda k: (q.get(k) or [None])[0]
        try:
            if u.path == "/":
                self._send(200, HTML, "text/html; charset=utf-8")
            elif u.path == "/api/config":
                self._send(200, json.dumps(config_status(), ensure_ascii=False))
            elif u.path == "/api/listdir":
                self._send(200, json.dumps(listdir(arg("path"), want_files=bool(arg("files"))), ensure_ascii=False))
            elif u.path == "/api/preview":
                self._send(200, json.dumps(preview_data(arg("xls")), ensure_ascii=False))
            elif u.path == "/api/verify":
                self._send(200, json.dumps({"text": capture(EX.cmd_verify, _args(arg("xls")))}, ensure_ascii=False))
            else:
                self._send(404, json.dumps({"error": "not found"}))
        except Exception:
            self._send(200, json.dumps({"error": clean(traceback.format_exc())}, ensure_ascii=False))

    def do_POST(self):
        global XLS_FILE
        u = urlparse(self.path)
        n = int(self.headers.get("Content-Length", 0))
        data = json.loads(self.rfile.read(n) or b"{}")
        try:
            if u.path == "/api/config":
                if "db_dir" in data:  EX.DB_DIR = os.path.abspath(os.path.expanduser(data["db_dir"]))
                if "xls_file" in data:
                    XLS_FILE = os.path.abspath(os.path.expanduser(data["xls_file"]))
                    EX.XLS_DIR = os.path.dirname(XLS_FILE)
                save_config()
                self._send(200, json.dumps(config_status(), ensure_ascii=False))
            elif u.path == "/api/commit":
                self._send(200, json.dumps({"text": capture(EX.cmd_preview, _args(data.get("xls")), commit=True)}, ensure_ascii=False))
            else:
                self._send(404, json.dumps({"error": "not found"}))
        except Exception:
            self._send(200, json.dumps({"error": clean(traceback.format_exc())}, ensure_ascii=False))

def main():
    load_config()
    s = socket.socket(); s.bind(("127.0.0.1", 0)); port = s.getsockname()[1]; s.close()
    url = f"http://127.0.0.1:{port}/"
    httpd = ThreadingHTTPServer(("127.0.0.1", port), H)
    print("="*54)
    print("  โปรแกรมนำเข้าข้อมูลศูนย์บริการ → Express")
    print(f"  เปิดใช้งานที่: {url}")
    print(f"  ฐานข้อมูล: {EX.DB_DIR}")
    print(f"  ไฟล์ xls : {XLS_FILE or '(ยังไม่เลือก)'}")
    print("  (ปิดหน้าต่างนี้ = ปิดโปรแกรม)   กด Ctrl+C เพื่อหยุด")
    print("="*54)
    threading.Timer(0.8, lambda: webbrowser.open(url)).start()
    try: httpd.serve_forever()
    except KeyboardInterrupt: print("\nปิดโปรแกรมแล้ว")

if __name__ == "__main__":
    main()
