#!/usr/bin/env python3
"""COSMIC Excel → Word Web 服务 v4.1"""
import os, sys, uuid, subprocess, tempfile, threading
from flask import Flask, request, send_file, jsonify, render_template_string

PIPELINE = r"C:\Users\小黑\.workbuddy\skills\cosmic-pipeline\scripts\pipeline.py"
if not os.path.exists(PIPELINE):
    PIPELINE = os.path.join(os.path.expanduser("~"), ".workbuddy", "skills", "cosmic-pipeline", "scripts", "pipeline.py")
TEMPLATE_PATH = "COSMIC模板.xlsx"
PYTHON = sys.executable

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024
tasks = {}

HTML = r'''<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>COSMIC 需求文档生成器</title>
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&display=swap');
*{margin:0;padding:0;box-sizing:border-box}
:root{--bg:#06070a;--s1:#0e1117;--s2:#161b22;--s3:#1c2333;--b1:rgba(255,255,255,.06);--b2:rgba(255,255,255,.1);
  --t1:#f0f6fc;--t2:#8b949e;--t3:#484f58;--ac:#58a6ff;--ac2:#bc8cff;--gn:#3fb950;--or:#d29922;--rd:#f85149;
  --grad:linear-gradient(135deg,#58a6ff,#bc8cff,#f778ba)}
body{font-family:'Inter','Microsoft YaHei',system-ui,sans-serif;background:var(--bg);color:var(--t1);
  min-height:100vh;background-image:radial-gradient(ellipse 80% 50% at 50% -20%,rgba(88,166,255,.08),transparent)}
.app{max-width:920px;margin:0 auto;padding:40px 20px 60px}
.hero{text-align:center;padding:40px 0 44px}
.badge{display:inline-flex;align-items:center;gap:6px;padding:5px 14px;background:rgba(88,166,255,.12);
  border:1px solid rgba(88,166,255,.2);border-radius:20px;font-size:11px;color:var(--ac);font-weight:600;letter-spacing:.5px;margin-bottom:16px}
.badge::before{content:'';width:6px;height:6px;border-radius:50%;background:var(--ac);box-shadow:0 0 8px var(--ac);animation:blink 2s infinite}
h1{font-size:clamp(28px,4vw,40px);font-weight:800;background:var(--grad);-webkit-background-clip:text;-webkit-text-fill-color:transparent;margin-bottom:10px;letter-spacing:-.02em}
.hero p{color:var(--t2);font-size:14px;max-width:440px;margin:0 auto;line-height:1.6}
.card{background:var(--s1);border:1px solid var(--b1);border-radius:12px;padding:24px;position:relative;overflow:hidden;margin-top:16px}
.card::before{content:'';position:absolute;top:0;left:0;right:0;height:1px;background:linear-gradient(90deg,transparent,rgba(88,166,255,.15),transparent)}
.tag{display:flex;align-items:center;gap:8px;font-size:11px;font-weight:600;color:var(--t2);text-transform:uppercase;letter-spacing:1px;margin-bottom:18px}
.tag i{width:18px;height:18px;border-radius:5px;display:flex;align-items:center;justify-content:center;font-size:9px;font-weight:700;font-style:normal;color:var(--bg)}
.tag .b{background:var(--ac)}.tag .g{background:var(--gn)}.tag .p{background:var(--ac2)}
.upload{border:1.5px dashed var(--b2);border-radius:12px;padding:44px 20px;text-align:center;cursor:pointer;transition:.3s}
.upload:hover,.upload.over{border-color:var(--ac);background:rgba(88,166,255,.05);box-shadow:inset 0 0 40px rgba(88,166,255,.02)}
.upload .ic{font-size:40px;margin-bottom:10px;opacity:.5}.upload .tt{font-size:15px;font-weight:600;margin-bottom:4px}
.upload .tt em{font-style:normal;color:var(--ac)}.upload .ht{font-size:12px;color:var(--t3)}
input[type=file]{display:none}
.fbadge{display:none;margin-top:12px;padding:10px 14px;background:var(--s2);border:1px solid var(--b1);border-radius:10px;font-size:13px;align-items:center;gap:10px}
.fbadge .fi{width:30px;height:30px;border-radius:7px;background:rgba(88,166,255,.1);display:flex;align-items:center;justify-content:center;font-size:14px}
.fbadge .fn{font-weight:600}.fbadge .fs{font-size:11px;color:var(--t3);margin-top:2px}
.acts{display:flex;gap:10px;flex-wrap:wrap;margin-top:18px}
.btn{display:inline-flex;align-items:center;gap:8px;padding:10px 20px;border-radius:10px;font-size:13px;font-weight:600;cursor:pointer;transition:.25s;border:none;font-family:inherit}
.btn:active{transform:scale(.97)}.btn-go{background:var(--grad);color:#fff;box-shadow:0 4px 16px rgba(88,166,255,.2)}
.btn-go:hover{box-shadow:0 6px 24px rgba(88,166,255,.3);transform:translateY(-1px)}
.btn-go:disabled{opacity:.3;cursor:not-allowed;transform:none;box-shadow:none}
.spin{width:16px;height:16px;border:2px solid rgba(255,255,255,.3);border-top-color:#fff;border-radius:50%;animation:sp .7s linear infinite;display:inline-block}
.btn-g{background:var(--s2);color:var(--t2);border:1px solid var(--b1)}.btn-g:hover{background:var(--s3);color:var(--t1)}
.btn-g:disabled{opacity:.25;cursor:not-allowed}
.btn-g.go{color:var(--or);border-color:rgba(210,153,34,.2)}.btn-g.go:hover{background:rgba(210,153,34,.08)}
.btn-g.em{color:var(--gn);border-color:rgba(63,185,80,.2)}.btn-g.em:hover{background:rgba(63,185,80,.08)}
.pw{margin-top:18px}.pb{width:100%;height:5px;background:var(--s3);border-radius:3px;overflow:hidden}
.pf{height:100%;border-radius:3px;background:var(--grad);width:0;transition:width .8s;box-shadow:0 0 12px rgba(88,166,255,.25)}
.pt{display:flex;justify-content:space-between;margin-top:8px;font-size:12px;color:var(--t2)}
.pt .pct{font-weight:700;color:var(--ac);font-size:13px}
.st{margin-top:20px}
.si{display:flex;gap:14px;padding:10px 0 10px 24px;position:relative;border-left:2px solid var(--b1);margin-left:12px}
.si::before{content:'';position:absolute;left:-6px;top:14px;width:10px;height:10px;border-radius:50%;border:2px solid var(--b1);background:var(--s1);transition:.4s;z-index:1}
.si.ac::before{border-color:var(--ac);background:rgba(88,166,255,.15);box-shadow:0 0 10px rgba(88,166,255,.4);animation:pulse 2s infinite}
.si.dn::before{border-color:var(--gn);background:var(--gn);box-shadow:0 0 6px rgba(63,185,80,.3)}
.sb{flex:1}.sh{display:flex;align-items:center;gap:8px}
.idx{font-size:10px;font-weight:700;color:var(--t3);min-width:16px}.idx.ac{color:var(--ac)}.idx.dn{color:var(--gn)}
.stt{font-size:12px;font-weight:600;color:var(--t2);transition:.3s}.si.ac .stt{color:var(--t1)}.si.dn .stt{color:var(--gn)}
.sti{font-size:10px;color:var(--t3);margin-left:auto;font-family:monospace}
.subs{display:flex;flex-wrap:wrap;gap:5px 10px;margin-top:7px;font-size:11px}
.sg{padding:2px 8px;border-radius:4px;background:var(--s2);color:var(--t3);border:1px solid var(--b1);transition:.3s}
.sg.ok{background:rgba(63,185,80,.1);color:var(--gn);border-color:rgba(63,185,80,.15)}
.sg.wk{background:rgba(88,166,255,.1);color:var(--ac);border-color:rgba(88,166,255,.15);animation:glow 1.5s infinite}
.lc{display:none;margin-top:16px}
.lh{display:flex;align-items:center;justify-content:space-between;padding:11px 14px;background:var(--s2);border:1px solid var(--b1);border-radius:10px 10px 0 0;cursor:pointer}
.lh:hover{background:var(--s3)}.lh span{font-size:12px;color:var(--t2);font-weight:600;letter-spacing:.5px}
.lh .ar{font-size:10px;color:var(--t3);transition:.3s}
.lb{background:#08090d;border:1px solid var(--b1);border-top:none;border-radius:0 0 10px 10px;max-height:300px;overflow-y:auto;padding:12px;
  font-family:'Cascadia Code',Consolas,monospace;font-size:11px;line-height:1.8;white-space:pre-wrap;word-break:break-all}
.ll{display:block}.ll::before{content:'› ';color:var(--t3)}.ll.s{color:var(--ac);font-weight:600}.ll.d{color:var(--gn)}.ll.e{color:var(--rd)}.ll.i{color:var(--or)}.ll.m{color:var(--t3)}
.toast{position:fixed;top:24px;right:24px;padding:12px 20px;border-radius:10px;font-size:13px;font-weight:500;z-index:999;display:none;
  backdrop-filter:blur(16px);border:1px solid var(--b1);animation:si .4s cubic-bezier(.16,1,.3,1)}
.toast.ok{background:rgba(63,185,80,.12);color:var(--gn);border-color:rgba(63,185,80,.15)}
.toast.er{background:rgba(248,81,73,.12);color:var(--rd);border-color:rgba(248,81,73,.15)}
.ft{text-align:center;padding:28px 0 0;font-size:11px;color:var(--t3)}
@keyframes pulse{0%,100%{box-shadow:0 0 4px rgba(88,166,255,.2)}50%{box-shadow:0 0 16px rgba(88,166,255,.5)}}
@keyframes glow{0%,100%{box-shadow:0 0 3px rgba(88,166,255,.1)}50%{box-shadow:0 0 10px rgba(88,166,255,.3)}}
@keyframes sp{to{transform:rotate(360deg)}}@keyframes si{from{transform:translateY(-16px);opacity:0}to{transform:translateY(0);opacity:1}}
@keyframes blink{0%,100%{opacity:1}50%{opacity:.3}}
</style>
</head>
<body>
<div class="app">
  <div class="hero">
    <div class="badge">COSMIC PIPELINE</div>
    <h1>需求文档生成器</h1>
    <p>上传功能拆分 Excel，自动生成含结构图与时序图的完整工程需求文档</p>
  </div>
  <div class="card">
    <div class="tag"><i class="b">1</i> 选择文件</div>
    <div class="upload" id="dropZone">
      <div class="ic">📄</div>
      <div class="tt">拖拽 Excel 到这里，或 <em>点击选择</em></div>
      <div class="ht">.xlsx · 最大 50MB · 支持 7列/10列/14列</div>
    </div>
    <input type="file" id="fileInput" accept=".xlsx">
    <div class="fbadge" id="fb"><div class="fi">📊</div><div><div class="fn" id="fn"></div><div class="fs" id="fs"></div></div></div>
    <div class="acts">
      <button class="btn btn-go" id="go" disabled><span id="gi">⚡</span> 开始生成</button>
      <button class="btn btn-g go" id="tpl">📥 下载模板</button>
      <button class="btn btn-g em" id="dl" disabled>📄 下载文档</button>
    </div>
  </div>
  <div class="card" id="pc" style="display:none">
    <div class="tag"><i class="g">2</i> 生成进度</div>
    <div class="pw"><div class="pb"><div class="pf" id="pf"></div></div>
      <div class="pt"><span id="pl">准备中...</span><span class="pct" id="pp">0%</span></div></div>
    <div class="st" id="sl"></div>
  </div>
  <div class="card lc" id="lc">
    <div class="lh" id="lh"><span>📋 运行日志</span><span class="ar" id="la">▼</span></div>
    <div class="lb" id="lb"></div>
  </div>
  <div class="ft">COSMIC Pipeline · Excel → Word 需求文档自动化</div>
</div>
<div class="toast" id="toast"></div>
<script>
var file=null,tid=null,poller=null,logOpen=true,ss=[];
var SD=[
  {n:"数据预处理",s:["读取Excel","检测列格式","填充合并"],w:5},
  {n:"生成Word",s:["构建HTML","转换Word"],w:25},
  {n:"增强PRD",s:["提取字段","生成段落","清理术语"],w:10},
  {n:"功能结构图",s:["构建映射","写入Excel"],w:5},
  {n:"结构图→PNG",s:["渲染图片"],w:15},
  {n:"时序图",s:["参与者","消息","PNG","描述"],w:15},
  {n:"组装Word",s:["插图","时序图","行距"],w:10},
  {n:"H5 PRD重写",s:["提取过程","构建段落","替换"],w:10},
  {n:"统一样式",s:["字体","缩进间距"],w:5}
];
var TW=0;for(var i=0;i<SD.length;i++)TW+=SD[i].w;

function initS(){
  var el=document.getElementById('sl');el.innerHTML='';ss=[];
  for(var i=0;i<SD.length;i++){
    ss.push({s:[],d:false,t0:0});
    for(var j=0;j<SD[i].s.length;j++)ss[i].s.push(0);
    var d=document.createElement('div');d.className='si';d.id='s'+i;
    d.innerHTML='<div class="sb"><div class="sh"><span class="idx">'+('0'+(i+1)).slice(-2)+'</span><span class="stt">'+SD[i].n+'</span><span class="sti" id="t'+i+'"></span></div><div class="subs" id="u'+i+'"></div></div>';
    el.appendChild(d);
    renderSubs(i);
  }
}
function renderSubs(i){
  var u=document.getElementById('u'+i);if(!u)return;u.innerHTML='';
  for(var j=0;j<SD[i].s.length;j++){
    var s=document.createElement('span');s.className='sg';
    if(ss[i].s[j])s.className+=' ok';
    s.textContent=SD[i].s[j];u.appendChild(s);
  }
}
function setStep(i,st,ui){
  var el=document.getElementById('s'+i);if(!el)return;
  el.className='si'+(st==='active'?' ac':st==='done'?' dn':'');
  var idx=el.querySelector('.idx');if(idx)idx.className='idx'+(st==='active'?' ac':st==='done'?' dn':'');
  if(st==='active'&&ui!==undefined){ss[i].s[ui]=1;renderSubs(i);
    var tags=el.querySelectorAll('.sg');if(tags[ui])tags[ui].className='sg wk';}
  if(st==='done'&&!ss[i].d){ss[i].d=true;for(var j=0;j<SD[i].s.length;j++)ss[i].s[j]=1;renderSubs(i);}
  if(st==='active'&&!ss[i].t0)ss[i].t0=Date.now();
  if(ss[i].t0){var sec=Math.round((Date.now()-ss[i].t0)/1000);if(sec>0){var te=document.getElementById('t'+i);if(te)te.textContent=sec+'s';}}
  var w=0;for(var k=0;k<SD.length;k++){if(ss[k].d)w+=SD[k].w;
    else if(k===i){var ds=0;for(var m=0;m<ss[k].s.length;m++)if(ss[k].s[m])ds++;w+=SD[k].w*(ds/SD[k].s.length)}}
  var pct=Math.round(w/TW*100);
  document.getElementById('pf').style.width=pct+'%';document.getElementById('pp').textContent=pct+'%';
  document.getElementById('pl').textContent=st==='done'?'✅ 完成':SD[i].n;
}
function toast(m,t){var e=document.getElementById('toast');e.textContent=m;e.className='toast '+(t||'ok');e.style.display='block';setTimeout(function(){e.style.display='none'},4000)}
function addLog(t){var b=document.getElementById('lb');var d=document.createElement('span');d.className='ll';
  if(t.indexOf('[步骤')>=0)d.className+=' s';else if(t.indexOf('完成')>=0||t.indexOf('✅')>=0)d.className+=' d';
  else if(t.indexOf('[ERR]')>=0||t.indexOf('错误')>=0)d.className+=' e';else if(t.indexOf('进度:')>=0)d.className+=' i';else d.className+=' m';
  d.textContent=t;b.appendChild(d);b.scrollTop=b.scrollHeight;}
function classify(l){
  if(!l)return;var s=l.toLowerCase();
  for(var i=0;i<9;i++){if(l.indexOf('[步骤 '+i+']')>=0){setStep(i,'active',0);return}}
  if(s.indexOf('生成纯文字')>=0)setStep(1,'active',0);
  else if(s.indexOf('完成:')>=0&&s.indexOf('docx')>=0)setStep(1,'done');
  else if(s.indexOf('增强 prd')>=0)setStep(2,'active',0);
  else if(s.indexOf('增强完成')>=0)setStep(2,'done');
  else if(s.indexOf('功能结构图')>=0&&s.indexOf('excel')>=0)setStep(3,'active',0);
  else if(s.indexOf('(7个一级')>=0)setStep(3,'done');
  else if(s.indexOf('png')>=0&&s.indexOf('图片')>=0)setStep(4,'active',0);
  else if(s.indexOf(')张')>=0)setStep(4,'done');
  else if(s.indexOf('时序图')>=0){if(s.indexOf('进度:')>=0){var m=l.match(/(\d+)\/(\d+)/);if(m)setStep(5,'active',Math.min(3,Math.floor(parseInt(m[1])/parseInt(m[2])*4)))}
    else if(s.indexOf('完成:')>=0)setStep(5,'done');else setStep(5,'active',0)}
  else if(s.indexOf('组装')>=0)setStep(6,'active',0);
  else if(s.indexOf('完整版.docx')>=0&&s.indexOf('[步骤')<0)setStep(6,'done');
  else if(s.indexOf('h5')>=0||s.indexOf('rewrite')>=0){setStep(7,'active',0);
    if(s.indexOf('extracted')>=0)setStep(7,'active',1);else if(s.indexOf('rewritten')>=0)setStep(7,'done')}
  else if(s.indexOf('统一样式')>=0)setStep(8,'active',0);
  else if(s.indexOf('规范化完成')>=0)setStep(8,'done');
}
async function startGen(){
  if(!file)return;
  document.getElementById('go').disabled=true;document.getElementById('gi').innerHTML='<span class="spin"></span>';
  document.getElementById('go').childNodes[1].textContent=' 生成中...';
  document.getElementById('pc').style.display='block';document.getElementById('lc').style.display='block';
  document.getElementById('lb').innerHTML='';document.getElementById('dl').disabled=true;initS();
  var fd=new FormData();fd.append('file',file);
  var r=await fetch('/api/start',{method:'POST',body:fd});var d=await r.json();
  if(d.error){toast(d.error,'er');reset();return}
  tid=d.task_id;addLog('任务启动: '+tid);poller=setInterval(doPoll,800);
}
function doPoll(){
  if(!tid)return;
  fetch('/api/status/'+tid).then(function(r){return r.json()}).then(function(d){
    if(d.logs)for(var i=0;i<d.logs.length;i++){addLog(d.logs[i]);classify(d.logs[i])}
    if(d.done){clearInterval(poller);
      if(d.success){for(var i=0;i<SD.length;i++)setStep(i,'done');
        document.getElementById('pf').style.width='100%';document.getElementById('pp').textContent='100%';
        document.getElementById('pl').textContent='✅ 完成';document.getElementById('dl').disabled=false;
        window.open('/api/download/'+tid);toast('生成完成, 已自动下载');
      }else{document.getElementById('pl').textContent='❌ 失败';toast('生成失败','er')}reset()}
  }).catch(function(){})}
function reset(){document.getElementById('go').disabled=!file;document.getElementById('gi').innerHTML='⚡';document.getElementById('go').childNodes[1].textContent=' 开始生成'}
document.addEventListener('DOMContentLoaded',function(){
  var dz=document.getElementById('dropZone'),fi=document.getElementById('fileInput');
  var gb=document.getElementById('go'),db=document.getElementById('dl'),tb=document.getElementById('tpl');
  var lh=document.getElementById('lh'),lb=document.getElementById('lb'),la=document.getElementById('la');
  if(dz){
    dz.addEventListener('click',function(){if(fi)fi.click()});
    dz.addEventListener('dragover',function(e){e.preventDefault();dz.classList.add('over')});
    dz.addEventListener('dragleave',function(){dz.classList.remove('over')});
    dz.addEventListener('drop',function(e){e.preventDefault();dz.classList.remove('over');
      if(e.dataTransfer.files[0])setF(e.dataTransfer.files[0])});
  }
  if(fi)fi.addEventListener('change',function(e){if(e.target.files[0])setF(e.target.files[0])});
  if(gb)gb.addEventListener('click',startGen);
  if(db)db.addEventListener('click',function(){if(tid)window.open('/api/download/'+tid)});
  if(tb)tb.addEventListener('click',function(){var a=document.createElement('a');a.href='/api/template';a.download='COSMIC功能点拆分_模板.xlsx';document.body.appendChild(a);a.click();a.remove();toast('模板下载中...')});
  if(lh)lh.addEventListener('click',function(){logOpen=!logOpen;lb.style.display=logOpen?'block':'none';la.style.transform=logOpen?'rotate(0)':'rotate(-90deg)'});
});
function setF(f){file=f;document.getElementById('go').disabled=false;
  document.getElementById('fb').style.display='flex';document.getElementById('fn').textContent=f.name;
  document.getElementById('fs').textContent=(f.size/1024).toFixed(1)+' KB';document.getElementById('dl').disabled=true}
</script>
</body>
</html>'''

@app.route('/')
def index():
    return render_template_string(HTML)

@app.route('/api/template')
def template():
    if os.path.exists(TEMPLATE_PATH):
        return send_file(TEMPLATE_PATH, as_attachment=True, download_name='COSMIC功能点拆分_模板.xlsx')
    return jsonify({"error": "模板文件不存在"}), 404

@app.route('/api/start', methods=['POST'])
def start():
    f = request.files.get('file')
    if not f or not f.filename.endswith('.xlsx'):
        return jsonify({"error": "请上传 .xlsx 文件"}), 400
    tid = uuid.uuid4().hex[:12]
    work_dir = os.path.join(tempfile.gettempdir(), f"cosmic_{tid}")
    os.makedirs(work_dir, exist_ok=True)
    xlsx_path = os.path.join(work_dir, f.filename)
    f.save(xlsx_path)
    out_dir = os.path.join(work_dir, "output")
    tasks[tid] = {"status": "running", "progress": 0, "logs": [], "step": "启动中",
                  "xlsx": xlsx_path, "out": out_dir, "result": None, "_sent_idx": 0}
    threading.Thread(target=run_pipeline, args=(tid, xlsx_path, out_dir), daemon=True).start()
    return jsonify({"task_id": tid})

def run_pipeline(tid, xlsx, out_dir):
    t = tasks[tid]
    try:
        t['logs'].append(f'📁 文件: {os.path.basename(xlsx)} ({os.path.getsize(xlsx)/1024:.1f} KB)')
        env = os.environ.copy()
        env['PYTHONIOENCODING'] = 'utf-8'
        proc = subprocess.Popen(
            [PYTHON, "-u", PIPELINE, xlsx, out_dir],
            stdout=subprocess.PIPE, stderr=subprocess.PIPE, env=env,
            creationflags=subprocess.CREATE_NO_WINDOW if sys.platform == 'win32' else 0)
        import concurrent.futures
        def read_pipe(pipe, is_err=False):
            for raw in pipe:
                try: line = raw.decode('utf-8', errors='replace').strip()
                except: line = raw.decode('gbk', errors='replace').strip()
                if line:
                    t['logs'].append(('[ERR] ' if is_err else '') + str(line))
        with concurrent.futures.ThreadPoolExecutor(max_workers=2) as ex:
            ex.submit(read_pipe, proc.stdout); ex.submit(read_pipe, proc.stderr, True)
            ex.shutdown(wait=True)
        proc.wait()
        if proc.returncode == 0:
            for f in os.listdir(out_dir):
                if f.endswith('_需求描述_完整版.docx'):
                    t['result'] = os.path.join(out_dir, f); break
            t['status'] = 'done'; t['success'] = bool(t['result']); t['progress'] = 100
        else:
            t['status'] = 'failed'; t['success'] = False
    except Exception as e:
        t['logs'].append(f'[ERR] {type(e).__name__}: {e}')
        t['status'] = 'failed'; t['success'] = False

@app.route('/api/status/<tid>')
def status(tid):
    t = tasks.get(tid)
    if not t: return jsonify({"error": "任务不存在"}), 404
    last = t.get('_sent_idx', 0)
    fresh = t['logs'][last:]
    t['_sent_idx'] = len(t['logs'])
    return jsonify({"done": t['status'] in ('done','failed'), "success": t.get('success',False),
                    "step": t.get('step',''), "pct": t.get('progress',0), "logs": fresh})

@app.route('/api/download/<tid>')
def download(tid):
    t = tasks.get(tid)
    if not t or not t.get('result'): return jsonify({"error": "文档不存在"}), 404
    return send_file(t['result'], as_attachment=True, download_name=os.path.basename(t['result']))

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=8080, debug=False)
