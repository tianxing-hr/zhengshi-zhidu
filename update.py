"""
天星公司制度 - 自动更新脚本
扫描 H:\正式制度\ 中的 .doc 文件，如有新增则重新生成 index.html 并推送到 GitHub Pages
"""
import olefile, re, os, json, base64, subprocess, sys
from datetime import datetime

DOC_DIR = r"H:\正式制度\人事制度"
OUT_DIR = r"H:\正式制度"
HTML_PATH = os.path.join(OUT_DIR, "index.html")
MANIFEST_PATH = os.path.join(OUT_DIR, ".docs_manifest.json")
REPO_DIR = OUT_DIR  # git repo is in H:\正式制度


def extract_text(filepath):
    """Extract text from WPS .doc (OLE2 format)"""
    ole = olefile.OleFileIO(filepath)
    data = ole.openstream('WordDocument').read()
    text = data.decode('utf-16-le', errors='ignore')
    # Remove binary garbage chars
    clean = re.sub(r'[^\u4e00-\u9fff\u3000-\u303f\uff00-\uffefa-zA-Z0-9\s\d.,;:!?()（）、。，；：！？""''…—\-\n\r《》【】第条〇℃％＋－×÷]', '', text)
    clean = re.sub(r'\s+', ' ', clean)
    # Remove leading garbage: find first meaningful content
    for starter in ['关于', '天星公司', '一、', '第一条']:
        idx = clean.find(starter)
        if 0 <= idx < 200:
            clean = clean[idx:]
            break
    # Remove trailing garbage: cut at PAGE/MERGEFORMAT markers
    for marker in ['PAGE', 'MERGEFORMAT']:
        idx = clean.find(marker)
        if idx > 200:
            clean = clean[:idx].strip()
            break
    # Also strip known binary start patterns
    for gs in ['粗噩', '醧敻', '椀嗸']:
        idx = clean.find(gs)
        if idx > 200:
            clean = clean[:idx].strip()
            break
    ole.close()
    return clean.strip()


def build_html(docs):
    """Build the complete HTML page from document list [(title, content), ...]"""
    sections = []
    toc = []
    for i, (title, content) in enumerate(docs):
        doc_id = f"doc{i}"
        paragraphs = content.split('。')
        formatted = ''.join(f'<p>{p.strip()}。</p>\n' for p in paragraphs if p.strip())
        
        toc.append(f'''
            <a href="#{doc_id}" class="toc-item" onclick="toggleDoc('{doc_id}')">
                <span class="toc-num">{i+1}</span>
                <span class="toc-title">{title}</span>
                <span class="toc-arrow">&#9662;</span>
            </a>''')
        sections.append(f'''
        <div class="doc-section" id="{doc_id}">
            <div class="doc-header" onclick="toggleDoc('{doc_id}')">
                <h2>{title}</h2>
                <span class="doc-arrow">&#9662;</span>
            </div>
            <div class="doc-body">{formatted}</div>
        </div>''')

    return f'''<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no">
<title>天星公司正式制度</title>
<style>
* {{ margin:0; padding:0; box-sizing:border-box; }}
body {{ font-family:-apple-system,BlinkMacSystemFont,"Segoe UI","PingFang SC","Microsoft YaHei",sans-serif; background:#f5f5f5; color:#333; font-size:15px; line-height:1.7; -webkit-text-size-adjust:100%; }}
.header {{ background:linear-gradient(135deg,#1a3a5c 0%,#2d6a9f 100%); color:white; padding:18px 16px; text-align:center; position:sticky; top:0; z-index:100; box-shadow:0 2px 8px rgba(0,0,0,0.15); }}
.header h1 {{ font-size:18px; font-weight:600; }}
.header .sub {{ font-size:12px; opacity:0.8; margin-top:4px; }}
.search-box {{ margin:12px 12px 0; position:relative; }}
.search-box input {{ width:100%; padding:10px 14px 10px 36px; border:1px solid #ddd; border-radius:8px; font-size:14px; background:white; outline:none; }}
.search-box input:focus {{ border-color:#2d6a9f; }}
.search-icon {{ position:absolute; left:12px; top:50%; transform:translateY(-50%); color:#999; font-size:16px; }}
.toc {{ margin:10px 12px; background:white; border-radius:10px; overflow:hidden; box-shadow:0 1px 3px rgba(0,0,0,0.08); }}
.toc-item {{ display:flex; align-items:center; padding:13px 14px; border-bottom:1px solid #f0f0f0; text-decoration:none; color:#333; font-size:14px; }}
.toc-item:last-child {{ border-bottom:none; }}
.toc-item:active {{ background:#f0f5fa; }}
.toc-num {{ width:24px; height:24px; background:#2d6a9f; color:white; border-radius:50%; display:flex; align-items:center; justify-content:center; font-size:12px; font-weight:600; margin-right:10px; flex-shrink:0; }}
.toc-title {{ flex:1; }}
.toc-arrow {{ color:#999; font-size:10px; }}
.doc-section {{ margin:10px 12px; background:white; border-radius:10px; overflow:hidden; box-shadow:0 1px 3px rgba(0,0,0,0.08); }}
.doc-header {{ padding:14px; background:#f8fafc; border-bottom:1px solid #eee; display:flex; justify-content:space-between; align-items:center; cursor:pointer; user-select:none; }}
.doc-header h2 {{ font-size:15px; font-weight:600; color:#1a3a5c; flex:1; }}
.doc-arrow {{ color:#999; font-size:12px; transition:transform 0.3s; margin-left:8px; }}
.doc-section.open .doc-arrow {{ transform:rotate(180deg); }}
.doc-body {{ padding:0 14px; max-height:0; overflow:hidden; transition:max-height 0.4s ease; }}
.doc-section.open .doc-body {{ max-height:20000px; padding:14px; }}
.doc-body p {{ margin-bottom:10px; text-indent:2em; }}
.doc-body p:last-child {{ margin-bottom:0; }}
.no-results {{ text-align:center; padding:30px; color:#999; font-size:14px; display:none; }}
.footer {{ text-align:center; padding:20px; color:#999; font-size:12px; }}
.highlight {{ background:#fff3b0; padding:1px 2px; border-radius:2px; }}
</style>
</head>
<body>
<div class="header"><h1>天星公司正式制度</h1><div class="sub">沈阳天星试验仪器股份有限公司</div></div>
<div class="search-box"><span class="search-icon">&#128269;</span><input type="text" id="search" placeholder="搜索制度内容..." oninput="doSearch()"></div>
<div class="toc" id="toc">{''.join(toc)}</div>
<div id="noResults" class="no-results">未找到匹配内容</div>
<div id="content">{''.join(sections)}</div>
<div class="footer">更新日期：{datetime.now().strftime('%Y年%m月%d日')} &nbsp;|&nbsp; 如有疑问请联系综合管理部</div>
<script>
function toggleDoc(d){{ var e=document.getElementById(d); e.classList.toggle('open'); }}
function doSearch(){{ var q=document.getElementById('search').value.trim().toLowerCase(); var s=document.querySelectorAll('.doc-section'); var t=document.getElementById('toc'); var n=document.getElementById('noResults'); var c=document.getElementById('content'); if(q.length===0){{ t.style.display=''; n.style.display='none'; c.style.display=''; s.forEach(function(e){{ e.style.display=''; e.classList.remove('open'); }}); document.querySelectorAll('.highlight').forEach(function(h){{ h.outerHTML=h.innerHTML; }}); return; }} t.style.display='none'; var f=0; s.forEach(function(e){{ var b=e.querySelector('.doc-body'); var tx=b.textContent.toLowerCase(); if(tx.indexOf(q)>=0){{ e.style.display=''; e.classList.add('open'); f++; var html=b.innerHTML; html=html.replace(/<span class="highlight">([^<]*)<\\/span>/g,'$1'); var rx=new RegExp('('+q.replace(/[.*+?^${{}}()|[\\]\\\\]/g,'\\\\$&')+')','gi'); html=html.replace(rx,'<span class="highlight">$1</span>'); b.innerHTML=html; }}else{{ e.style.display='none'; e.classList.remove('open'); }} }}); c.style.display=f>0?'':'none'; n.style.display=f>0?'none':''; if(f===0){{ n.innerHTML='未找到包含"<b>'+q+'</b>"的内容'; }} }}
document.addEventListener('DOMContentLoaded',function(){{ document.getElementById('doc0').classList.add('open'); }});
</script>
</body>
</html>'''


def main():
    # 1. Scan .doc files
    existing_docs = []
    for fname in os.listdir(DOC_DIR):
        if fname.lower().endswith('.doc') and not fname.startswith('~'):
            existing_docs.append(fname)

    # 2. Read manifest (previous state)
    prev_files = []
    if os.path.exists(MANIFEST_PATH):
        with open(MANIFEST_PATH, 'r', encoding='utf-8') as f:
            prev_files = json.load(f)

    # 3. Check for changes
    current_set = set(existing_docs)
    prev_set = set(prev_files)
    new_files = current_set - prev_set
    removed_files = prev_set - current_set

    if not new_files and not removed_files:
        print(f"[{datetime.now()}] 无变化，跳过更新")
        return

    if new_files:
        print(f"[{datetime.now()}] 发现新文件: {new_files}")
    if removed_files:
        print(f"[{datetime.now()}] 文件已移除: {removed_files}")

    # 4. Extract all docs
    print(f"[{datetime.now()}] 提取 {len(existing_docs)} 个文档...")
    docs = []
    for fname in existing_docs:
        fpath = os.path.join(DOC_DIR, fname)
        # Use filename as title (remove extension)
        title = os.path.splitext(fname)[0]
        content = extract_text(fpath)
        docs.append((title, content))
        print(f"  {fname}: {len(content)} chars")

    # 5. Build HTML
    print(f"[{datetime.now()}] 生成 HTML...")
    html = build_html(docs)
    with open(HTML_PATH, 'w', encoding='utf-8') as f:
        f.write(html)
    print(f"  HTML: {len(html)} bytes")

    # 6. Git commit & push
    print(f"[{datetime.now()}] 提交到 GitHub...")
    os.chdir(REPO_DIR)
    subprocess.run(['git', 'add', 'index.html'], check=True, capture_output=True)
    result = subprocess.run(['git', 'commit', '-m', 
        f'Auto update: {len(existing_docs)} docs, {datetime.now().strftime("%Y-%m-%d %H:%M")}'],
        capture_output=True, text=True)
    
    if 'nothing to commit' in result.stdout + result.stderr:
        print("  HTML 无实质变化，跳过推送")
    else:
        subprocess.run(['git', 'push', 'origin', 'master'], check=True, timeout=90)
        print("  推送成功！")

    # 7. Save manifest
    with open(MANIFEST_PATH, 'w', encoding='utf-8') as f:
        json.dump(existing_docs, f, ensure_ascii=False)

    print(f"[{datetime.now()}] 更新完成")
    print(f"  链接: https://tianxing-hr.github.io/zhengshi-zhidu/")


if __name__ == '__main__':
    main()
