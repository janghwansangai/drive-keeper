"""사용 안내 영상용: 구글 화면 예시 그림(따라 하기 페이지의 그림)과 홈페이지 화면을 고화질로 찍는다.

실행: .venv\\Scripts\\python tools\\video_pages.py   (PowerShell에서. Bash에서는 Edge가 찍지 못함)
결과: build/video/caps/m_*.png, page_*.png, page_rects.json
"""
from __future__ import annotations

import json
import re
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
DOCS = ROOT / "docs"
WORK = ROOT / "build" / "video" / "pages"
OUT = ROOT / "build" / "video" / "caps"
EDGE = next(p for p in (r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
                        r"C:\Program Files\Microsoft\Edge\Application\msedge.exe") if Path(p).exists())
SIZE = (1874, 922)
# (이름, 따라 하기 단계, 그 단계의 몇 번째 그림)
MOCKS = [("m_download", "a1", 0), ("m_extract", "a2", 1), ("m_smart", "a3", 1), ("m_desk", "a4", 1),
         ("m_b1", "b1", 0), ("m_b2", "b2", 0), ("m_b3", "b3", 0), ("m_b4", "b4", 0), ("m_b5", "b5", 0), ("m_b6", "b6", 0),
         ("m_choose", "c2", 0), ("m_warn", "c3", 0), ("m_allow", "c4", 0), ("m_done", "c4", 1),
         ("m_e1", "e1", 0), ("m_other", "e2", 0)]

MOCK_JS = """
<script>
(function(){
  const q=new URLSearchParams(location.search), sel=q.get('s'), idx=+q.get('i')||0;
  const el=document.querySelectorAll('#'+sel+' .m')[idx];
  document.body.innerHTML='';
  const st=document.createElement('style');
  st.textContent='html,body{height:100%;margin:0;background:#EEF2F9;overflow:hidden}'
   +'#wrap{position:absolute;container-type:inline-size}'
   +'.tap{animation:none!important;box-shadow:0 0 0 4px #F97316!important}';
  document.head.append(st);
  const w=document.createElement('div'); w.id='wrap'; w.append(el); document.body.append(w);
  let fs=sel==='a4'?30:27, width=(el.classList.contains('msgbox')||el.classList.contains('smart'))?1050:1500;
  w.style.width=width+'px';
  for(;fs>12;fs-=1){ el.style.setProperty('font-size',fs+'px','important'); if(w.offsetHeight<=800&&el.offsetHeight<=800) break; }
  w.style.left=Math.round((innerWidth-w.offsetWidth)/2)+'px';
  w.style.top=Math.round((innerHeight-w.offsetHeight)/2)+'px';
  const rects=[...document.querySelectorAll('.tap')].map(t=>{const r=t.getBoundingClientRect();
    return [Math.round(r.x),Math.round(r.y),Math.round(r.width),Math.round(r.height),t.dataset.tag||''];});
  const p=document.createElement('pre'); p.id='out'; p.hidden=true;
  p.textContent=JSON.stringify({rects,fs,vw:innerWidth,vh:innerHeight}); document.body.append(p);
})();
</script>
"""
PAGE_JS = """
<script>
window.addEventListener('load',function(){setTimeout(function(){
  const out={vw:innerWidth,vh:innerHeight,rects:{}};
  document.querySelectorAll('[data-rect]').forEach(e=>{const r=e.getBoundingClientRect();
    out.rects[e.dataset.rect]=[Math.round(r.x),Math.round(r.y),Math.round(r.width),Math.round(r.height)];});
  const p=document.createElement('pre'); p.id='out'; p.hidden=true; p.textContent=JSON.stringify(out); document.body.append(p);
},300);});
</script>
"""


def run(url: str, shot: Path | None) -> str:
    prof = ROOT / "build" / "edge_profile"
    args = [EDGE, "--headless", "--disable-gpu", "--no-first-run", f"--user-data-dir={prof}", "--hide-scrollbars",
            "--force-device-scale-factor=1", f"--window-size={SIZE[0]},{SIZE[1]}", "--virtual-time-budget=7000"]
    args += [f"--screenshot={shot}"] if shot else ["--dump-dom"]
    r = subprocess.run(args + [url], capture_output=True, timeout=120)
    return r.stdout.decode("utf-8", "replace")


def data_of(dom: str) -> dict:
    m = re.search(r'<pre id="out"[^>]*>(.*?)</pre>', dom, re.S)
    return json.loads(m.group(1).replace("&amp;", "&").replace("&lt;", "<").replace("&gt;", ">")) if m else {}


def main() -> int:
    WORK.mkdir(parents=True, exist_ok=True)
    OUT.mkdir(parents=True, exist_ok=True)
    base = f'<base href="{DOCS.as_uri()}/">'
    guide = (DOCS / "start.html").read_text(encoding="utf-8")
    html = re.sub(r'<script src="assets/shots/spots.js"></script>\s*<script>.*?</script>', "", guide, flags=re.S)
    html = re.sub(r'<script src="https://script\.google\.com[^>]*></script>', "", html)   # 오류 신고 단추 빼기
    html = html.replace("<head>", "<head>" + base, 1).replace("</body>", MOCK_JS + "</body>")
    page = WORK / "mock.html"
    page.write_text(html, encoding="utf-8")
    rects: dict = {}
    for name, step, i in MOCKS:
        url = f"{page.as_uri()}?s={step}&i={i}"
        d = data_of(run(url, None))
        run(url, OUT / f"{name}.png")
        rects[name] = d
        print(name, d.get("fs"), d.get("vw"), d.get("vh"), len(d.get("rects", [])), flush=True)

    home = (DOCS / "index.html").read_text(encoding="utf-8")
    home = home.replace("<head>", "<head>" + base, 1).replace('id="dl-hero"', 'id="dl-hero" data-rect="download"', 1)
    home = home.replace("</body>", PAGE_JS + "</body>")
    (WORK / "home.html").write_text(home, encoding="utf-8")
    d = data_of(run((WORK / "home.html").as_uri(), None))
    run((WORK / "home.html").as_uri(), OUT / "page_home.png")
    rects["page_home"] = d
    print("home", d, flush=True)
    g = guide.replace("<head>", "<head>" + base, 1).replace("</body>", PAGE_JS + "</body>")
    (WORK / "guide.html").write_text(g, encoding="utf-8")
    run((WORK / "guide.html").as_uri(), OUT / "page_guide.png")
    (OUT / "page_rects.json").write_text(json.dumps(rects, ensure_ascii=False), encoding="utf-8")
    return 0


if __name__ == "__main__":
    sys.exit(main())
