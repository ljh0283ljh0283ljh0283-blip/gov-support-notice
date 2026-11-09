"""경기도 지원사업 공고 모아보기

경기도경제과학진흥원이 운영하는 '경기기업비서(이지비즈, egbiz.or.kr)'의
지원사업 목록을 가져와 분야별로 정리한 HTML 리포트와 엑셀용 CSV를 만듭니다.
지난번 실행 이후 새로 올라온 공고에는 NEW 표시를 붙이고, 원하면 메일로 보냅니다.

실행:  python gov_support_watch.py          (설정은 config.json)
"""
import csv
import html
import json
import os
import re
import smtplib
import sys
import urllib.parse
import urllib.request
from datetime import date, datetime, timedelta, timezone
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from pathlib import Path

BASE = Path(__file__).resolve().parent
LIST_API = "https://www.egbiz.or.kr/sp/selectSupportPrjListAjax.do"
DETAIL_URL = "https://www.egbiz.or.kr/sp/supportPrjDtl.do?listUrl=supportPrjCatList&bizCyclId={}"
GBSA_NAME = "경기도경제과학진흥원"
CATEGORY_ORDER = ["기술", "경영", "수출", "금융", "인력", "내수", "창업", "제도", "동반성장", "소상공", "기타"]
STATUS_KO = {"apply": "접수중", "ready": "준비중"}
PAGE_UNIT = 200
KST = timezone(timedelta(hours=9))
# 경기도 31개 시·군 (공고 제목에 이름이 있으면 그 지역 전용 공고로 본다)
SIGUN = ["가평", "고양", "과천", "광명", "광주", "구리", "군포", "김포", "남양주", "동두천", "부천", "성남",
         "수원", "시흥", "안산", "안성", "안양", "양주", "양평", "여주", "연천", "오산", "용인", "의왕",
         "의정부", "이천", "파주", "평택", "포천", "하남", "화성"]
SIGUN_RE = {s: re.compile(("(?<!남)" if s == "양주" else "") + re.escape(s)) for s in SIGUN}


# ---------------------------------------------------------------- 수집
def fetch_all():
    """접수중·준비중 공고 전체를 가져온다."""
    items, page = [], 0
    while True:
        body = urllib.parse.urlencode({
            "pageIndex": page, "bizNm": "", "areaCode": "", "categoryId": "",
            "sortCd": "bizCyclId", "prjStatus": "all", "gginsttYn": "Y",
            "part": "category", "pageUnit": PAGE_UNIT,
        }).encode()
        req = urllib.request.Request(LIST_API, data=body, headers={
            "User-Agent": "Mozilla/5.0 (gov-support-watch)",
            "X-Requested-With": "XMLHttpRequest",
            "Content-Type": "application/x-www-form-urlencoded; charset=UTF-8",
        })
        with urllib.request.urlopen(req, timeout=60) as r:
            batch = json.loads(r.read().decode("utf-8")).get("value") or []
        items += batch
        if len(batch) < PAGE_UNIT or page > 50:
            break
        page += 1

    out, seen_ids = [], set()
    for x in items:
        if x.get("prjStatus") not in STATUS_KO or x["bizCyclId"] in seen_ids:
            continue
        seen_ids.add(x["bizCyclId"])
        out.append({
            "id": x["bizCyclId"],
            "title": (x.get("bizNm") or "").strip(),
            "category": x.get("categoryNm") or "기타",
            "org": x.get("outsdInstNm") or "",
            "status": STATUS_KO[x["prjStatus"]],
            "start": x.get("aplyBgngDt") or "",
            "end": x.get("aplyEndDt") or "",
            "until_budget": x.get("bgtExhRcptYn") == "Y",
            "views": x.get("pbancInqCnt") or 0,
            "url": DETAIL_URL.format(x["bizCyclId"]),
        })
    return out


# ---------------------------------------------------------------- 정리
def d_day(item, today):
    if item["until_budget"] or not item["end"]:
        return None
    try:
        return (datetime.strptime(item["end"], "%Y-%m-%d").date() - today).days
    except ValueError:
        return None


def organize(items, cfg, seen, today):
    cats = set(cfg.get("관심분야") or CATEGORY_ORDER)
    excl = [k.lower() for k in cfg.get("제외키워드", []) if k]
    kws = [k for k in cfg.get("관심키워드", []) if k]
    first_run = not seen

    kept = []
    for it in items:
        t = it["title"].lower()
        if it["category"] in CATEGORY_ORDER and it["category"] not in cats:
            continue
        if cfg.get("진흥원_공고만") and GBSA_NAME not in it["org"]:
            continue
        if any(k in t for k in excl):
            continue
        it["dday"] = d_day(it, today)
        if it["dday"] is not None and it["dday"] < 0:
            continue
        it["keywords"] = [k for k in kws if k.lower() in t]
        it["gbsa"] = GBSA_NAME in it["org"]
        it["regions"] = [s for s, rx in SIGUN_RE.items() if rx.search(it["title"])]
        if first_run:  # 첫 실행: 최근 7일 안에 접수 시작한 공고만 NEW로 본다
            it["new"] = it["start"] >= (today - timedelta(days=7)).isoformat()
        else:
            it["new"] = it["id"] not in seen
        kept.append(it)

    kept.sort(key=lambda i: (not i["new"], i["dday"] if i["dday"] is not None else 9999))
    return kept


# ---------------------------------------------------------------- 리포트
def esc(s):
    return html.escape(str(s))


def dday_badge(it, soon):
    if it["until_budget"]:
        return '<span class="b gray">예산소진시</span>'
    if it["dday"] is None:
        return ""
    cls = "red" if it["dday"] <= soon else "gray"
    label = "오늘마감" if it["dday"] == 0 else f"D-{it['dday']}"
    return f'<span class="b {cls}">{label}</span>'


def is_soon(it, soon):
    return it["dday"] is not None and it["dday"] <= soon


def in_region(it, region):
    """지역 탭에 보일 공고: 특정 시·군이 적히지 않은 경기 공통 공고 + 그 시·군 공고"""
    return not it["regions"] or region in it["regions"]


def row(it, soon):
    period = "예산 소진 시까지" if it["until_budget"] else f'{esc(it["start"])} ~ {esc(it["end"])}'
    tags = "".join(f'<span class="kw">{esc(k)}</span>' for k in it["keywords"])
    new = '<span class="b new">NEW</span>' if it["new"] else ""
    gbsa = '<span class="b gbsa">진흥원</span>' if it["gbsa"] else ""
    reg = "".join(f'<span class="reg">{esc(r)}</span>' for r in it["regions"])
    return (f'<tr data-cat="{esc(it["category"])}" data-new="{int(it["new"])}" data-reg="{esc(",".join(it["regions"]))}"'
            f' data-soon="{int(is_soon(it, soon))}" data-kw="{int(bool(it["keywords"]))}">'
            f'<td class="t"><a href="{esc(it["url"])}" target="_blank">{esc(it["title"])}</a>'
            f'<div class="meta">{new}{gbsa}{reg}<span class="cat">{esc(it["category"])}</span>{tags}</div></td>'
            f'<td class="org">{esc(it["org"])}</td><td class="per">{period}</td>'
            f'<td class="dd">{dday_badge(it, soon)}</td></tr>')


def table(items, soon, empty="해당 공고가 없습니다."):
    tbl = ('<div class="tw"><table><thead><tr><th>사업명</th><th>기관</th><th>접수기간</th><th>마감</th></tr></thead><tbody>'
           + "".join(row(i, soon) for i in items) + "</tbody></table></div>") if items else ""
    return f'<div class="blk">{tbl}<p class="empty"{" hidden" if items else ""}>{empty}</p></div>'


def build_html(items, cfg, today, for_mail=False, share=False, stamp=None, live_tab=False):
    soon = int(cfg.get("마감임박_일수", 7))
    new = [i for i in items if i["new"]]
    matched = [i for i in items if i["keywords"]]
    closing = sorted([i for i in items if i["dday"] is not None and i["dday"] <= soon], key=lambda i: i["dday"])
    by_cat = {}
    for i in items:
        by_cat.setdefault(i["category"], []).append(i)
    cats = [c for c in CATEGORY_ORDER if c in by_cat] + [c for c in by_cat if c not in CATEGORY_ORDER]

    scope = "경기도경제과학진흥원 직접 공고" if cfg.get("진흥원_공고만") else "경기기업비서(이지비즈) 분야별지원사업"
    stats = "".join(f'<div class="stat"><b id="{sid}">{n}</b><span>{l}</span></div>' for sid, n, l in
                    [("s-total", len(items), "접수중·준비중"), ("s-new", len(new), "새 공고"),
                     ("s-soon", len(closing), f"{soon}일 내 마감"), ("s-kw", len(matched), "맞춤 키워드")])
    regions = [r for r in cfg.get("지역탭", []) if r in SIGUN]
    rtabs = ""
    if not for_mail:
        quick = "".join(f'<button type="button" class="kwq" data-q="{esc(k)}">{esc(k)}</button>'
                        for k in cfg.get("관심키워드", []) if k)
        rtabs = ('<div class="rtabs" role="tablist" aria-label="보기">'
                 f'<button class="rtab on" data-r="*" role="tab" aria-selected="true">경기 전체 <span>{len(items)}</span></button>'
                 + "".join(f'<button class="rtab" data-r="{esc(r)}" role="tab" aria-selected="false">{esc(r)} '
                           f'<span>{sum(in_region(i, r) for i in items)}</span></button>' for r in regions)
                 + '<button class="rtab" data-r="kw" role="tab" aria-selected="false">키워드 검색</button>'
                 + ('<button class="rtab live" data-live="1" type="button" title="이지비즈에서 지금 공고를 다시 가져옵니다">'
                    '최신 정보 불러오기</button>' if live_tab else '')
                 + '<p class="rnote" id="rnote">지역 탭은 경기도 공통 공고와 그 시·군 이름이 들어간 공고를 보여주고, '
                   '다른 시·군 전용 공고는 뺍니다.</p></div>'
                 + '<div class="kwbox" id="kwbox" hidden>'
                   '<label for="kwq">찾을 단어</label>'
                   '<input id="kwq" type="search" autocomplete="off" placeholder="예: 수출지원, 스마트공장">'
                   f'<div class="kwquick"><span>자주 찾는 단어</span>{quick}</div>'
                   '<p class="rnote" id="kwres">사업명·분야·기관에서 찾습니다. \'수출지원\'처럼 끝에 붙은 지원·사업은 떼고 찾고, '
                   '띄어 쓰면 모든 단어가 들어간 공고, 쉼표(,)로 나누면 하나라도 들어간 공고를 보여줍니다.</p></div>')

    if for_mail:  # 메일 본문은 새 공고만 분야별로 짧게
        body = f'<div class="stats">{stats}</div>'
        new_by_cat = {}
        for i in new:
            new_by_cat.setdefault(i["category"], []).append(i)
        if not new:
            body += '<p class="empty">새로 올라온 공고가 없습니다.</p>'
        for c in [c for c in cats if c in new_by_cat]:
            body += f'<h2>{esc(c)} <small>{len(new_by_cat[c])}건</small></h2>' + table(new_by_cat[c], soon)
        body += '<p class="foot">전체 목록은 첨부한 리포트 파일을 열어 보세요.</p>'
    else:
        chips = f'<button class="chip on" data-f="*">전체 <span class="n">{len(items)}</span></button>' + "".join(
            f'<button class="chip" data-f="{esc(c)}">{esc(c)} <span class="n">{len(by_cat[c])}</span></button>' for c in cats)
        body = (f'{rtabs}<div class="stats">{stats}</div>'
                f'<h2>우리 회사 맞춤 <small>관심 키워드: {esc(", ".join(cfg.get("관심키워드", [])))}</small></h2>'
                + table(matched, soon, "관심 키워드에 맞는 공고가 없습니다.")
                + f'<h2>마감 임박 <small>{soon}일 이내</small></h2>' + table(closing, soon)
                + f'<h2>분야별 전체</h2><div class="chips">{chips}'
                  f'<label class="onlynew"><input type="checkbox" id="onlynew"> 새 공고만</label></div>'
                + '<div id="all">' + "".join(
                    f'<section data-cat="{esc(c)}"><h3>{esc(c)} <small><span class="n">{len(by_cat[c])}</span>건</small></h3>'
                    + table(by_cat[c], soon) + "</section>" for c in cats)
                + '<p class="empty" id="none" hidden>조건에 맞는 공고가 없습니다.</p></div>')

    script = "" if for_mail else """<script>
const $$=s=>[...document.querySelectorAll(s)];
$$('.rtab[data-live]').forEach(b=>b.onclick=()=>{b.textContent='불러오는 중…';b.disabled=true;
  location.href=location.pathname+'?live='+Date.now();});
const tabs=$$('.rtab:not([data-live])'),chips=$$('.chip'),nb=document.getElementById('onlynew');
let reg='*',cat='*';
const kwq=document.getElementById('kwq'),kwbox=document.getElementById('kwbox'),
  kwres=document.getElementById('kwres'),rnote=document.getElementById('rnote'),kwHelp=kwres?kwres.textContent:'';
const norm=s=>s.replace(/\\s+/g,'').toLowerCase();
$$('tbody tr').forEach(r=>{r.dataset.t=norm(r.querySelector('td.t a').textContent+' '+r.dataset.cat+' '+r.querySelector('td.org').textContent);});
/* 쉼표 = 또는, 띄어쓰기 = 그리고. '수출지원'처럼 끝에 붙은 지원·지원사업·사업은 떼고 찾는다 */
const stem=w=>{const s=w.replace(/(지원사업|지원|사업)$/,'');return s.length>=2?s:w;};
const terms=()=>kwq?kwq.value.split(',').map(g=>g.trim().split(/\\s+/).map(w=>stem(norm(w))).filter(Boolean)).filter(g=>g.length):[];
const inReg=r=>{
  if(reg==='kw'){const t=terms();return !t.length||t.some(g=>g.every(x=>r.dataset.t.includes(x)));}
  return reg==='*'||!r.dataset.reg||r.dataset.reg.split(',').includes(reg);};
const setN=(id,n)=>{const e=document.getElementById(id);if(e)e.textContent=n;};
function apply(){
  $$('.blk').forEach(b=>{if(b.closest('#all'))return;let v=0;
    b.querySelectorAll('tbody tr').forEach(r=>{const ok=inReg(r);r.hidden=!ok;if(ok)v++;});
    const tw=b.querySelector('.tw');if(tw)tw.hidden=!v;b.querySelector('.empty').hidden=!!v;});
  const rows=$$('#all tbody tr').filter(inReg);
  if(kwres){const t=kwq.value.trim();kwres.textContent=t?`'${t}' 검색 결과 ${rows.length}건`:kwHelp;}
  try{localStorage.setItem('kwq',kwq.value);}catch(e){}
  setN('s-total',rows.length);setN('s-new',rows.filter(r=>r.dataset.new==='1').length);
  setN('s-soon',rows.filter(r=>r.dataset.soon==='1').length);setN('s-kw',rows.filter(r=>r.dataset.kw==='1').length);
  chips.forEach(c=>{const n=c.querySelector('.n');if(n)n.textContent=c.dataset.f==='*'?rows.length:rows.filter(r=>r.dataset.cat===c.dataset.f).length;});
  $$('#all section').forEach(s=>{let v=0;
    s.querySelectorAll('tbody tr').forEach(r=>{const ok=inReg(r)&&(!nb.checked||r.dataset.new==='1');r.hidden=!ok;if(ok)v++;});
    s.querySelector('h3 .n').textContent=v;s.hidden=!((cat==='*'||s.dataset.cat===cat)&&v);});
  const none=document.getElementById('none');if(none)none.hidden=$$('#all section').some(s=>!s.hidden);
}
tabs.forEach(t=>t.onclick=()=>{tabs.forEach(x=>{x.classList.remove('on');x.setAttribute('aria-selected','false');});
  t.classList.add('on');t.setAttribute('aria-selected','true');reg=t.dataset.r;
  if(kwbox){kwbox.hidden=reg!=='kw';rnote.hidden=reg==='kw';if(reg==='kw')kwq.focus();}
  apply();});
if(kwq){try{kwq.value=localStorage.getItem('kwq')||'';}catch(e){}
  kwq.addEventListener('input',apply);
  $$('.kwq').forEach(b=>b.onclick=()=>{kwq.value=b.dataset.q;apply();kwq.focus();});}
chips.forEach(c=>c.onclick=()=>{chips.forEach(x=>x.classList.remove('on'));c.classList.add('on');cat=c.dataset.f;apply();});
nb.onchange=apply;
</script>"""

    css = """
/* 공고 게시판 느낌: 차분한 청록 잉크 + 행정문서 표 */
:root{--bg:#f4f6f5;--panel:#ffffff;--line:#dfe5e3;--fg:#1c2624;--muted:#5d6b68;--accent:#0f6b5c;
 --accent-fg:#ffffff;--soft:#e3f0ec;--warn:#b42318;--warn-fg:#ffffff;--kw-bg:#fdf1d3;--kw-fg:#6f4e00;
 --gbsa:#4b4fa8;--chip-bg:#ffffff;--gray-bg:#e9edec}
@media (prefers-color-scheme: dark){:root:not([data-theme="light"]){--bg:#121816;--panel:#1a2220;--line:#2c3734;
 --fg:#e4ebe9;--muted:#9aa8a5;--accent:#4cc2a8;--accent-fg:#0d1513;--soft:#1f3631;--warn:#f0705f;--warn-fg:#1a0d0b;
 --kw-bg:#3a3018;--kw-fg:#f2d48b;--gbsa:#9da1ef;--chip-bg:#1a2220;--gray-bg:#26302d;color-scheme:dark}}
:root[data-theme="dark"]{--bg:#121816;--panel:#1a2220;--line:#2c3734;--fg:#e4ebe9;--muted:#9aa8a5;--accent:#4cc2a8;
 --accent-fg:#0d1513;--soft:#1f3631;--warn:#f0705f;--warn-fg:#1a0d0b;--kw-bg:#3a3018;--kw-fg:#f2d48b;--gbsa:#9da1ef;
 --chip-bg:#1a2220;--gray-bg:#26302d;color-scheme:dark}
body{font-family:'Pretendard','Malgun Gothic','Apple SD Gothic Neo',sans-serif;background:var(--bg);color:var(--fg);
 margin:0;padding-block:24px 40px;padding-left:max(16px,env(safe-area-inset-left,0px));padding-right:max(16px,env(safe-area-inset-right,0px))}
.wrap{max-width:1080px;margin:0 auto}
h1{font-size:24px;margin:0 0 6px;text-wrap:balance} .sub{color:var(--muted);font-size:13px;margin-bottom:18px}
h2{font-size:17px;margin:30px 0 10px;border-left:4px solid var(--accent);padding-left:8px;text-wrap:balance}
h3{font-size:15px;margin:18px 0 8px} small{color:var(--muted);font-weight:normal;font-size:12px}
.stats{display:grid;grid-template-columns:repeat(auto-fit,minmax(130px,1fr));gap:10px}
.stat{background:var(--panel);border:1px solid var(--line);border-radius:10px;padding:12px 16px}
.stat b{display:block;font-size:26px;color:var(--accent);font-variant-numeric:tabular-nums} .stat span{font-size:12px;color:var(--muted)}
.tw{overflow-x:auto;border:1px solid var(--line);border-radius:8px;background:var(--panel)}
table{width:100%;border-collapse:collapse;font-size:13px}
th{background:var(--soft);text-align:left;padding:8px;font-weight:600;color:var(--fg)}
td{padding:9px 8px;border-top:1px solid var(--line);vertical-align:top}
td.t a{color:var(--fg);text-decoration:none;font-weight:600} td.t a:hover,td.t a:focus-visible{color:var(--accent);text-decoration:underline}
td.org{color:var(--muted);width:200px} td.per{white-space:nowrap;color:var(--muted);width:170px;font-variant-numeric:tabular-nums}
td.dd{width:80px;text-align:center}
.meta{margin-top:4px;display:flex;gap:4px;flex-wrap:wrap}
.b,.kw,.cat{display:inline-block;font-size:11px;padding:1px 6px;border-radius:4px}
.b.new{background:var(--accent);color:var(--accent-fg)} .b.gbsa{background:var(--gbsa);color:var(--accent-fg)}
.b.red{background:var(--warn);color:var(--warn-fg)} .b.gray{background:var(--gray-bg);color:var(--muted)}
.kw{background:var(--kw-bg);color:var(--kw-fg)} .cat{background:var(--soft);color:var(--accent)}
.chips{display:flex;gap:6px;flex-wrap:wrap;align-items:center;margin-bottom:8px}
.chip{border:1px solid var(--line);background:var(--chip-bg);color:var(--fg);border-radius:16px;padding:5px 12px;cursor:pointer;font-size:13px;font-family:inherit}
.chip.on{background:var(--accent);color:var(--accent-fg);border-color:var(--accent)}
.chip:focus-visible{outline:2px solid var(--accent);outline-offset:2px}
.onlynew{margin-left:auto;font-size:13px}
[hidden]{display:none!important}
.rtabs{display:flex;flex-wrap:wrap;gap:0;align-items:flex-end;border-bottom:2px solid var(--line);margin-bottom:16px}
.rtab{border:0;background:none;color:var(--muted);font:inherit;font-size:15px;font-weight:600;padding:8px 18px;cursor:pointer;
 border-bottom:3px solid transparent;margin-bottom:-2px}
.rtab span{font-weight:normal;font-size:12px;font-variant-numeric:tabular-nums}
.rtab.on{color:var(--accent);border-bottom-color:var(--accent)}
.rtab.live{margin-left:auto;align-self:center;margin-bottom:6px;border:1px solid var(--accent);border-radius:16px;
 padding:5px 14px;font-size:13px;color:var(--accent);background:var(--panel)}
.rtab.live:hover{background:var(--accent);color:var(--accent-fg)} .rtab.live:disabled{opacity:.7;cursor:wait}
.rtab:focus-visible{outline:2px solid var(--accent);outline-offset:-2px}
.rnote{flex-basis:100%;margin:6px 0 8px;font-size:12px;color:var(--muted)}
.kwbox{display:flex;flex-direction:column;gap:8px;background:var(--panel);border:1px solid var(--line);border-radius:10px;
 padding:14px 16px;margin-bottom:16px}
.kwbox label{font-size:13px;font-weight:600}
.kwbox input{font:inherit;font-size:15px;padding:9px 12px;border:1px solid var(--line);border-radius:8px;
 background:var(--bg);color:var(--fg);width:100%;max-width:480px;box-sizing:border-box}
.kwbox input:focus{outline:2px solid var(--accent);outline-offset:1px}
.kwquick{display:flex;flex-wrap:wrap;gap:6px;align-items:center;font-size:12px;color:var(--muted)}
.kwq{border:1px solid var(--line);background:var(--kw-bg);color:var(--kw-fg);border-radius:14px;padding:3px 10px;
 font:inherit;font-size:12px;cursor:pointer}
.kwq:focus-visible{outline:2px solid var(--accent);outline-offset:2px}
.kwbox .rnote{margin:0}
.reg{display:inline-block;font-size:11px;padding:1px 6px;border-radius:4px;border:1px solid var(--line);color:var(--fg)}
.empty{color:var(--muted);font-size:13px;background:var(--panel);border:1px dashed var(--line);padding:12px;border-radius:8px}
.foot{color:var(--muted);font-size:12px;margin-top:20px}
@media(max-width:700px){td.org,td.per,th:nth-child(2),th:nth-child(3){display:none}}
"""
    inner = f"""<title>경기 지원사업 공고</title>
<meta name="claude:viewport-fit-sides" content="cover">
<style>{css}</style><div class="wrap">
<h1>경기 지원사업 공고 모아보기</h1>
<div class="sub">{esc(stamp or today.strftime('%Y년 %m월 %d일'))} 기준 · 출처: {esc(scope)} · NEW = 지난번 확인 이후 새로 올라온 공고</div>
{body}
<p class="foot">사업명을 누르면 이지비즈 공고 상세 페이지로 이동합니다.</p>
</div>{script}"""
    if share:  # claude.ai 공유 페이지용: 문서 골격은 게시할 때 자동으로 붙는다
        return inner
    return f'<!doctype html><html lang="ko"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"></head><body>{inner}</body></html>'


def write_csv(items, path):
    with open(path, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        w.writerow(["새공고", "분야", "지역", "사업명", "기관", "상태", "접수시작", "접수마감", "남은일수", "맞춤키워드", "링크"])
        for i in items:
            w.writerow(["NEW" if i["new"] else "", i["category"], ", ".join(i["regions"]) or "경기 공통", i["title"], i["org"], i["status"],
                        i["start"], "예산소진시" if i["until_budget"] else i["end"],
                        "" if i["dday"] is None else i["dday"], ", ".join(i["keywords"]), i["url"]])


# ---------------------------------------------------------------- 메일
def send_mail(cfg, items, today, attach_path):
    m = cfg.get("메일", {})
    new_count = sum(i["new"] for i in items)
    if m.get("새공고_없으면_안보냄", True) and new_count == 0:
        print("새 공고가 없어 메일은 보내지 않았습니다.")
        return
    pw = os.environ.get("GOV_MAIL_PASSWORD")
    to = [a for a in m.get("받는사람", []) if a]
    if not (pw and m.get("보내는사람") and to):
        print("메일 설정(보내는사람/받는사람/환경변수 GOV_MAIL_PASSWORD)이 비어 있어 메일은 건너뜁니다.")
        return
    msg = MIMEMultipart()
    msg["Subject"] = f"[지원사업] {today:%m/%d} 새 공고 {new_count}건"
    msg["From"], msg["To"] = m["보내는사람"], ", ".join(to)
    msg.attach(MIMEText(build_html(items, cfg, today, for_mail=True), "html", "utf-8"))
    part = MIMEApplication(attach_path.read_bytes(), Name=attach_path.name)
    part["Content-Disposition"] = f'attachment; filename="{attach_path.name}"'
    msg.attach(part)
    with smtplib.SMTP(m.get("smtp서버", "smtp.gmail.com"), int(m.get("포트", 587)), timeout=60) as s:
        s.starttls()
        s.login(m["보내는사람"], pw)
        s.send_message(msg)
    print(f"메일 발송 완료 → {', '.join(to)}")


# ---------------------------------------------------------------- 실시간 페이지 (Vercel)
def render_live():
    """이지비즈에서 지금 공고를 가져와 웹페이지 HTML을 만든다. 기록 파일은 읽기만 한다."""
    cfg = json.loads((BASE / "config.json").read_text(encoding="utf-8"))
    seen_path = BASE / "data" / "seen.json"
    seen = json.loads(seen_path.read_text(encoding="utf-8")) if seen_path.exists() else {}
    now = datetime.now(KST)
    items = organize(fetch_all(), cfg, seen, now.date())
    return build_html(items, cfg, now.date(), stamp=now.strftime("%Y년 %m월 %d일 %H:%M"), live_tab=True)


# ---------------------------------------------------------------- main
def main():
    cfg = json.loads((BASE / "config.json").read_text(encoding="utf-8"))
    seen_path = BASE / "data" / "seen.json"
    seen_path.parent.mkdir(exist_ok=True)
    seen = json.loads(seen_path.read_text(encoding="utf-8")) if seen_path.exists() else {}
    today = datetime.now(KST).date()

    raw = fetch_all()
    items = organize(raw, cfg, seen, today)

    out = BASE / "reports"
    out.mkdir(exist_ok=True)
    report = out / f"지원사업_{today:%Y%m%d}.html"
    report.write_text(build_html(items, cfg, today), encoding="utf-8")
    (out / "최신_지원사업.html").write_text(build_html(items, cfg, today), encoding="utf-8")
    write_csv(items, out / f"지원사업_{today:%Y%m%d}.csv")
    (BASE / "share").mkdir(exist_ok=True)  # 공유 페이지(claude.ai) 게시용
    (BASE / "share" / "gov-support.html").write_text(build_html(items, cfg, today, share=True), encoding="utf-8")
    docs = BASE / "docs"  # GitHub Pages 웹사이트용
    docs.mkdir(exist_ok=True)
    (docs / "index.html").write_text(build_html(items, cfg, today), encoding="utf-8")
    write_csv(items, docs / "list.csv")
    (docs / ".nojekyll").write_text("", encoding="utf-8")

    for it in raw:  # 필터와 상관없이 본 공고는 모두 기록 (다음 실행의 NEW 판단용)
        seen.setdefault(it["id"], today.isoformat())
    seen_path.write_text(json.dumps(seen, ensure_ascii=False, indent=0), encoding="utf-8")

    new_count = sum(i["new"] for i in items)
    print(f"수집 {len(raw)}건 → 조건에 맞는 공고 {len(items)}건 (새 공고 {new_count}건)")
    print(f"리포트: {report}")
    if cfg.get("메일", {}).get("보내기"):
        send_mail(cfg, items, today, report)
    if "--open" in sys.argv and hasattr(os, "startfile"):
        os.startfile(out / "최신_지원사업.html")


if __name__ == "__main__":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass
    main()
