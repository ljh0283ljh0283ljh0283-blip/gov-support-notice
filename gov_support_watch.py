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
# 경기도 밖 시·도 (제목이나 기관 이름에 있으면 그 지역 공고로 본다)
# '부산물', '대전환' 같은 낱말과 경기도 광주시는 다른 지역으로 잘못 보지 않게 걸러낸다
PROVINCE_RE = {
    "서울": re.compile(r"서울"), "인천": re.compile(r"인천"), "부산": re.compile(r"부산(?!물)"),
    "대구": re.compile(r"대구(?!경)"), "대전": re.compile(r"대전(?!환)"), "울산": re.compile(r"울산"),
    "세종": re.compile(r"세종(?!대왕|학당)"), "광주광역": re.compile(r"광주광역|\[광주\]|광주(?:테크노|정보|과학|경제|창조)"),
    "강원": re.compile(r"강원"), "충북": re.compile(r"충북|충청북"), "충남": re.compile(r"충남|충청남"),
    "전북": re.compile(r"전북|전라북"), "전남": re.compile(r"전남|전라남"), "경북": re.compile(r"경북|경상북"),
    "경남": re.compile(r"경남|경상남"), "제주": re.compile(r"제주"),
    "비수도권": re.compile(r"비수도권"),  # 경기도 기업은 신청할 수 없는 공고
}
# 경기도 안의 권역 공고: 안산이 속한 권역을 확인하기 전까지는 빼지 않고 표시만 한다
ZONE_RE = re.compile(r"(서부|동부|남부|북부|북동부|중부)권역?")


def detect_regions(title, org):
    """공고의 지역 표시: 제목의 경기도 시·군 + 제목·기관의 다른 시·도. 없으면 경기도 전체·전국 공고."""
    found = [s for s, rx in SIGUN_RE.items() if rx.search(title)]
    text = f"{title} {org}"
    found += [p for p, rx in PROVINCE_RE.items() if rx.search(text)]
    return found


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
        it["regions"] = detect_regions(it["title"], it["org"])
        z = ZONE_RE.search(it["title"])
        it["zone"] = f"경기 {z.group(1)}권" if z else ""
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
        return '<span class="due">예산 소진 시</span>'
    if it["dday"] is None:
        return ""
    cls = "due soon" if it["dday"] <= soon else "due"
    label = "오늘 마감" if it["dday"] == 0 else f"D-{it['dday']}"
    return f'<span class="{cls}">{label}</span>'


def short_date(s):
    return s[5:].replace("-", ".") if len(s) == 10 else s


def is_soon(it, soon):
    return it["dday"] is not None and it["dday"] <= soon


def in_region(it, region):
    """지역 탭에 보일 공고: 그 시·군 이름이 있는 공고 + 지역 표시가 없는 경기도 전체·전국 공고"""
    return not it["regions"] or region in it["regions"]


def row(it, soon):
    period = "예산 소진 시까지" if it["until_budget"] else f'{short_date(it["start"])} ~ {short_date(it["end"])}'
    marks = ('<span class="mk new">[새글]</span>' if it["new"] else "") + \
            ('<span class="mk">[진흥원]</span>' if it["gbsa"] else "") + \
            "".join(f'<span class="mk">[{esc(r)}]</span>' for r in it["regions"]) + \
            (f'<span class="mk">[{esc(it["zone"])}]</span>' if it.get("zone") else "")
    kw = f'<span class="kw">관심: {esc("·".join(it["keywords"]))}</span>' if it["keywords"] else ""
    return (f'<tr data-cat="{esc(it["category"])}" data-new="{int(it["new"])}" data-reg="{esc(",".join(it["regions"]))}"'
            f' data-soon="{int(is_soon(it, soon))}" data-kw="{int(bool(it["keywords"]))}">'
            f'<td class="no"></td><td class="c">{esc(it["category"])}</td>'
            f'<td class="t">{marks}<a href="{esc(it["url"])}" target="_blank" rel="noopener">{esc(it["title"])}</a>'
            f'<span class="sm"><span class="c2">{esc(it["category"])}</span>{esc(period)}</span>{kw}</td>'
            f'<td class="org">{esc(it["org"])}</td><td class="per">{esc(period)}</td>'
            f'<td class="dd">{dday_badge(it, soon)}</td></tr>')


def table(items, soon, empty="해당 공고가 없습니다."):
    tbl = ('<div class="tw"><table><colgroup><col class="w-no"><col class="w-c"><col><col class="w-org"><col class="w-per">'
           '<col class="w-dd"></colgroup><thead><tr><th class="no">번호</th><th class="c">분야</th><th>사업명</th>'
           '<th class="org">기관</th><th class="per">접수기간</th><th class="dd">마감</th></tr></thead><tbody>'
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
    stats = (f'<p class="summary">접수 중인 공고 <b id="s-total">{len(items)}</b>건 '
             f'<span class="sep">·</span> 새 공고 <b id="s-new">{len(new)}</b>건 '
             f'<span class="sep">·</span> {soon}일 안에 마감 <b id="s-soon" class="red">{len(closing)}</b>건 '
             f'<span class="sep">·</span> 관심 키워드 <b id="s-kw">{len(matched)}</b>건</p>')
    regions = [r for r in cfg.get("지역탭", []) if r in SIGUN]
    rtabs = ""
    if not for_mail:
        quick = '<span class="sep">·</span>'.join(f'<button type="button" class="kwq" data-q="{esc(k)}">{esc(k)}</button>'
                                                  for k in cfg.get("관심키워드", []) if k)
        rtabs = ('<div class="tabbar"><div class="rtabs" role="tablist" aria-label="보기">'
                 f'<button class="rtab on" data-r="*" role="tab" aria-selected="true">경기 전체 <span>({len(items)})</span></button>'
                 + "".join(f'<button class="rtab" data-r="{esc(r)}" role="tab" aria-selected="false">{esc(r)} '
                           f'<span>({sum(in_region(i, r) for i in items)})</span></button>' for r in regions)
                 + '<button class="rtab" data-r="kw" role="tab" aria-selected="false">키워드 검색</button></div>'
                 + ('<button class="rtab live" data-live="1" type="button" title="이지비즈에서 지금 공고를 다시 가져옵니다">'
                    '최신 정보 불러오기</button>' if live_tab else '')
                 + '</div>'
                 + '<p class="rnote" id="rnote">※ 지역 탭에는 제목에 그 시·군 이름이 있는 공고와, 지역 표시가 없는 '
                   '경기도 전체·전국 공고만 나옵니다. 다른 시·군이나 다른 시·도(서울·인천 등) 공고는 빠집니다.</p>'
                 + '<form class="kwbox" id="kwbox" hidden onsubmit="return false">'
                   '<div class="kwrow"><label for="kwq">검색어</label>'
                   '<input id="kwq" type="search" autocomplete="off" placeholder="예: 수출지원, 스마트공장">'
                   '<button type="submit" class="kwgo">검색</button></div>'
                   f'<p class="kwquick">추천 검색어 : {quick}</p>'
                   '<p class="rnote" id="kwres">※ 사업명·분야·기관에서 찾습니다. \'수출지원\'처럼 끝에 붙은 지원·사업은 떼고 찾고, '
                   '띄어 쓰면 모든 단어가 들어간 공고, 쉼표(,)로 나누면 하나라도 들어간 공고를 보여줍니다.</p></form>')

    if for_mail:  # 메일 본문은 새 공고만 분야별로 짧게
        body = stats
        new_by_cat = {}
        for i in new:
            new_by_cat.setdefault(i["category"], []).append(i)
        if not new:
            body += '<p class="empty">새로 올라온 공고가 없습니다.</p>'
        for c in [c for c in cats if c in new_by_cat]:
            body += f'<h2>{esc(c)} <small>{len(new_by_cat[c])}건</small></h2>' + table(new_by_cat[c], soon)
        body += '<p class="foot">전체 목록은 첨부한 리포트 파일을 열어 보세요.</p>'
    else:
        chips = f'<button class="chip on" data-f="*">전체(<span class="n">{len(items)}</span>)</button>' + "".join(
            f'<button class="chip" data-f="{esc(c)}">{esc(c)}(<span class="n">{len(by_cat[c])}</span>)</button>' for c in cats)
        body = (f'{rtabs}{stats}'
                f'<h2>우리 회사 관심 공고 <small>관심 키워드 : {esc(", ".join(cfg.get("관심키워드", [])))}</small></h2>'
                + table(matched, soon, "관심 키워드에 맞는 공고가 없습니다.")
                + f'<h2>마감 임박 공고 <small>{soon}일 이내</small></h2>' + table(closing, soon)
                + f'<h2>분야별 전체 공고</h2><div class="chips">{chips}'
                  f'<label class="onlynew"><input type="checkbox" id="onlynew"> 새 공고만 보기</label></div>'
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
/* 관공서 공고 게시판: 흰 바탕, 진한 윗선 표, 파란 링크, 빨간 마감 표시. 장식 없이 표와 글자로만 구분한다 */
:root{--bg:#ffffff;--fg:#222222;--muted:#6b6b6b;--line:#dcdcdc;--rule:#333333;--head:#f6f6f4;--tab:#f1f1ef;
 --link:#1d4e89;--visited:#5a3d86;--red:#c4161c;--hl:#fff7d6;
 --serif:'Noto Serif KR','Batang','AppleMyungjo',serif;
 --sans:'Malgun Gothic','맑은 고딕','Apple SD Gothic Neo',sans-serif}
@media (prefers-color-scheme: dark){:root:not([data-theme="light"]){--bg:#17181a;--fg:#e4e4e1;--muted:#9c9c98;
 --line:#34363a;--rule:#b9b9b5;--head:#202124;--tab:#1d1e21;--link:#8db4ea;--visited:#c0a6e6;--red:#ff6b6b;
 --hl:#3a3420;color-scheme:dark}}
:root[data-theme="dark"]{--bg:#17181a;--fg:#e4e4e1;--muted:#9c9c98;--line:#34363a;--rule:#b9b9b5;--head:#202124;
 --tab:#1d1e21;--link:#8db4ea;--visited:#c0a6e6;--red:#ff6b6b;--hl:#3a3420;color-scheme:dark}
body{font-family:var(--sans);font-size:14px;line-height:1.55;background:var(--bg);color:var(--fg);margin:0;
 padding-block:28px 48px;padding-left:max(16px,env(safe-area-inset-left,0px));padding-right:max(16px,env(safe-area-inset-right,0px))}
[hidden]{display:none!important}
.wrap{max-width:1100px;margin:0 auto}
.crumb{font-size:12px;color:var(--muted);margin:0 0 6px}
h1{font-family:var(--serif);font-size:26px;font-weight:700;margin:0;letter-spacing:-.5px}
.sub{color:var(--muted);font-size:13px;margin:6px 0 22px;padding-bottom:14px;border-bottom:1px solid var(--line)}
h2{font-family:var(--serif);font-size:18px;font-weight:700;margin:34px 0 8px;letter-spacing:-.3px}
h3{font-size:14px;font-weight:700;margin:22px 0 6px}
small{color:var(--muted);font-weight:normal;font-size:12px;font-family:var(--sans);margin-left:6px;letter-spacing:0}
.summary{margin:14px 0 4px;font-size:14px}
.summary b{font-variant-numeric:tabular-nums} .summary b.red{color:var(--red)}
.sep{color:var(--muted);margin:0 6px}
/* 탭: 서류철 모양 */
.tabbar{display:flex;flex-wrap:wrap;align-items:flex-end;gap:8px;border-bottom:1px solid var(--rule)}
.rtabs{display:flex;flex-wrap:wrap;gap:2px}
.rtab{font:inherit;font-size:14px;color:var(--muted);background:var(--tab);border:1px solid var(--line);border-bottom:0;
 padding:7px 16px;cursor:pointer;margin-bottom:-1px}
.rtab span{font-size:12px;font-variant-numeric:tabular-nums}
.rtab.on{background:var(--bg);color:var(--fg);font-weight:700;border-color:var(--rule);border-bottom:1px solid var(--bg)}
.rtab:hover{color:var(--fg)}
.rtab:focus-visible,.chip:focus-visible,.kwq:focus-visible,.kwgo:focus-visible{outline:2px solid var(--link);outline-offset:1px}
.rtab.live{margin:0 0 6px auto;background:var(--bg);border:1px solid var(--line);font-size:13px;padding:4px 12px;color:var(--fg)}
.rtab.live:hover{border-color:var(--rule)} .rtab.live:disabled{color:var(--muted);cursor:wait}
.rnote{margin:8px 0 0;font-size:12px;color:var(--muted)}
/* 검색 */
.kwbox{margin:14px 0 0;padding:14px 16px;background:var(--head);border:1px solid var(--line)}
.kwrow{display:flex;gap:6px;align-items:center;flex-wrap:wrap}
.kwrow label{font-weight:700;font-size:13px;margin-right:4px}
.kwrow input{font:inherit;font-size:14px;padding:6px 8px;border:1px solid #999;border-radius:0;background:var(--bg);
 color:var(--fg);flex:1;min-width:0;max-width:420px}
.kwgo{font:inherit;font-size:13px;padding:6px 16px;border:1px solid var(--rule);background:var(--rule);color:var(--bg);cursor:pointer}
.kwquick{margin:10px 0 0;font-size:12px;color:var(--muted)}
.kwq{font:inherit;font-size:12px;border:0;background:none;padding:0;color:var(--link);cursor:pointer}
.kwq:hover{text-decoration:underline}
/* 표: 게시판 */
.tw{overflow-x:auto}
table{width:100%;border-collapse:collapse;border-top:2px solid var(--rule);table-layout:fixed}
col.w-no{width:52px} col.w-c{width:58px} col.w-org{width:190px} col.w-per{width:118px} col.w-dd{width:78px}
th{background:var(--head);font-weight:700;font-size:13px;padding:9px 8px;border-bottom:1px solid var(--line);text-align:center}
td{padding:9px 8px;border-bottom:1px solid var(--line);vertical-align:top}
tbody{counter-reset:n} tbody tr{counter-increment:n}
td.no::before{content:counter(n)} td.no,td.c,td.dd{text-align:center;color:var(--muted);font-size:13px}
td.t{word-break:keep-all;overflow-wrap:anywhere}
td.t a{color:var(--link);text-decoration:none} td.t a:visited{color:var(--visited)}
td.t a:hover{text-decoration:underline}
tbody tr:hover td{background:var(--head)}
.mk{font-size:12px;color:var(--muted);margin-right:4px} .mk.new{color:var(--red);font-weight:700}
.kw{display:block;font-size:12px;color:var(--muted);margin-top:2px}
.sm{display:none}
td.org{font-size:13px;color:var(--muted);word-break:keep-all}
td.per{font-size:13px;color:var(--muted);text-align:center;white-space:nowrap;font-variant-numeric:tabular-nums}
.due{font-size:13px;font-variant-numeric:tabular-nums;white-space:nowrap} .due.soon{color:var(--red);font-weight:700}
.empty{margin:0;padding:22px 8px;text-align:center;color:var(--muted);border-top:2px solid var(--rule);border-bottom:1px solid var(--line)}
/* 분야 선택: 글자 링크 */
.chips{display:flex;flex-wrap:wrap;align-items:center;gap:2px 0;margin:4px 0 6px;font-size:13px}
.chip{font:inherit;border:0;background:none;color:var(--muted);padding:2px 8px;cursor:pointer;border-left:1px solid var(--line)}
.chip:first-child{border-left:0;padding-left:0}
.chip.on{color:var(--fg);font-weight:700;text-decoration:underline;text-underline-offset:3px}
.chip:hover{color:var(--fg)}
.onlynew{margin-left:auto;font-size:13px;color:var(--muted)}
.foot{color:var(--muted);font-size:12px;margin-top:36px;padding-top:12px;border-top:1px solid var(--line)}
@media(max-width:760px){
 col.w-org,col.w-per,col.w-c,th.org,th.per,th.c,td.org,td.per,td.c{display:none}
 col.w-no{width:44px} col.w-dd{width:62px} th{white-space:nowrap;padding-inline:4px}
 .sm{display:block;font-size:12px;color:var(--muted);margin-top:2px} .sm .c2{margin-right:6px}
 .onlynew{margin-left:0;flex-basis:100%;margin-top:6px}
 .rtab.live{margin-left:0}
}
"""
    inner = f"""<title>경기 지원사업 공고</title>
<meta name="claude:viewport-fit-sides" content="cover">
<link rel="preconnect" href="https://fonts.googleapis.com"><link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=Noto+Serif+KR:wght@700&display=swap">
<style>{css}</style><div class="wrap">
<p class="crumb">경기도 기업지원 공고 모음</p>
<h1>지원사업 공고</h1>
<p class="sub">{esc(stamp or today.strftime('%Y년 %m월 %d일'))} 기준 &nbsp;|&nbsp; 자료 : {esc(scope)} &nbsp;|&nbsp; [새글] 지난번 확인 뒤 새로 올라온 공고</p>
{body}
<p class="foot">사업명을 누르면 경기기업비서(egbiz.or.kr)의 공고 원문으로 이동합니다. 신청 자격과 마감일은 원문에서 꼭 다시 확인하세요.</p>
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
def _load_json(rel, remote_base, default):
    """GitHub에 올라간 최신 파일을 먼저 읽고, 안 되면 함께 배포된 파일을 읽는다."""
    if remote_base:
        try:
            req = urllib.request.Request(remote_base + rel, headers={"User-Agent": "gov-support-watch"})
            with urllib.request.urlopen(req, timeout=8) as r:
                return json.loads(r.read().decode("utf-8"))
        except Exception:
            pass
    path = BASE / rel
    return json.loads(path.read_text(encoding="utf-8")) if path.exists() else default


def render_live(remote_base=None):
    """이지비즈에서 지금 공고를 가져와 웹페이지 HTML을 만든다. 기록 파일은 읽기만 한다."""
    cfg = _load_json("config.json", remote_base, {})
    seen = _load_json("data/seen.json", remote_base, {})
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
