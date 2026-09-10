"""우체국 택배 접수 목록 — 웹에서 채우고 우체국 양식(.xls)으로 내보낸다.

받는 분을 화면에서 하나씩 넣어두었다가, 다 되면 '내보내기'를 눌러
인터넷우체국 창구소포 파일접수 양식에 맞는 .xls 파일을 만든다.

새 파일을 만들지 않고 우체국이 준 양식 파일을 열어 줄만 채운다.
시트 이름과 머리글이 우체국이 기대하는 것과 한 글자도 달라지면 안 되기
때문이다.
"""
import json
import os
import re
import subprocess
import sys
import webbrowser
from datetime import datetime
from threading import Timer

from flask import Flask, request, jsonify, send_file

BASE = os.path.dirname(os.path.abspath(__file__))
STORE = os.path.join(BASE, "storage")
DATA = os.path.join(STORE, "entries.json")
BOOK = os.path.join(STORE, "addressbook.json")   # 보내는 분·받는 분 주소록
HISTORY = os.path.join(STORE, "history.json")    # 접수가 끝난 손님 기록
TEMPLATE = os.path.join(BASE, "template.xls")
OUTDIR = os.path.expanduser("~/Downloads")
# 보내는 물품이 늘 같아서 고정한다. 바뀌면 이 값만 고치면 된다.
DEFAULT_CODE = "농/수/축산물(일반)"
DEFAULT_VOLUME = "80"       # 상자 크기가 늘 같아서 고정한다
PORT = int(os.environ.get("PARCEL_PORT", "5063"))

# 우체국 양식의 칸 순서. 이 순서와 개수를 바꾸면 접수가 거부된다.
COLUMNS = [
    "받는분", "우편번호", "주소", "상세주소", "일반전화", "휴대전화",
    "중량", "부피", "내용품코드", "내용물", "배달방식", "요청사항",
    "분할여부", "분할1중량", "분할1부피", "분할2중량", "분할2부피",
]

app = Flask(__name__)


def load():
    try:
        with open(DATA, encoding="utf-8") as f:
            return json.load(f)
    except (OSError, json.JSONDecodeError):
        return []


def save(rows):
    os.makedirs(STORE, exist_ok=True)
    with open(DATA, "w", encoding="utf-8") as f:
        json.dump(rows, f, ensure_ascii=False, indent=2)


def clean_memo(text):
    """요청사항에서 우체국이 거부하는 글자를 뺀다.

    '#', 괄호 같은 기호가 들어가면 접수 화면이 '비정상적인 한글조합 문자'라며
    파일을 통째로 거부한다. 완성형 한글·영숫자·공백과 쉼표·마침표·하이픈만 남긴다.
    """
    keep = []
    for ch in str(text):
        if ("가" <= ch <= "힣") or ch.isalnum() or ch in " ,.-":
            keep.append(ch)
        else:
            keep.append(" ")
    return " ".join("".join(keep).split())


# 화면 정렬과 같은 순서로 뽑는다. 창구에서 상자를 종류별로 쌓아 놓고 내기 편하다.
BOX_ORDER = {"청향": 0, "캠벨": 1, "켐벨": 1}


def box_key(row):
    p = str(row.get("상품", ""))
    variety = BOX_ORDER.get(p[:2], 9)
    m = re.search(r"(\d+)\s*kg", p)
    kg = int(m.group(1)) if m else 0
    return (variety, kg, 2 if "2개" in p else 1, row.get("받는분", ""))


def export_xls(rows, sender=""):
    rows = sorted(rows, key=box_key)
    """우체국 양식 파일을 열어 예시 줄을 지우고 우리 줄을 넣는다."""
    import xlrd
    from xlutils.copy import copy as xl_copy

    book = xlrd.open_workbook(TEMPLATE, formatting_info=True)
    out = xl_copy(book)
    sheet = out.get_sheet(0)

    # 양식에 들어 있는 예시 줄(홍길동1, 홍길동2)을 지운다.
    src = book.sheet_by_index(0)
    for r in range(1, src.nrows):
        for c in range(src.ncols):
            sheet.write(r, c, "")

    for i, row in enumerate(rows, start=1):
        for c, key in enumerate(COLUMNS):
            val = row.get(key, "")
            # 접수는 늘 시골농원 이름으로 한다. 실제로 보내는 손님은
            # 받는 분 이름 뒤에 '누구보냄'으로 적어 구분한다.
            if key == "받는분":
                who = row.get("보낸이", "")
                if who and who != "시골농원":
                    val = f"{val} {who}보냄"
            # 늘 같은 값이라 화면에서 묻지 않고 여기서 채운다
            if key == "내용품코드":
                val = DEFAULT_CODE
            elif key == "부피":
                val = DEFAULT_VOLUME
            elif key == "요청사항":
                val = clean_memo(val)
            elif key == "상세주소" and not str(val).strip():
                # 우체국은 상세주소를 필수로 본다. 단독주택처럼 없을 때는 마침표를 넣는다.
                val = "."
            sheet.write(i, c, val)

    who = sender if sender else "전체"
    name = f"보내는사람 {who}_{datetime.now():%Y%m%d_%H%M}.xls"
    path = os.path.join(OUTDIR, name)
    out.save(path)
    return path


@app.route("/")
def index():
    return HTML


@app.route("/rows", methods=["GET", "POST"])
def rows_route():
    if request.method == "POST":
        save(request.get_json() or [])
    return jsonify(load())


def load_book():
    """주소록. 보내는 분마다 그 사람이 보냈던 받는 분을 함께 기억한다.

    같은 손님이 해마다 같은 분들께 보내는 일이 많아서, 한 번 받아 적어 두면
    다음부터는 고르기만 하면 된다.
    """
    try:
        with open(BOOK, encoding="utf-8") as f:
            return json.load(f)
    except (OSError, json.JSONDecodeError):
        return {"senders": [], "recipients": []}


def save_book(b):
    os.makedirs(STORE, exist_ok=True)
    with open(BOOK, "w", encoding="utf-8") as f:
        json.dump(b, f, ensure_ascii=False, indent=2)


def remember(sender, phone, recipient):
    """이번에 넣은 사람을 주소록에 새겨 둔다. 이미 있으면 최신 값으로 고친다."""
    b = load_book()
    if sender:
        for x in b["senders"]:
            if x["이름"] == sender:
                if phone:
                    x["전화"] = phone
                break
        else:
            b["senders"].append({"이름": sender, "전화": phone})
    if recipient.get("받는분"):
        key = (sender, recipient["받는분"])
        for i, x in enumerate(b["recipients"]):
            if (x.get("보낸이"), x.get("받는분")) == key:
                b["recipients"][i] = dict(recipient, 보낸이=sender)
                break
        else:
            b["recipients"].append(dict(recipient, 보낸이=sender))
    save_book(b)


@app.route("/book", methods=["GET", "POST"])
def book_route():
    if request.method == "POST":
        d = request.get_json() or {}
        if d.get("action") == "delete":
            b = load_book()
            b["recipients"] = [x for x in b["recipients"]
                               if not (x.get("보낸이") == d.get("보낸이")
                                       and x.get("받는분") == d.get("받는분"))]
            save_book(b)
        else:
            remember(d.get("보내는분", ""), d.get("보내는분전화", ""), d.get("받는분정보") or {})
    return jsonify(load_book())


@app.route("/export", methods=["POST"])
def export_route():
    rows = load()
    if not rows:
        return jsonify({"error": "넣은 것이 없습니다"}), 400
    missing = []
    for i, r in enumerate(rows, 1):
        for key in ("받는분", "우편번호", "주소", "중량"):
            if not str(r.get(key, "")).strip():
                missing.append(f"{i}번째 줄: {key}")
    if missing:
        return jsonify({"error": "빠진 칸이 있습니다", "missing": missing[:10]}), 400
    # 손님이 여럿이어도 접수는 시골농원 이름으로 한 번에 한다. 그래서 파일도 하나다.
    return jsonify({"ok": True, "names": [os.path.basename(export_xls(rows, "시골농원"))]})


def load_history():
    """접수가 끝난 손님 기록. 목록에서 내보낸 뒤 이리로 옮긴다."""
    try:
        with open(HISTORY, encoding="utf-8") as f:
            return json.load(f)
    except (OSError, json.JSONDecodeError):
        return []


@app.route("/hist", methods=["GET", "POST"])
def hist_route():
    """POST면 지금 목록을 통째로 기록에 넣고 목록을 비운다."""
    hist = load_history()
    if request.method == "POST":
        stamp = f"{datetime.now():%Y-%m-%d %H:%M}"
        for r in load():
            hist.append(dict(r, 접수일=stamp))
        save([])
        os.makedirs(STORE, exist_ok=True)
        with open(HISTORY, "w", encoding="utf-8") as f:
            json.dump(hist, f, ensure_ascii=False, indent=2)
    return jsonify(hist)


@app.route("/history")
def history_page():
    """지난 접수 기록을 보는 화면."""
    return HISTORY_HTML


@app.route("/contacts")
def contacts_page():
    """주소록에 쌓인 연락처를 보고 복사하는 화면. 카톡으로 안내 보낼 때 쓴다."""
    return CONTACTS_HTML


@app.route("/form")
def customer_form():
    """손님이 손으로 적을 종이 양식. 브라우저에서 인쇄한다."""
    return send_file(os.path.join(BASE, "손님용_양식.html"))


@app.route("/open_folder", methods=["POST"])
def open_folder():
    if sys.platform == "darwin":
        subprocess.run(["open", OUTDIR])
    elif os.name == "nt":
        os.startfile(OUTDIR)      # noqa: winonly
    return jsonify({"ok": True})


HTML = r"""
<!doctype html>
<html lang="ko"><head>
<meta charset="utf-8">
<link rel="icon" href="data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 100 100'%3E%3Ctext y='.9em' font-size='90'%3E%F0%9F%93%A6%3C/text%3E%3C/svg%3E">
<title>우체국 접수 목록</title>
<script src="https://t1.daumcdn.net/mapjsapi/bundle/postcode/prod/postcode.v2.js"></script>
<style>
 :root{ --bg:#0d0f14; --panel:#16191f; --elev:#1d2129; --line:#262b33;
        --txt:#e8eaf0; --muted:#949aa6; --accent:#e8503a; --ok:#3ddc84; --r:14px; }
 *{ box-sizing:border-box; }
 body{ font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",sans-serif;
       background:var(--bg); color:var(--txt); margin:0; }
 .wrap{ max-width:1180px; margin:0 auto; padding:32px 24px 64px; }
 h1{ font-size:22px; margin:0 0 4px; letter-spacing:-.02em; }
 p.sub{ color:var(--muted); margin:8px 0 22px; font-size:13.5px; }
 .card{ background:var(--panel); border:1px solid var(--line); border-radius:var(--r); padding:18px; }
 label{ display:block; font-size:12px; color:var(--muted); margin-bottom:5px; }
 input,select,textarea{ width:100%; background:var(--elev); color:var(--txt);
   border:1px solid var(--line); border-radius:9px; padding:9px 11px; font-size:13.5px;
   font-family:inherit; outline:none; }
 input:focus,select:focus,textarea:focus{ border-color:var(--accent); }
 .grid{ display:grid; gap:12px; }
 .g4{ grid-template-columns:1fr 1fr 1.8fr 1.3fr; }
 .g5{ grid-template-columns:.9fr .9fr 2.1fr 1.3fr; }
 .unit{ font-size:11.5px; color:var(--muted); white-space:nowrap; }
 .g3{ grid-template-columns:1fr 1fr 2fr; }
 button{ color:#fff; border:none; padding:10px 18px; border-radius:9px; font-size:13.5px;
   font-weight:650; cursor:pointer; background:var(--accent); }
 button.ghost{ background:var(--elev); color:var(--txt); border:1px solid var(--line); }
 button:disabled{ opacity:.5; cursor:default; }
 button:hover{ filter:brightness(1.08); }
 button:active{ transform:translateY(1px) scale(.98); filter:brightness(.85); }
 .row{ display:flex; gap:9px; align-items:center; margin-top:14px; flex-wrap:wrap; }
 table{ width:100%; border-collapse:collapse; margin-top:20px; font-size:13px; }
 th,td{ text-align:left; padding:9px 8px; border-bottom:1px solid var(--line); }
 th{ color:var(--muted); font-weight:600; font-size:11.5px; }
 td.num{ color:var(--muted); }
 .del{ color:var(--muted); cursor:pointer; }
 .del:hover{ color:var(--accent); }
 .grp td{ background:#20242c; font-weight:600; }
 .mini{ float:right; padding:3px 10px; font-size:11.5px; font-weight:600; }
 .sortbtn{ padding:6px 13px; font-size:12.5px; }
 .sortbtn.on{ background:var(--accent); color:#fff; border-color:var(--accent); }
 .gotoPost{ background:var(--accent); color:#fff; text-decoration:none; font-size:14px;
   font-weight:700; padding:12px 22px; border-radius:9px; white-space:nowrap; }
 .gotoPost:hover{ filter:brightness(1.12); }
 .gotoPost:active{ transform:translateY(1px) scale(.98); filter:brightness(.85); }
 .edit{ color:var(--muted); cursor:pointer; font-size:12.5px; }
 .edit:hover{ color:var(--accent); }
 .msg{ margin-top:14px; padding:11px 14px; border-radius:9px; font-size:13px; display:none; }
 .msg.ok{ background:rgba(61,220,132,.12); color:var(--ok); display:block; }
 .msg.err{ background:rgba(232,80,58,.12); color:#ff8f7a; display:block; white-space:pre-line; }
 .foot{ margin-top:26px; color:var(--muted); font-size:12.5px; }
 .foot a{ color:var(--accent); }
</style></head><body>
<div id="donePanel" style="display:none;position:fixed;inset:0;z-index:1000;
     background:rgba(0,0,0,.55)">
  <div style="position:absolute;left:50%;top:50%;transform:translate(-50%,-50%);
       width:min(460px,92vw);background:var(--panel);border:1px solid var(--line);
       border-radius:12px;padding:22px;box-shadow:0 20px 60px rgba(0,0,0,.5)">
    <div style="font-size:16px;font-weight:700;margin-bottom:10px">📄 xls 파일을 만들었습니다</div>
    <div id="doneList" style="font-size:13.5px;line-height:1.7;word-break:break-all"></div>
    <div class="row">
      <button id="doneOpen">다운로드 폴더 열기</button>
      <button class="ghost" id="doneClose">닫기</button>
    </div>
  </div>
</div>

<div id="addrPanel" style="display:none;position:fixed;inset:0;z-index:999;
     background:rgba(0,0,0,.55)">
  <div style="position:absolute;left:50%;top:50%;transform:translate(-50%,-50%);
       width:min(540px,94vw);height:min(580px,88vh);background:#fff;border-radius:12px;
       overflow:hidden;box-shadow:0 20px 60px rgba(0,0,0,.5)">
    <div style="display:flex;justify-content:space-between;align-items:center;
         padding:10px 14px;background:var(--accent);color:#fff;font-size:14px;font-weight:650">
      <span>주소찾기</span>
      <span id="addrClose" style="cursor:pointer;font-size:19px;line-height:1">&times;</span>
    </div>
    <div id="addrBox" style="width:100%;height:calc(100% - 41px)"></div>
  </div>
</div>
<div class="wrap">
 <div style="display:flex;justify-content:space-between;align-items:center;gap:18px;flex-wrap:wrap">
  <div>
   <h1>📦 우체국 접수 목록</h1>
   <p class="sub">받는 분을 넣어두고, 다 되면 내보내기를 누르세요. 우체국 창구소포 파일접수 양식(.xls)으로 나옵니다.<br>내용품코드 <b>농/수/축산물(일반)</b>, 부피 <b>80</b> 은 자동으로 들어갑니다. 상자는 둘씩 묶어 한 단계 위로 올립니다.</p>
  </div>
  <a class="gotoPost" href="https://epost.kr/main.retrieveMainPage.comm" target="_blank"
     rel="noopener">우체국 사이트 ↗</a>
 </div>

 <div class="card">
  <div class="grid" style="grid-template-columns:1.2fr 1fr 2fr;margin-bottom:12px">
   <div><label>보내는 손님 <span style="color:var(--muted)">(안 적으면 시골농원)</span></label>
     <input id="s_name" list="senders" placeholder="이름 (전에 온 분이면 골라짐)"></div>
   <div><label>보내는 분 전화 <span style="color:var(--muted)">(비워도 됨)</span></label>
     <input id="s_tel" placeholder="010-1234-5678"></div>
  </div>
  <datalist id="senders"></datalist>

  <div class="grid g4">
   <div><label>받는 분 *</label><input id="f_name" placeholder="홍길동"></div>
   <div><label>우편번호 *</label>
     <div style="display:flex;gap:6px">
       <input id="f_zip" placeholder="눌러서 검색" maxlength="5" readonly
              style="cursor:pointer;background:#20242c">
       <button class="ghost" id="findAddr" style="white-space:nowrap;padding:9px 12px">주소찾기</button>
     </div></div>
   <div><label>주소 *</label><input id="f_addr" placeholder="주소찾기로 채워집니다" readonly
        style="background:#20242c"></div>
   <div><label>상세주소 (동·호수)</label><input id="f_addr2" placeholder="101동 1502호"></div>
  </div>
  <div class="grid g5" style="margin-top:12px">
   <div><label>휴대전화</label><input id="f_mobile" placeholder="010-1234-5678"></div>
   <div><label>일반전화</label><input id="f_tel" placeholder="02-1234-5678"></div>
   <div><label>상자 개수 * <span style="color:var(--muted)">(둘씩 묶어 자동 계산)</span></label>
     <div style="display:flex;gap:5px;align-items:center">
       <input id="c_3" type="number" min="0" placeholder="0" style="width:52px">
       <span class="unit">3kg</span>
       <input id="c_5" type="number" min="0" placeholder="0" style="width:52px">
       <span class="unit">5kg</span>
     </div></div>
   <div><label>배송시 요청사항</label><input id="f_memo" placeholder="부재시 경비실에 맡겨주세요"></div>
  </div>
  <div class="row">
   <button id="add">목록에 추가</button>
   <button class="ghost" id="clearForm">칸 비우기</button>
   <button class="ghost" id="cancelEdit" style="display:none">수정 취소</button>
   <span style="color:var(--muted);font-size:12.5px">* 표시는 반드시 넣어야 합니다</span>
  </div>
 </div>

 <div class="row" style="margin-top:18px">
  <span style="color:var(--muted);font-size:12.5px">정렬</span>
  <button class="ghost sortbtn" id="sortBox">상자별</button>
  <button class="ghost sortbtn" id="sortEntry">접수 순서</button>
 </div>

 <table id="tbl"><thead><tr>
   <th>#</th><th>받는 분</th><th>상자</th><th>접수 중량</th><th>우편번호</th><th>주소</th>
   <th>연락처</th><th>요청사항</th><th></th><th></th>
 </tr></thead><tbody></tbody></table>

 <div class="row">
  <button id="exp">xls 내보내기 (시골농원 이름으로 한 파일)</button>
  <button class="ghost" id="archive">접수 끝냄 (기록으로 보내기)</button>
  <button class="ghost" id="clearAll">목록 전체 비우기</button>
  <span id="count" style="color:var(--muted);font-size:12.5px"></span>
 </div>
 <div class="msg" id="msg"></div>

 <div class="foot">저장 위치: <a href="#" id="openFolder">다운로드</a> 폴더 ·
   <a href="/form" target="_blank">손님용 종이 양식 인쇄</a> ·
   <a href="/contacts" target="_blank">연락처 보기</a> ·
   <a href="/history" target="_blank">지난 접수 기록</a><br>
   목록은 자동 저장되니 창을 닫아도 남아 있습니다</div>
</div>
<script>
const $ = id => document.getElementById(id);

// 주소찾기 — 우체국 화면과 같은 방식. 검색해서 고르면 우편번호와 주소가
// 자동으로 채워지고, 상세주소 칸으로 커서가 간다.
// 다음(카카오) 우편번호 서비스를 쓴다. 무료이고 키가 필요 없다.
function openAddrSearch(){
  if(typeof daum === 'undefined'){
    show('err','주소찾기를 불러오지 못했습니다. 인터넷 연결을 확인하세요.');
    return;
  }
  const panel = $('addrPanel'), box = $('addrBox');
  box.innerHTML = '';
  panel.style.display = 'block';
  new daum.Postcode({
    oncomplete: function(data){
      // 도로명이 있으면 도로명, 없으면 지번. 우체국 양식이 도로명 기준이다.
      const addr = data.userSelectedType === 'J' ? data.jibunAddress
                 : (data.roadAddress || data.jibunAddress);
      $('f_zip').value  = data.zonecode;
      $('f_addr').value = addr;
      panel.style.display = 'none';
      $('f_addr2').focus();
      hide();
    },
    onresize: function(size){ box.style.height = size.height + 'px'; },
    width: '100%', height: '100%',
  }).embed(box);
}
let rows = [];

// 보이는 이름과 실제 넣는 값이 다르다. 상자 이름으로 고르면 실제 무게가 들어간다.
const KGLABEL = { '5':'3kg', '7':'5kg', '15':'10kg' };
// 상자를 둘씩 묶어 한 단계 위로 올린다. 건수가 줄어 요금이 적게 나온다.
//   3kg 1개  -> 5kg 1건, 3kg 2개 -> 7kg 1건 (우체국에 적는 무게)
//   5kg 2개  -> 10kg 1건
//   10kg     -> 묶지 않고 개수만큼
// 묶고 남은 한 개는 그대로 한 건이 된다.
// 상자를 둘씩 묶어 한 단계 위로 올린다. 건수가 줄어 요금이 적게 나온다.
//   3kg 1개 -> 5kg 접수, 3kg 2개 -> 7kg 접수
//   5kg 1개 -> 7kg 접수, 5kg 2개 -> 15kg 접수
// 화면에 보이는 상자 이름과 우체국에 적는 무게가 다르므로 둘을 같이 들고 다닌다.
function packBoxes(c3, c5){
  const out = [];
  const put = (n, 중량, 상품) => { for(let i = 0; i < n; i++) out.push({중량, 상품}); };
  put(Math.floor(c5 / 2), '15', '캠벨 5kg 2개');
  put(c5 % 2,             '7',  '캠벨 5kg');
  put(Math.floor(c3 / 2), '7',  '청향 3kg 2개');
  put(c3 % 2,             '5',  '청향 3kg');
  return out;
}

const FIELDS = {
  받는분:'f_name', 우편번호:'f_zip', 주소:'f_addr', 상세주소:'f_addr2',
  일반전화:'f_tel', 휴대전화:'f_mobile', 요청사항:'f_memo',
};

// 주소에서 도로명+건물번호만 눈에 띄게. 시/도·구는 흐리게 둔다.
function addrHtml(a){
  const t = String(a || '').trim().split(/\s+/);
  if(t.length < 2) return esc(a);
  const head = t.slice(0, -2).join(' ');
  const core = t.slice(-2).join(' ');
  return (head ? esc(head) + ' ' : '')
       + `<b style="color:var(--accent)">${esc(core)}</b>`;
}

function esc(s){ return String(s??'').replace(/[&<>"]/g, c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c])); }

// 표 정렬: 'box'는 품종·크기 순, 'entry'는 넣은 순서 그대로.
let sortMode = 'box';

// 표는 품종(청향 먼저), 그다음 상자 크기 순으로 세운다. 눈으로 세기 편하다.
const VORDER = { '청향': 0, '캠벨': 1, '켐벨': 1 };
function boxKey(r){
  const p = r.상품 || '';
  const v = (p.slice(0,2) in VORDER) ? VORDER[p.slice(0,2)] : 9;
  const m = p.match(/(\d+)\s*kg/);
  return [v, m ? +m[1] : 0, /2개/.test(p) ? 2 : 1];
}

// 우체국은 보내는 분 1명당 접수 1건이라 화면도 보내는 분별로 묶어 보여준다.
function groupRows(){
  const out = [];
  rows.forEach((r,i)=>{
    const name = r.보낸이 || '시골농원';
    let g = out.find(x => x.name === name);
    if(!g){ g = {name, items:[]}; out.push(g); }
    g.items.push(i);
  });
  for(const g of out){
    g.items.sort((a,b)=>{
      const x = boxKey(rows[a]), y = boxKey(rows[b]);
      return (x[0]-y[0]) || (x[1]-y[1]) || (x[2]-y[2])
          || rows[a].받는분.localeCompare(rows[b].받는분, 'ko');
    });
  }
  return out;
}

function render(){
  const tb = $('tbl').querySelector('tbody');
  // 손님별로 나누지 않고 한 줄로 세운다. 손님이 있으면 받는 분 옆에 적는다.
  const order = rows.map((r,i)=>i);
  if(sortMode === 'box'){
    order.sort((a,b)=>{
      const x = boxKey(rows[a]), y = boxKey(rows[b]);
      return (x[0]-y[0]) || (x[1]-y[1]) || (x[2]-y[2])
          || rows[a].받는분.localeCompare(rows[b].받는분, 'ko');
    });
  }
  $('sortBox').classList.toggle('on', sortMode === 'box');
  $('sortEntry').classList.toggle('on', sortMode === 'entry');
  // 번호는 송장 개수를 세려고 붙인다. 보이는 순서대로 1,2,3...
  tb.innerHTML = order.map((i, n) => { const r = rows[i]; return `<tr>
    <td class="num">${n+1}</td>
    <td>${esc(r.받는분)}${r.보낸이 ? ` <span style="color:var(--muted)">${esc(r.보낸이)}보냄</span>` : ''}</td>
    <td class="num">${esc(r.상품 || KGLABEL[r.중량] || r.중량)}</td>
    <td class="num">${esc(r.중량)}kg</td>
    <td class="num">${esc(r.우편번호)}</td>
    <td>${addrHtml(r.주소)} ${esc(r.상세주소||'')}</td>
    <td class="num">${esc(r.휴대전화||r.일반전화||'')}</td>
    <td>${esc(r.요청사항||'')}</td>
    <td class="edit" data-i="${i}">수정</td>
    <td class="del" data-i="${i}">✕</td></tr>`; }).join('');
  tb.querySelectorAll('.del').forEach(el=>el.onclick=()=>{
    if(editing === +el.dataset.i) stopEdit();
    rows.splice(+el.dataset.i,1); persist();
  });
  tb.querySelectorAll('.edit').forEach(el=>el.onclick=()=>startEdit(+el.dataset.i));
  $('count').textContent = rows.length ? `${rows.length}건` : '';
}

async function persist(){
  await fetch('/rows',{method:'POST',headers:{'Content-Type':'application/json'},
                       body:JSON.stringify(rows)});
  render();
}

// 목록에 넣은 뒤에도 고칠 수 있다. 손글씨 양식을 옮겨 적다 보면 이름이나
// 호수가 틀리는 일이 잦다. 상자 개수는 이미 건별로 나뉘어 있어 건드리지 않고
// 받는 분 정보만 고친다. 개수를 바꾸려면 지우고 다시 넣는다.
let editing = null;

// 화면에 적어둔 상자 이름을 개수로 되돌린다. 고칠 때 칸을 채워 주기 위해서다.
const BOXBACK = { '3kg':[1,0], '3kg 2개':[2,0], '5kg':[0,1], '5kg 2개':[0,2],
                  '6kg':[0,1], '10kg':[0,2] };

// '캠벨 5kg 2개'처럼 품종이 붙어 있어도 개수를 읽어낸다.
function backCount(상품){
  const p = String(상품 || '').replace(/^(청향|캠벨|켐벨)\s*/, '');
  return BOXBACK[p] || [0, 0];
}

// 한 사람 앞으로 나간 줄은 여러 개일 수 있다(상자를 둘씩 묶어 나눠 놓아서).
// 고칠 때는 그 사람 줄을 통째로 다룬다.
function sameRows(r){
  return rows.map((x,i)=>i).filter(i => {
    const x = rows[i];
    return (x.보낸이||'') === (r.보낸이||'') && x.받는분 === r.받는분
        && x.주소 === r.주소 && (x.상세주소||'') === (r.상세주소||'')
        && (x.상품||'') === (r.상품||'');   // 상자가 다르면 따로 고친다
  });
}

function startEdit(i){
  const r = rows[i];
  editing = i;
  $('s_name').value = r.보낸이 || '';
  for(const [k,id] of Object.entries(FIELDS)) $(id).value = r[k] || '';
  // 상자 개수도 되돌려 채운다. 숫자를 바꾸면 그 사람 줄이 다시 짜인다.
  let c3 = 0, c5 = 0;
  for(const j of sameRows(r)){
    const back = backCount(rows[j].상품);
    c3 += back[0]; c5 += back[1];
  }
  $('c_3').value = c3 || '';
  $('c_5').value = c5 || '';
  $('add').textContent = '수정 저장';
  $('cancelEdit').style.display = '';
  show('ok', `${esc(r.받는분)} 건을 고치는 중입니다. 상자 개수도 바꿀 수 있습니다.`);
  $('f_name').focus();
}

function stopEdit(){
  editing = null;
  $('add').textContent = '목록에 추가';
  $('cancelEdit').style.display = 'none';
}

$('cancelEdit').onclick = ()=>{ stopEdit(); clearForm(); hide(); };

$('add').onclick = ()=>{
  const r = {};
  for(const [k,id] of Object.entries(FIELDS)) r[k] = $(id).value.trim();
  r.분할여부 = 'N';
  r.분할1중량 = r.분할1부피 = r.분할2중량 = r.분할2부피 = '';
  for(const k of ['받는분','우편번호','주소']){
    if(!r[k]){ show('err', k + ' 칸을 채워주세요'); return; }
  }
  // 접수는 늘 시골농원 이름으로 한다. 손님 이름은 받는 분 옆에 적어 구분한다.
  const sender = $('s_name').value.trim();
  r.보낸이 = (sender && sender !== '시골농원') ? sender : '';
  r.보내는분 = '시골농원';
  const n = id => Math.max(0, parseInt($(id).value || '0', 10) || 0);
  if(editing !== null){
    const old = rows[editing];
    const mine = sameRows(old);
    const weights = packBoxes(n('c_3'), n('c_5'));
    if(!weights.length){ show('err', '상자 개수를 넣어주세요'); return; }
    // 그 사람 줄을 지우고 바뀐 개수대로 다시 만든다.
    const first = mine[0];
    rows = rows.filter((x,i) => !mine.includes(i));
    rows.splice(first, 0, ...weights.map(w => Object.assign({}, r, w)));
    stopEdit(); persist(); clearForm();
    show('ok', `${r.받는분} 건을 고쳤습니다 (${weights.length}건)`);
    $('f_name').focus();
    return;
  }
  const weights = packBoxes(n('c_3'), n('c_5'));
  if(!weights.length){ show('err', '상자 개수를 넣어주세요'); return; }
  // 한 손님이 여러 건이 될 수 있다. 건마다 한 줄씩 만든다.
  for(const w of weights) rows.push(Object.assign({}, r, w));
  // 주소록에 새겨 둔다. 다음에 같은 분이 오면 고르기만 하면 된다.
  fetch('/book', {method:'POST', headers:{'Content-Type':'application/json'},
    body: JSON.stringify({보내는분: sender, 보내는분전화: $('s_tel').value.trim(),
                          받는분정보: r})}).then(()=>loadBook());
  persist(); clearForm(); hide();
  $('f_name').focus();
};

function clearForm(){
  // 보내는 사람이 같은 물건을 여러 명에게 보내는 일이 많다.
  // 물품 정보는 남겨 두고 받는 분 정보만 비운다.
  ['f_name','f_zip','f_addr','f_addr2','f_tel','f_mobile',
   'c_3','c_5'].forEach(id=>$(id).value='');
}
$('clearForm').onclick = ()=>{ Object.values(FIELDS).forEach(id=>$(id).value=''); };

// 창구에 내고 나면 목록을 기록으로 옮긴다. 다음 손님과 섞이지 않게.
$('archive').onclick = async ()=>{
  if(!rows.length) return;
  if(!confirm(`${rows.length}건을 접수 기록으로 옮기고 목록을 비웁니다. 계속할까요?`)) return;
  await fetch('/hist', {method:'POST', headers:{'Content-Type':'application/json'},
                        body: JSON.stringify({})});
  rows = []; render();
  show('ok', '기록으로 옮겼습니다. 지난 접수 기록에서 볼 수 있습니다.');
};

$('clearAll').onclick = ()=>{
  if(!rows.length) return;
  if(!confirm(`목록 ${rows.length}건을 모두 지울까요?`)) return;
  rows = []; persist();
};

function show(kind, text){ const m=$('msg'); m.className='msg '+kind; m.textContent=text; }
function hide(){ $('msg').className='msg'; }

// sender를 주면 그 사람 것만, 안 주면 전체를 뽑는다.
async function exportXls(){
  hide();
  const res = await fetch('/export', {method:'POST',
    headers:{'Content-Type':'application/json'}, body: '{}'});
  const d = await res.json();
  if(!res.ok){
    show('err', d.error + (d.missing ? '\n' + d.missing.join('\n') : ''));
    return;
  }
  $('doneList').innerHTML = d.names.map(n=>esc(n)).join('<br>')
    + '<div style="color:var(--muted);margin-top:8px">다운로드 폴더에 저장했습니다</div>';
  $('donePanel').style.display = 'block';
}

$('doneClose').onclick = ()=>{ $('donePanel').style.display='none'; };
$('donePanel').onclick = e=>{ if(e.target.id==='donePanel') $('donePanel').style.display='none'; };
$('doneOpen').onclick = ()=>{ fetch('/open_folder',{method:'POST'}); };
$('exp').onclick = ()=> exportXls();
$('sortBox').onclick = ()=>{ sortMode = 'box'; render(); };
$('sortEntry').onclick = ()=>{ sortMode = 'entry'; render(); };


$('addrClose').onclick = ()=>{ $('addrPanel').style.display='none'; };
$('addrPanel').onclick = e=>{ if(e.target.id==='addrPanel') $('addrPanel').style.display='none'; };
document.addEventListener('keydown', e=>{
  if(e.key==='Escape'){ $('addrPanel').style.display='none'; $('donePanel').style.display='none'; }
});
$('findAddr').onclick = e=>{ e.preventDefault(); openAddrSearch(); };
$('f_zip').onclick = ()=> openAddrSearch();
$('f_addr').onclick = ()=> openAddrSearch();
$('openFolder').onclick = e=>{ e.preventDefault(); fetch('/open_folder',{method:'POST'}); };

let book = {senders:[], recipients:[]};

function refreshSenders(){
  $('senders').innerHTML = book.senders
    .map(x=>`<option value="${esc(x.이름)}">`).join('');
}

// 보내는 분 전화는 접수에 안 쓰인다(우체국 양식에 칸이 없다). 비워 둬도 된다.
function onSenderChange(){}
$('s_name').oninput = onSenderChange;
$('s_name').onchange = onSenderChange;

async function loadBook(){
  book = await (await fetch('/book')).json();
  refreshSenders(); onSenderChange();
}

fetch('/rows').then(r=>r.json()).then(d=>{ rows=d; render(); });
loadBook();
</script></body></html>
"""


HISTORY_HTML = r"""
<!doctype html><html lang="ko"><head><meta charset="utf-8">
<title>지난 접수 기록</title>
<style>
 :root{ --bg:#0d0f14; --panel:#16191f; --elev:#1d2129; --line:#262b33;
        --txt:#e8eaf0; --muted:#949aa6; --accent:#e8503a; --ok:#3ddc84; }
 *{box-sizing:border-box}
 body{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",sans-serif;
      background:var(--bg);color:var(--txt);margin:0}
 .wrap{max-width:1000px;margin:0 auto;padding:28px 22px 60px}
 h1{font-size:20px;margin:0 0 6px}
 p.sub{color:var(--muted);font-size:13px;margin:6px 0 18px}
 table{width:100%;border-collapse:collapse;font-size:13.5px}
 th,td{text-align:left;padding:8px;border-bottom:1px solid var(--line)}
 th{color:var(--muted);font-size:11.5px;font-weight:600}
 td.num{color:var(--muted)}
 input{background:var(--elev);border:1px solid var(--line);color:var(--txt);
       padding:9px 11px;border-radius:8px;font-size:13.5px;width:260px}
 .row{display:flex;gap:8px;align-items:center;margin:12px 0}
 .cnt{color:var(--muted);font-size:12.5px}
 a{color:var(--accent)}
</style></head><body><div class="wrap">
 <h1>🧾 지난 접수 기록</h1>
 <p class="sub">접수를 끝낸 손님들입니다. 누가 언제 어디로 보냈는지 여기서 찾습니다.
   · <a href="/">접수 화면으로</a></p>
 <div class="row">
   <input id="q" placeholder="이름·주소·전화로 찾기">
   <span class="cnt" id="cnt"></span>
 </div>
 <table id="t"><thead><tr>
   <th>접수일</th><th>보내는 분</th><th>받는 분</th><th>연락처</th>
   <th>주소</th><th>상자</th><th>요청사항</th>
 </tr></thead><tbody></tbody></table>
</div>
<script>
const $ = id => document.getElementById(id);
function esc(s){ return String(s??'').replace(/[&<>"]/g, c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c])); }
let hist = [];

function render(){
  const q = $('q').value.trim();
  // 최근 것이 위로 오게 뒤집는다.
  const list = hist.slice().reverse()
    .filter(r => !q || JSON.stringify(r).includes(q));
  $('t').querySelector('tbody').innerHTML = list.map(r=>`<tr>
    <td class="num">${esc(r.접수일||'')}</td>
    <td>${esc(r.보내는분||'')}</td>
    <td>${esc(r.받는분||'')}</td>
    <td class="num">${esc(r.휴대전화||r.일반전화||'')}</td>
    <td>${esc(r.주소||'')} ${esc(r.상세주소||'')}</td>
    <td class="num">${esc(r.상품||r.중량||'')}</td>
    <td>${esc(r.요청사항||'')}</td></tr>`).join('');
  $('cnt').textContent = `${list.length}건` + (q ? ` (전체 ${hist.length}건)` : '');
}

$('q').oninput = render;
(async ()=>{ hist = await (await fetch('/hist')).json(); render(); })();
</script></body></html>
"""


CONTACTS_HTML = r"""
<!doctype html><html lang="ko"><head><meta charset="utf-8">
<title>연락처</title>
<style>
 :root{ --bg:#0d0f14; --panel:#16191f; --elev:#1d2129; --line:#262b33;
        --txt:#e8eaf0; --muted:#949aa6; --accent:#e8503a; --ok:#3ddc84; }
 *{box-sizing:border-box}
 body{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",sans-serif;
      background:var(--bg);color:var(--txt);margin:0}
 .wrap{max-width:900px;margin:0 auto;padding:28px 22px 60px}
 h1{font-size:20px;margin:0 0 6px} h2{font-size:15px;margin:26px 0 8px}
 p.sub{color:var(--muted);font-size:13px;margin:6px 0 18px}
 table{width:100%;border-collapse:collapse;font-size:13.5px}
 th,td{text-align:left;padding:8px;border-bottom:1px solid var(--line)}
 th{color:var(--muted);font-size:11.5px;font-weight:600}
 td.num{color:var(--muted)}
 button{background:var(--accent);color:#fff;border:none;padding:8px 15px;
        border-radius:8px;font-size:13px;font-weight:650;cursor:pointer}
 button.ghost{background:var(--elev);color:var(--txt);border:1px solid var(--line)}
 .row{display:flex;gap:8px;align-items:center;margin:12px 0}
 .msg{color:var(--ok);font-size:13px}
 a{color:var(--accent)}
</style></head><body><div class="wrap">
 <h1>📇 연락처</h1>
 <p class="sub">접수하면서 쌓인 연락처입니다. 카톡으로 안내 보낼 때 복사해 쓰세요.
   · <a href="/">접수 화면으로</a></p>

 <h2>보내는 분</h2>
 <div class="row">
   <button id="copyS">전화번호만 복사</button>
   <button class="ghost" id="copySAll">이름+전화 복사</button>
   <span class="msg" id="msgS"></span>
 </div>
 <table id="tS"><thead><tr><th>이름</th><th>전화</th></tr></thead><tbody></tbody></table>

 <h2>받는 분</h2>
 <div class="row">
   <button id="copyR">전화번호만 복사</button>
   <button class="ghost" id="copyRAll">이름+전화 복사</button>
   <span class="msg" id="msgR"></span>
 </div>
 <table id="tR"><thead><tr>
   <th>보낸 분</th><th>받는 분</th><th>전화</th><th>주소</th></tr></thead><tbody></tbody></table>
</div>
<script>
const $=id=>document.getElementById(id);
const esc=s=>String(s??'').replace(/[&<>"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));
let book={senders:[],recipients:[]};

function phoneOf(x){ return x.휴대전화 || x.일반전화 || x.전화 || ''; }

async function copy(text, where){
  try{ await navigator.clipboard.writeText(text); $(where).textContent='복사했습니다'; }
  catch(e){ $(where).textContent='복사 실패 — 직접 긁어서 복사하세요'; }
  setTimeout(()=>$(where).textContent='', 2500);
}

fetch('/book').then(r=>r.json()).then(d=>{
  book=d;
  $('tS').querySelector('tbody').innerHTML = book.senders.map(x=>
    `<tr><td>${esc(x.이름)}</td><td class="num">${esc(x.전화||'')}</td></tr>`).join('')
    || '<tr><td colspan="2" style="color:#949aa6">아직 없습니다</td></tr>';
  $('tR').querySelector('tbody').innerHTML = book.recipients.map(x=>
    `<tr><td class="num">${esc(x.보낸이||'')}</td><td>${esc(x.받는분)}</td>
     <td class="num">${esc(phoneOf(x))}</td>
     <td>${esc((x.주소||'')+' '+(x.상세주소||''))}</td></tr>`).join('')
    || '<tr><td colspan="4" style="color:#949aa6">아직 없습니다</td></tr>';
});

$('copyS').onclick=()=>copy(book.senders.map(x=>x.전화).filter(Boolean).join('\n'),'msgS');
$('copySAll').onclick=()=>copy(book.senders.map(x=>`${x.이름} ${x.전화||''}`.trim()).join('\n'),'msgS');
$('copyR').onclick=()=>copy(book.recipients.map(phoneOf).filter(Boolean).join('\n'),'msgR');
$('copyRAll').onclick=()=>copy(book.recipients.map(x=>`${x.받는분} ${phoneOf(x)}`.trim()).join('\n'),'msgR');
</script></body></html>
"""


if __name__ == "__main__":
    site = f"http://127.0.0.1:{PORT}"
    print(f"우체국 접수 목록 실행 중 -> {site}  (종료: Ctrl+C)")
    if not os.environ.get("NO_BROWSER"):
        Timer(1.2, lambda: webbrowser.open(site)).start()
    app.run(host="127.0.0.1", port=PORT, debug=False)
