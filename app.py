import streamlit as st
import streamlit.components.v1 as components
import openpyxl
import json
import io
import re
import urllib.request
import urllib.error
from pathlib import Path
from datetime import datetime

# ─── 연간 목표매출 (VND) ───────────────────────────────────────────────────────
ANNUAL_TARGETS = {
    1:  {'젖병': 1754469504, '젖꼭지': 403546437, '바디워시': 481810370, '물티슈': 1410549521, '젖병세정제': 514027234, '청소솔': 836906151, '기저귀': 200000000},
    2:  {'젖병': 1270020456, '젖꼭지': 282661378, '바디워시': 348516244, '물티슈': 1201577622, '젖병세정제': 327117563, '청소솔': 619352366, '기저귀': 180000000},
    3:  {'젖병': 2076477719, '젖꼭지': 474242542, '바디워시': 639039590, '물티슈': 2378230343, '젖병세정제': 540891541, '청소솔': 1198454158, '기저귀': 400000000},
    4:  {'젖병': 1686502756, '젖꼭지': 413887703, '바디워시': 675997924, '물티슈': 2283146524, '젖병세정제': 468675303, '청소솔': 951979456, '기저귀': 600000000},
    5:  {'젖병': 1700000000, '젖꼭지': 434582088, '바디워시': 680000000, '물티슈': 2900000000, '젖병세정제': 650000000, '청소솔': 1000000000, '기저귀': 850000000},
    6:  {'젖병': 1740000000, '젖꼭지': 607200000, '바디워시': 690000000, '물티슈': 3700000000, '젖병세정제': 750000000, '청소솔': 1200000000, '기저귀': 1200000000},
    7:  {'젖병': 2040000000, '젖꼭지': 811440000, '바디워시': 862500000, '물티슈': 4500000000, '젖병세정제': 920000000, '청소솔': 1285000000, '기저귀': 1800000000},
    8:  {'젖병': 2400000000, '젖꼭지': 1104000000, '바디워시': 1035000000, '물티슈': 4500000000, '젖병세정제': 977500000, '청소솔': 1350000000, '기저귀': 3000000000},
    9:  {'젖병': 2680000000, '젖꼭지': 1420250000, '바디워시': 1150000000, '물티슈': 5300000000, '젖병세정제': 1150000000, '청소솔': 1450000000, '기저귀': 5000000000},
    10: {'젖병': 2800000000, '젖꼭지': 1700000000, '바디워시': 1150000000, '물티슈': 5300000000, '젖병세정제': 1150000000, '청소솔': 1550000000, '기저귀': 8000000000},
    11: {'젖병': 2880000000, '젖꼭지': 2100000000, '바디워시': 1035000000, '물티슈': 5000000000, '젖병세정제': 1092500000, '청소솔': 1500000000, '기저귀': 9000000000},
    12: {'젖병': 2880000000, '젖꼭지': 2200000000, '바디워시': 1035000000, '물티슈': 6000000000, '젖병세정제': 1092500000, '청소솔': 1600000000, '기저귀': 10000000000},
}

CATEGORIES = ['젖병', '젖꼭지', '바디워시', '물티슈', '젖병세정제', '청소솔', '기저귀']
DATA_DIR = Path(__file__).parent / 'data'
DATA_DIR.mkdir(exist_ok=True)


# ─── 카테고리 분류 ─────────────────────────────────────────────────────────────
def categorize(name):
    n = str(name).lower()
    if 'tã quần' in n or 'ta quan' in n:
        return '기저귀'
    if ('khăn ướt' in n or 'khan uot' in n) \
            and 'nước rửa' not in n and 'nuoc rua' not in n \
            and 'sữa tắm' not in n and 'sua tam' not in n \
            and 'nước giặt' not in n and 'nuoc giat' not in n:
        return '물티슈'
    if 'nước giặt' in n or 'nuoc giat' in n:
        return '세탁세제'
    if 'sữa tắm' in n or 'sua tam' in n or 'bọt tắm' in n or 'bot tam' in n \
            or 'bọt rửa tay' in n or 'bot rua tay' in n:
        return '바디워시'
    if 'dụng cụ' in n or 'dung cu' in n:
        return '청소솔'
    if 'nước rửa bình' in n or 'nuoc rua binh' in n:
        return '젖병세정제'
    if 'núm' in n or ('num' in n and 'binh' not in n and 'nuoc' not in n) \
            or 'ống hút thông minh' in n or 'ong hut thong minh' in n:
        return '젖꼭지'
    is_bottle = ('bình sữa' in n or 'binh sua' in n
                 or 'bình ống hút' in n or 'binh ong hut' in n
                 or 'nắp đậy' in n or 'nap day' in n
                 or 'phễu rót' in n or 'pheu rot' in n
                 or ('ống hút' in n and 'thông minh' not in n))
    if is_bottle:
        return '젖병'
    return '기타'


# ─── 엑셀 시트 파싱 ────────────────────────────────────────────────────────────
def parse_sheet(ws):
    rows = list(ws.values)
    if not rows:
        return {}

    header_idx = None
    name_col = None
    rev_col = None

    for i, row in enumerate(rows):
        row_lower = [str(c).lower().strip() if c is not None else '' for c in row]
        has_name = any('tên' in c or ('ten' in c and 'san' in c) for c in row_lower)
        has_rev = any('doanh' in c for c in row_lower)
        if has_name and has_rev:
            header_idx = i
            for j, h in enumerate(row_lower):
                if ('tên' in h or ('ten' in h and 'san' in h)) and name_col is None:
                    name_col = j
                if 'doanh' in h and rev_col is None:
                    rev_col = j
            break

    if header_idx is None or name_col is None or rev_col is None:
        return {}

    result = {cat: 0.0 for cat in CATEGORIES}

    for row in rows[header_idx + 1:]:
        if not row or len(row) <= max(name_col, rev_col):
            continue
        name = row[name_col]
        rev = row[rev_col]
        if name is None:
            continue
        try:
            rev_val = float(str(rev).replace(',', '').replace(' ', '')) if rev is not None else 0.0
        except (ValueError, TypeError):
            rev_val = 0.0
        if rev_val <= 0:
            continue
        cat = categorize(str(name))
        if cat in result:
            result[cat] += rev_val

    return result


# ─── Gist / 로컬 저장소 ────────────────────────────────────────────────────────
def _use_gist():
    try:
        return bool(st.secrets.get('GIST_TOKEN') and st.secrets.get('GIST_ID'))
    except Exception:
        return False

def _gist_request(method, data=None):
    token = st.secrets['GIST_TOKEN']
    gist_id = st.secrets['GIST_ID']
    url = f'https://api.github.com/gists/{gist_id}'
    body = json.dumps(data).encode() if data else None
    req = urllib.request.Request(url, data=body, method=method, headers={
        'Authorization': f'token {token}',
        'Accept': 'application/vnd.github.v3+json',
        'Content-Type': 'application/json',
        'User-Agent': 'sales-dashboard',
    })
    with urllib.request.urlopen(req) as r:
        return json.loads(r.read())

def _load_gist_records():
    try:
        res = _gist_request('GET')
        content = res['files']['sales_records.json']['content']
        return {int(k): v for k, v in json.loads(content).items()}
    except Exception:
        return {}

def _save_gist_records(records):
    payload = {str(k): v for k, v in records.items()}
    _gist_request('PATCH', {
        'files': {'sales_records.json': {'content': json.dumps(payload, ensure_ascii=False, indent=2)}}
    })

def save_record(month, date_label, actual):
    target = ANNUAL_TARGETS[month]
    pct = {cat: round(actual.get(cat, 0) / target[cat] * 100, 1) if target[cat] else 0.0 for cat in CATEGORIES}
    record = {
        'month': month,
        'date_label': date_label,
        'saved_at': datetime.now().isoformat(),
        'actual': {k: float(v) for k, v in actual.items()},
        'target': {k: float(v) for k, v in target.items()},
        'pct': pct,
    }
    if _use_gist():
        all_records = _load_gist_records()
        all_records[month] = record
        _save_gist_records(all_records)
    else:
        path = DATA_DIR / f'2026_{month:02d}.json'
        with open(path, 'w', encoding='utf-8') as f:
            json.dump(record, f, ensure_ascii=False, indent=2)

def load_history():
    if _use_gist():
        return _load_gist_records()
    records = {}
    for fp in sorted(DATA_DIR.glob('2026_*.json')):
        m = re.match(r'2026_(\d+)', fp.stem)
        if m:
            with open(fp, encoding='utf-8') as f:
                records[int(m.group(1))] = json.load(f)
    return records


# ─── 포맷 ──────────────────────────────────────────────────────────────────────
def fmt_vnd(val):
    val = float(val)
    if val >= 1e9:
        return f"{val / 1e9:.2f} tỷ"
    if val >= 1e6:
        return f"{val / 1e6:.0f} triệu"
    return f"{val:,.0f}"


# ─── 달성률 바 HTML ────────────────────────────────────────────────────────────
def render_bars(actual_data, target_data, date_label=""):
    items_html = ""
    total_actual = 0.0
    total_target = 0.0
    under_count = 0

    for cat in CATEGORIES:
        actual = float(actual_data.get(cat, 0))
        target = float(target_data.get(cat, 1))
        pct = actual / target * 100 if target else 0
        bar_w = min(pct, 100)
        total_actual += actual
        total_target += target

        if pct >= 100:
            color = '#27AE60'
        elif pct >= 70:
            color = '#E67E22'
        else:
            color = '#E74C3C'
            under_count += 1

        over_tag = (
            '<span style="background:#27AE60;color:#fff;padding:1px 7px;'
            'border-radius:10px;font-size:11px;margin-left:6px;">OVER</span>'
            if pct > 100 else ''
        )

        items_html += f"""
        <div style="margin-bottom:18px;">
          <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:5px;">
            <span style="font-weight:600;font-size:15px;color:#2C3E50;">{cat}</span>
            <span style="font-weight:700;font-size:16px;color:{color};">{pct:.1f}%{over_tag}</span>
          </div>
          <div style="background:#ECF0F1;border-radius:10px;height:28px;overflow:hidden;">
            <div style="width:{bar_w:.1f}%;background:{color};height:100%;border-radius:10px;
                        display:flex;align-items:center;padding-left:8px;box-sizing:border-box;">
              {'<span style="color:white;font-size:12px;font-weight:600;">'+fmt_vnd(actual)+'</span>' if bar_w > 15 else ''}
            </div>
          </div>
          <div style="display:flex;justify-content:space-between;margin-top:4px;font-size:12px;color:#95A5A6;">
            <span>실적: {fmt_vnd(actual)} VND</span>
            <span>목표: {fmt_vnd(target)} VND</span>
          </div>
        </div>
        """

    total_pct = total_actual / total_target * 100 if total_target else 0
    total_bar_w = min(total_pct, 100)
    total_color = '#27AE60' if total_pct >= 100 else '#E67E22' if total_pct >= 70 else '#E74C3C'

    items_html += f"""
    <div style="margin-top:10px;padding-top:14px;border-top:2px solid #ECF0F1;">
      <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:6px;">
        <span style="font-weight:700;font-size:17px;color:#2C3E50;">전체 합계</span>
        <span style="font-weight:800;font-size:20px;color:{total_color};">{total_pct:.1f}%</span>
      </div>
      <div style="background:#ECF0F1;border-radius:10px;height:36px;overflow:hidden;">
        <div style="width:{total_bar_w:.1f}%;background:{total_color};height:100%;border-radius:10px;
                    display:flex;align-items:center;padding-left:10px;box-sizing:border-box;">
          {'<span style="color:white;font-size:13px;font-weight:700;">'+fmt_vnd(total_actual)+'</span>' if total_bar_w > 10 else ''}
        </div>
      </div>
      <div style="display:flex;justify-content:space-between;margin-top:4px;font-size:12px;color:#95A5A6;">
        <span>실적: {fmt_vnd(total_actual)} VND</span>
        <span>목표: {fmt_vnd(total_target)} VND</span>
      </div>
    </div>
    """

    subtitle = f'<p style="color:#7F8C8D;font-size:13px;margin:0 0 16px;">{date_label} 기준</p>' if date_label else ''
    total_pct_val = total_actual / total_target * 100 if total_target else 0
    under_color = '#E74C3C' if under_count > 0 else '#27AE60'
    pct_color = '#27AE60' if total_pct_val >= 100 else '#E67E22' if total_pct_val >= 70 else '#E74C3C'

    summary_html = f"""
    <div style="display:grid;grid-template-columns:repeat(4,1fr);gap:10px;margin-bottom:20px;">
      <div style="background:#1E1E2E;border-radius:10px;padding:16px;">
        <div style="font-size:12px;color:#95A5A6;margin-bottom:6px;">전체 달성률</div>
        <div style="font-size:24px;font-weight:700;color:{pct_color};">{total_pct_val:.1f}%</div>
      </div>
      <div style="background:#1E1E2E;border-radius:10px;padding:16px;">
        <div style="font-size:12px;color:#95A5A6;margin-bottom:6px;">실적 합계</div>
        <div style="font-size:24px;font-weight:700;color:#ECF0F1;">{fmt_vnd(total_actual)}</div>
      </div>
      <div style="background:#1E1E2E;border-radius:10px;padding:16px;">
        <div style="font-size:12px;color:#95A5A6;margin-bottom:6px;">목표 합계</div>
        <div style="font-size:24px;font-weight:700;color:#ECF0F1;">{fmt_vnd(total_target)}</div>
      </div>
      <div style="background:#1E1E2E;border-radius:10px;padding:16px;">
        <div style="font-size:12px;color:#95A5A6;margin-bottom:6px;">미달 카테고리</div>
        <div style="font-size:24px;font-weight:700;color:{under_color};">{under_count}개</div>
      </div>
    </div>
    """

    return f"""
    <div style="background:white;padding:24px;border-radius:14px;
                box-shadow:0 2px 12px rgba(0,0,0,0.08);margin-top:8px;">
      {subtitle}
      {summary_html}
      {items_html}
    </div>
    """


# ─── 앱 레이아웃 ───────────────────────────────────────────────────────────────
st.set_page_config(page_title="2026 매출 대시보드", layout="wide", page_icon="📊")

st.markdown("""
<style>
  .stApp { background: #F0F2F5; }
  section[data-testid="stSidebar"] { background: #2C3E50; }
  section[data-testid="stSidebar"] * { color: #ECF0F1 !important; }
  section[data-testid="stSidebar"] .stFileUploader label { color: #BDC3C7 !important; }
  .block-container { padding-top: 1.5rem; }
</style>
""", unsafe_allow_html=True)

st.title("📊 2026년 온라인 매출 달성 현황")

# ── 사이드바 ──────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("## 📂 데이터 업로드")
    st.markdown("---")
    uploaded = st.file_uploader("SKU 매출 엑셀 파일 (.xlsx)", type=['xlsx'], label_visibility='collapsed')

    if uploaded:
        file_bytes = uploaded.read()
        wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
        sheets = wb.sheetnames

        thang_sheets = [s for s in sheets if re.search(r'th[aá]ng', s, re.IGNORECASE)]

        if not thang_sheets:
            st.error("THÁNG XX 시트를 찾을 수 없습니다.")
        else:
            selected_sheet = st.selectbox("조회할 시트", thang_sheets)
            m_match = re.search(r'(\d+)', selected_sheet)
            month_num = int(m_match.group(1)) if m_match else None

            if month_num:
                default_date = f"{month_num}월 {datetime.now().day}일"
                date_label = st.text_input("기준일 입력", value=default_date)

                st.markdown("")
                if st.button("📊 분석 실행", use_container_width=True, type="primary"):
                    ws = wb[selected_sheet]
                    actual = parse_sheet(ws)
                    if not actual or all(v == 0 for v in actual.values()):
                        st.error("데이터를 읽지 못했습니다. 시트 구조를 확인해 주세요.")
                    else:
                        st.session_state['result'] = {
                            'month': month_num,
                            'date_label': date_label,
                            'actual': actual,
                        }
                        st.session_state['view_month'] = None
                        st.success(f"✅ {month_num}월 분석 완료!")

                if 'result' in st.session_state and st.session_state['result']['month'] == month_num:
                    st.markdown("")
                    if st.button("💾 이달 기록 저장", use_container_width=True):
                        r = st.session_state['result']
                        save_record(r['month'], r['date_label'], r['actual'])
                        st.success("저장 완료!")
                        st.rerun()

    st.markdown("---")
    st.markdown("### 📌 목표매출 기준\n2026년 연간 목표매출표")

# ── 월별 히스토리 버튼 ────────────────────────────────────────────────────────
history = load_history()

st.markdown("### 📅 월별 기록")
cols = st.columns(12)
for month, col in zip(range(1, 13), cols):
    with col:
        if month in history:
            pcts = list(history[month]['pct'].values())
            avg = sum(pcts) / len(pcts) if pcts else 0
            icon = "✅" if avg >= 90 else "🟡" if avg >= 70 else "🔴"
            label = f"{icon} {month}월"
        else:
            label = f"⬜ {month}월"

        if st.button(label, key=f"m{month}", use_container_width=True):
            st.session_state['view_month'] = month

st.markdown("---")

# ── 메인 결과 표시 ────────────────────────────────────────────────────────────
view_month = st.session_state.get('view_month')
result = st.session_state.get('result')

if view_month and view_month in history:
    rec = history[view_month]
    st.subheader(f"📈 {view_month}월 달성 현황")
    html = render_bars(rec['actual'], rec['target'], rec.get('date_label', ''))
    components.html(html, height=720, scrolling=False)

elif result:
    month = result['month']
    target = ANNUAL_TARGETS[month]
    st.subheader(f"📈 {month}월 달성 현황")
    html = render_bars(result['actual'], target, result['date_label'])
    components.html(html, height=720, scrolling=False)

else:
    st.info("왼쪽에서 파일을 업로드하고 분석을 실행하세요. 또는 위 월 버튼을 클릭하면 저장된 기록을 볼 수 있습니다.")
