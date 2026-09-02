import io
import datetime
import json
import numpy as np
import pandas as pd
import streamlit as st
import warnings
import urllib.request
from zoneinfo import ZoneInfo
import xlwt

warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')

# ==========================================================
# 0. 구글 시트 통신 엔진 🤖 
# ==========================================================
def load_from_cloud():
    try:
        if "WEB_APP_URL" not in st.secrets: 
            st.warning("⚠️ Secrets에 WEB_APP_URL이 없습니다. (저장 기능 꺼짐)")
            return False
        url = st.secrets["WEB_APP_URL"]
        req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
        with urllib.request.urlopen(req, timeout=10) as response:
            data = response.read().decode('utf-8')
            if not data or not data.strip():
                st.error("🚨 구글이 텅 빈 데이터를 보냈습니다.")
                return False
            try:
                parsed = json.loads(data)
                st.session_state['inventory_loaded'] = parsed.get('inventory_loaded', False)
                st.session_state['stock_seosan'] = parsed.get('stock_seosan', {})
                st.session_state['stock_yongma'] = parsed.get('stock_yongma', {})
                st.session_state['order_count'] = parsed.get('order_count', 0)
                st.session_state['history'] = parsed.get('history', [])
                return True
            except json.JSONDecodeError:
                st.error(f"🚨 구글 응답 해석 실패! 구글이 보낸 진짜 내용:\n{data[:500]}")
                return False
    except Exception as e: 
        st.error(f"🚨 DB 통신 자체 실패: {e}")
    return False

def save_to_cloud():
    try:
        if "WEB_APP_URL" not in st.secrets: return False
        url = st.secrets["WEB_APP_URL"]
        data = {
            'inventory_loaded': st.session_state.get('inventory_loaded', False),
            'stock_seosan': {str(k): int(v) for k, v in st.session_state.get('stock_seosan', {}).items()},
            'stock_yongma': {str(k): int(v) for k, v in st.session_state.get('stock_yongma', {}).items()},
            'order_count': int(st.session_state.get('order_count', 0)),
            'history': st.session_state.get('history', []),
            'last_updated': datetime.datetime.now(ZoneInfo("Asia/Seoul")).strftime("%Y-%m-%d %H:%M:%S")
        }
        json_payload = json.dumps(data, ensure_ascii=False).encode('utf-8')
        req = urllib.request.Request(url, data=json_payload, headers={'Content-Type': 'text/plain;charset=utf-8'})
        with urllib.request.urlopen(req, timeout=10) as response:
            res_text = response.read().decode('utf-8')
            if "SUCCESS" not in res_text:
                st.error(f"🚨 DB 저장 오류 응답: {res_text}")
            return "SUCCESS" in res_text
    except Exception as e: 
        st.error(f"🚨 DB 저장 통신 실패: {e}") 
        return False

def df_to_xls_bytes(df):
    wb = xlwt.Workbook(encoding='utf-8')
    ws = wb.add_sheet('Sheet1')
    for col_idx, col_name in enumerate(df.columns):
        ws.write(0, col_idx, str(col_name))
    for row_idx, row in enumerate(df.itertuples(index=False)):
        for col_idx, val in enumerate(row):
            if pd.isna(val) or val == "": ws.write(row_idx + 1, col_idx, "")
            elif isinstance(val, (int, float, np.integer, np.floating)): ws.write(row_idx + 1, col_idx, val)
            else: ws.write(row_idx + 1, col_idx, str(val))
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()

# ==========================================================
# 1. Web UI 구성 및 기본 세팅 
# ==========================================================
st.set_page_config(page_title="폴레드 주문분배 시스템", page_icon="🍶", layout="wide")
SIDEBAR_LOGO_URL = "https://cdn-pro-web-223-233.cdn-nhncommerce.com/poled0304_godomall_com/data/skin/front/db_poled_C/img/dimg/about_logo02.png"

st.title("🍶 MADE BY DS ")
st.caption("Seosan & Yongma Multi-Warehouse Engine (v9.7 - Rollback & Count Mode)")
st.markdown("---")

ALLOWED_8DIGIT_CODES = [
    '10101101', '10101102', '10101105', '10101106', '10101108', '10101109', '10101110', '10101111', 
    '10101112', '10101113', '10101103', '10101104', '10101107', '10101114', '10101115', '10101116', 
    '10111108', '10111110', '10111112', '10111106', '10102102', '10102101'
]

def clean_product_code(series):
    s = series.fillna("").astype(str).str.strip().str.replace(r'\.0$', '', regex=True)
    def remove_fake_zero(val):
        val_str = str(val).strip()
        if val_str.endswith('.0'): val_str = val_str[:-2]
        if val_str == "" or val_str.lower() == "nan": return ""
        if len(val_str) == 8 and val_str not in ALLOWED_8DIGIT_CODES and val_str[-1] == '0': return val_str[:-1]
        elif len(val_str) == 6 and val_str[-1] == '0': return val_str[:-1]
        return val_str
    return s.apply(remove_fake_zero)

def get_pack_stats(df):
    if df is None or df.empty or '주문번호' not in df.columns: return {'단포': 0, '단수합포': 0, '이종합포': 0}
    stats = {'단포': 0, '단수합포': 0, '이종합포': 0}
    for _, group in df.groupby('주문번호'):
        if group['제품코드'].nunique() > 1: stats['이종합포'] += 1
        elif group['수량'].sum() > 1: stats['단수합포'] += 1
        else: stats['단포'] += 1
    return stats

# ==========================================================
# 2. 세션 금고 
# ==========================================================
if 'inventory_loaded' not in st.session_state:
    st.session_state['inventory_loaded'] = False
    st.session_state['stock_seosan'] = {}
    st.session_state['stock_yongma'] = {}
    st.session_state['order_count'] = 0
    st.session_state['history'] = []
    if load_from_cloud(): st.toast("☁️ 마지막 작업 상태를 불러왔습니다!", icon="✅")

# ==========================================================
# 3. 사이드바 (재고 업로드)
# ==========================================================
with st.sidebar:
    st.image(SIDEBAR_LOGO_URL, width="stretch")
    st.markdown("---")
    st.header("🏢 1단계: 창고 재고 업로드")
    is_disabled = st.session_state['inventory_loaded']
    file_seosan = st.file_uploader("📂 서산창고 (.xls)", type=['xlsx', 'xls'], disabled=is_disabled)
    file_yongma = st.file_uploader("📂 용마창고 (.xls)", type=['xlsx', 'xls'], disabled=is_disabled)
    
    if st.button("📥 재고 확정", type="primary", disabled=is_disabled) and file_seosan and file_yongma:
        try:
            df_s = pd.read_excel(file_seosan, usecols="B,L", engine='xlrd' if file_seosan.name.endswith('.xls') else None)
            df_s.columns = ['제품코드', '재고수량']
            df_s['제품코드'] = clean_product_code(df_s['제품코드'])
            df_s['재고수량'] = pd.to_numeric(df_s['재고수량'], errors='coerce').fillna(0)
            st.session_state['stock_seosan'] = df_s[df_s['제품코드'] != ""].groupby('제품코드')['재고수량'].sum().to_dict()
            
            df_y = pd.read_excel(file_yongma, usecols="B,H", engine='xlrd' if file_yongma.name.endswith('.xls') else None)
            df_y.columns = ['제품코드', '재고수량']
            df_y['제품코드'] = clean_product_code(df_y['제품코드'])
            df_y['재고수량'] = pd.to_numeric(df_y['재고수량'], errors='coerce').fillna(0)
            st.session_state['stock_yongma'] = df_y[df_y['제품코드'] != ""].groupby('제품코드')['재고수량'].sum().to_dict()
            
            st.session_state['inventory_loaded'] = True
            if save_to_cloud():
                st.toast("✅ 구글 시트 저장 성공!", icon="☁️")
            st.rerun()
        except Exception as e: st.error(f"⚠️ 재고 로딩 에러: {e}")
                
    st.markdown("---")
    if st.button("🚨 당일 마감 & 초기화", type="secondary"):
        st.session_state.clear()
        save_to_cloud()
        st.rerun()

# ==========================================================
# 4. 메인 화면 (지휘관 대시보드)
# ==========================================================
c1, c2 = st.columns(2)
c1.info(f"🍶 **서산 잔여 품목:** {len(st.session_state['stock_seosan'])}개")
c2.info(f"🍶 **용마 잔여 품목:** {len(st.session_state['stock_yongma'])}개")

if not st.session_state['inventory_loaded']:
    st.warning("👈 좌측에서 재고를 먼저 등록해주세요.")
    st.stop()

if st.session_state['history']:
    st.markdown("---")
    st.subheader("📈 누적 배정 히스토리")
    st.dataframe(pd.DataFrame(st.session_state['history']), hide_index=True, width="stretch")

st.markdown("---")
st.header("📋 2단계: 발주서 분배 (실시간 대시보드)")
file_order = st.file_uploader(f"📑 발주서 ({st.session_state['order_count']+1}차 - .xlsx, .xls 가능)", type=['xlsx', 'xls'])

if 'latest_result' in st.session_state:
    res = st.session_state['latest_result']
    st.success(f"🎉 {res['order_cnt']}차 배정 완료! (서산 배정 목표 {res['target_seosan_boxes']}박스 중 {res['current_file_seosan_alloc']}박스 할당됨)")
    
    st.markdown("### 💾 개별 결과 파일 다운로드")
    d1, d2, d3, d4 = st.columns(4)
    with d1: st.download_button("🏢 서산창고 (.xls)", res['df_s_bytes'], f"{res['today_str']}_{res['order_cnt']}차_서산.xls", "application/vnd.ms-excel", use_container_width=True)
    with d2: st.download_button("🏢 용마창고 (.xls)", res['df_y_bytes'], f"{res['today_str']}_{res['order_cnt']}차_용마.xls", "application/vnd.ms-excel", use_container_width=True)
    with d3: st.download_button("⚠️ 미배정 (.xls)", res['df_un_bytes'], f"{res['today_str']}_{res['order_cnt']}차_미배정.xls", "application/vnd.ms-excel", use_container_width=True)
    with d4: st.download_button("📊 모니터링 (.xls)", res['df_mon_bytes'], f"{res['today_str']}_{res['order_cnt']}차_모니터링.xls", "application/vnd.ms-excel", use_container_width=True)
    
    st.markdown("---")
    st.subheader(f"📊 {res['order_cnt']}차 포장 유형 분석")
    rc1, rc2, rc3 = st.columns(3)
    with rc1: st.write("**📑 전체**"); st.write(f"단포: `{res['total_stats']['단포']}` / 단수: `{res['total_stats']['단수합포']}` / 이종: `{res['total_stats']['이종합포']}`")
    with rc2: st.write("**🏢 서산**"); st.write(f"단포: `{res['s_stats']['단포']}` / 단수: `{res['s_stats']['단수합포']}` / 이종: `{res['s_stats']['이종합포']}`")
    with rc3: st.write("**🏢 용마**"); st.write(f"단포: `{res['y_stats']['단포']}` / 단수: `{res['y_stats']['단수합포']}` / 이종: `{res['y_stats']['이종합포']}`")
    
    st.markdown("---")
    btn_col1, btn_col2 = st.columns(2)
    with btn_col1:
        if st.button("✅ 배정 확정 (다음 차수 준비)", type="primary", use_container_width=True):
            del st.session_state['latest_result']
            if 'snapshot' in st.session_state:
                del st.session_state['snapshot']
            st.rerun()
            
    with btn_col2:
        if st.button("🔙 배정 취소 (결과 롤백 후 재배정)", type="secondary", use_container_width=True):
            if 'snapshot' in st.session_state:
                st.session_state['stock_seosan'] = st.session_state['snapshot']['stock_seosan'].copy()
                st.session_state['stock_yongma'] = st.session_state['snapshot']['stock_yongma'].copy()
                st.session_state['history'] = st.session_state['snapshot']['history'].copy()
                st.session_state['order_count'] = st.session_state['snapshot']['order_count']
                del st.session_state['snapshot']
            
            del st.session_state['latest_result']
            save_to_cloud()
            st.toast("🔙 배정이 취소되고 재고가 원래대로 복구되었습니다!", icon="🔄")
            st.rerun()

elif file_order:
    try: orders_df = pd.read_excel(file_order, engine='xlrd' if file_order.name.endswith('.xls') else None)
    except: orders_df = pd.read_excel(file_order)

    orders_df.columns = orders_df.columns.str.strip()
    orig_columns = orders_df.columns.tolist()
    qty_col_name = orig_columns[18]
    
    col_B_name = orig_columns[1]; col_A_name = orig_columns[0]
    orders_df[col_B_name] = orders_df[col_B_name].astype(str).str.replace(r'_사은품.*', '', regex=True).str.strip()
    orders_df[col_A_name] = orders_df[col_A_name].astype(str).str.replace(r'_사은품.*', '', regex=True).str.strip()
    
    col_A_str = orders_df[col_A_name]; col_B_str = orders_df[col_B_name]
    is_type1 = col_A_str.str.contains(r'\d{6}[a-zA-Z]{2}\d{3}', na=False, regex=True)
    orders_df['주문번호'] = np.where(is_type1, col_A_str, col_B_str)
    
    orig_pcode_col_name = orig_columns[9]
    orders_df[orig_pcode_col_name] = clean_product_code(orders_df.iloc[:, 9])
    orders_df['제품코드'] = orders_df[orig_pcode_col_name]
    orders_df['수량'] = pd.to_numeric(orders_df.iloc[:, 18], errors='coerce').fillna(0)
    
    orders_df = orders_df[orders_df['제품코드'] != ""].reset_index(drop=True)
    orders_df['_orig_idx'] = orders_df.index
    
    grouped = list(orders_df.groupby('주문번호', sort=False))
    total_boxes = len(grouped)
    
    cat_counts = {'단포': 0, '동종': 0, '이종': 0}
    ordered_groups = []
    
    for oid, group in grouped:
        items = group.to_dict('records')
        reqs_check = set(it['제품코드'] for it in items)
        total_qty = sum(it['수량'] for it in items)
        if len(reqs_check) > 1: cat = '이종'
        elif total_qty > 1: cat = '동종'
        else: cat = '단포'
        cat_counts[cat] += 1
        ordered_groups.append((oid, items, cat))

    st.success(f"📊 **현재 발주서 요약:** 총 **{total_boxes}** 박스 (단포: `{cat_counts['단포']}` / 동종합포: `{cat_counts['동종']}` / 이종합포: `{cat_counts['이종']}`)")

    with st.form("allocation_form"):
        st.markdown("### 🎯 서산창고 배정 타겟 설정")
        st.caption("설정된 타겟만큼 **[단포 ➔ 동종합포 ➔ 이종합포]** 순서로 서산에 우선 배정, 나머지는 용마로 배정됩니다.")
        t_col1, t_col2 = st.columns(2)
        with t_col1: target_mode = st.radio("설정 방식", ["건수(박스)로 설정", "비율(%)로 설정"], horizontal=True)
        with t_col2: target_val = st.number_input("목표 값 입력 (건수는 박스 수, 비율은 %)", min_value=0, value=total_boxes, step=10)
        submitted = st.form_submit_button("🚀 자동 분배 실행", type="primary")

    if submitted:
        with st.spinner("지휘관 목표치에 맞춰 최적화 배정 중..."):
            
            st.session_state['snapshot'] = {
                'stock_seosan': st.session_state['stock_seosan'].copy(),
                'stock_yongma': st.session_state['stock_yongma'].copy(),
                'history': st.session_state['history'].copy(),
                'order_count': st.session_state['order_count']
            }
            
            if target_mode == "비율(%)로 설정": target_seosan_boxes = int(total_boxes * (target_val / 100.0))
            else: target_seosan_boxes = int(target_val)
                
            total_stats = get_pack_stats(orders_df)

            def sort_key(x):
                if x[2] == '단포': return 1
                elif x[2] == '동종': return 2
                else: return 3
            ordered_groups.sort(key=sort_key)
            
            temp_s = st.session_state['stock_seosan'].copy()
            temp_y = st.session_state['stock_yongma'].copy()
            results_map = {}
            current_file_seosan_alloc = 0 
            
            for oid, items, cat in ordered_groups:
                reqs = {}
                for it in items: reqs[it['제품코드']] = reqs.get(it['제품코드'], 0) + it['수량']
                
                if current_file_seosan_alloc < target_seosan_boxes: base_pri = '서산'
                else: base_pri = '용마'
                    
                pri_name = '서산' if base_pri == '서산' else '용마'
                sec_name = '용마' if base_pri == '서산' else '서산'
                pri_stock = temp_s if base_pri == '서산' else temp_y
                sec_stock = temp_y if base_pri == '서산' else temp_s
                
                if all(pri_stock.get(it['제품코드'], 0) >= it['수량'] for it in items):
                    for it in items:
                        pc, q, idx = it['제품코드'], it['수량'], it['_orig_idx']
                        pri_stock[pc] = pri_stock.get(pc, 0) - q
                        results_map[idx] = {'주문번호': oid, '제품코드': pc, '수량': q, '서산배정': q if pri_name=='서산' else 0, '용마배정': q if pri_name=='용마' else 0, '상태': f'{pri_name} 완배'}
                    if pri_name == '서산': current_file_seosan_alloc += 1
                    continue
                    
                if all(sec_stock.get(it['제품코드'], 0) >= it['수량'] for it in items):
                    for it in items:
                        pc, q, idx = it['제품코드'], it['수량'], it['_orig_idx']
                        sec_stock[pc] = sec_stock.get(pc, 0) - q
                        results_map[idx] = {'주문번호': oid, '제품코드': pc, '수량': q, '서산배정': q if sec_name=='서산' else 0, '용마배정': q if sec_name=='용마' else 0, '상태': f'{sec_name} 완배'}
                    if sec_name == '서산': current_file_seosan_alloc += 1
                    continue
                    
                if all(reqs[pc] <= (temp_s.get(pc, 0) + temp_y.get(pc, 0)) for pc in reqs):
                    seosan_contributed = False
                    for it in items:
                        pc, q, idx = it['제품코드'], it['수량'], it['_orig_idx']
                        av_s, av_y = temp_s.get(pc, 0), temp_y.get(pc, 0)
                        if base_pri == '서산':
                            t_s = min(q, av_s); temp_s[pc] = temp_s.get(pc, 0) - t_s
                            t_y = q - t_s; temp_y[pc] = temp_y.get(pc, 0) - t_y
                        else:
                            t_y = min(q, av_y); temp_y[pc] = temp_y.get(pc, 0) - t_y
                            t_s = q - t_y; temp_s[pc] = temp_s.get(pc, 0) - t_s
                        results_map[idx] = {'주문번호': oid, '제품코드': pc, '수량': q, '서산배정': t_s, '용마배정': t_y, '상태': '분할배정'}
                        if t_s > 0: seosan_contributed = True
                    if seosan_contributed: current_file_seosan_alloc += 1
                else:
                    for it in items:
                        idx = it['_orig_idx']
                        pc = it['제품코드']
                        total_avail = temp_s.get(pc, 0) + temp_y.get(pc, 0)
                        reason_str = '실재고부족' if reqs[pc] > total_avail else '합배송품절'
                        results_map[idx] = {'주문번호': oid, '제품코드': pc, '수량': it['수량'], '서산배정': 0, '용마배정': 0, '상태': reason_str}
            
            results_list = [results_map[i] for i in range(len(orders_df))]
            st.session_state['stock_seosan'] = temp_s
            st.session_state['stock_yongma'] = temp_y
            st.session_state['order_count'] += 1
            
            list_s, list_y, list_un = [], [], []
            for i, row in enumerate(results_list):
                orig_row = orders_df.iloc[i].to_dict()
                if row['서산배정'] > 0:
                    r = orig_row.copy(); r[qty_col_name] = row['서산배정']; r['수량'] = row['서산배정']; list_s.append(r)
                if row['용마배정'] > 0:
                    r = orig_row.copy(); r[qty_col_name] = row['용마배정']; r['수량'] = row['용마배정']; list_y.append(r)
                if row['서산배정'] == 0 and row['용마배정'] == 0:
                    r = orig_row.copy(); r['[사유]'] = row['상태']; list_un.append(r)
            
            df_s = pd.DataFrame(list_s) if list_s else pd.DataFrame(columns=orders_df.columns)
            df_y = pd.DataFrame(list_y) if list_y else pd.DataFrame(columns=orders_df.columns)
            df_un = pd.DataFrame(list_un) if list_un else pd.DataFrame(columns=orders_df.columns.tolist() + ['[사유]'])
            
            s_stats = get_pack_stats(df_s); y_stats = get_pack_stats(df_y)
            
            st.session_state['history'].append({
                '차수': f"{st.session_state['order_count']}차",
                '서산 단포': s_stats['단포'], '서산 단수합포': s_stats['단수합포'], '서산 이종합포': s_stats['이종합포'],
                '용마 단포': y_stats['단포'], '용마 단수합포': y_stats['단수합포'], '용마 이종합포': y_stats['이종합포'],
                '미배정': df_un['주문번호'].nunique() if not df_un.empty else 0
            })
            
            save_to_cloud()
            
            st.session_state['latest_result'] = {
                'df_s_bytes': df_to_xls_bytes(df_s[orig_columns] if not df_s.empty else df_s),
                'df_y_bytes': df_to_xls_bytes(df_y[orig_columns] if not df_y.empty else df_y),
                'df_un_bytes': df_to_xls_bytes(df_un),
                'df_mon_bytes': df_to_xls_bytes(pd.DataFrame(results_list)),
                'target_seosan_boxes': target_seosan_boxes,
                'current_file_seosan_alloc': current_file_seosan_alloc,
                'total_stats': total_stats,
                's_stats': s_stats,
                'y_stats': y_stats,
                'order_cnt': st.session_state['order_count'],
                'today_str': datetime.datetime.now(ZoneInfo("Asia/Seoul")).strftime("%m%d")
            }
            st.rerun()