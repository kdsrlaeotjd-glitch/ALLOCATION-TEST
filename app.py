import io
import zipfile
import datetime
import json
import numpy as np
import pandas as pd
import streamlit as st
import warnings
import urllib.request
import urllib.parse
from zoneinfo import ZoneInfo

warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')

# ==========================================================
# 0. 구글 시트 통신 및 .xls 파일 생성 엔진 🤖
# ==========================================================
def load_from_cloud():
    try:
        if "WEB_APP_URL" not in st.secrets:
            return False
        url = st.secrets["WEB_APP_URL"]
        req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
        with urllib.request.urlopen(req, timeout=10) as response:
            data = response.read().decode('utf-8')
            if data and data.strip():
                parsed = json.loads(data)
                st.session_state['inventory_loaded'] = parsed.get('inventory_loaded', False)
                st.session_state['stock_seosan'] = parsed.get('stock_seosan', {})
                st.session_state['stock_yongma'] = parsed.get('stock_yongma', {})
                st.session_state['order_count'] = parsed.get('order_count', 0)
                st.session_state['history'] = parsed.get('history', [])
                return True
    except Exception:
        pass
    return False

def save_to_cloud():
    try:
        if "WEB_APP_URL" not in st.secrets:
            st.error("🚨 Streamlit Secrets에 'WEB_APP_URL'이 설정되어 있지 않습니다.")
            return False
            
        url = st.secrets["WEB_APP_URL"]
        s_dict = {str(k): int(v) for k, v in st.session_state.get('stock_seosan', {}).items()}
        y_dict = {str(k): int(v) for k, v in st.session_state.get('stock_yongma', {}).items()}
        
        data = {
            'inventory_loaded': st.session_state.get('inventory_loaded', False),
            'stock_seosan': s_dict,
            'stock_yongma': y_dict,
            'order_count': int(st.session_state.get('order_count', 0)),
            'history': st.session_state.get('history', []),
            'last_updated': datetime.datetime.now(ZoneInfo("Asia/Seoul")).strftime("%Y-%m-%d %H:%M:%S")
        }
        json_payload = json.dumps(data, ensure_ascii=False).encode('utf-8')
        
        req = urllib.request.Request(url, data=json_payload, headers={'Content-Type': 'text/plain;charset=utf-8'})
        with urllib.request.urlopen(req, timeout=10) as response:
            res_text = response.read().decode('utf-8')
            if "SUCCESS" in res_text:
                return True
            else:
                st.error(f"🚨 구글 시트 저장 서버 응답: {res_text}")
                return False
    except Exception as e:
        st.error(f"🚨 구글 시트 저장 실패: {repr(e)}")
        return False

def df_to_xls_bytes(df):
    """DataFrame을 WMS 및 MS Excel 호환 .xls 바이너리 스트림으로 변환"""
    xml_str = '<?xml version="1.0" encoding="utf-8"?>\n'
    xml_str += '<?mso-application progid="Excel.Sheet"?>\n'
    xml_str += '<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"\n'
    xml_str += ' xmlns:o="urn:schemas-microsoft-com:office:office"\n'
    xml_str += ' xmlns:x="urn:schemas-microsoft-com:office:excel"\n'
    xml_str += ' xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet">\n'
    xml_str += ' <Worksheet ss:Name="Sheet1">\n'
    xml_str += '  <Table>\n'
    
    xml_str += '   <Row>\n'
    for col in df.columns:
        clean_col = str(col).replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
        xml_str += f'    <Cell><Data ss:Type="String">{clean_col}</Data></Cell>\n'
    xml_str += '   </Row>\n'
    
    for _, row in df.iterrows():
        xml_str += '   <Row>\n'
        for val in row:
            if pd.isna(val) or val is None:
                cell_str = ''
                cell_type = 'String'
            elif isinstance(val, (int, float, np.integer, np.floating)):
                cell_str = str(val)
                cell_type = 'Number'
            else:
                cell_str = str(val).replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
                cell_type = 'String'
            xml_str += f'    <Cell><Data ss:Type="{cell_type}">{cell_str}</Data></Cell>\n'
        xml_str += '   </Row>\n'
        
    xml_str += '  </Table>\n'
    xml_str += ' </Worksheet>\n'
    xml_str += '</Workbook>'
    
    return xml_str.encode('utf-8')

# ==========================================================
# 1. Web UI 구성 및 기본 세팅 (v8.1 - CAPA 1400 Routing)
# ==========================================================
st.set_page_config(page_title="폴레드 주문분배 시스템", page_icon="🍶", layout="wide")

SIDEBAR_LOGO_URL = "https://cdn-pro-web-223-233.cdn-nhncommerce.com/poled0304_godomall_com/data/skin/front/db_poled_C/img/dimg/about_logo02.png"

st.title("🍶 MADE BY DS ")
st.caption("Seosan & Yongma Multi-Warehouse Allocation Engine (v8.1 - Dynamic CAPA 1400)")
st.markdown("---")

ALLOWED_8DIGIT_CODES = [
    '10101101', '10101102', '10101105', '10101106', '10101108',
    '10101109', '10101110', '10101111', '10101112', '10101113',
    '10101103', '10101104', '10101107', '10101114', '10101115',
    '10101116', '10111108', '10111110', '10111112', '10111106',
    '10102102', '10102101'
]

def clean_product_code(series):
    s = series.fillna("").astype(str).str.strip()
    s = s.str.replace(r'\.0$', '', regex=True)
    def remove_fake_zero(val):
        val_str = str(val).strip()
        if val_str.endswith('.0'): val_str = val_str[:-2]
        if val_str == "" or val_str.lower() == "nan": return ""
        if len(val_str) == 8 and val_str not in ALLOWED_8DIGIT_CODES:
            if val_str[-1] == '0': return val_str[:-1]
        elif len(val_str) == 6:
            if val_str[-1] == '0': return val_str[:-1]
        return val_str
    return s.apply(remove_fake_zero)

def get_pack_stats(df):
    if df is None or df.empty or '주문번호' not in df.columns: return {'단포': 0, '단수합포': 0, '이종합포': 0}
    needed = ['주문번호', '제품코드', '수량']
    for col in needed:
        if col not in df.columns: return {'단포': 0, '단수합포': 0, '이종합포': 0}
    grouped = df.groupby('주문번호')
    stats = {'단포': 0, '단수합포': 0, '이종합포': 0}
    for _, group in grouped:
        sku_cnt = group['제품코드'].nunique(); total_qty = group['수량'].sum()
        if sku_cnt > 1: stats['이종합포'] += 1
        elif total_qty > 1: stats['단수합포'] += 1
        else: stats['단포'] += 1
    return stats

# ==========================================================
# 2. 세션 금고 & 클라우드 자동 불러오기
# ==========================================================
if 'inventory_loaded' not in st.session_state:
    st.session_state['inventory_loaded'] = False
    st.session_state['stock_seosan'] = {}
    st.session_state['stock_yongma'] = {}
    st.session_state['order_count'] = 0
    st.session_state['history'] = []
    
    if load_from_cloud():
        st.toast("☁️ 구글 시트(DB)에서 마지막 작업 상태를 불러왔습니다!", icon="✅")

# ==========================================================
# 3. 사이드바 (CAPA 기능 1400건 적용)
# ==========================================================
with st.sidebar:
    st.image(SIDEBAR_LOGO_URL, width="stretch")
    st.markdown("---")
    
    # 💡 [핵심] 일일 서산 CAPA(한도) 컨트롤러 (기본값 1400)
    st.header("🎯 서산창고 일일 CAPA 설정")
    seosan_capa = st.number_input("서산 하루 최대 배정(건)", value=1400, step=100)
    
    # 현재까지 누적된 서산 주문건수(포장건수) 계산
    current_seosan_alloc = sum(h.get('서산 단포', 0) + h.get('서산 단수합포', 0) + h.get('서산 이종합포', 0) for h in st.session_state['history'])
    
    progress_val = min(current_seosan_alloc / seosan_capa, 1.0) if seosan_capa > 0 else 0.0
    st.progress(progress_val)
    if current_seosan_alloc >= seosan_capa:
        st.error(f"📦 현재 누적 배정량: **{current_seosan_alloc}** / {seosan_capa}건 (CAPA 초과!)")
    else:
        st.caption(f"📦 현재 누적 배정량: **{current_seosan_alloc}** / {seosan_capa}건")

    st.markdown("---")
    st.header("🏢 1단계: 창고 재고 업로드")
    
    is_disabled = st.session_state['inventory_loaded']
    file_seosan = st.file_uploader("📂 서산창고 (.xlsx, .xls 가능)", type=['xlsx', 'xls'], disabled=is_disabled)
    file_yongma = st.file_uploader("📂 용마창고 (.xlsx, .xls 가능)", type=['xlsx', 'xls'], disabled=is_disabled)
    
    if st.button("📥 재고 확정", type="primary", disabled=is_disabled):
        if file_seosan and file_yongma:
            try:
                df_s_check = pd.read_excel(file_seosan, nrows=0, engine='xlrd' if file_seosan.name.endswith('.xls') else None)
                if '제품코드' in df_s_check.columns and '재고수량' in df_s_check.columns:
                    df_s = pd.read_excel(file_seosan, usecols=['제품코드', '재고수량'], engine='xlrd' if file_seosan.name.endswith('.xls') else None)
                else:
                    df_s = pd.read_excel(file_seosan, usecols="B,L", engine='xlrd' if file_seosan.name.endswith('.xls') else None)
                    df_s.columns = ['제품코드', '재고수량']
                df_s['제품코드'] = clean_product_code(df_s['제품코드'])
                df_s['재고수량'] = pd.to_numeric(df_s['재고수량'], errors='coerce').fillna(0)
                df_s = df_s[df_s['제품코드'] != ""]
                st.session_state['stock_seosan'] = df_s.groupby('제품코드')['재고수량'].sum().to_dict()
                
                df_y_check = pd.read_excel(file_yongma, nrows=0, engine='xlrd' if file_yongma.name.endswith('.xls') else None)
                if '제품코드' in df_y_check.columns and '재고수량' in df_y_check.columns:
                    df_y = pd.read_excel(file_yongma, usecols=['제품코드', '재고수량'], engine='xlrd' if file_yongma.name.endswith('.xls') else None)
                else:
                    df_y = pd.read_excel(file_yongma, usecols="B,H", engine='xlrd' if file_yongma.name.endswith('.xls') else None)
                    df_y.columns = ['제품코드', '재고수량']
                df_y['제품코드'] = clean_product_code(df_y['제품코드'])
                df_y['재고수량'] = pd.to_numeric(df_y['재고수량'], errors='coerce').fillna(0)
                df_y = df_y[df_y['제품코드'] != ""]
                st.session_state['stock_yongma'] = df_y.groupby('제품코드')['재고수량'].sum().to_dict()
                
                st.session_state['inventory_loaded'] = True
                
                if save_to_cloud():
                    st.toast("☁️ 재고 데이터가 클라우드 DB에 무사히 저장되었습니다!", icon="✅")
                    
                st.success("✅ 재고 등록 완료!")
                st.rerun()
            except Exception as e:
                st.error(f"⚠️ 재고 로딩 에러: {e}")
                
    if st.session_state['inventory_loaded']:
        st.markdown("---")
        st.header("💾 잔여 재고 수동 백업")
        df_s_bk = pd.DataFrame(list(st.session_state['stock_seosan'].items()), columns=['제품코드', '재고수량'])
        df_y_bk = pd.DataFrame(list(st.session_state['stock_yongma'].items()), columns=['제품코드', '재고수량'])
        bk_zip = io.BytesIO()
        with zipfile.ZipFile(bk_zip, 'w', zipfile.ZIP_DEFLATED) as zf:
            zf.writestr("백업_서산창고.xls", df_to_xls_bytes(df_s_bk))
            zf.writestr("백업_용마창고.xls", df_to_xls_bytes(df_y_bk))
        st.download_button("💾 이중 백업 (ZIP) 다운로드", bk_zip.getvalue(), f"잔여재고_수동백업_{datetime.datetime.now(ZoneInfo('Asia/Seoul')).strftime('%m%d_%H%M')}.zip", "application/zip", type="secondary")

    st.markdown("---")
    if st.button("🚨 당일 마감 & 초기화", type="secondary"):
        st.session_state.clear()
        save_to_cloud()
        st.success("🔄 초기화 및 클라우드 청소 완료.")
        st.rerun()

# ==========================================================
# 4. 메인 화면
# ==========================================================
c1, c2 = st.columns(2)
c1.info(f"🍶 **서산 잔여 품목:** {len(st.session_state['stock_seosan'])}개")
c2.info(f"🍶 **용마 잔여 품목:** {len(st.session_state['stock_yongma'])}개")

if not st.session_state['inventory_loaded']:
    st.warning("👈 좌측에서 재고를 먼저 등록해주세요.")
    st.stop()

st.header("📋 2단계: 발주서 분배 (연속 차감)")

file_order = st.file_uploader(f"📑 발주서 ({st.session_state['order_count']+1}차 - .xlsx, .xls 가능)", type=['xlsx', 'xls'])

current_kst_time = datetime.datetime.now(ZoneInfo("Asia/Seoul"))
is_morning = current_kst_time.hour < 12

if is_morning:
    # --- [오전 12시 이전: 스마트 혼합 배정 모드] ---
    priority_options = [
        '서산창고 우선 (모든 건 서산 ➔ 용마)', 
        '용마창고 우선 (모든 건 용마 ➔ 서산)', 
        '스마트 혼합 (단포·단수: 서산우선 / 이종: 용마우선)'
    ]
    default_priority_idx = 1 # 기본값 용마
    
    if file_order:
        if "프랭클린" in file_order.name:
            try:
                temp_df = pd.read_excel(file_order, usecols="Q", engine='xlrd' if file_order.name.endswith('.xls') else None)
                unique_cnt = temp_df.iloc[:, 0].nunique()
                
                if unique_cnt >= 1200:
                    default_priority_idx = 2  # 스마트 혼합
                    st.info(f"💡 **[오전 모드] '프랭클린' 파일 감지! (Q열 고유: {unique_cnt}건)** ➔ 1,200건 이상이므로 **[스마트 혼합]**이 자동 선택되었습니다.")
                else:
                    default_priority_idx = 0  # 서산 우선
                    st.info(f"💡 **[오전 모드] '프랭클린' 파일 감지! (Q열 고유: {unique_cnt}건)** ➔ 1,200건 미만이므로 **[서산창고 우선]**이 자동 선택되었습니다.")
            except Exception as e:
                st.warning("⚠️ 데이터 분석 실패 ➔ 기본값(용마 우선) 세팅")
            finally:
                file_order.seek(0)
        else:
            default_priority_idx = 1
            st.info("💡 **[오전 모드] 일반 발주서 감지 (프랭클린 아님)** ➔ **[용마창고 우선]**이 자동 선택되었습니다.")
else:
    # --- [오후 12시 이후: 심플 수동 배정 모드] ---
    st.info("🕒 **오후 12시가 지나 [수동 심플 배정] 모드로 전환되었습니다.** (스마트 자동 선택 해제)")
    priority_options = [
        '서산창고 우선 (모든 건 서산 ➔ 용마)', 
        '용마창고 우선 (모든 건 용마 ➔ 서산)'
    ]
    default_priority_idx = 1

priority_choice = st.radio("🍶 **우선 순위 설정:**", priority_options, index=default_priority_idx, horizontal=True)


if file_order and st.button("🚀 자동 분배 실행", type="primary"):
    try:
        with st.spinner("배정 로직 가동 중... (실시간 CAPA 감지 중)"):
            try: orders_df = pd.read_excel(file_order, engine='xlrd' if file_order.name.endswith('.xls') else None)
            except: orders_df = pd.read_excel(file_order)

            orders_df.columns = orders_df.columns.str.strip()
            orig_columns = orders_df.columns.tolist()
            qty_col_name = orig_columns[18]
            
            col_B_name = orig_columns[1]; col_A_name = orig_columns[0]
            orders_df[col_B_name] = orders_df[col_B_name].astype(str).str.replace(r'_사은품.*', '', regex=True).str.strip()
            orders_df[col_A_name] = orders_df[col_A_name].astype(str).str.replace(r'_사은품.*', '', regex=True).str.strip()
            
            col_A_str = orders_df[col_A_name]; col_B_str = orders_df[col_B_name]
            pattern = r'\d{6}[a-zA-Z]{2}\d{3}'
            is_type1 = col_A_str.str.contains(pattern, na=False, regex=True)
            orders_df['주문번호'] = np.where(is_type1, col_A_str, col_B_str)
            
            orig_pcode_col_name = orig_columns[9]
            orders_df[orig_pcode_col_name] = clean_product_code(orders_df.iloc[:, 9])
            orders_df['제품코드'] = orders_df[orig_pcode_col_name]
            orders_df['수량'] = pd.to_numeric(orders_df.iloc[:, 18], errors='coerce').fillna(0)
            
            orders_df = orders_df[orders_df['제품코드'] != ""].reset_index(drop=True)
            orders_df['_orig_idx'] = orders_df.index
            
            total_stats = get_pack_stats(orders_df)
            
            temp_s = st.session_state['stock_seosan'].copy()
            temp_y = st.session_state['stock_yongma'].copy()
            results_map = {}
                
            grouped = list(orders_df.groupby('주문번호', sort=False))
            
            # 💡 [핵심] 현재 작업 중 실시간으로 서산 CAPA를 트래킹하기 위한 변수
            running_seosan_alloc = current_seosan_alloc 
            capa_routed_count = 0 # 한도 초과로 우회된 건수
            
            # 💡 [Order-by-Order 딥러닝급 순차 처리 로직]
            for oid, group in grouped:
                items = group.to_dict('records')
                reqs = {}
                for it in items: reqs[it['제품코드']] = reqs.get(it['제품코드'], 0) + it['수량']
                
                # 1. 사용자가 선택한 라디오 버튼에 따른 '기본 우선순위' 판단
                if "서산창고 우선" in priority_choice:
                    base_pri = '서산'
                elif "용마창고 우선" in priority_choice:
                    base_pri = '용마'
                else:
                    is_multi_sku = len(reqs) > 1
                    base_pri = '용마' if is_multi_sku else '서산'
                
                # 2. [CAPA 1,400건 강제 우회 로직] 
                # 서산으로 가야 할 운명인데 한도(1400건)가 넘었으면 무조건 용마로 꺾어버림!
                if base_pri == '서산' and running_seosan_alloc >= seosan_capa:
                    curr_pri = '용마'
                    capa_routed_count += 1
                else:
                    curr_pri = base_pri
                    
                pri_name = '서산' if curr_pri == '서산' else '용마'
                sec_name = '용마' if curr_pri == '서산' else '서산'
                pri_stock = temp_s if curr_pri == '서산' else temp_y
                sec_stock = temp_y if curr_pri == '서산' else temp_s
                
                # 시도 1: 우선 창고 완배정
                if all(pri_stock.get(it['제품코드'], 0) >= it['수량'] for it in items):
                    for it in items:
                        pc, q, idx = it['제품코드'], it['수량'], it['_orig_idx']
                        pri_stock[pc] -= q
                        res = {'주문번호': oid, '제품코드': pc, '수량': q, '서산배정': 0, '용마배정': 0, '상태': f'{pri_name} 완배'}
                        res[f'{pri_name}배정'] = q
                        results_map[idx] = res
                    if pri_name == '서산': running_seosan_alloc += 1
                    continue
                    
                # 시도 2: 2순위 창고(대체 창고) 완배정
                if all(sec_stock.get(it['제품코드'], 0) >= it['수량'] for it in items):
                    for it in items:
                        pc, q, idx = it['제품코드'], it['수량'], it['_orig_idx']
                        sec_stock[pc] -= q
                        res = {'주문번호': oid, '제품코드': pc, '수량': q, '서산배정': 0, '용마배정': 0, '상태': f'{sec_name} 완배'}
                        res[f'{sec_name}배정'] = q
                        results_map[idx] = res
                    if sec_name == '서산': running_seosan_alloc += 1
                    continue
                    
                # 시도 3: 분할 배정 (양쪽 다 끌어모으기) 또는 품절
                if all(reqs[pc] <= (temp_s.get(pc, 0) + temp_y.get(pc, 0)) for pc in reqs):
                    for it in items:
                        pc, q, idx = it['제품코드'], it['수량'], it['_orig_idx']
                        av_s, av_y = temp_s.get(pc, 0), temp_y.get(pc, 0)
                        if curr_pri == '서산':
                            t_s = min(q, av_s); temp_s[pc] = av_s - t_s; t_y = q - t_s; temp_y[pc] = av_y - t_y
                        else:
                            t_y = min(q, av_y); temp_y[pc] = av_y - t_y; t_s = q - t_y; temp_s[pc] = av_s - t_s
                        results_map[idx] = {'주문번호': oid, '제품코드': pc, '수량': q, '서산배정': t_s, '용마배정': t_y, '상태': '분할배정'}
                    # 서산에서 1개라도 나갔으므로 건수 증가 처리 (안전빵)
                    running_seosan_alloc += 1
                else:
                    for it in items:
                        idx = it['_orig_idx']
                        pc = it['제품코드']
                        total_avail = temp_s.get(pc, 0) + temp_y.get(pc, 0)
                        reason_str = '실재고부족' if reqs[pc] > total_avail else '합배송품절'
                        results_map[idx] = {'주문번호': oid, '제품코드': it['제품코드'], '수량': it['수량'], '서산배정': 0, '용마배정': 0, '상태': reason_str}
            
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
            
            if save_to_cloud():
                st.toast("☁️ 변경된 재고량이 클라우드 DB에 무사히 저장되었습니다!", icon="💾")
            
            today_str = datetime.datetime.now(ZoneInfo("Asia/Seoul")).strftime("%m%d")
            order_cnt = st.session_state['order_count']
            
            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zf:
                for fn_label, dfd in [
                    (f"{today_str}_{order_cnt}차 서산.xls", df_s[orig_columns] if not df_s.empty else df_s), 
                    (f"{today_str}_{order_cnt}차 용마.xls", df_y[orig_columns] if not df_y.empty else df_y), 
                    (f"{today_str}_{order_cnt}차 미배정.xls", df_un), 
                    (f"{today_str}_{order_cnt}차 모니터링.xls", pd.DataFrame(results_list))
                ]:
                    zf.writestr(fn_label, df_to_xls_bytes(dfd))
            
            st.success(f"🎉 {st.session_state['order_count']}차 배정 완료!")
            
            # 💡 한도 초과 우회 안내 메시지 띄우기
            if capa_routed_count > 0:
                st.warning(f"🚨 **CAPA 알림:** 작업 도중 서산창고 일일 한도({seosan_capa}건)를 초과하여, **{capa_routed_count}건의 주문이 용마창고로 자동 우회(Routing)** 배정되었습니다!")
            
            st.subheader(f"📊 {st.session_state['order_count']}차 포장 유형 분석")
            rc1, rc2, rc3 = st.columns(3)
            with rc1: st.write("**📑 전체**"); st.write(f"단포: `{total_stats['단포']}` / 단수: `{total_stats['단수합포']}` / 이종: `{total_stats['이종합포']}`")
            with rc2: st.write("**🏢 서산**"); st.write(f"단포: `{s_stats['단포']}` / 단수: `{s_stats['단수합포']}` / 이종: `{s_stats['이종합포']}`")
            with rc3: st.write("**🏢 용마**"); st.write(f"단포: `{y_stats['단포']}` / 단수: `{y_stats['단수합포']}` / 이종: `{y_stats['이종합포']}`")
            
            zip_filename = f"{today_str}_{order_cnt}차.zip"
            st.download_button("💾 통합 다운로드 (.xls 묶음)", zip_buffer.getvalue(), zip_filename, "application/zip", width="stretch")
    except Exception as e:
        st.error(f"🚨 배정 중 중단됨: {e}")

if st.session_state['history']:
    st.markdown("---")
    st.subheader("📈 누적 배정 히스토리")
    st.dataframe(pd.DataFrame(st.session_state['history']), hide_index=True, width="stretch")
