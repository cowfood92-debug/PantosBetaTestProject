import streamlit as st
import pandas as pd
import pdfplumber
import io
import re

st.set_page_config(page_title="판토스 물류비용 정산 시스템", layout="wide")

st.title("📦 판토스 물류비용 정산 시스템")
st.write("마감내역서(PDF) 청구 건을 기준으로 반입계획서와 매칭하여 100% 동일한 엑셀을 만들고, **분기 계약 단가와 일치하는지 자동 검증(Audit)**합니다.")

st.sidebar.header("📂 파일 업로드")
uploaded_plan = st.sidebar.file_uploader("1. 반입계획서 (엑셀)", type=["xlsx"])
uploaded_pantos = st.sidebar.file_uploader("2. 판토스 마감내역서 (PDF)", type=["pdf"])

STANDARD_TRUCKING_RATE = 699000
BASE_CPT_RATE = 120.30

# 💡 분기 계약 단가표
CONTRACT_RATES = {
    "OCEAN FREIGHT": 200,
    "WHARFAGE": 9504,
    "CONTAINER CLEANING FEE": 50000,
    "CHASSIS CHARGE": 90,
    "EMERGENCY BUNKER SURCHARGE": 140,
    "PREPULL CHARGE": 150,
    "TERMINAL HANDLING CHARGE": 210000,
    "TRANSPORTATION CHARGE": 450,
    "DOCUMENT FEE": 50000,
    "DOCUMENT FEE AT ORIGIN PORT": 40
}

def to_num(s):
    if pd.isna(s) or s is None:
        return 0
    s = str(s).replace(",", "").strip()
    if s in ("", "-"):
        return 0
    try:
        return float(s)
    except ValueError:
        return 0

def normalize_lot(text):
    if pd.isna(text) or text is None:
        return ""
    t = str(text).upper()
    t = re.sub(r"\(.*?\)", "", t)
    t = re.sub(r"[^A-Z0-9]", "", t)
    return t

@st.cache_data(show_spinner=False)
def parse_pantos_pdf(file_bytes):
    rows_all = []
    with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
        for page in pdf.pages:
            for table in page.extract_tables():
                if any(r and "FREIGHT" in [str(c) for c in r] for r in table):
                    rows_all.extend(table)

    lots = {}
    order = []
    current_ref = None

    for r in rows_all:
        if r is None or len(r) < 18:
            continue
        if r[0] == "NO." or (r[0] and "T O T A L" in str(r[0])):
            continue

        if r[0] is not None and str(r[0]).strip().isdigit():
            current_ref = r[4]
            if current_ref not in lots:
                lots[current_ref] = {
                    "__EXRATE__": 0.0, "__QTY__": 0.0, "__TRUCK_RATE__": 0.0, "총액": 0.0,
                    "해상운임_USD": 0.0, "샤시운임_USD": 0.0, "선적지서류_USD": 0.0, "유류할증료_USD": 0.0, 
                    "프리풀_USD": 0.0, "선적지내륙운송_USD": 0.0, "스토리지_USD": 0.0, "기타_USD": 0.0,
                    "트러킹_KRW": 0.0, "부두사용료_KRW": 0.0, "세척비_KRW": 0.0, "터미널비_KRW": 0.0, "서류발급비_KRW": 0.0,
                    "__AUDIT_WARNINGS__": [] 
                }
                order.append(current_ref)

        freight_name = r[8]
        if current_ref is None or not freight_name or freight_name in ("FREIGHT",):
            continue
            
        qty = to_num(r[6])
        rate = to_num(r[9])
        curr = r[10]
        exrate = to_num(r[11])
        amt_t = to_num(r[12])
        amt_l = to_num(r[13])
        proxy_l = to_num(r[16])
        sum_krw = to_num(r[17])

        val_usd = amt_t if curr == "USD" else 0
        val_krw = amt_l if amt_l else proxy_l
        d = lots[current_ref]

        if "40HC" in str(r[5]):
            d["__QTY__"] = max(d["__QTY__"], qty)

        f_name_clean = str(freight_name).strip().upper()
        if f_name_clean in CONTRACT_RATES:
            expected_rate = CONTRACT_RATES[f_name_clean]
            if rate and rate != expected_rate:
                msg = f"🛑 [{freight_name}] 단가 불일치! (계약단가: {expected_rate:,.0f} / 청구단가: {rate:,.0f})"
                if msg not in d["__AUDIT_WARNINGS__"]:
                    d["__AUDIT_WARNINGS__"].append(msg)
                    
        elif "TRUCKING" in f_name_clean:
            if rate and rate not in (699000, 780400):
                msg = f"🛑 [TRUCKING CHARGE] 규정 외 단가 청구! (청구단가: {rate:,.0f}원)"
                if msg not in d["__AUDIT_WARNINGS__"]:
                    d["__AUDIT_WARNINGS__"].append(msg)

        if "OCEAN FREIGHT" in freight_name:
            d["해상운임_USD"] += val_usd
        elif "CHASSIS" in freight_name:
            d["샤시운임_USD"] += val_usd
        elif "DOCUMENT FEE AT ORIGIN" in freight_name:
            d["선적지서류_USD"] += val_usd
        elif "EMERGENCY BUNKER" in freight_name:
            d["유류할증료_USD"] += val_usd
        elif "PREPULL" in freight_name:
            d["프리풀_USD"] += val_usd
        elif "TRANSPORTATION" in freight_name:
            d["선적지내륙운송_USD"] += val_usd
        elif freight_name == "DOCUMENT FEE":
            d["서류발급비_KRW"] += val_krw
        elif "WHARFAGE" in freight_name:
            d["부두사용료_KRW"] += val_krw
        elif "CLEANING" in freight_name:
            d["세척비_KRW"] += val_krw
        elif "TERMINAL" in freight_name:
            d["터미널비_KRW"] += val_krw
        elif "TRUCKING" in freight_name:
            d["트러킹_KRW"] += val_krw
            d["__TRUCK_RATE__"] = rate

        if curr == "USD" and exrate and d["__EXRATE__"] == 0:
            d["__EXRATE__"] = exrate
        if sum_krw:
            d["총액"] = sum_krw

    return lots, order

if uploaded_plan and uploaded_pantos:
    with st.spinner("단가 검증 및 결과 엑셀을 조립하는 중입니다..."):
        pdf_lots, pdf_order = parse_pantos_pdf(uploaded_pantos.getvalue())
        
        try:
            df_temp = pd.read_excel(uploaded_plan, header=None)
            header_row_idx = 0
            for i, row in df_temp.iterrows():
                if any("Lot" in str(val) or "서류발송" in str(val) for val in row.values):
                    header_row_idx = i
                    break
            df_plan = pd.read_excel(uploaded_plan, header=header_row_idx)
        except Exception as e:
            st.error(f"반입계획서 엑셀 리딩 오류: {e}")
            st.stop()

        cols = df_plan.columns.astype(str)
        col_map = {
            "lot": next((c for c in cols if "Lot" in c or "오더" in c), None),
            "ship_date": next((c for c in cols if "선적일" in c), None),
            "arr_date": next((c for c in cols if "입항일" in c), None),
            "roll": next((c for c in cols if "ROLL" in c.upper()), None),
            "kg": next((c for c in cols if "KG" in c.upper() or "중 량" in c or "중량" in c), None),
            "sqm": next((c for c in cols if "SQM" in c.upper() or "S Q" in c or "수량" in c), None),
            "amt": next((c for c in cols if "금 액" in c or "외화" in c), None),
            "month": next((c for c in cols if "발주월" in c), None),
            "factory_date": next((c for c in cols if "배차" in c or "입고" in c), None),
            "clearance": next((c for c in cols if "통관" in c), None),
        }

        plan_map = {}
        for idx, row in df_plan.iterrows():
            raw = str(row.get(col_map["lot"], "")).strip()
            if raw in ("nan", "None", "") or "Lot" in raw:
                continue
            norm = normalize_lot(raw)
            plan_map[norm] = {"raw": raw, "row": row}

        results = []
        notes = [] 
        
        for pdf_ref in pdf_order:
            d = pdf_lots[pdf_ref]
            norm_pdf = normalize_lot(pdf_ref)
            
            audit_warnings = d.get("__AUDIT_WARNINGS__", [])
            for w in audit_warnings:
                notes.append(f"[{pdf_ref}] {w}")
            
            matched_plan = plan_map.get(norm_pdf)
            if not matched_plan:
                for p_norm, p_data in plan_map.items():
                    if norm_pdf in p_norm or p_norm in norm_pdf:
                        matched_plan = p_data
                        break
            
            if matched_plan:
                row = matched_plan["row"]
                raw_order = matched_plan["raw"] 
                p_ship = row.get(col_map["ship_date"], "")
                p_arr = row.get(col_map["arr_date"], "")
                p_roll = to_num(row.get(col_map["roll"], 0))
                p_kg = to_num(row.get(col_map["kg"], 0))
                p_sqm = to_num(row.get(col_map["sqm"], 0))
                p_amt = to_num(row.get(col_map["amt"], 0))
                p_month = row.get(col_map["month"], "")
                p_fac = row.get(col_map["factory_date"], "")
                p_clear = row.get(col_map["clearance"], "")
            else:
                raw_order = pdf_ref
                p_ship = p_arr = p_month = p_fac = p_clear = ""
                p_roll = p_kg = p_sqm = p_amt = 0
                notes.append(f"⚠️ [{raw_order}]: 반입계획서 엑셀에서 찾을 수 없어 날짜/수량이 빈칸 처리되었습니다.")

            out = [None] * 42
            out[0] = len(results) + 1
            out[1] = raw_order
            out[2] = p_ship if pd.notna(p_ship) else ""
            out[3] = p_arr if pd.notna(p_arr) else ""
            out[4] = p_roll
            out[5] = p_kg
            out[6] = p_sqm
            out[7] = p_amt
            out[8] = p_month
            out[9] = p_fac if pd.notna(p_fac) else ""
            out[10] = p_clear if pd.notna(p_clear) else ""

            qty = d.get("__QTY__", 0)

            out[12] = d.get("__EXRATE__", 0)
            out[13] = d.get("해상운임_USD", 0)
            out[14] = d.get("샤시운임_USD", 0)
            out[15] = d.get("선적지서류_USD", 0)
            out[16] = d.get("유류할증료_USD", 0)
            out[17] = d.get("프리풀_USD", 0)
            out[18] = d.get("선적지내륙운송_USD", 0)
            out[19] = d.get("스토리지_USD", 0)
            out[20] = d.get("기타_USD", 0)
            out[11] = sum(out[13:21])
            out[21] = out[11] * out[12]

            truck_rate = d.get("__TRUCK_RATE__", 0)
            if truck_rate == 780400: 
                safe_road = 200400
                safe_rail = 0
                notes.append(f"🚚 [{raw_order}]: 육송 배차건 (단가 780,400원 / 안전운임제 200,400원 적용)")
            elif truck_rate == 699000:
                safe_road = 0
                safe_rail = 119000
            else: 
                safe_road = 0
                safe_rail = 0

            out[22] = safe_road
            out[23] = safe_rail
            out[24] = 580000
            out[25] = d.get("트러킹_KRW", 0)
            out[26] = d.get("부두사용료_KRW", 0) + d.get("세척비_KRW", 0) + d.get("터미널비_KRW", 0) + d.get("서류발급비_KRW", 0)
            out[27] = out[25] + out[26]

            out[28] = out[7] / out[5] if out[5] else 0
            out[29] = out[7] / out[6] if out[6] else 0
            out[30] = out[21] + out[27]
            out[31] = out[30] / out[6] if out[6] else 0
            out[32] = qty
            out[33] = out[30] / qty if qty else 0
            
            out[40] = qty * (safe_road + safe_rail)
            out[41] = out[16] * out[12]

            results.append(out)

            bunker_usd = d.get("유류할증료_USD", 0)
            if bunker_usd > 0:
                notes.append(f"⛽ [{raw_order}]: 유류할증료(BUNKER) ${bunker_usd:,.0f} 달러 발생")
                
            calculated_total = out[30]
            pdf_total = d.get("총액", 0)
            if abs(calculated_total - pdf_total) > 10:
                notes.append(f"❌ [{raw_order}]: 총액 불일치 (시스템산출: {calculated_total:,.0f}원 vs PDF청구: {pdf_total:,.0f}원)")

        total_ocean_usd = sum([r[13] for r in results])
        total_ocean_krw = sum([r[13] * r[12] for r in results])
        total_kg = sum([r[5] for r in results])
        total_amount_krw = sum([r[30] for r in results])

        avg_exrate = total_ocean_krw / total_ocean_usd if total_ocean_usd else 0
        avg_exrate = round(avg_exrate, 1) 
        total_usd_equiv = total_amount_krw / avg_exrate if avg_exrate else 0
        total_tons = total_kg / 1000 if total_kg else 0
        base_rate = round(total_usd_equiv / total_tons, 2) if total_tons else 0
        calculated_fca_rate = base_rate + 28

        for out in results:
            kg_tons = out[5] / 1000 if out[5] else 0
            out[34] = BASE_CPT_RATE * kg_tons
            out[35] = calculated_fca_rate * kg_tons 
            out[36] = out[34] * out[12] 
            out[37] = out[35] * out[12] 
            out[38] = out[37] - out[36] 

    # ---------------------------------------------------------
    # 📊 화면 출력: 4개 항목 깔끔한 대시보드
    # ---------------------------------------------------------
    total_orders = len(results)
    total_qty = sum([r[32] for r in results])

    st.subheader("📊 이번 달 물류 운송 종합 브리핑")
    
    # 4칸으로 나누어 핵심 정보만 깔끔하게 표시
    col1, col2, col3, col4 = st.columns(4)
    col1.metric("총 진행 오더", f"{total_orders} 건")
    col2.metric("총 운송 컨테이너", f"{total_qty:,.0f} 대")
    col3.metric("총 톤수", f"{total_kg / 1000:,.1f} 톤")
    col4.metric("총 비용", f"{total_amount_krw:,.0f} 원")
    st.write("---")

    st.subheader("💡 정산 특이사항 요약")
    if notes:
        unique_notes = list(set(notes))
        for note in unique_notes:
            if "🛑" in note or "❌" in note:
                st.error(note)
            elif "🚚" in note:
                st.warning(note)
            else:
                st.info(note)
    else:
        st.success("🎉 모든 건이 분기 계약 단가와 100% 일치하며 완벽하게 정산되었습니다!")

    # ---------------------------------------------------------
    # 📥 엑셀 작성 및 다운로드
    # ---------------------------------------------------------
    headers_row1 = [
        "NO.", "Lot (서류발송)", "선적일", "입항일", "ROLL", "kg", "SQM", "외화물품대", "발주월", "공장입고", "통관",
        "해상운임", "", "", "", "", "", "", "", "", "", "",
        "국내운송", "", "", "", "", "",
        "외화", "", "총계", "", "컨테이너", ""
    ]
    headers_row2 = [
        "", "", "", "", "", "", "", "", "", "", "",
        "외화($)", "환율", "해상운임", "샤시운임", "선적지 서류", "유류할증료", "프리풀", "선적지 내륙운송", "스토리지", "기타", "원화(\\)",
        "컨테이너 1대 당\n안전운임제(철송X)", "컨테이너 1대 당\n안전운임제(철송)", "컨테이너 1대당\n트러킹(기본)", "총 컨테이너 포함\n트러킹 총비용", "국내운송비(트러킹제외)", "계",
        "kg", "SQM", "총 금액", "SQM", "수량", "대당 평균 비용"
    ]

    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
        workbook = writer.book
        worksheet = workbook.add_worksheet("6월_완성본")
        
        fmt_header = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'border': 1, 'bg_color': '#D9E1F2', 'text_wrap': True})
        fmt_yellow_header = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'border': 1, 'bg_color': '#FFFF00'})
        fmt_blue_header = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'border': 1, 'bg_color': '#DDEBF7'})
        fmt_red_currency = workbook.add_format({'num_format': '"$"#,##0.00', 'align': 'center', 'valign': 'vcenter', 'border': 1, 'font_color': 'red'})
        
        fmt_num = workbook.add_format({'num_format': '#,##0', 'border': 1})
        fmt_float = workbook.add_format({'num_format': '#,##0.00', 'border': 1})
        fmt_date = workbook.add_format({'num_format': 'yyyy-mm-dd', 'border': 1, 'align': 'center'})
        fmt_center = workbook.add_format({'align': 'center', 'border': 1})
        
        fmt_total = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'border': 1, 'bg_color': '#FFFF00'})
        fmt_total_num = workbook.add_format({'bold': True, 'num_format': '#,##0', 'border': 1, 'bg_color': '#FFFF00'})
        fmt_total_float = workbook.add_format({'bold': True, 'num_format': '#,##0.00', 'border': 1, 'bg_color': '#FFFF00'})

        worksheet.set_row(0, 30)
        worksheet.set_row(1, 30)
        
        for i in range(11):
            worksheet.merge_range(0, i, 1, i, headers_row1[i], fmt_header)
        
        worksheet.merge_range(0, 11, 0, 21, "해상운임", fmt_header)
        for i in range(11, 22):
            worksheet.write(1, i, headers_row2[i], fmt_header)

        worksheet.merge_range(0, 22, 0, 27, "국내운송", fmt_header)
        for i in range(22, 28):
            worksheet.write(1, i, headers_row2[i], fmt_header)

        worksheet.merge_range(0, 28, 0, 29, "외화", fmt_header)
        for i in range(28, 30):
            worksheet.write(1, i, headers_row2[i], fmt_header)

        worksheet.merge_range(0, 30, 0, 31, "총계", fmt_header)
        for i in range(30, 32):
            worksheet.write(1, i, headers_row2[i], fmt_header)

        worksheet.merge_range(0, 32, 0, 33, "컨테이너", fmt_header)
        for i in range(32, 34):
            worksheet.write(1, i, headers_row2[i], fmt_header)

        worksheet.write(0, 34, "톤당 CPT 운임", fmt_yellow_header)
        worksheet.write(0, 35, "톤당 FCA 운임", fmt_yellow_header)
        
        worksheet.write_number(1, 34, BASE_CPT_RATE, fmt_red_currency)
        
        last_data_row = len(results) + 2 
        fca_formula = f"=IFERROR(ROUND((SUM(AE3:AE{last_data_row})/ROUND(SUM(V3:V{last_data_row})/SUM(L3:L{last_data_row}),1))/(SUM(F3:F{last_data_row})/1000),2)+28, 0)"
        worksheet.write_formula(1, 35, fca_formula, fmt_red_currency, calculated_fca_rate)

        worksheet.merge_range(0, 36, 1, 36, "CPT 원화 환산", fmt_blue_header)
        worksheet.merge_range(0, 37, 1, 37, "FCA 원화 환산", fmt_blue_header)
        worksheet.merge_range(0, 38, 1, 38, "운임 차액", fmt_blue_header)

        worksheet.write(0, 39, "")
        worksheet.write(1, 39, "")

        worksheet.merge_range(0, 40, 1, 40, "안전운임제 합계", fmt_yellow_header)
        worksheet.merge_range(0, 41, 1, 41, "유류할증료 합계", fmt_yellow_header)

        for r_idx, row_data in enumerate(results):
            xls_r = r_idx + 2
            excel_row = xls_r + 1 
            
            for c_idx, val in enumerate(row_data):
                if isinstance(val, float):
                    if pd.isna(val) or val == float('inf') or val == float('-inf'):
                        val = 0
                elif pd.isna(val):
                    val = ""

                if c_idx == 34:
                    worksheet.write_formula(xls_r, c_idx, f"=F{excel_row}/1000*AI$2", fmt_float, float(val) if val else 0)
                elif c_idx == 35:
                    worksheet.write_formula(xls_r, c_idx, f"=F{excel_row}/1000*AJ$2", fmt_float, float(val) if val else 0)
                elif c_idx == 36:
                    worksheet.write_formula(xls_r, c_idx, f"=AI{excel_row}*M{excel_row}", fmt_num, float(val) if val else 0)
                elif c_idx == 37:
                    worksheet.write_formula(xls_r, c_idx, f"=AJ{excel_row}*M{excel_row}", fmt_num, float(val) if val else 0)
                elif c_idx == 38:
                    worksheet.write_formula(xls_r, c_idx, f"=AL{excel_row}-AK{excel_row}", fmt_num, float(val) if val else 0)
                
                elif val == "":
                    worksheet.write(xls_r, c_idx, "", fmt_center)
                elif c_idx in (2, 3, 9, 10): 
                    try:
                        if hasattr(val, 'strftime'):
                            worksheet.write_datetime(xls_r, c_idx, val, fmt_date)
                        elif val and isinstance(val, str):
                            dt = pd.to_datetime(val)
                            worksheet.write_datetime(xls_r, c_idx, dt, fmt_date)
                        else:
                            worksheet.write(xls_r, c_idx, val, fmt_center)
                    except:
                        worksheet.write(xls_r, c_idx, val, fmt_center)
                elif c_idx in (12, 28, 29, 31):
                    worksheet.write_number(xls_r, c_idx, float(val), fmt_float)
                elif isinstance(val, (int, float)):
                    worksheet.write_number(xls_r, c_idx, float(val), fmt_num)
                else:
                    worksheet.write(xls_r, c_idx, str(val), fmt_center)

        total_r = len(results) + 2 
        for i in range(42):
            worksheet.write(total_r, i, "", fmt_total) 
            
        worksheet.write(total_r, 1, "총계", fmt_total)
        
        sums_mapping = {
            4: "E",   # ROLL
            5: "F",   # KG
            6: "G",   # SQM
            7: "H",   # 외화물품대
            11: "L",  # 해상운임 외화($)
            21: "V",  # 해상운임 원화(\)
            27: "AB", # 국내운송 계
            30: "AE", # 총계 총금액
            34: "AI", # 톤당 CPT 운임 합
            35: "AJ", # 톤당 FCA 운임 합
            36: "AK", # CPT 원화 환산 합
            37: "AL", # FCA 원화 환산 합
            38: "AM"  # 운임 차액 합 (총 세이브 금액)
        }
        
        for c_idx, col_letter in sums_mapping.items():
            formula = f"=SUM({col_letter}3:{col_letter}{last_data_row})"
            fmt = fmt_total_float if c_idx in [11, 34, 35] else fmt_total_num
            worksheet.write_formula(total_r, c_idx, formula, fmt)

        worksheet.set_column(0, 0, 5)
        worksheet.set_column(1, 1, 15)
        worksheet.set_column(2, 3, 11)
        worksheet.set_column(4, 8, 10)
        worksheet.set_column(9, 10, 11)
        worksheet.set_column(11, 21, 11)
        worksheet.set_column(22, 27, 13)
        worksheet.set_column(28, 33, 11)
        worksheet.set_column(34, 38, 14)
        worksheet.set_column(39, 39, 3)
        worksheet.set_column(40, 41, 15)

        worksheet_audit = workbook.add_worksheet("검증_리포트")
        fmt_audit_title = workbook.add_format({'bold': True, 'valign': 'vcenter', 'bg_color': '#FFC000', 'border': 1, 'font_size': 12})
        fmt_audit_text = workbook.add_format({'valign': 'vcenter', 'border': 1})
        
        worksheet_audit.set_column(0, 0, 120)
        worksheet_audit.set_row(0, 30)
        worksheet_audit.write(0, 0, "💡 정산 특이사항 요약", fmt_audit_title)
        
        if notes:
            unique_notes = list(set(notes))
            for idx, note in enumerate(unique_notes):
                clean_note = note.replace("**", "") 
                worksheet_audit.write(idx + 1, 0, clean_note, fmt_audit_text)
                worksheet_audit.set_row(idx + 1, 25)
        else:
            worksheet_audit.write(1, 0, "🎉 특이사항 없음: 모든 건이 분기 계약 단가와 완벽하게 일치합니다.", fmt_audit_text)
            worksheet_audit.set_row(1, 25)

    st.write("---")
    st.subheader("📥 최종 자료 다운로드")
    st.download_button(
        label="📊 클릭하여 최종 엑셀 파일 다운로드",
        data=buffer.getvalue(),
        file_name="판토스_물류비용정산_최종본.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
else:
    st.info("👈 1. 반입계획서(엑셀)와 2. 마감내역서(PDF)를 순서대로 올려주시면 완성됩니다.")
