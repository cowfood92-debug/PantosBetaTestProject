import streamlit as st
import pandas as pd
import pdfplumber
import io
import re

# 브라우저 탭 이름과 아이콘 설정
st.set_page_config(page_title="판토스 물류비용 자동 정산 시스템", page_icon="📦", layout="wide")

# ---------------------------------------------------------
# 🎨 커스텀 CSS 적용
# ---------------------------------------------------------
st.markdown("""
<style>
#MainMenu {visibility: hidden;}
.stDeployButton {display:none;}
footer {visibility: hidden;}
header {visibility: hidden;}
.main-title { font-size: 2.8rem; font-weight: 900; color: #1E3A8A; margin-bottom: 0.1rem; }
.sub-title { font-size: 1.1rem; color: #6B7280; margin-bottom: 2rem; }
.guide-box { background-color: #F9FAFB; padding: 25px 20px; border-radius: 12px; text-align: center; border: 1px solid #E5E7EB; box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05); height: 100%; }
.guide-step { font-size: 1.2rem; font-weight: 700; color: #374151; margin-bottom: 10px; }
.guide-desc { font-size: 1rem; color: #4B5563; }
</style>
""", unsafe_allow_html=True)

# 💡 V46 업데이트 마크
st.markdown(
    '<div class="main-title">📦 판토스 물류비용 자동 정산 시스템 <span style="font-size: 1.2rem; color: #10B981; vertical-align: middle;">[V46 11월 평택 인하 에디션]</span></div>',
    unsafe_allow_html=True)
st.markdown(
    '<div class="sub-title">마감내역서(PDF)와 반입계획서(Excel) 매칭부터 <strong>분기 계약 단가 자동 검증(Audit)</strong>까지 10초 만에 완벽하게 끝내세요.</div>',
    unsafe_allow_html=True)

st.sidebar.markdown("### 📂 이번 달 정산 데이터 업로드")
uploaded_plan = st.sidebar.file_uploader("1. 반입계획서 (엑셀)", type=["xlsx"])
uploaded_pantos_list = st.sidebar.file_uploader("2. 판토스 개별 인보이스 (PDF 여러 개 드래그 가능)", type=["pdf"],
                                                accept_multiple_files=True)


def to_num(s):
    if pd.isna(s) or s is None: return 0
    s = str(s).replace(",", "").strip()
    if s in ("", "-"): return 0
    try:
        return float(s)
    except ValueError:
        return 0


def normalize_lot(text):
    if pd.isna(text) or text is None: return ""
    t = str(text).upper()
    t = re.sub(r"\(.*?\)", "", t)
    t = re.sub(r"[^A-Z0-9]", "", t)
    return t


def smart_extract_row_indiv(r):
    row_str = " ".join([str(x) for x in r if x is not None])
    is_usd = "USD" in row_str.upper()

    tokens = re.findall(r"(?<!\S)\d{1,3}(?:,\d{3})*(?:\.\d+)?(?!\S)|(?<!\S)\d+(?:\.\d+)?(?!\S)", row_str)
    nums = [float(t.replace(",", "")) for t in tokens if float(t.replace(",", "")) > 0]

    if len(nums) > 1 and nums[-1] == nums[-2]:
        nums.pop()

    val_usd, val_krw, exrate = 0, 0, 0
    if not nums: return 0, 0, 0

    if is_usd:
        if len(nums) >= 2:
            val_krw = nums[-1]
            val_usd = nums[-2]
            for n in nums:
                if abs(val_usd * n - val_krw) <= 10:
                    exrate = n
                    break
            if exrate == 0:
                for n in nums:
                    if 1000 < n < 2000:
                        exrate = n
                        break
                if exrate > 0:
                    val_usd = round(val_krw / exrate, 2)
    else:
        val_krw = nums[-1]

    return val_usd, val_krw, exrate


@st.cache_data(show_spinner=False)
def parse_pantos_pdf_individual_batch(file_bytes_list):
    lots = {}
    order = []

    for file_bytes in file_bytes_list:
        with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
            for page in pdf.pages:
                text = page.extract_text()
                if not text: continue

                ref_match = re.search(r"REF\s*NO\.?\s*([A-Za-z0-9\-]+)", text)
                if not ref_match: continue
                current_ref = normalize_lot(ref_match.group(1).strip())

                if current_ref not in lots:
                    lots[current_ref] = {
                        "__EXRATE__": 0.0, "__QTY__": 0.0, "__TRUCK_RATE__": 0.0, "총액": 0.0,
                        "__POL__": "", "__POD__": "", "__ON_BOARD__": "",
                        "해상운임_USD": 0.0, "샤시운임_USD": 0.0, "선적지서류_USD": 0.0, "유류할증료_USD": 0.0,
                        "프리풀_USD": 0.0, "선적지내륙운송_USD": 0.0, "스토리지_USD": 0.0, "기타_USD": 0.0,
                        "트러킹_KRW": 0.0, "부두사용료_KRW": 0.0, "세척비_KRW": 0.0, "터미널비_KRW": 0.0, "서류발급비_KRW": 0.0,
                        "기타_KRW": 0.0, "__RAW_RATES__": {}, "__RAW_NAMES__": {}, "__AUDIT_WARNINGS__": []
                    }
                    if current_ref not in order:
                        order.append(current_ref)

                d = lots[current_ref]

                pol_m = re.search(r"POL\s+(.*?)\s+(?:INCOTERMS|POD|F\.DEST)", text)
                pod_m = re.search(r"POD\s+(.*?)\s+(?:VESSEL/FLIGHT|VESSEL|F\.DEST|GWT)", text)
                onboard_m = re.search(r"ON-BOARD\s+(\d{4}-\d{2}-\d{2})", text)
                
                if pol_m and not d["__POL__"]: d["__POL__"] = pol_m.group(1).upper()
                if pod_m and not d["__POD__"]: d["__POD__"] = pod_m.group(1).upper()
                if onboard_m and not d["__ON_BOARD__"]: d["__ON_BOARD__"] = onboard_m.group(1)

                qty_match = re.search(r"40HC\s*[xX*]\s*(\d+)", text, re.IGNORECASE)
                if qty_match:
                    d["__QTY__"] = max(d["__QTY__"], float(qty_match.group(1)))

                for line in text.split('\n'):
                    if "GRAND TOTAL" in line.upper():
                        tokens = re.findall(r"(?<!\S)\d{1,3}(?:,\d{3})*(?:\.\d+)?(?!\S)|(?<!\S)\d+(?:\.\d+)?(?!\S)",
                                            line)
                        nums = [float(t.replace(",", "")) for t in tokens if float(t.replace(",", "")) > 0]
                        if nums:
                            d["총액"] = max(d.get("총액", 0), nums[-1])
                            break

                tables = page.extract_tables()
                for table in tables:
                    is_freight = False
                    for r in table[:3]:
                        if any(r) and any(str(c).upper().find("FREIGHT") != -1 for c in r if c):
                            is_freight = True
                            break
                    if not is_freight:
                        continue

                    for r in table:
                        if not r or not any(r): continue

                        row_str_upper = " ".join([str(x) for x in r if x is not None]).upper()

                        if "GRAND TOTAL" in row_str_upper:
                            break

                        freight_name = str(r[0]).strip()
                        f_name_clean = freight_name.upper()

                        if not f_name_clean or any(skip in f_name_clean for skip in
                                                   ["FREIGHT & CHARGE", "SUB TOTAL", "영세", "세금무관", "과세", "TAX",
                                                    "REMARK"]):
                            continue

                        rate = 0
                        rate_match = re.search(r"(?:USD|KRW)\s+([\d,\.]+)", row_str_upper)
                        if rate_match:
                            rate = float(rate_match.group(1).replace(",", ""))

                        val_usd, val_krw, exrate = smart_extract_row_indiv(r)

                        lookup_name = f_name_clean
                        if "LSS" in f_name_clean or "LOW SULPHUR" in f_name_clean or "LOW SULFUR" in f_name_clean:
                            lookup_name = "LSS"

                        d["__RAW_RATES__"][lookup_name] = rate
                        d["__RAW_NAMES__"][lookup_name] = freight_name

                        is_mapped = False
                        if "OCEAN FREIGHT" in f_name_clean:
                            d["해상운임_USD"] += val_usd;
                            is_mapped = True
                        elif "CHASSIS" in f_name_clean:
                            d["샤시운임_USD"] += val_usd;
                            is_mapped = True
                        elif "DOCUMENT FEE AT ORIGIN" in f_name_clean:
                            d["선적지서류_USD"] += val_usd;
                            is_mapped = True
                        elif "EMERGENCY BUNKER" in f_name_clean or "LSS" in f_name_clean or "LOW SULPHUR" in f_name_clean or "LOW SULFUR" in f_name_clean:
                            d["유류할증료_USD"] += val_usd;
                            is_mapped = True
                        elif "PREPULL" in f_name_clean:
                            d["프리풀_USD"] += val_usd;
                            is_mapped = True
                        elif "TRANSPORTATION" in f_name_clean:
                            d["선적지내륙운송_USD"] += val_usd;
                            is_mapped = True
                        elif f_name_clean == "DOCUMENT FEE":
                            d["서류발급비_KRW"] += val_krw;
                            is_mapped = True
                        elif "WHARFAGE" in f_name_clean:
                            d["부두사용료_KRW"] += val_krw;
                            is_mapped = True
                        elif "CLEANING" in f_name_clean:
                            d["세척비_KRW"] += val_krw;
                            is_mapped = True
                        elif "TERMINAL" in f_name_clean:
                            d["터미널비_KRW"] += val_krw;
                            is_mapped = True
                        elif "TRUCKING" in f_name_clean:
                            d["트러킹_KRW"] += val_krw
                            d["__TRUCK_RATE__"] = rate
                            is_mapped = True

                        if not is_mapped:
                            if val_usd > 0:
                                d["기타_USD"] += val_usd
                                msg = f"⚠️ [신규항목] '{freight_name}' (${val_usd:,.0f}) 발견 ➡️ '기타 외화' 합산"
                                if msg not in d["__AUDIT_WARNINGS__"]: d["__AUDIT_WARNINGS__"].append(msg)
                            elif val_krw > 0:
                                d["기타_KRW"] += val_krw
                                msg = f"⚠️ [신규항목] '{freight_name}' ({val_krw:,.0f}원) 발견 ➡️ '내륙 기타' 합산"
                                if msg not in d["__AUDIT_WARNINGS__"]: d["__AUDIT_WARNINGS__"].append(msg)

                        if exrate > 0 and d["__EXRATE__"] == 0: d["__EXRATE__"] = exrate

    return lots, order


buffer = io.BytesIO()
is_processed = False
total_orders = 0
total_qty = 0

if uploaded_plan and uploaded_pantos_list:
    with st.spinner("엑셀 및 여러 개의 개별 인보이스를 동시 분석하고 있습니다..."):
        file_bytes_list = [f.getvalue() for f in uploaded_pantos_list]
        pdf_lots, pdf_order = parse_pantos_pdf_individual_batch(file_bytes_list)

        try:
            xls = pd.ExcelFile(uploaded_plan)
            plan_map = {}
            for sheet_name in xls.sheet_names:
                df_temp = pd.read_excel(xls, sheet_name=sheet_name, header=None)
                header_row_idx = -1
                for i, row in df_temp.iterrows():
                    if any(isinstance(val, str) and ("Lot" in val or "서류발송" in val) for val in row.values):
                        header_row_idx = i; break
                
                if header_row_idx == -1: continue

                country = "CN" if "홍타" in sheet_name else "US/EU"
                df_plan = pd.read_excel(xls, sheet_name=sheet_name, header=header_row_idx)
                cols = df_plan.columns.astype(str)
                
                col_map = {
                    "lot": next((c for c in cols if "Lot" in c or "오더" in c), None),
                    "ship_date": next((c for c in cols if "선적일" in c or "ETD" in c.upper()), None),
                    "arr_date": next((c for c in cols if "입항일" in c or "ETA" in c.upper()), None),
                    "roll": next((c for c in cols if "ROLL" in c.upper()), None),
                    "kg": next((c for c in cols if "KG" in c.upper() or "중 량" in c or "중량" in c), None),
                    "sqm": next((c for c in cols if "SQM" in c.upper() or "S Q" in c or "수량" in c), None),
                    "amt": next((c for c in cols if "금 액" in c or "외화" in c or "금액" in c or "물품대" in c), None),
                    "month": next((c for c in cols if "발주월" in c), None),
                    "factory_date": next((c for c in cols if "배차" in c or "입고" in c), None),
                    "clearance": next((c for c in cols if "통관" in c), None),
                }
                for idx, row in df_plan.iterrows():
                    if not col_map["lot"]: continue
                    raw = str(row.get(col_map["lot"], "")).strip()
                    if raw in ("nan", "None", "") or "Lot" in raw or "총계" in raw or "합계" in raw: continue
                    norm = normalize_lot(raw)
                    plan_map[norm] = {"raw": raw, "row": row, "country": country, "col_map": col_map}
        except Exception as e:
            st.error(f"반입계획서 엑셀 리딩 오류: {e}")
            st.stop()

        # 💡 [V46 업데이트] 평택 11월 인하 단가(388,100원) 프리패스 추가!
        MASTER_TRUCK_RATES = [
            699000, 780400, 785400,
            321000, 829900, 1060300, 587200,
            283700, 739700, 943900, 518700,
            208500, 394200, 388100
        ]
        
        # 💡 [V46 업데이트] 평택항 육송 단가 이름 매핑 추가
        TRUCK_ROUTE_MAP = {
            283700: "삼륭 안산공장-인천항(20')",
            321000: "삼륭 안산공장-인천항(40')",
            739700: "삼륭 안산공장-부산항(20')",
            943900: "SR 천안공장-부산항(20')",
            1060300: "SR 천안공장-부산항(40')",
            518700: "SR 천안공장-인천항(20')",
            587200: "SR 천안공장-인천항(40')",
            388100: "평택항 육송 (11월 인하 단가)"
        }

        results, notes = [], []
        for pdf_ref in pdf_order:
            d = pdf_lots[pdf_ref]
            norm_pdf = normalize_lot(pdf_ref)

            matched_plan = plan_map.get(norm_pdf)
            if not matched_plan:
                for p_norm, p_data in plan_map.items():
                    if norm_pdf in p_norm or p_norm in norm_pdf:
                        matched_plan = p_data;
                        break

            if matched_plan:
                row, col_map, raw_order = matched_plan["row"], matched_plan["col_map"], matched_plan["raw"]
                p_ship = row.get(col_map["ship_date"], "")
                p_arr = row.get(col_map["arr_date"], "")
                p_roll, p_kg, p_sqm, p_amt = to_num(row.get(col_map["roll"], 0)), to_num(row.get(col_map["kg"], 0)), to_num(row.get(col_map["sqm"], 0)), to_num(row.get(col_map["amt"], 0))
                p_month = row.get(col_map["month"], "")
                p_fac = row.get(col_map["factory_date"], "")
                p_clear = row.get(col_map["clearance"], "")
            else:
                raw_order, p_ship, p_arr, p_month, p_fac, p_clear = pdf_ref, "", "", "", "", ""
                p_roll, p_kg, p_sqm, p_amt = 0, 0, 0, 0
                notes.append(f"⚠️ [{raw_order}]: 반입계획서 엑셀에서 찾을 수 없어 날짜/수량이 빈칸 처리되었습니다.")

            order_country = "US/EU"
            if "HT" in norm_pdf:
                order_country = "CN"
            elif matched_plan:
                order_country = matched_plan["country"]

            if order_country == "US/EU":
                ORDER_CPT_RATE = 120.30
                CONTRACT_RATES = {
                    "OCEAN FREIGHT": 200, "WHARFAGE": 9504, "CONTAINER CLEANING FEE": 50000,
                    "CHASSIS CHARGE": 90, "EMERGENCY BUNKER SURCHARGE": 140, "PREPULL CHARGE": 150,
                    "TERMINAL HANDLING CHARGE": 210000, "TRANSPORTATION CHARGE": 450,
                    "DOCUMENT FEE": 50000, "DOCUMENT FEE AT ORIGIN PORT": 40
                }
            else:
                ORDER_CPT_RATE = 35.00
                CONTRACT_RATES = {
                    "OCEAN FREIGHT": [150, 200, 600, 950], 
                    "LSS": 400, "WHARFAGE": 9504, "CONTAINER CLEANING FEE": 50000,
                    "TERMINAL HANDLING CHARGE": 210000, "DOCUMENT FEE": 50000
                }

            for lookup_name, rate in d.get("__RAW_RATES__", {}).items():
                original_name = d["__RAW_NAMES__"].get(lookup_name, lookup_name)
                if lookup_name in CONTRACT_RATES:
                    expected_rate = CONTRACT_RATES[lookup_name]
                    if rate:
                        if isinstance(expected_rate, list):
                            if not any(abs(rate - valid_r) < 10 for valid_r in expected_rate):
                                d["__AUDIT_WARNINGS__"].append(f"🛑 [{original_name}] 단가 불일치! (허용단가: {expected_rate} / 청구단가: {rate:,.0f})")
                        else:
                            if abs(rate - expected_rate) >= 10:
                                d["__AUDIT_WARNINGS__"].append(f"🛑 [{original_name}] 단가 불일치! (계약단가: {expected_rate:,.0f} / 청구단가: {rate:,.0f})")
                elif "TRUCKING" in lookup_name:
                    if rate:
                        is_valid_truck = any(abs(rate - safe_rate) < 10 for safe_rate in MASTER_TRUCK_RATES)
                        if not is_valid_truck:
                            d["__AUDIT_WARNINGS__"].append(f"🛑 [TRUCKING CHARGE] 규정 외 단가 청구! (청구단가: {rate:,.0f}원)")

            for w in d.get("__AUDIT_WARNINGS__", []): notes.append(f"[{pdf_ref}] {w}")

            out = [None] * 43
            out[0], out[1] = len(results) + 1, raw_order
            out[2], out[3] = p_ship if pd.notna(p_ship) else "", p_arr if pd.notna(p_arr) else ""
            out[4], out[5], out[6], out[7] = p_roll, p_kg, p_sqm, p_amt
            out[8], out[9], out[10] = p_month, p_fac if pd.notna(p_fac) else "", p_clear if pd.notna(p_clear) else ""

            qty = d.get("__QTY__", 0)
            out[12] = d.get("__EXRATE__", 0)
            out[13], out[14], out[15] = d.get("해상운임_USD", 0), d.get("샤시운임_USD", 0), d.get("선적지서류_USD", 0)
            out[16], out[17], out[18] = d.get("유류할증료_USD", 0), d.get("프리풀_USD", 0), d.get("선적지내륙운송_USD", 0)
            out[19], out[20] = d.get("스토리지_USD", 0), d.get("기타_USD", 0)

            out[11] = sum(out[13:21])
            out[21] = out[11] * out[12]

            truck_rate = d.get("__TRUCK_RATE__", 0)
            base_trucking = 580000

            is_master = any(abs(truck_rate - tr) < 10 for tr in MASTER_TRUCK_RATES)

            if any(abs(truck_rate - tr) < 10 for tr in [780400, 785400, 829900]):
                safe_road, safe_rail = truck_rate - 580000, 0
                route_name = "삼륭 안산공장-부산항(40')" if abs(truck_rate - 829900) < 10 else "육송 배차건"
                notes.append(f"🚚 [{raw_order}]: {route_name} (단가 {truck_rate:,.0f}원 / 안전운임제 {safe_road:,.0f}원 적용)")
            elif abs(truck_rate - 699000) < 10:
                safe_road, safe_rail = 0, 119000
            else:
                matched_route = None
                for price, route in TRUCK_ROUTE_MAP.items():
                    if abs(truck_rate - price) < 10:
                        matched_route = route
                        break
                
                if matched_route:
                    safe_road, safe_rail = 0, 0
                    base_trucking = truck_rate
                    notes.append(f"🚚 [{raw_order}]: {matched_route} (단가 {truck_rate:,.0f}원)")
                elif is_master:
                    safe_road, safe_rail = 0, 0
                    base_trucking = truck_rate
                    notes.append(f"🚚 [{raw_order}]: 기타 예외 노선 육송 (단가 {truck_rate:,.0f}원)")
                else:
                    safe_road, safe_rail = 0, 0

            out[22], out[23], out[24] = safe_road, safe_rail, base_trucking
            out[25] = d.get("트러킹_KRW", 0)
            out[26] = d.get("부두사용료_KRW", 0) + d.get("세척비_KRW", 0) + d.get("터미널비_KRW", 0) + d.get("서류발급비_KRW",
                                                                                                 0) + d.get("기타_KRW", 0)
            out[27] = out[25] + out[26]
            out[28], out[29] = out[7] / out[5] if out[5] else 0, out[7] / out[6] if out[6] else 0

            out[30] = out[21] + out[27]
            out[31], out[32], out[33] = out[30] / out[6] if out[6] else 0, qty, out[30] / qty if qty else 0
            out[40], out[41] = qty * (safe_road + safe_rail), out[16] * out[12]

            out[42] = ORDER_CPT_RATE

            results.append(out)

            if abs(out[30] - d.get("총액", 0)) > 100:
                notes.append(f"❌ [{raw_order}]: 총액 불일치 (시스템: {out[30]:,.0f}원 vs PDF: {d.get('총액', 0):,.0f}원)")
            else:
                notes.append(f"✅ [{raw_order}]: 시스템 검증 완료 (100% 일치)")

        total_orders = len(results)
        total_qty = sum([r[32] for r in results])
        
        total_usd_all = sum([r[11] for r in results])
        total_krw_of_usd = sum([r[21] for r in results])
        total_kg = sum([r[5] for r in results])
        total_amount_krw = sum([r[30] for r in results])

        avg_exrate_raw = total_krw_of_usd / total_usd_all if total_usd_all else 0
        avg_exrate = int(avg_exrate_raw * 10 + 0.5) / 10.0

        total_tons = total_kg / 1000 if total_kg else 0
        base_rate_raw = (total_amount_krw / avg_exrate) / total_tons if avg_exrate and total_tons else 0
        base_rate = int(base_rate_raw * 100 + 0.5) / 100.0

        calculated_fca_rate = base_rate + 28

        total_cpt_krw, total_fca_krw = 0, 0
        for out in results:
            kg_tons = out[5] / 1000 if out[5] else 0
            out[34], out[35] = out[42] * kg_tons, calculated_fca_rate * kg_tons
            out[36], out[37] = out[34] * out[12], out[35] * out[12]
            out[38] = out[37] - out[36]
            total_cpt_krw += out[36]
            total_fca_krw += out[37]

        is_processed = True

        with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
            workbook = writer.book
            worksheet = workbook.add_worksheet("월별_완성본")
            fmt_header = workbook.add_format(
                {'bold': True, 'align': 'center', 'valign': 'vcenter', 'border': 1, 'bg_color': '#D9E1F2',
                 'text_wrap': True})
            fmt_yellow_header = workbook.add_format(
                {'bold': True, 'align': 'center', 'valign': 'vcenter', 'border': 1, 'bg_color': '#FFFF00'})
            fmt_blue_header = workbook.add_format(
                {'bold': True, 'align': 'center', 'valign': 'vcenter', 'border': 1, 'bg_color': '#DDEBF7'})
            fmt_red_currency = workbook.add_format(
                {'num_format': '"$"#,##0.00', 'align': 'center', 'valign': 'vcenter', 'border': 1, 'font_color': 'red'})
            fmt_num = workbook.add_format({'num_format': '#,##0', 'border': 1})
            fmt_float = workbook.add_format({'num_format': '#,##0.00', 'border': 1})
            fmt_date = workbook.add_format({'num_format': 'yyyy-mm-dd', 'border': 1, 'align': 'center'})
            fmt_center = workbook.add_format({'align': 'center', 'border': 1})
            fmt_total = workbook.add_format(
                {'bold': True, 'align': 'center', 'valign': 'vcenter', 'border': 1, 'bg_color': '#FFFF00'})
            fmt_total_num = workbook.add_format(
                {'bold': True, 'num_format': '#,##0', 'border': 1, 'bg_color': '#FFFF00'})
            fmt_total_float = workbook.add_format(
                {'bold': True, 'num_format': '#,##0.00', 'border': 1, 'bg_color': '#FFFF00'})

            worksheet.set_row(0, 30);
            worksheet.set_row(1, 30)
            headers_row1 = ["NO.", "Lot (서류발송)", "선적일", "입항일", "ROLL", "kg", "SQM", "외화물품대", "발주월", "공장입고", "통관",
                            "해상운임", "", "", "", "", "", "", "", "", "", "", "국내운송", "", "", "", "", "", "외화", "", "총계",
                            "", "컨테이너", ""]
            headers_row2 = ["", "", "", "", "", "", "", "", "", "", "", "외화($)", "환율", "해상운임", "샤시운임", "선적지 서류",
                            "유류할증료", "프리풀", "선적지 내륙운송", "스토리지", "기타", "원화(\\)", "컨테이너 1대 당\n안전운임제(철송X)",
                            "컨테이너 1대 당\n안전운임제(철송)", "컨테이너 1대당\n트러킹(기본)", "총 컨테이너 포함\n트러킹 총비용", "국내운송비(트러킹제외)", "계",
                            "kg", "SQM", "총 금액", "SQM", "수량", "대당 평균 비용"]

            for i in range(11): worksheet.merge_range(0, i, 1, i, headers_row1[i], fmt_header)
            worksheet.merge_range(0, 11, 0, 21, "해상운임", fmt_header)
            for i in range(11, 22): worksheet.write(1, i, headers_row2[i], fmt_header)
            worksheet.merge_range(0, 22, 0, 27, "국내운송", fmt_header)
            for i in range(22, 28): worksheet.write(1, i, headers_row2[i], fmt_header)
            worksheet.merge_range(0, 28, 0, 29, "외화", fmt_header)
            for i in range(28, 30): worksheet.write(1, i, headers_row2[i], fmt_header)
            worksheet.merge_range(0, 30, 0, 31, "총계", fmt_header)
            for i in range(30, 32): worksheet.write(1, i, headers_row2[i], fmt_header)
            worksheet.merge_range(0, 32, 0, 33, "컨테이너", fmt_header)
            for i in range(32, 34): worksheet.write(1, i, headers_row2[i], fmt_header)

            worksheet.write(0, 34, "톤당 CPT 운임", fmt_yellow_header)
            worksheet.write(0, 35, "톤당 FCA 운임", fmt_yellow_header)
            worksheet.write(1, 34, "오더별 적용", fmt_header)
            last_data_row = len(results) + 2
            fca_formula = f"=IFERROR(ROUND((SUM(AE3:AE{last_data_row})/ROUND(SUM(V3:V{last_data_row})/SUM(L3:L{last_data_row}),1))/(SUM(F3:F{last_data_row})/1000),2)+28, 0)"
            worksheet.write_formula(1, 35, fca_formula, fmt_red_currency, calculated_fca_rate)
            worksheet.merge_range(0, 36, 1, 36, "CPT 원화 환산", fmt_blue_header)
            worksheet.merge_range(0, 37, 1, 37, "FCA 원화 환산", fmt_blue_header)
            worksheet.merge_range(0, 38, 1, 38, "운임 차액", fmt_blue_header)
            worksheet.write(0, 39, "");
            worksheet.write(1, 39, "")
            worksheet.merge_range(0, 40, 1, 40, "안전운임제 합계", fmt_yellow_header)
            worksheet.merge_range(0, 41, 1, 41, "유류할증료 합계", fmt_yellow_header)

            for r_idx, row_data in enumerate(results):
                xls_r = r_idx + 2;
                excel_row = xls_r + 1
                for c_idx, val in enumerate(row_data):
                    if c_idx >= 42: continue

                    if isinstance(val, float):
                        val = 0 if pd.isna(val) or val == float('inf') or val == float('-inf') else val
                    elif pd.isna(val):
                        val = ""

                    if c_idx == 34:
                        worksheet.write_formula(xls_r, c_idx, f"=F{excel_row}/1000*{row_data[42]}", fmt_float,
                                                float(val) if val else 0)
                    elif c_idx == 35:
                        worksheet.write_formula(xls_r, c_idx, f"=F{excel_row}/1000*AJ$2", fmt_float,
                                                float(val) if val else 0)
                    elif c_idx == 36:
                        worksheet.write_formula(xls_r, c_idx, f"=AI{excel_row}*M{excel_row}", fmt_num,
                                                float(val) if val else 0)
                    elif c_idx == 37:
                        worksheet.write_formula(xls_r, c_idx, f"=AJ{excel_row}*M{excel_row}", fmt_num,
                                                float(val) if val else 0)
                    elif c_idx == 38:
                        worksheet.write_formula(xls_r, c_idx, f"=AL{excel_row}-AK{excel_row}", fmt_num,
                                                float(val) if val else 0)
                    elif val == "":
                        worksheet.write(xls_r, c_idx, "", fmt_center)
                    elif c_idx in (2, 3, 9, 10):
                        try:
                            if hasattr(val, 'strftime'):
                                worksheet.write_datetime(xls_r, c_idx, val, fmt_date)
                            elif val and isinstance(val, str):
                                dt = pd.to_datetime(val); worksheet.write_datetime(xls_r, c_idx, dt, fmt_date)
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
            for i in range(42): worksheet.write(total_r, i, "", fmt_total)
            worksheet.write(total_r, 1, "총계", fmt_total)
            sums_mapping = {4: "E", 5: "F", 6: "G", 7: "H", 11: "L", 21: "V", 27: "AB", 30: "AE", 34: "AI", 35: "AJ",
                            36: "AK", 37: "AL", 38: "AM"}
            for c_idx, col_letter in sums_mapping.items():
                formula = f"=SUM({col_letter}3:{col_letter}{last_data_row})"
                fmt = fmt_total_float if c_idx in [11, 34, 35] else fmt_total_num
                worksheet.write_formula(total_r, c_idx, formula, fmt)

            worksheet.set_column(0, 0, 5);
            worksheet.set_column(1, 1, 15);
            worksheet.set_column(2, 3, 11);
            worksheet.set_column(4, 8, 10)
            worksheet.set_column(9, 10, 11);
            worksheet.set_column(11, 21, 11);
            worksheet.set_column(22, 27, 13);
            worksheet.set_column(28, 33, 11)
            worksheet.set_column(34, 38, 14);
            worksheet.set_column(39, 39, 3);
            worksheet.set_column(40, 41, 15)

            worksheet_audit = workbook.add_worksheet("검증_리포트")
            fmt_audit_title = workbook.add_format(
                {'bold': True, 'valign': 'vcenter', 'bg_color': '#FFC000', 'border': 1, 'font_size': 12})
            fmt_audit_text = workbook.add_format({'valign': 'vcenter', 'border': 1})
            worksheet_audit.set_column(0, 0, 120);
            worksheet_audit.set_row(0, 30)
            worksheet_audit.write(0, 0, "💡 정산 특이사항 요약", fmt_audit_title)
            if notes:
                for idx, note in enumerate(list(set(notes))):
                    worksheet_audit.write(idx + 1, 0, note.replace("**", ""), fmt_audit_text)
                    worksheet_audit.set_row(idx + 1, 25)
            else:
                worksheet_audit.write(1, 0, "🎉 특이사항 없음: 모든 건이 분기 계약 단가와 완벽하게 일치합니다.", fmt_audit_text)
                worksheet_audit.set_row(1, 25)

if is_processed:
    st.subheader(f"📊 당월 물류 운송 브리핑 [자동 감지: 다국가 오더 혼합 처리 완료]")

    col1, col2, col3, col4 = st.columns(4)
    col1.metric("당월 진행 오더", f"{total_orders} 건")
    col2.metric("당월 운송 컨테이너", f"{total_qty:,.0f} 대")
    col3.metric("당월 운송 톤수", f"{total_kg / 1000:,.1f} 톤")
    col4.metric("당월 발생 비용", f"{total_amount_krw:,.0f} 원")
    st.write("---")

    st.subheader("💡 정산 특이사항 요약")
    if notes:
        for note in list(set(notes)):
            if "🛑" in note or "❌" in note:
                st.error(note)
            elif "✅" in note:
                st.success(note)
            elif "🚚" in note:
                st.warning(note)
            else:
                st.info(note)
    else:
        st.success("🎉 특이사항 없음: 모든 건이 분기 계약 단가와 100% 일치합니다!")

    st.write("---")
    st.subheader("📥 최종 자료 다운로드")
    st.download_button(
        label="📊 클릭하여 최종 엑셀 파일 다운로드",
        data=buffer.getvalue(),
        file_name="판토스_물류비용정산_최종본.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
else:
    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown("### 💡 처음 오셨나요? 이렇게 사용해 보세요!")
    st.markdown("<br>", unsafe_allow_html=True)
    col1, col2 = st.columns(2)
    with col1:
        st.markdown(
            '<div class="guide-box"><div class="guide-step">📝 Step 1</div><div class="guide-desc"><b>반입계획서 (엑셀)</b><br>업로드</div></div>',
            unsafe_allow_html=True)
    with col2:
        st.markdown(
            '<div class="guide-box"><div class="guide-step">📄 Step 2</div><div class="guide-desc"><b>개별 인보이스 (PDF 여러 개)</b><br>마우스로 드래그 앤 드롭</div></div>',
            unsafe_allow_html=True)
