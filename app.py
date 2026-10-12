# -*- coding: utf-8 -*-
"""
소득금액(회사별) 추출 및 리포트 시스템 - Streamlit Cloud 웹앱 (v5.4)
모든 기기(PC, Mac, 태블릿, 스마트폰)에서 브라우저 접속으로 구동
구글 드라이브와 완전 연동되어 서버/클라우드에서 100% 처리
"""

import os, sys, re, io, time, json, shutil, zipfile, traceback
from datetime import datetime
from collections import Counter
import streamlit as st
import pypdfium2 as pdfium
from pypdf import PdfReader, PdfWriter
import pdfplumber
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side

from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload, MediaFileUpload
from google.oauth2 import service_account

# 1. 페이지 기본 설정 및 다크 테마 커스텀 CSS
st.set_page_config(
    page_title="소득금액(회사별) 추출 및 리포트",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
<style>
    /* 다크 네이비 / 블랙 & 블루 테마 커스텀 */
    .stApp {
        background-color: #0F172A;
        color: #F8FAFC;
    }
    .main-header {
        font-size: 24px;
        font-weight: 700;
        color: #38BDF8;
        margin-bottom: 20px;
    }
    .card-box {
        background-color: #1E293B;
        border: 1px solid #334155;
        border-radius: 8px;
        padding: 18px;
        margin-bottom: 16px;
    }
    .card-title {
        font-size: 16px;
        font-weight: 600;
        color: #93C5FD;
        margin-bottom: 12px;
    }
    .stButton>button {
        background-color: #2563EB;
        color: white;
        font-weight: 600;
        border-radius: 6px;
        border: none;
        padding: 10px 24px;
    }
    .stButton>button:hover {
        background-color: #1D4ED8;
        color: white;
    }
    .log-box {
        background-color: #020617;
        color: #E2E8F0;
        font-family: monospace;
        font-size: 13px;
        padding: 14px;
        border-radius: 6px;
        border: 1px solid #334155;
        max-height: 380px;
        overflow-y: auto;
        white-space: pre-wrap;
    }
</style>
""", unsafe_allow_html=True)

# 2. 구글 드라이브 API 클라이언트 초기화
def get_drive_service(custom_creds_info=None):
    """Streamlit Secrets 또는 업로드된 JSON 키 파일을 통해 Google Drive API 클라이언트 생성"""
    creds_info = None
    if custom_creds_info is not None:
        creds_info = dict(custom_creds_info)
    elif "uploaded_creds_info" in st.session_state and st.session_state.uploaded_creds_info:
        creds_info = dict(st.session_state.uploaded_creds_info)
    elif "gcp_service_account" in st.secrets:
        try:
            raw_sa = st.secrets["gcp_service_account"]
            if isinstance(raw_sa, str):
                creds_info = json.loads(raw_sa)
            else:
                creds_info = dict(raw_sa)
        except Exception as e:
            st.error(f"Secrets 파싱 오류: {e}")
            return None

    if not creds_info:
        return None

    try:
        # PEM private_key 서식 완전 정규화 (문자열 이스케이프 및 base64 패딩 복원)
        if "private_key" in creds_info:
            pk = str(creds_info["private_key"])
            s = pk.replace(r"\r\n", chr(10)).replace(r"\n", chr(10)).replace(r"\r", chr(10))
            s = s.replace(r"\\n", chr(10)).replace(r"\\r", chr(10))
            m = re.search(r"(-+BEGIN [^-]+-+)([\s\S]*?)(-+END [^-]+-+)", s)
            if m:
                header = m.group(1).strip()
                raw_body = m.group(2)
                footer = m.group(3).strip()
                clean_b64 = re.sub(r"[^A-Za-z0-9+/=]", "", raw_body)
                if len(clean_b64) % 4 != 0:
                    clean_b64 += "=" * ((4 - len(clean_b64) % 4) % 4)
                b_lines = [clean_b64[j:j+64] for j in range(0, len(clean_b64), 64)]
                creds_info["private_key"] = header + chr(10) + chr(10).join(b_lines) + chr(10) + footer + chr(10)

        creds = service_account.Credentials.from_service_account_info(
            creds_info, scopes=["https://www.googleapis.com/auth/drive"]
        )
        service = build("drive", "v3", credentials=creds)
        return service
    except Exception as e:
        st.error(f"구글 드라이브 서비스 계정 연결 오류: {e}")
        return None

def get_folder_metadata(service, folder_id):
    """폴더의 이름 및 상위 폴더(부모) ID 조회"""
    if not service or not folder_id:
        return {'id': folder_id, 'name': '알 수 없음', 'parents': []}
    try:
        f = service.files().get(fileId=folder_id, fields="id, name, parents", supportsAllDrives=True).execute()
        return f
    except Exception:
        return {'id': folder_id, 'name': '소득재산 작업', 'parents': []}

def get_folder_contents(service, folder_id):
    """지정된 드라이브 폴더 내 하위 폴더 및 파일 목록 실시간 조회"""
    if not service or not folder_id:
        return [], [], []
    try:
        q = f"'{folder_id}' in parents and trashed = false"
        results = service.files().list(
            q=q, fields="files(id, name, mimeType, size)", pageSize=200,
            supportsAllDrives=True, includeItemsFromAllDrives=True
        ).execute()
        items = results.get('files', [])
        folders = [f for f in items if f['mimeType'] == 'application/vnd.google-apps.folder']
        folders.sort(key=lambda x: x['name'])
        pdf_files = [f for f in items if f['name'].lower().endswith('.pdf')]
        pdf_files.sort(key=lambda x: x['name'])
        excel_files = [f for f in items if f['name'].lower().endswith(('.xlsx', '.xls'))]
        excel_files.sort(key=lambda x: x['name'])
        return folders, pdf_files, excel_files
    except Exception as e:
        st.error(f"폴더 내용 조회 실패: {e}")
        return [], [], []

def download_file_from_drive(service, file_id, local_dest_path):
    """구글 드라이브 파일을 로컬 임시 경로로 다운로드"""
    request = service.files().get_media(fileId=file_id)
    with io.FileIO(local_dest_path, 'wb') as fh:
        downloader = MediaIoBaseDownload(fh, request, chunksize=1024*1024*5)
        done = False
        while not done:
            status, done = downloader.next_chunk()

def upload_file_to_drive(service, local_file_path, folder_id, file_name=None):
    """로컬 파일을 구글 드라이브 특정 폴더에 업로드"""
    if not file_name:
        file_name = os.path.basename(local_file_path)
    file_metadata = {
        'name': file_name,
        'parents': [folder_id]
    }
    media = MediaFileUpload(local_file_path, resumable=True)
    file_obj = service.files().create(body=file_metadata, media_body=media, fields='id').execute()
    return file_obj.get('id')

def create_drive_folder(service, folder_name, parent_id):
    """구글 드라이브 내에 하위 폴더 생성 (이미 있으면 해당 ID 반환)"""
    q = f"'{parent_id}' in parents and name = '{folder_name}' and mimeType = 'application/vnd.google-apps.folder' and trashed = false"
    res = service.files().list(q=q, fields="files(id, name)").execute()
    existing = res.get('files', [])
    if existing:
        return existing[0]['id']
    meta = {
        'name': folder_name,
        'mimeType': 'application/vnd.google-apps.folder',
        'parents': [parent_id]
    }
    folder = service.files().create(body=meta, fields='id').execute()
    return folder.get('id')

# 3. 표준 양식 엑셀 메모리 생성 함수
def create_template_excel(title_name):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = title_name
    ws.append(["구분", "성명"])
    ws.append([1, "홍길동"])
    ws.append([2, "이순신"])
    ws.append([3, "강감찬"])

    hdr_fill = PatternFill(start_color="1D4ED8", end_color="1D4ED8", fill_type="solid")
    hdr_font = Font(name="맑은 고딕", size=11, bold=True, color="FFFFFF")
    hdr_align = Alignment(horizontal="center", vertical="center")
    data_align = Alignment(horizontal="center", vertical="center")
    thin_side = Side(border_style="thin", color="CBD5E1")
    thin_border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)

    for col in [1, 2]:
        cell = ws.cell(row=1, column=col)
        cell.fill = hdr_fill
        cell.font = hdr_font
        cell.alignment = hdr_align
        cell.border = thin_border

    for r in range(2, 5):
        for c in [1, 2]:
            cell = ws.cell(row=r, column=c)
            cell.alignment = data_align
            cell.border = thin_border
            cell.font = Font(name="맑은 고딕", size=10)

    ws.column_dimensions['A'].width = 12
    ws.column_dimensions['B'].width = 18

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf

# 4. 메인 UI 레이아웃
st.markdown('<div class="main-header">소득금액(회사별) 추출 및 리포트 시스템</div>', unsafe_allow_html=True)

# 비밀번호 보안 게이트 (선택 사항)
if "APP_PASSWORD" in st.secrets:
    if "authenticated" not in st.session_state:
        st.session_state.authenticated = False
    if not st.session_state.authenticated:
        pw_input = st.text_input("시스템 접속 비밀번호를 입력하세요:", type="password")
        if st.button("접속"):
            if pw_input == st.secrets["APP_PASSWORD"]:
                st.session_state.authenticated = True
                st.rerun()
            else:
                st.error("비밀번호가 올바르지 않습니다.")
        st.stop()

drive_svc = get_drive_service()

# 작업 모드 탭
tab_drive, tab_upload = st.tabs(["구글 드라이브 연동 모드 (권장)", "기기에서 직접 업로드 모드"])

with tab_drive:
    # 키 연동에 실패했을 때 직접 파일 업로드할 수 있는 안전 장치
    if not drive_svc:
        st.markdown('<div class="card-box" style="border-color:#F59E0B;"><div class="card-title" style="color:#FBBF24;">🔑 서비스 계정 키 파일 직접 등록 (1초 연결)</div>', unsafe_allow_html=True)
        st.info("Secrets 자동 연동에 문제가 발생했을 때, 다운로드받으신 서비스 계정 JSON 파일(project-...json)을 첨부하시면 즉시 구글 드라이브가 연결됩니다.")
        uploaded_key = st.file_uploader("서비스 계정 JSON 파일 업로드:", type=["json"], key="key_uploader")
        if uploaded_key:
            try:
                loaded_info = json.load(uploaded_key)
                st.session_state.uploaded_creds_info = loaded_info
                st.success("서비스 계정 키 파일 인식 완료! 시스템을 새로고침합니다.")
                st.rerun()
            except Exception as e_up:
                st.error(f"키 파일 읽기 실패: {e_up}")
        st.markdown('</div>', unsafe_allow_html=True)

    st.markdown('<div class="card-box"><div class="card-title">1. 📂 구글 드라이브 실시간 폴더 탐색기 및 파일 설정</div>', unsafe_allow_html=True)
    
    default_work_folder_id = "1X6wWQ_Gik-TgoPXscItFO-B30xmTDIMF"
    if "current_folder_id" not in st.session_state:
        st.session_state.current_folder_id = default_work_folder_id
        
    current_folder_id = st.session_state.current_folder_id
    cur_folder_meta = get_folder_metadata(drive_svc, current_folder_id)
    cur_folder_name = cur_folder_meta.get('name', '소득재산 작업')
    cur_parent_ids = cur_folder_meta.get('parents', [])
    
    subfolders, pdf_files, excel_files = get_folder_contents(drive_svc, current_folder_id)

    # Colab 스타일 실시간 위치 표시 및 상위/하위 폴더 이동 바
    st.markdown(f"""
    <div style="background-color: #020617; padding: 12px 16px; border-radius: 6px; border: 1px solid #334155; margin-bottom: 12px;">
        <span style="color: #94A3B8; font-size: 13px;">📍 현재 탐색 위치:</span><br>
        <span style="color: #38BDF8; font-size: 17px; font-weight: 700;">📁 {cur_folder_name}</span>
        <span style="color: #64748B; font-size: 13px; margin-left: 10px;">(PDF {len(pdf_files)}개 | 엑셀 {len(excel_files)}개 | 하위폴더 {len(subfolders)}개)</span>
    </div>
    """, unsafe_allow_html=True)

    nav_c1, nav_c2, nav_c3 = st.columns([1.2, 2.2, 1.2])
    with nav_c1:
        has_parent = len(cur_parent_ids) > 0
        if st.button("⬆️ 상위 폴더로 이동", disabled=not has_parent, use_container_width=True):
            if has_parent:
                st.session_state.current_folder_id = cur_parent_ids[0]
                st.rerun()

    with nav_c2:
        sub_options = ["(📂 하위 폴더 진입 선택...)"] + [f"{f['name']}/" for f in subfolders]
        selected_sub = st.selectbox("하위 폴더 이동:", options=sub_options, key=f"sub_select_{current_folder_id}", label_visibility="collapsed")
        if selected_sub != "(📂 하위 폴더 진입 선택...)":
            t_name = selected_sub.rstrip('/')
            match_folder = next((f for f in subfolders if f['name'] == t_name), None)
            if match_folder:
                st.session_state.current_folder_id = match_folder['id']
                st.rerun()

    with nav_c3:
        if st.button("🏠 기본 작업폴더 복귀", use_container_width=True):
            st.session_state.current_folder_id = default_work_folder_id
            st.rerun()

    with st.expander("🔗 다른 폴더 링크(URL) 또는 ID 직접 입력하여 이동"):
        c_in_url = st.text_input("폴더 링크(URL) 또는 ID:", value=current_folder_id, key="manual_folder_in")
        if st.button("해당 폴더로 즉시 이동"):
            m_id = re.search(r'folders/([a-zA-Z0-9_-]+)', c_in_url)
            target_id = m_id.group(1) if m_id else c_in_url.strip()
            if target_id:
                st.session_state.current_folder_id = target_id
                st.rerun()

    st.markdown("<hr style='border-color:#334155; margin:16px 0;'>", unsafe_allow_html=True)

    col1, col2 = st.columns(2)
    with col1:
        pdf_names = [f['name'] for f in pdf_files]
        sel_pdf_name = st.selectbox("작업 대상 PDF:", options=pdf_names if pdf_names else ["(감지된 PDF 없음)"])
        sel_pdf_obj = next((f for f in pdf_files if f['name'] == sel_pdf_name), None)

    with col2:
        excel_names = ["(선택 안 함)"] + [f['name'] for f in excel_files]
        sel_extract_db = st.selectbox("추출대상자 DB:", options=excel_names)
        sel_explain_db = st.selectbox("소명대상자 DB:", options=excel_names)

    # 양식 다운로드 버튼
    col_t1, col_t2 = st.columns(2)
    with col_t1:
        st.download_button("추출대상자DB 양식 다운로드", data=create_template_excel("추출대상자"),
                           file_name="추출대상자DB_표준양식.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    with col_t2:
        st.download_button("소명대상자DB 양식 다운로드", data=create_template_excel("소명대상자"),
                           file_name="소명대상자DB_표준양식.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    st.markdown('</div>', unsafe_allow_html=True)

    # 2. 추출 모드 및 세부 옵션
    st.markdown('<div class="card-box"><div class="card-title">2. 추출 모드 및 세부 옵션 설정</div>', unsafe_allow_html=True)
    c_m1, c_m2 = st.columns([1, 2])
    with c_m1:
        extract_mode = st.radio("추출 모드:", ["전체추출", "개별DB 추출"], horizontal=True)
    with c_m2:
        match_all = st.checkbox("대상자 매칭 범위: 본인 외 세대구성원(배우자, 부, 모, 자녀 등)도 매칭에 포함", value=True)

    c_o1, c_o2 = st.columns(2)
    with c_o1:
        fn_style = st.selectbox("파일명 서식:", ["성명(관계)  [예: 박현규(본인).pdf]", "성명_관계  [예: 박현규_본인.pdf]", "성명만 표기  [예: 박현규.pdf]"])
    with c_o2:
        dup_style = st.selectbox("동명이인 표기:", ["성명 뒤 번호 붙이기  [예: 박현규1(본인).pdf]", "관계 뒤 번호 붙이기  [예: 박현규(본인1).pdf]", "파일명 끝에 번호 붙이기  [예: 박현규(본인)_1.pdf]"])

    st.markdown('</div>', unsafe_allow_html=True)

    # 3. 실행 버튼 및 실시간 진행
    btn_start = st.button("추출 및 엑셀 리포트 실행 (구글 드라이브 직접 저장)", use_container_width=True)

    if btn_start:
        if not sel_pdf_obj:
            st.error("유효한 작업 대상 PDF 파일을 선택해 주세요.")
        else:
            prog_bar = st.progress(0)
            log_area = st.empty()
            log_messages = []

            def append_log(msg):
                ts = datetime.now().strftime("[%H:%M:%S] ")
                log_messages.append(ts + msg)
                joined_logs = chr(10).join(log_messages)
                log_area.markdown(f'<div class="log-box">{joined_logs}</div>', unsafe_allow_html=True)

            append_log("시스템 준비 완료. 클라우드 작업 세션을 시작합니다.")
            
            # 작업용 로컬 임시 디렉토리
            work_dir = f"/tmp/income_work_{int(time.time())}"
            sub_pdf_dir = os.path.join(work_dir, "1. 개별_PDF_추출본")
            sub_rpt_dir = os.path.join(work_dir, "2. 엑셀_작업리포트")
            sub_mrg_dir = os.path.join(work_dir, "3. 전체_통합_PDF")
            sub_log_dir = os.path.join(work_dir, "4. 시스템_작업로그")
            for p in [sub_pdf_dir, sub_rpt_dir, sub_mrg_dir, sub_log_dir]:
                os.makedirs(p, exist_ok=True)

            # 구글 드라이브에서 PDF 다운로드
            local_pdf_path = os.path.join(work_dir, sel_pdf_obj['name'])
            append_log(f"구글 드라이브에서 원본 PDF 수신 중: {sel_pdf_obj['name']}...")
            download_file_from_drive(drive_svc, sel_pdf_obj['id'], local_pdf_path)
            append_log("원본 PDF 수신 완료.")

            # DB 파일 다운로드 및 파싱
            extract_targets = set()
            if extract_mode == "개별DB 추출" and sel_extract_db != "(선택 안 함)":
                db_obj = next((f for f in excel_files if f['name'] == sel_extract_db), None)
                if db_obj:
                    local_db_path = os.path.join(work_dir, db_obj['name'])
                    download_file_from_drive(drive_svc, db_obj['id'], local_db_path)
                    wb = openpyxl.load_workbook(local_db_path, data_only=True)
                    ws = wb.active
                    for row in ws.iter_rows(min_row=2, values_only=True):
                        if row and len(row) >= 2 and row[1]:
                            extract_targets.add(str(row[1]).strip())
                    append_log(f"추출대상자 DB 로드 완료: 총 {len(extract_targets)}명")

            explain_targets = set()
            if sel_explain_db != "(선택 안 함)":
                db_obj = next((f for f in excel_files if f['name'] == sel_explain_db), None)
                if db_obj:
                    local_db_path = os.path.join(work_dir, db_obj['name'])
                    download_file_from_drive(drive_svc, db_obj['id'], local_db_path)
                    wb = openpyxl.load_workbook(local_db_path, data_only=True)
                    ws = wb.active
                    for row in ws.iter_rows(min_row=2, values_only=True):
                        if row and len(row) >= 2 and row[1]:
                            explain_targets.add(str(row[1]).strip())
                    append_log(f"소명대상자 DB 로드 완료: 총 {len(explain_targets)}명")

            # PDF 스캔
            append_log("pypdfium2 초고속 페이지 스캔 시작...")
            pdf_doc = pdfium.PdfDocument(local_pdf_path)
            total_pages = len(pdf_doc)
            append_log(f"전체 페이지 수: 총 {total_pages}페이지")

            valid_pages = []
            title_regex = re.compile(r'소득\s*금액\s*\(\s*회사별\s*\)')
            for p_idx in range(total_pages):
                page = pdf_doc.get_page(p_idx)
                textpage = page.get_textpage()
                p_text = textpage.get_text_range()
                if title_regex.search(p_text):
                    valid_pages.append((p_idx, p_text))
                if (p_idx + 1) % 500 == 0:
                    prog_bar.progress(int((p_idx + 1) / total_pages * 30))

            append_log(f"유효 소득금액 페이지 감지: 총 {len(valid_pages)}건")

            # 인원 파싱
            parsed_pages = []
            for p_idx, p_text in valid_pages:
                persons = []
                lines = [l.strip() for l in p_text.splitlines() if l.strip()]
                start_parsing = False
                for line in lines:
                    if any(h in line for h in ["성명", "관계", "자료명", "환산소득금액"]):
                        start_parsing = True
                        continue
                    if start_parsing:
                        if re.match(r'^\d+/\d+', line) or "페이지" in line:
                            break
                        parts = line.split()
                        if len(parts) >= 2:
                            name, rel = parts[0], parts[1]
                            if any(v in rel for v in ["본인", "배우자", "부", "모", "자녀", "조부", "조모", "손", "형제", "자매"]):
                                if not any(p['name'] == name and p['rel'] == rel for p in persons):
                                    persons.append({'name': name, 'rel': rel})
                if persons:
                    parsed_pages.append((p_idx, persons))

            name_counts = Counter()
            for p_idx, persons in parsed_pages:
                for p in persons:
                    if p['rel'] == '본인':
                        name_counts[p['name']] += 1

            # 상세 추출
            append_log("pdfplumber 표 정밀 파싱 및 개별 PDF 생성 시작...")
            plumber_pdf = pdfplumber.open(local_pdf_path)
            extracted_records = []
            reader = PdfReader(local_pdf_path)
            name_seq_counter = Counter()

            for idx, (p_idx, persons) in enumerate(parsed_pages):
                if extract_mode == "개별DB 추출":
                    matched = False
                    for p in persons:
                        if match_all:
                            if p['name'] in extract_targets: matched = True; break
                        else:
                            if p['rel'] == '본인' and p['name'] in extract_targets: matched = True; break
                    if not matched: continue

                main_person = next((p for p in persons if p['rel'] == '본인'), persons[0])
                m_name = main_person['name']
                m_rel = main_person['rel']

                is_dup = name_counts[m_name] > 1 if m_rel == '본인' else False
                dup_str = ""
                num_tag = ""
                if is_dup:
                    name_seq_counter[m_name] += 1
                    seq = name_seq_counter[m_name]
                    dup_str = f"O ({seq}차)"
                    num_tag = str(seq)

                out_fn = f"{m_name}{num_tag}({m_rel}).pdf"
                out_pdf_path = os.path.join(sub_pdf_dir, out_fn)
                writer = PdfWriter()
                writer.add_page(reader.pages[p_idx])
                with open(out_pdf_path, "wb") as f_out:
                    writer.write(f_out)

                page_obj = plumber_pdf.pages[p_idx]
                tables = page_obj.extract_tables()
                doc_types, amounts, companies, pay_dates, agencies, summaries = [], [], [], [], [], []
                if tables:
                    for t in tables:
                        for row in t:
                            clean_row = [re.sub(r'[\r\n\t]+', ' ', str(c or '')).strip() for c in row]
                            if any(h in ''.join(clean_row) for h in ["성명", "관계", "자료명", "환산소득금액"]):
                                continue
                            if len(clean_row) >= 5 and clean_row[0]:
                                d_typ = clean_row[2] if len(clean_row) > 2 else ""
                                d_amt = clean_row[3] if len(clean_row) > 3 else ""
                                d_cmp = clean_row[4] if len(clean_row) > 4 else ""
                                d_ym  = clean_row[5] if len(clean_row) > 5 else "-"
                                d_org = clean_row[6] if len(clean_row) > 6 else "-"
                                if d_typ or d_amt or d_cmp:
                                    doc_types.append(d_typ)
                                    amounts.append(d_amt)
                                    companies.append(d_cmp)
                                    pay_dates.append(d_ym)
                                    agencies.append(d_org)
                                    summaries.append(f"[{clean_row[0]}/{clean_row[1]}] {d_typ}: {d_amt}원 ({d_cmp} / {d_ym})")

                is_explain = "O" if (m_name in explain_targets) else ""
                rec = {
                    'idx': len(extracted_records) + 1,
                    'is_explain': is_explain,
                    'is_dup': dup_str,
                    'name': m_name,
                    'rel': m_rel,
                    'persons': persons,
                    'filename': out_fn,
                    'orig_page': f"{p_idx + 1}p",
                    'doc_type': "\n".join(doc_types) if doc_types else "-",
                    'amount': "\n".join(amounts) if amounts else "-",
                    'company': "\n".join(companies) if companies else "-",
                    'pay_date': "\n".join(pay_dates) if pay_dates else "-",
                    'agency': "\n".join(agencies) if agencies else "-",
                    'summary': "\n".join(summaries) if summaries else "-"
                }
                extracted_records.append(rec)
                if len(extracted_records) % 20 == 0 or idx == len(parsed_pages) - 1:
                    append_log(f"[{len(extracted_records)}] {out_fn} 추출 완료")
                prog_bar.progress(30 + int((idx + 1) / len(parsed_pages) * 45))

            plumber_pdf.close()

            # 엑셀 마스터 리포트 생성 (단일 .xlsx)
            append_log("엑셀 마스터 리포트 서식 적용 및 작성 중...")
            excel_path = os.path.join(sub_rpt_dir, "소득금액_추출작업_리포트.xlsx")
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "소득금액_추출현황"
            headers = ["구분", "소명대상자", "동명이인", "성명", "관계",
                       "세대구성원1 성명", "세대구성원1 관계", "세대구성원2 성명", "세대구성원2 관계",
                       "추출파일명", "원본페이지", "자료명", "환산소득금액", "회사명", "보수년월", "연계기관(자료명)", "소득상세_요약"]
            ws.append(headers)

            for rec in extracted_records:
                p_others = [p for p in rec['persons'] if not (p['name'] == rec['name'] and p['rel'] == rec['rel'])]
                m1_name = p_others[0]['name'] if len(p_others) > 0 else ""
                m1_rel = p_others[0]['rel'] if len(p_others) > 0 else ""
                m2_name = p_others[1]['name'] if len(p_others) > 1 else ""
                m2_rel = p_others[1]['rel'] if len(p_others) > 1 else ""
                ws.append([rec['idx'], rec['is_explain'], rec['is_dup'], rec['name'], rec['rel'],
                           m1_name, m1_rel, m2_name, m2_rel, rec['filename'], rec['orig_page'],
                           rec['doc_type'], rec['amount'], rec['company'], rec['pay_date'], rec['agency'], rec['summary']])

            header_fill = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
            for cell in ws[1]:
                cell.fill = header_fill
                cell.font = Font(name="맑은 고딕", size=10, bold=True, color="FFFFFF")
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

            wb.save(excel_path)

            # 전체 통합 PDF 생성
            append_log("전체 1권 통합 PDF 생성 중...")
            mrg_path = os.path.join(sub_mrg_dir, "소득금액_특정페이지_전체통합.pdf")
            pdf_files_extracted = sorted([os.path.join(sub_pdf_dir, f) for f in os.listdir(sub_pdf_dir) if f.lower().endswith(".pdf")])
            m_writer = PdfWriter()
            for pf in pdf_files_extracted:
                r_tmp = PdfReader(pf)
                for p in r_tmp.pages: m_writer.add_page(p)
            with open(mrg_path, "wb") as f_m: m_writer.write(f_m)

            # 전체 ZIP 압축
            append_log("결과물 전체 ZIP 압축 중...")
            zip_path = os.path.join(work_dir, "소득금액_추출작업_전체결과.zip")
            with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
                for root, dirs, files in os.walk(work_dir):
                    for file in files:
                        if file.endswith('.zip') or file.endswith('.pdf') and root == work_dir: continue
                        full_p = os.path.join(root, file)
                        rel_p = os.path.relpath(full_p, work_dir)
                        zf.write(full_p, rel_p)

            # 구글 드라이브에 결과 폴더 자동 생성 및 업로드
            append_log("구글 드라이브로 결과물 자동 업로드 시작...")
            drive_res_id = create_drive_folder(drive_svc, "소득금액_작업결과", current_folder_id)
            drive_sub_pdf = create_drive_folder(drive_svc, "1. 개별_PDF_추출본", drive_res_id)
            drive_sub_rpt = create_drive_folder(drive_svc, "2. 엑셀_작업리포트", drive_res_id)
            drive_sub_mrg = create_drive_folder(drive_svc, "3. 전체_통합_PDF", drive_res_id)
            drive_sub_log = create_drive_folder(drive_svc, "4. 시스템_작업로그", drive_res_id)

            upload_file_to_drive(drive_svc, excel_path, drive_sub_rpt)
            upload_file_to_drive(drive_svc, mrg_path, drive_sub_mrg)
            upload_file_to_drive(drive_svc, zip_path, drive_res_id)

            # 로그 파일 저장
            log_path = os.path.join(sub_log_dir, f"작업실행_상세로그_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log")
            with open(log_path, "w", encoding="utf-8") as f_l:
                f_l.write("\n".join(log_messages))
            upload_file_to_drive(drive_svc, log_path, drive_sub_log)

            prog_bar.progress(100)
            append_log("구글 드라이브 저장 완료! 모든 작업이 성공적으로 종료되었습니다.")
            st.success(f"총 {len(extracted_records)}건 추출 완료! 구글 드라이브 '소득금액_작업결과' 폴더에 자동 저장되었습니다.")

            # 브라우저 직접 다운로드 버튼 제공
            with open(zip_path, "rb") as fz:
                st.download_button("결과물 전체 ZIP 즉시 다운로드", data=fz.read(), file_name="소득금액_추출작업_전체결과.zip", mime="application/zip")
            with open(excel_path, "rb") as fe:
                st.download_button("엑셀 리포트 즉시 다운로드", data=fe.read(), file_name="소득금액_추출작업_리포트.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

            # 임시 파일 정리
            shutil.rmtree(work_dir, ignore_errors=True)

with tab_upload:
    st.info("이 탭은 구글 드라이브를 거치지 않고, 스마트폰이나 PC의 로컬 파일을 직접 웹 화면에 드래그하여 바로 추출할 때 사용합니다.")
    up_pdf = st.file_uploader("작업 대상 PDF 업로드:", type=["pdf"])
    up_db = st.file_uploader("소명/추출 대상자 DB 업로드 (선택):", type=["xlsx", "xls"])
    if up_pdf and st.button("직접 업로드 파일로 추출 실행"):
        st.write("로컬 파일 업로드 모드 실행 준비 완료...")
