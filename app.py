import eventlet
eventlet.monkey_patch()
# -*- coding: utf-8 -*-
import os
import socket
import re
import time
import random
import copy
import zipfile
from io import BytesIO

from pypdf import PdfReader, PdfWriter
from flask import Flask, render_template, request, jsonify, send_file, redirect
from flask_socketio import SocketIO, emit, join_room, leave_room
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

app = Flask(__name__)
app.config['SECRET_KEY'] = 'lan_exam_secret_key_2026'
socketio = SocketIO(app, cors_allowed_origins="*", ping_timeout=60, ping_interval=25)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
SUBMISSION_ROOT = os.path.join(BASE_DIR, 'bailamthisinh')
UPLOAD_DIR = os.path.join(BASE_DIR, 'uploads')
STATIC_PDF_DIR = os.path.join(BASE_DIR, 'static', 'pdf_exams')

for d in [SUBMISSION_ROOT, UPLOAD_DIR, STATIC_PDF_DIR]:
    if not os.path.exists(d):
        os.makedirs(d)

rooms = {}
excel_dirty_rooms = set()

# Thang điểm Đúng/Sai chuẩn theo công văn Bộ GD&ĐT
DEFAULT_TF_SCALE = {"1": 0.1, "2": 0.25, "3": 0.5, "4": 1.0}


def get_or_create_room(room_id):
    room_id = str(room_id).strip() if room_id else "default"
    if room_id not in rooms:
        room_dir = os.path.join(SUBMISSION_ROOT, room_id)
        if not os.path.exists(room_dir):
            os.makedirs(room_dir)
        rooms[room_id] = {
            "room_id": room_id,
            "class_name": "",
            "raw_banks": {},
            "active_exam": {
                "exam_type": "exam",
                "num_mcq": 0, "score_mcq": 0.0, "score_per_mcq": 0.0,
                "num_tf": 0, "score_tf": 0.0, "score_per_tf": 0.0,
                "num_sa": 0, "score_sa": 0.0, "score_per_sa": 0.0,
                "duration": 15,
                "tf_scale": copy.deepcopy(DEFAULT_TF_SCALE),
                "status": "waiting",
                "start_time": None
            },
            "students": {},
            "submitted_ips": set()
        }
    return rooms[room_id]


def get_submission_dir(room_id):
    d = os.path.join(SUBMISSION_ROOT, str(room_id))
    if not os.path.exists(d):
        os.makedirs(d)
    return d


def clear_submission_folder(room_id):
    room_dir = get_submission_dir(room_id)
    if os.path.exists(room_dir):
        for fname in os.listdir(room_dir):
            fpath = os.path.join(room_dir, fname)
            try:
                if os.path.isfile(fpath):
                    os.remove(fpath)
            except Exception:
                pass


# ==============================================================================
# HÀM BÓC TÁCH ĐÁP ÁN: MẶC ĐỊNH LẤY ĐẦU TRANG CUỐI CÙNG & CẮT TRANG CUỐI
# ==============================================================================
def parse_single_pdf_file(pdf_path, room_id, exam_code):
    reader = PdfReader(pdf_path)
    total_pages = len(reader.pages)
    
    if total_pages == 0:
        return None

    answer_page_idx = total_pages - 1
    raw_answer_content = reader.pages[answer_page_idx].extract_text() or ""

    writer = PdfWriter()
    for idx in range(total_pages - 1):
        writer.add_page(reader.pages[idx])

    clean_filename = f"clean_exam_{room_id}_{exam_code}.pdf"
    clean_filepath = os.path.join(STATIC_PDF_DIR, clean_filename)
    with open(clean_filepath, "wb") as f_out:
        writer.write(f_out)

    clean_pdf_url = f"/static/pdf_exams/{clean_filename}"

    mcq_questions = []
    tf_questions = []
    sa_questions = []

    lines = [line.strip() for line in raw_answer_content.splitlines() if line.strip()]
    labels_tf = ['a', 'b', 'c', 'd']

    for line in lines:
        clean_line = line.replace('\xa0', ' ').strip()
        
        if re.search(r'^(?:ĐÁP\s*ÁN|DAP\s*AN)', clean_line, re.IGNORECASE):
            continue

        m = re.match(r'^(?:câu\s*)?(\d+)\s*[\.\:\-]?\s*(.*)', clean_line, re.IGNORECASE)
        if not m:
            continue

        q_id = int(m.group(1))
        content = m.group(2).strip()

        if not content and '.' in clean_line:
            parts = clean_line.split('.', 1)
            if len(parts) == 2:
                q_id = int(re.sub(r'\D', '', parts[0]))
                content = parts[1].strip()

        if not content:
            continue

        # DẠNG 1: ĐÚNG / SAI
        chars_only = re.sub(r'[^A-ZĐa-zđ]', '', content).upper()
        if len(chars_only) == 4 and all(c in ['Đ', 'S', 'D', 'T', 'F'] for c in chars_only):
            sub_items = []
            for i, val in enumerate(chars_only):
                is_true = (val in ['Đ', 'D', 'T'])
                sub_items.append({"key": labels_tf[i], "text": "", "correct": is_true})
            tf_questions.append({
                "id": q_id,
                "type": "tf",
                "question": f"Câu {q_id}",
                "sub_items": sub_items
            })
            continue

        # DẠNG 2: TRẮC NGHIỆM ABCD
        clean_ans = re.sub(r'[^A-Za-z]', '', content).strip().upper()
        if len(clean_ans) == 1 and clean_ans in ['A', 'B', 'C', 'D']:
            chosen_key = clean_ans
            options = []
            for k in ['A', 'B', 'C', 'D']:
                options.append({"key": k, "text": "", "correct": (k == chosen_key)})
            mcq_questions.append({
                "id": q_id,
                "type": "mcq",
                "question": f"Câu {q_id}",
                "options": options
            })
            continue

        # DẠNG 3: TRẢ LỜI NGẮN
        sa_questions.append({
            "id": q_id,
            "type": "sa",
            "question": f"Câu {q_id}",
            "answers": [content.strip()]
        })

    mcq_questions.sort(key=lambda x: x["id"])
    tf_questions.sort(key=lambda x: x["id"])
    sa_questions.sort(key=lambda x: x["id"])

    print(f"\n--- [KẾT QUẢ BÓC TÁCH FILE: {exam_code}] ---")
    print(f"Tổng số trang: {total_pages} (Trang đáp án: {answer_page_idx + 1})")
    print(f"Phần 1 - Trắc nghiệm ({len(mcq_questions)} câu): {[q['id'] for q in mcq_questions]}")
    print(f"Phần 2 - Đúng/Sai     ({len(tf_questions)} câu): {[q['id'] for q in tf_questions]}")
    print(f"Phần 3 - Trả lời ngắn ({len(sa_questions)} câu): {[q['id'] for q in sa_questions]}")
    print(f"-> TỔNG CỘNG: {len(mcq_questions) + len(tf_questions) + len(sa_questions)} CÂU\n")

    return {
        "exam_code": exam_code,
        "pdf_url": clean_pdf_url,
        "mcq": mcq_questions,
        "tf": tf_questions,
        "sa": sa_questions
    }


# ==============================================================================
# HỆ THỐNG GIÁM SÁT THỜI GIAN & GHI FILE EXCEL GOM NHÓM CHỐNG NGHẼN
# ==============================================================================
def auto_exam_watcher():
    tick = 0
    while True:
        eventlet.sleep(1)
        tick += 1

        if tick % 2 == 0 and excel_dirty_rooms:
            for r_id in list(excel_dirty_rooms):
                excel_dirty_rooms.discard(r_id)
                try:
                    update_summary_excel(r_id)
                except Exception as e:
                    print(f"[!] Lỗi cập nhật Excel phòng {r_id}: {e}")

        for room_id, rdata in list(rooms.items()):
            active = rdata["active_exam"]
            if active["status"] == "running" and active["start_time"]:
                elapsed = time.time() - active["start_time"]
                total_duration = active["duration"] * 60

                if elapsed >= total_duration + 3:
                    has_updates = False
                    for sid, st in list(rdata["students"].items()):
                        if st.get("submitted") != "Đã nộp":
                            st["time_left"] = 0
                            process_grading(room_id, sid, st.get("answers", {}), st.get("name"), emit_socket=False)
                            has_updates = True

                    if has_updates:
                        active["status"] = "finished"
                        socketio.emit('update_teacher_list', list(rdata["students"].values()), to=room_id)


eventlet.spawn(auto_exam_watcher)


def get_client_ip(req):
    if req.headers.get('X-Forwarded-For'):
        return req.headers.get('X-Forwarded-For').split(',')[0].strip()
    return req.remote_addr or '127.0.0.1'


def get_cloudflare_url():
    log_path = os.path.join(BASE_DIR, "cloudflare.log")
    if os.path.exists(log_path):
        try:
            with open(log_path, "r", encoding="utf-8", errors="ignore") as f:
                content = f.read()
                matches = re.findall(r'https://[a-zA-Z0-9-]+\.trycloudflare\.com', content)
                if matches:
                    return matches[-1]
        except Exception:
            pass
    return None


def get_base_url():
    cf_url = get_cloudflare_url()
    if cf_url:
        return cf_url

    render_url = os.environ.get('RENDER_EXTERNAL_URL')
    if render_url:
        return render_url.rstrip('/')

    s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
    try:
        s.connect(('8.8.8.8', 80))
        ip = s.getsockname()[0]
    except Exception:
        ip = '127.0.0.1'
    finally:
        s.close()
    return f"http://{ip}:5000"


def normalize_sa_string(val):
    if val is None:
        return ""
    val_str = str(val).replace('\xa0', ' ').strip().lower().replace(',', '.')
    val_str = re.sub(r'[\.\:\;\s]+$', '', val_str)
    return val_str


def generate_individual_exam(room_data, assigned_code=None):
    raw_banks = room_data.get("raw_banks", {})
    available_codes = list(raw_banks.keys())

    if not available_codes:
        return {"is_pdf": False, "mcq": [], "tf": [], "sa": [], "duration": 15}, {}

    if not assigned_code or assigned_code not in raw_banks:
        assigned_code = random.choice(available_codes)

    chosen_bank = raw_banks[assigned_code]

    all_mcq = chosen_bank["mcq"]
    all_tf = chosen_bank["tf"]
    all_sa = chosen_bank["sa"]

    mcq_send, mcq_server = [], []
    for q in all_mcq:
        client_opts, server_opts = [], []
        for opt in q.get('options', []):
            lbl = opt["key"]
            client_opts.append({"label": lbl, "orig_key": lbl})
            server_opts.append({"label": lbl, "orig_key": lbl, "correct": opt.get("correct", False)})

        mcq_send.append({"id": q["id"], "type": "mcq", "options": client_opts})
        mcq_server.append({"id": q["id"], "type": "mcq", "options": server_opts})

    tf_send, tf_server = [], []
    for q in all_tf:
        client_subs = [{"key": s["key"]} for s in q.get("sub_items", [])]
        server_subs = [{"key": s["key"], "correct": s.get("correct", False)} for s in q.get("sub_items", [])]
        tf_send.append({"id": q["id"], "type": "tf", "sub_items": client_subs})
        tf_server.append({"id": q["id"], "type": "tf", "sub_items": server_subs})

    sa_send, sa_server = [], []
    for q in all_sa:
        sa_send.append({"id": q["id"], "type": "sa"})
        sa_server.append({"id": q["id"], "type": "sa", "answers": q.get("answers", [])})

    client_payload = {
        "is_pdf": True,
        "exam_code": assigned_code,
        "pdf_url": chosen_bank.get("pdf_url", ""),
        "mcq": mcq_send,
        "tf": tf_send,
        "sa": sa_send,
        "duration": room_data["active_exam"]["duration"]
    }
    server_exam = {
        "exam_code": assigned_code,
        "pdf_url": chosen_bank.get("pdf_url", ""),
        "mcq": mcq_server,
        "tf": tf_server,
        "sa": sa_server
    }
    return client_payload, server_exam


def update_summary_excel(room_id):
    room_data = get_or_create_room(room_id)
    room_dir = get_submission_dir(room_id)

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Bang_Diem_Tong_Hop"

    thin_border = Border(
        left=Side(style='thin', color='B0B0B0'),
        right=Side(style='thin', color='B0B0B0'),
        top=Side(style='thin', color='B0B0B0'),
        bottom=Side(style='thin', color='B0B0B0')
    )
    navy_fill = PatternFill(start_color="0B3C6D", end_color="0B3C6D", fill_type="solid")
    header_font = Font(name="Arial", size=10, bold=True, color="FFFFFF")
    bold_font = Font(name="Arial", size=10, bold=True)
    regular_font = Font(name="Arial", size=10)

    headers = ["TT", "Mã đề", "Địa chỉ IP", "Họ và tên", "Thời gian làm bài", "Vi phạm", "Trạng thái", "Điểm số"]
    ws.append(headers)
    for col_idx in range(1, len(headers) + 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.fill = navy_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = thin_border

    scores_list = []
    idx = 1
    current_row = 2

    for sid, st in list(room_data["students"].items()):
        spent_secs = st.get('time_spent', 0)
        mins = spent_secs // 60
        secs = spent_secs % 60
        score_val = float(st.get("score", 0.0))
        scores_list.append(score_val)

        ws.append([
            idx,
            st.get("exam_code", "–"),
            st["ip"],
            st["name"],
            f"{mins:02d}:{secs:02d}",
            st["violations"],
            st["submitted"],
            score_val
        ])

        for col_idx in range(1, 9):
            c = ws.cell(row=current_row, column=col_idx)
            c.font = regular_font
            c.border = thin_border
            if col_idx == 4:
                c.alignment = Alignment(horizontal="left", vertical="center")
            elif col_idx == 8:
                c.alignment = Alignment(horizontal="center", vertical="center")
                c.font = bold_font
            else:
                c.alignment = Alignment(horizontal="center", vertical="center")

        idx += 1
        current_row += 1

    total_students = len(scores_list)
    cnt_duoi_5 = sum(1 for s in scores_list if s <= 5.0)
    cnt_duoi_65 = sum(1 for s in scores_list if 5.0 < s <= 6.5)
    cnt_duoi_8 = sum(1 for s in scores_list if 6.5 < s <= 8.0)
    cnt_tren_8 = sum(1 for s in scores_list if s > 8.0)

    def calc_rate(count):
        if total_students == 0:
            return "0.0%"
        return f"{round((count / total_students) * 100, 1)}%"

    stat_start_row = current_row + 2
    stat_headers = ["Mức điểm", "Tỉ lệ%", "Số lượng"]
    stat_data = [
        ["0 <= điểm <= 5.0", calc_rate(cnt_duoi_5), cnt_duoi_5],
        ["5.0 < điểm <= 6.5", calc_rate(cnt_duoi_65), cnt_duoi_65],
        ["6.5 < điểm <= 8.0", calc_rate(cnt_duoi_8), cnt_duoi_8],
        ["8.0 < điểm <= 10", calc_rate(cnt_tren_8), cnt_tren_8]
    ]

    for c_idx, h_text in enumerate(stat_headers, 1):
        cell = ws.cell(row=stat_start_row, column=c_idx, value=h_text)
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = thin_border
        cell.fill = navy_fill

    for r_offset, row_values in enumerate(stat_data, 1):
        r_idx = stat_start_row + r_offset
        for c_idx, val in enumerate(row_values, 1):
            cell = ws.cell(row=r_idx, column=c_idx, value=val)
            cell.font = regular_font
            cell.border = thin_border
            cell.alignment = Alignment(horizontal="center", vertical="center")

    col_widths = {1: 8, 2: 12, 3: 20, 4: 32, 5: 20, 6: 12, 7: 15, 8: 14}
    for col_idx, width in col_widths.items():
        col_letter = get_column_letter(col_idx)
        ws.column_dimensions[col_letter].width = width

    summary_path = os.path.join(room_dir, "Bang_Diem_Tong_Hop.xlsx")
    wb.save(summary_path)


def save_individual_submission(room_id, st, answers, score):
    room_dir = get_submission_dir(room_id)
    safe_name = re.sub(r'[^\w\s-]', '', st['name']).strip().replace(" ", "_")
    safe_ip = st["ip"].replace(":", "_")
    filename = f"{safe_name}_{safe_ip}.txt"
    filepath = os.path.join(room_dir, filename)

    exam_data = st.get("exam", {})
    NL = "\r\n"

    with open(filepath, "w", encoding="utf-8") as f:
        f.write(f" Họ và tên thí sinh : {st['name']}" + NL)
        f.write(f" Mã đề thi          : {st.get('exam_code', '–')}" + NL)
        f.write(f" Địa chỉ IP         : {st['ip']}" + NL)
        f.write(f" Trạng thái         : {st['submitted']}" + NL)
        f.write(f" Số lần vi phạm     : {st['violations']}" + NL)
        f.write(f" Tổng điểm          : {score} điểm" + NL)
        f.write("-" * 60 + NL)

        f.write(" PHẦN 1: TRẮC NGHIỆM NHIỀU LỰA CHỌN" + NL)
        f.write("-" * 60 + NL)
        for q in exam_data.get("mcq", []):
            ans = answers.get(f"mcq_{q['id']}", "Chưa làm")
            correct_opt = next((opt['label'] for opt in q.get('options', []) if opt.get('correct')), "Chưa rõ")
            f.write(f"Câu {q['id']}:" + NL)
            f.write(f"   Thí sinh chọn : {ans}" + NL)
            f.write(f"   Đáp án : {correct_opt}" + NL + NL)

        f.write(" PHẦN 2: TRẮC NGHIỆM ĐÚNG / SAI" + NL)
        f.write("-" * 60 + NL)
        for q in exam_data.get("tf", []):
            f.write(f"Câu {q['id']}:" + NL)
            for sub in q.get('sub_items', []):
                user_val = answers.get(f"tf_{q['id']}_{sub['key']}", "Chưa làm")
                expected = "Đ" if sub.get('correct') else "S"
                f.write(f"   {sub['key']}) Thí sinh: {user_val} | Đáp án: {expected}" + NL)
            f.write(NL)

        f.write(" PHẦN 3: CÂU HỎI TRẢ LỜI NGẮN" + NL)
        f.write("-" * 60 + NL)
        for q in exam_data.get("sa", []):
            user_val = answers.get(f"sa_{q['id']}", "Chưa làm")
            expected_str = " / ".join(q.get("answers", [])) or "Chưa rõ"
            f.write(f"Câu {q['id']}:" + NL)
            f.write(f"   Thí sinh    : {user_val}" + NL)
            f.write(f"   Đáp án chuẩn: {expected_str}" + NL + NL)


def find_student_record(room_data, sid=None, student_name=None, client_ip=None):
    if sid and sid in room_data["students"]:
        return sid, room_data["students"][sid]

    if student_name and client_ip:
        for s_id, s_data in room_data["students"].items():
            if s_data.get("name", "").strip().lower() == student_name.strip().lower() and s_data.get("ip") == client_ip:
                return s_id, s_data

    if student_name:
        for s_id, s_data in room_data["students"].items():
            if s_data.get("name", "").strip().lower() == student_name.strip().lower():
                return s_id, s_data

    return None, None


def process_grading(room_id, sid=None, answers=None, student_name=None, exam_code=None, client_ip=None, emit_socket=True):
    room_data = get_or_create_room(room_id)
    target_sid, st = find_student_record(room_data, sid, student_name, client_ip)

    if not st:
        return None

    if exam_code:
        st["exam_code"] = exam_code

    if not st.get("exam") and st.get("exam_code") in room_data.get("raw_banks", {}):
        _, server_exam = generate_individual_exam(room_data, st["exam_code"])
        st["exam"] = server_exam

    is_practice = (room_data["active_exam"].get("exam_type") == "practice")
    student_exam = st.get("exam") or {}

    correct_answers = {}
    if is_practice:
        for q in student_exam.get("mcq", []):
            c_opt = next((opt['label'] for opt in q.get('options', []) if opt.get('correct')), "")
            correct_answers[f"mcq_{q['id']}"] = c_opt
        for q in student_exam.get("tf", []):
            for sub in q.get('sub_items', []):
                correct_answers[f"tf_{q['id']}_{sub['key']}"] = "Đ" if sub.get('correct') else "S"
        for q in student_exam.get("sa", []):
            correct_answers[f"sa_{q['id']}"] = " / ".join(q.get("answers", []))

    ack_payload = {
        "score": st.get("score", 0.0),
        "show_answers": is_practice,
        "correct_answers": correct_answers if is_practice else {}
    }

    if st.get("submitted") == "Đã nộp":
        if emit_socket and sid:
            socketio.emit('exam_finished_ack', ack_payload, room=sid)
            if target_sid and target_sid != sid:
                socketio.emit('exam_finished_ack', ack_payload, room=target_sid)
        return ack_payload

    if not answers and st.get("answers"):
        answers = st["answers"]
    elif not answers:
        answers = {}

    total_allowed_secs = room_data["active_exam"]["duration"] * 60
    if room_data["active_exam"].get("start_time"):
        elapsed = int(time.time() - room_data["active_exam"]["start_time"])
        time_spent_secs = min(total_allowed_secs, max(0, elapsed))
    else:
        current_left = st.get("time_left", 0)
        time_spent_secs = max(0, total_allowed_secs - current_left)
    st["time_spent"] = time_spent_secs

    st["submitted"] = "Đã nộp"
    st["answers"] = answers
    st["time_left"] = 0
    room_data["submitted_ips"].add(st["ip"])

    score = 0.0

    # Chấm Phần 1: Trắc nghiệm ABCD
    mcq_pts = room_data["active_exam"].get("score_per_mcq", 0.25)
    for q in student_exam.get("mcq", []):
        ans = answers.get(f"mcq_{q['id']}")
        correct_opt = next((opt['label'] for opt in q.get('options', []) if opt.get('correct')), None)
        if ans and str(ans).strip().upper() == str(correct_opt).strip().upper():
            score += mcq_pts

    # Chấm Phần 2: Đúng / Sai
    tf_pts = room_data["active_exam"].get("score_per_tf", 1.0)
    scale = room_data["active_exam"].get("tf_scale", DEFAULT_TF_SCALE)
    for q in student_exam.get("tf", []):
        correct_cnt = 0
        for sub in q.get('sub_items', []):
            user_val = answers.get(f"tf_{q['id']}_{sub['key']}")
            expected = "Đ" if sub.get('correct') else "S"
            if user_val and str(user_val).strip() == expected:
                correct_cnt += 1
        ratio = float(scale.get(str(correct_cnt), 0.0))
        score += tf_pts * ratio

    # Chấm Phần 3: Trả lời ngắn / Điền số
    sa_pts = room_data["active_exam"].get("score_per_sa", 0.0)
    for q in student_exam.get("sa", []):
        user_val = answers.get(f"sa_{q['id']}", "")
        if user_val:
            cleaned_user = normalize_sa_string(user_val)
            raw_target_list = q.get("answers", [])
            valid_targets = []
            for t in raw_target_list:
                for sub_t in str(t).split('/'):
                    v = normalize_sa_string(sub_t)
                    if v:
                        valid_targets.append(v)

            if cleaned_user and cleaned_user in valid_targets:
                score += sa_pts

    st["score"] = round(score, 2)
    ack_payload["score"] = st["score"]

    save_individual_submission(room_id, st, answers, st["score"])
    excel_dirty_rooms.add(room_id)

    if emit_socket:
        socketio.emit('update_teacher_list', list(room_data["students"].values()), to=room_id)
        if sid:
            socketio.emit('exam_finished_ack', ack_payload, room=sid)
        if target_sid and target_sid != sid:
            socketio.emit('exam_finished_ack', ack_payload, room=target_sid)

    return ack_payload


# ==============================================================================
# ROUTES WEB & API DỰ PHÒNG NỘP BÀI HTTP
# ==============================================================================
@app.route('/')
def teacher_dashboard():
    return render_template('index.html')


@app.route('/kiemtra')
def preview_exam_page():
    return render_template('kiemtra.html')


@app.route('/student')
def student_view():
    return render_template('student.html')


@app.route('/r/<room_id>')
def short_room(room_id):
    return redirect(f"/student?room={room_id}")


@app.route('/ping')
def ping():
    return "pong", 200


@app.route('/api/submit_exam', methods=['POST'])
def api_submit_exam():
    data = request.json or {}
    room_id = str(data.get('room', 'default')).strip()
    name = str(data.get('name', '')).strip()
    exam_code = data.get('exam_code')
    answers = data.get('answers', {})
    client_ip = get_client_ip(request)

    ack = process_grading(
        room_id=room_id,
        sid=None,
        answers=answers,
        student_name=name,
        exam_code=exam_code,
        client_ip=client_ip,
        emit_socket=True
    )
    if ack:
        return jsonify({"success": True, "ack": ack})
    return jsonify({"success": False, "error": "Không tìm thấy thông tin thí sinh trong phòng!"})


@app.route('/api/get-tunnel-url', methods=['GET'])
def api_get_tunnel_url():
    cf_url = get_cloudflare_url()
    return jsonify({
        "url": cf_url if cf_url else get_base_url(),
        "is_tunnel": bool(cf_url)
    })


@app.route('/api/get_ip', methods=['GET'])
def api_get_ip():
    room_id = request.args.get('room', '')
    base_url = get_base_url()
    if room_id:
        url = f"{base_url}/r/{room_id}"
    else:
        url = f"{base_url}/student"
    return jsonify({
        "ip": base_url.replace("http://", "").replace("https://", "").split(":")[0],
        "port": 5000,
        "url": url
    })


@app.route('/api/upload_bank', methods=['POST'])
def upload_bank():
    room_id = request.form.get('room', 'default')
    room_data = get_or_create_room(room_id)

    files = request.files.getlist('file')
    if not files or len(files) == 0:
        return jsonify({"success": False, "error": "Chưa chọn file PDF!"})

    parsed_banks = {}
    for f in files:
        filename = f.filename
        if not filename.lower().endswith('.pdf'):
            continue

        exam_code = os.path.splitext(filename)[0].strip()
        save_pdf_path = os.path.join(UPLOAD_DIR, f"{room_id}_{filename}")
        f.save(save_pdf_path)

        res = parse_single_pdf_file(save_pdf_path, room_id, exam_code)
        if res and (res["mcq"] or res["tf"] or res["sa"]):
            res["filename"] = filename
            res["is_pdf"] = True
            parsed_banks[exam_code] = res

    if not parsed_banks:
        return jsonify({
            "success": False,
            "error": "Không tìm thấy nội dung đáp án ở trang cuối đề hoặc cấu trúc chưa đúng quy ước!"
        })

    room_data["raw_banks"] = parsed_banks

    first_code = list(parsed_banks.keys())[0]
    sample = parsed_banks[first_code]

    n_mcq = len(sample["mcq"])
    n_tf = len(sample["tf"])
    n_sa = len(sample["sa"])
    
    room_data["active_exam"]["num_mcq"] = n_mcq
    room_data["active_exam"]["num_tf"] = n_tf
    room_data["active_exam"]["num_sa"] = n_sa

    return jsonify({
        "success": True,
        "total_codes": len(parsed_banks),
        "codes_list": list(parsed_banks.keys()),
        "total_mcq": n_mcq,
        "total_tf": n_tf,
        "total_sa": n_sa,
        "filename": ", ".join(list(parsed_banks.keys())),
        "pdf_url": sample["pdf_url"],
        "bank_data": parsed_banks
    })


@app.route('/api/configure_exam', methods=['POST'])
def configure_exam():
    data = request.json or {}
    room_id = data.get("room", "default")
    room_data = get_or_create_room(room_id)

    room_data["submitted_ips"].clear()
    room_data["students"].clear()
    clear_submission_folder(room_id)

    raw_banks = room_data.get("raw_banks", {})
    if not raw_banks:
        return jsonify({"success": False, "error": "Chưa có đề nào được nạp!"})

    first_code = list(raw_banks.keys())[0]
    sample = raw_banks[first_code]

    n_mcq = len(sample["mcq"])
    n_tf = len(sample["tf"])
    n_sa = len(sample["sa"])

    s_mcq = float(data.get("score_mcq", 0.0))
    s_tf = float(data.get("score_tf", 0.0))
    s_sa = float(data.get("score_sa", 0.0))
    dur = int(data.get("duration", 15))

    tf_scale_input = data.get("tf_scale")
    if tf_scale_input and isinstance(tf_scale_input, dict):
        final_tf_scale = {
            "1": float(tf_scale_input.get("1", DEFAULT_TF_SCALE["1"])),
            "2": float(tf_scale_input.get("2", DEFAULT_TF_SCALE["2"])),
            "3": float(tf_scale_input.get("3", DEFAULT_TF_SCALE["3"])),
            "4": float(tf_scale_input.get("4", DEFAULT_TF_SCALE["4"]))
        }
    else:
        final_tf_scale = copy.deepcopy(DEFAULT_TF_SCALE)

    room_data["class_name"] = data.get("class_name", "")
    room_data["active_exam"].update({
        "exam_type": data.get("exam_type", "exam"),
        "num_mcq": n_mcq,
        "score_mcq": s_mcq,
        "score_per_mcq": round(s_mcq / max(1, n_mcq), 3),
        "num_tf": n_tf,
        "score_tf": s_tf,
        "score_per_tf": round(s_tf / max(1, n_tf), 3),
        "num_sa": n_sa,
        "score_sa": s_sa,
        "score_per_sa": round(s_sa / max(1, n_sa), 3),
        "duration": dur,
        "tf_scale": final_tf_scale,
        "status": "ready"
    })
    socketio.emit('update_teacher_list', [], to=room_id)
    return jsonify({"success": True})


@app.route('/api/reset_exam', methods=['POST'])
def api_reset_exam():
    data = request.json or {}
    room_id = data.get("room", "default")
    room_data = get_or_create_room(room_id)

    room_data["submitted_ips"].clear()
    room_data["students"].clear()
    clear_submission_folder(room_id)
    room_data["class_name"] = ""
    room_data["active_exam"] = {
        "exam_type": "exam",
        "num_mcq": 0, "score_mcq": 0.0, "score_per_mcq": 0.0,
        "num_tf": 0, "score_tf": 0.0, "score_per_tf": 0.0,
        "num_sa": 0, "score_sa": 0.0, "score_per_sa": 0.0,
        "duration": 15,
        "tf_scale": copy.deepcopy(DEFAULT_TF_SCALE),
        "status": "waiting",
        "start_time": None
    }
    socketio.emit('update_teacher_list', [], to=room_id)
    return jsonify({"success": True})


@app.route('/api/export_excel', methods=['GET'])
def export_excel():
    room_id = request.args.get('room', 'default')
    room_data = get_or_create_room(room_id)
    room_dir = get_submission_dir(room_id)
    summary_path = os.path.join(room_dir, "Bang_Diem_Tong_Hop.xlsx")
    update_summary_excel(room_id)

    raw_class = room_data.get("class_name", "").strip() or f"Phong_{room_id}"
    safe_class_name = re.sub(r'[^\w\s-]', '', raw_class).strip().replace(" ", "_")
    download_filename = f"Bang_diem_tong_hop_lop_{safe_class_name}.xlsx"

    return send_file(summary_path, as_attachment=True, download_name=download_filename)


@app.route('/api/download_all_results', methods=['GET'])
def download_all_results():
    room_id = request.args.get('room', 'default')
    room_data = get_or_create_room(room_id)
    room_dir = get_submission_dir(room_id)

    update_summary_excel(room_id)
    memory_file = BytesIO()
    with zipfile.ZipFile(memory_file, 'w', zipfile.ZIP_DEFLATED) as zf:
        for root, dirs, files in os.walk(room_dir):
            for file in files:
                file_path = os.path.join(root, file)
                zf.write(file_path, arcname=file)

    memory_file.seek(0)
    class_name = room_data.get("class_name", f"Phong_{room_id}").replace(" ", "_")
    zip_filename = f"KetQua_BaiThi_{class_name}.zip"

    return send_file(
        memory_file,
        mimetype='application/zip',
        as_attachment=True,
        download_name=zip_filename
    )


# ==============================================================================
# SOCKETIO SỰ KIỆN PHÒNG THI
# ==============================================================================
@socketio.on('join_teacher')
def handle_teacher_join(data):
    room_id = data.get('room', 'default')
    join_room(room_id)
    room_data = get_or_create_room(room_id)
    emit('update_teacher_list', list(room_data["students"].values()), room=request.sid)


@socketio.on('join_student')
def handle_student_join(data):
    sid = request.sid
    room_id = data.get('room', 'default')
    join_room(room_id)
    room_data = get_or_create_room(room_id)
    client_ip = get_client_ip(request)

    name = data.get('name', 'Thí sinh').strip() or 'Thí sinh'

    existing_sid, existing_record = find_student_record(room_data, None, name, client_ip)

    if existing_record and existing_record.get("submitted") == "Đã nộp":
        emit('exam_blocked', {'message': 'Bạn đã hoàn thành bài thi trong phòng này!'}, room=sid)
        return

    # KHÓA SỔ PHÒNG THI: Chặn người lạ tự ý đăng nhập sau khi giáo viên đã bấm Bắt đầu làm bài
    if room_data["active_exam"]["status"] == "running" and not existing_record:
        emit('exam_blocked', {
            'message': 'Phòng thi đã khóa sổ sau khi bắt đầu phát đề! Không thể tham gia từ bên ngoài.'
        }, room=sid)
        return

    is_running = (room_data["active_exam"]["status"] == "running")
    initial_status = "Đang thi" if is_running else "Chờ"

    time_left = room_data["active_exam"]["duration"] * 60
    if is_running and room_data["active_exam"]["start_time"]:
        elapsed = int(time.time() - room_data["active_exam"]["start_time"])
        time_left = max(0, time_left - elapsed)

    available_codes = list(room_data.get("raw_banks", {}).keys())
    if available_codes:
        assigned_code = available_codes[len(room_data["students"]) % len(available_codes)]
    else:
        assigned_code = None

    if existing_record:
        if existing_sid and existing_sid in room_data["students"]:
            room_data["students"].pop(existing_sid)
        existing_record["ip"] = client_ip
        existing_record["time_left"] = time_left
        if existing_record.get("submitted") != "Đã nộp":
            existing_record["submitted"] = initial_status
        if not existing_record.get("exam_code") and assigned_code:
            existing_record["exam_code"] = assigned_code
        room_data["students"][sid] = existing_record
    else:
        room_data["students"][sid] = {
            "ip": client_ip,
            "name": name,
            "room": room_id,
            "exam_code": assigned_code,
            "time_left": time_left,
            "time_spent": 0,
            "violations": 0,
            "submitted": initial_status,
            "score": 0.0,
            "answers": {},
            "exam": None
        }

    emit('update_teacher_list', list(room_data["students"].values()), to=room_id)

    if is_running:
        st_obj = room_data["students"][sid]
        client_payload, server_exam = generate_individual_exam(room_data, st_obj["exam_code"])
        st_obj["exam"] = server_exam
        client_payload["time_left"] = time_left
        client_payload["saved_answers"] = st_obj.get("answers", {})
        emit('start_exam_now', client_payload, room=sid)


@socketio.on('sync_answer')
def handle_sync_answer(data):
    sid = request.sid
    room_id = data.get('room', 'default')
    student_name = data.get('name', '')
    room_data = get_or_create_room(room_id)
    _, st = find_student_record(room_data, sid, student_name, get_client_ip(request))
    if st:
        key = data.get("key")
        val = data.get("val")
        if key:
            st["answers"][key] = val


@socketio.on('sync_timer')
def handle_timer_sync(data):
    sid = request.sid
    room_id = data.get('room', 'default')
    student_name = data.get('name', '')
    room_data = get_or_create_room(room_id)
    target_sid, st = find_student_record(room_data, sid, student_name, get_client_ip(request))
    if st:
        t_left = data.get("time_left", 0)
        st["time_left"] = t_left

        if t_left <= 0 and st["submitted"] != "Đã nộp":
            process_grading(room_id, target_sid or sid, st.get("answers", {}), st.get("name"))


# ==============================================================================
# BẪY VÀ XỬ LÝ VI PHẠM (GHI NHẬN CẢ TRONG LÚC CHỜ, LÚC THI VÀ SAU KHI NỘP BÀI)
# ==============================================================================
@socketio.on('report_violation')
def handle_violation(data):
    sid = request.sid
    room_id = str(data.get('room', 'default')).strip() if isinstance(data, dict) else 'default'
    student_name = data.get('name', '').strip() if isinstance(data, dict) else ''
    room_data = get_or_create_room(room_id)
    
    _, st = find_student_record(room_data, sid, student_name, get_client_ip(request))

    if st:
        st["violations"] += 1
        status_note = "SAU KHI NỘP BÀI" if st.get("submitted") == "Đã nộp" else ("LÚC CHỜ" if st.get("submitted") == "Chờ" else "TRONG LÚC THI")
        print(f"-> [CẢNH BÁO VI PHẠM - {status_note}] Thí sinh: {st['name']} (Phòng: {room_id}) vi phạm lần thứ: {st['violations']}")

        if st.get("submitted") == "Đã nộp":
            save_individual_submission(room_id, st, st.get("answers", {}), st.get("score", 0.0))
            excel_dirty_rooms.add(room_id)

        emit('update_teacher_list', list(room_data["students"].values()), to=room_id)
    else:
        print(f"-> [!] Nhận vi phạm nhưng không tìm thấy thí sinh (SID: {sid}, Tên: {student_name})")


@socketio.on('submit_exam')
def handle_submit(data):
    sid = request.sid
    room_id = data.get('room', 'default') if isinstance(data, dict) else 'default'
    client_ip = get_client_ip(request)
    if isinstance(data, dict):
        answers = data.get('answers', {})
        name = data.get('name')
        exam_code = data.get('exam_code')
        process_grading(room_id, sid, answers, name, exam_code, client_ip=client_ip)
    else:
        process_grading(room_id, sid, {}, client_ip=client_ip)


@socketio.on('teacher_start_exam')
def teacher_start(data):
    room_id = data.get('room', 'default') if isinstance(data, dict) else 'default'
    room_data = get_or_create_room(room_id)

    available_codes = list(room_data.get("raw_banks", {}).keys())
    if not available_codes:
        emit('teacher_error', {
            'message': 'Chưa có đề thi trên máy chủ! Thầy vui lòng bấm "Nạp đề thi (PDF)" lại trước khi bắt đầu.'
        }, room=request.sid)
        return

    room_data["submitted_ips"].clear()
    clear_submission_folder(room_id)

    room_data["active_exam"]["status"] = "running"
    room_data["active_exam"]["start_time"] = time.time()
    
    duration_secs = room_data["active_exam"]["duration"] * 60

    for idx, (sid, st) in enumerate(list(room_data["students"].items())):
        assigned_code = available_codes[idx % len(available_codes)]
        st["exam_code"] = assigned_code
        if st.get("submitted") != "Đã nộp":
            st["submitted"] = "Đang thi"

        client_payload, server_exam = generate_individual_exam(room_data, assigned_code)
        st["exam"] = server_exam
        client_payload["time_left"] = duration_secs
        emit('start_exam_now', client_payload, room=sid)

    socketio.emit('exam_started_broadcast', {'room': room_id}, to=room_id)
    emit('update_teacher_list', list(room_data["students"].values()), to=room_id)


@socketio.on('teacher_end_exam')
def handle_teacher_end_exam(data):
    room_id = data.get('room', 'default') if isinstance(data, dict) else 'default'
    room_data = get_or_create_room(room_id)
    room_data["active_exam"]["status"] = "finished"
    
    for sid, st in list(room_data["students"].items()):
        if st.get("submitted") != "Đã nộp":
            st["time_left"] = 0
            process_grading(room_id, sid, st.get("answers", {}), st.get("name"), emit_socket=False)
            
    emit('force_submit_exam', {'message': 'Giáo viên đã kết thúc ca thi!'}, to=room_id)
    emit('update_teacher_list', list(room_data["students"].values()), to=room_id)


@socketio.on('reset_exam_session')
def handle_reset_session(data):
    room_id = data.get('room', 'default') if isinstance(data, dict) else 'default'
    room_data = get_or_create_room(room_id)

    room_data["submitted_ips"].clear()
    room_data["students"].clear()
    clear_submission_folder(room_id)
    room_data["active_exam"]["status"] = "waiting"
    room_data["active_exam"]["start_time"] = None
    emit('update_teacher_list', [], to=room_id)


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    socketio.run(app, host='0.0.0.0', port=port, debug=False, use_reloader=False)