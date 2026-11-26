import base64
import os
import re
from bs4 import BeautifulSoup
import mammoth


def convert_image_to_base64(image):
  """Tự động trích xuất mọi ảnh (sơ đồ mạch điện Lý, đồ thị, hình vẽ Toán/Hóa/Sinh)

  từ file Word và chuyển đổi sang dạng base64 nhúng thẳng vào thẻ HTML. Giúp
  điện thoại học sinh dùng 4G xem được 100%, không bao giờ bị lỗi 404 hình ảnh.
  """
  with image.open() as image_bytes:
    encoded_src = base64.b64encode(image_bytes.read()).decode("ascii")
  return {"src": f"data:{image.content_type};base64,{encoded_src}"}


def convert_docx_to_html(docx_path):
  """Chuyển đổi file .docx thành HTML có nhúng ảnh base64 và giữ lại cấu trúc văn bản/ký hiệu."""
  with open(docx_path, "rb") as docx_file:
    result = mammoth.convert_to_html(
        docx_file,
        convert_image=mammoth.images.img_element(convert_image_to_base64),
    )
    return result.value


def parse_exam_docx(docx_path):
  """Hàm bóc tách đề thi toàn năng cho tất cả các môn:

  - Nhận diện câu hỏi theo chuẩn GDPT: Câu 1, Câu 2... hoặc [NB], [TH], [VD],
  [VDC]
  - Hỗ trợ công thức Toán, Lý, Hóa (dạng $...$ hoặc MathML)
  - Tự động gom các phương án A, B, C, D (cả chữ, công thức và hình vẽ đi kèm)
  - Nhận diện các thẻ HTML để trình duyệt render nguyên vẹn
  """
  html_content = convert_docx_to_html(docx_path)
  soup = BeautifulSoup(html_content, "html.parser")

  # Lấy toàn bộ các khối nội dung (<p>, <table>, <ul>, <ol>)
  elements = soup.find_all(["p", "table", "ul", "ol"])

  questions = []
  current_q = None

  # Biểu thức phát hiện câu hỏi: "Câu 1.", "Câu 1:", "[NB] Câu 1:", "Bài 1:"
  q_pattern = re.compile(
      r"^(?:\[(?:NB|TH|VD|VDC)\]\s*)?(?:Câu|Bài)\s*(\d+)[\.:\-\s]",
      re.IGNORECASE,
  )

  for el in elements:
    text = el.get_text().strip()
    el_html = str(el)

    # 1. Phát hiện bắt đầu câu hỏi mới
    match_q = q_pattern.match(text)
    if match_q:
      if current_q:
        questions.append(current_q)
      q_num = match_q.group(1)

      # Giữ nguyên cấu trúc HTML bên trong (chứa ảnh + công thức) và bỏ chữ "Câu X:"
      cleaned_content = re.sub(
          r"^(<p[^>]*>)?\s*(?:\[(?:NB|TH|VD|VDC)\]\s*)?(?:Câu|Bài)\s*\d+[\.:\-\s]*",
          r"\1",
          el_html,
          flags=re.IGNORECASE,
      )

      current_q = {
          "id": int(q_num) if q_num.isdigit() else len(questions) + 1,
          "content": cleaned_content,
          "options": {},
          "correct_answer": None,
          "type": "multiple_choice",
      }
      continue

    # 2. Xử lý các phương án A, B, C, D và hình minh họa đi kèm
    if current_q is not None:
      found_option = False
      for opt_key in ["A", "B", "C", "D"]:
        opt_regex = re.compile(
            rf"(?:^|\s)(<strong>\s*)?{opt_key}[\.:\)]\s*(\<\/strong\>)?",
            re.IGNORECASE,
        )
        if opt_regex.search(el_html):
          found_option = True
          break

      if found_option:
        current_q["options_raw"] = current_q.get("options_raw", "") + el_html
      else:
        # Nếu không chứa A, B, C, D thì đây là hình ảnh, sơ đồ mạch điện hoặc bảng số liệu của câu hỏi
        current_q["content"] += el_html

  if current_q:
    questions.append(current_q)

  # 3. Chuẩn hóa các phương án A, B, C, D
  for q in questions:
    raw_opt_html = q.pop("options_raw", "")
    if raw_opt_html:
      splits = re.split(
          r"(?:<[^>]+>|\s)*\b([A-D])[\.:\)](?:<[^>]+>|\s)*", raw_opt_html
      )
      for i in range(1, len(splits), 2):
        key = splits[i].upper()
        val = splits[i + 1] if i + 1 < len(splits) else ""
        val_soup = BeautifulSoup(val, "html.parser")
        q["options"][key] = str(val_soup).strip()

    for k in ["A", "B", "C", "D"]:
      if k not in q["options"]:
        q["options"][k] = ""

  return questions
