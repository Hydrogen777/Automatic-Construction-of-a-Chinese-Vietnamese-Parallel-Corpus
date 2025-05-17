import os
import fitz  # PyMuPDF
import easyocr
import pandas as pd
import re

reader = easyocr.Reader(['vi'], gpu=True)

pdf_path = "DuongHoaiMinh_Sắc-phong-triều-Nguyễn-trên-địa-bàn-Thừa-Thiên-Huế.pdf"
output_folder = "output_images"
zoom_factor = 2.75
start_page = 26
end_page = 456


def pdf_page_to_image(pdf_document, output_folder, page_number, zoom_factor, y0=None, y1=None):
    
    if page_number < 0 or page_number >= pdf_document.page_count:
        print(f"Số trang không hợp lệ. PDF chỉ có {pdf_document.page_count} trang.")
        return None

    pdf_page = pdf_document.load_page(page_number)
    width, height = pdf_page.rect.width, pdf_page.rect.height
    matrix = fitz.Matrix(zoom_factor, zoom_factor)
    
    
    if y0 is not None and y1 is not None:
        y0 /= zoom_factor
        y1 /= zoom_factor
        rect = fitz.Rect(0, y0, width, y1)
    elif y0 is not None:
        y0 /= zoom_factor
        rect = fitz.Rect(0, y0, width, height - 30)
    elif y1 is not None:
        y1 /= zoom_factor
        if y1 - 500 < 0:
            return None
        rect = fitz.Rect(0, y1 - 400, width, y1)
    else:
        rect = None

    
    
    
    pix = pdf_page.get_pixmap(matrix=matrix, clip=rect)
    output_file = os.path.join(output_folder, f"{'cropped_' if rect else 'full_'}page_{page_number}.png")
    pix.save(output_file)
    print(f'Saved {output_file}')
    return output_file

def is_vietnamese(text):
    vietnamese_pattern = r'^[a-zA-Z0-9àáảãạâầấẩẫậăằắẳẵặđèéẻẽẹêềếểễệìíỉĩịòóỏõọôồốổỗộơờớởỡợùúủũụưừứửữựỳýỷỹỵÀÁẢÃẠÂẦẤẨẪẬĂẰẮẲẴẶĐÈÉẺẼẸÊỀẾỂỄỆÌÍỈĨỊÒÓỎÕỌÔỒỐỔỖỘƠỜỚỞỠỢÙÚỦŨỤƯỪỨỬỮỰỲÝỶỸỴ.,!?-_ :;]+$'
    return bool(re.fullmatch(vietnamese_pattern, text))

def Is_y0(s):
    """Điều kiện để xác định loại y0."""
    return re.match(r'^(a\.|g\.|a,|g,|\d+\.$|^\d+\s\.$)$', s, re.IGNORECASE) or s in {"a", "g"}

def Is_y1(s):
    """Điều kiện để xác định loại y1."""
    return re.match(r'^(b\.)$', s, re.IGNORECASE) or s.startswith("b.") or s.startswith("b :") or s.startswith("b:") or s.startswith("b .") or s in {"b ", 
"b_", "b,", "b Phiên âm:", "b. Phiên âm:", "b.Phiên âm:", "b_Phiên âm:", "b Phiên âm", "b. Phiên âm", "b.Phiên âm", "b_Phiên âm", 
"b Phien am:", "b. Phien am:", "b.Phien am:", "b_Phien am:", "b Phien am", "b. Phien am", "b.Phien am", "b_Phien am", 
"Phiên âm:", "Phiên âm:", " Phiên âm:", "Phiên âm:", "Phiên âm", " Phiên âm", "Phiên âm", "Phiên âm", 
"Phien am:", " Phien am:", "Phien am:", "Phien am:", "Phien am", " Phien am", "Phien am", "Phien am"}

def Is_y2(s):
    """Điều kiện để xác định loại y1."""
    return re.match(r'^(c\.)$', s, re.IGNORECASE) or s.startswith("c.") or s.startswith("c .") or s.startswith("c :") or s.startswith("c:") or s in {"c ", "c_", "c,", "c Dịch nghĩa:", "c. Dịch nghĩa:", "c.Dịch nghĩa:", "c_Dịch nghĩa:", "c Dịch nghĩa", "c. Dịch nghĩa", "c.Dịch nghĩa", "c_Dịch nghĩa", 
"c Phien nghia:", "c. Phien nghia:", "c.Phien nghia:", "c_Phien nghia:", "c Phien nghia", "c. Phien nghia", "c.Phien nghia", "c_Phien nghia", 
"c Dich nghia:", "c. Dich nghia:", "c.Dich nghia:", "c_Dich nghia:", "c Dich nghia", "c. Dich nghia", "c.Dich nghia", "c_Dich nghia"}

def find_SinoNom(image_path, page_number, pdf_document=None):
    if not os.path.exists(image_path):
        print(f"Image file không tồn tại: {image_path}")
        return

    result = reader.readtext(image_path)
    filtered_result = [
    {
        "Text": re.sub(r'_', ' ', item[1]),
        "Box": [[int(point[0]), int(point[1])] for point in item[0]],
        "Type": "y0" if Is_y0(item[1]) else "y1" if Is_y1(item[1]) else "y2" if Is_y2(item[1]) else "unknown"
    }
    for item in result if Is_y0(item[1]) or Is_y1(item[1]) or Is_y2(item[1])
]

    for item in filtered_result:
        text = item["Text"]
        box = item["Box"]
        top_left = tuple(map(int, box[0]))
        bottom_right = tuple(map(int, box[2]))
        print(f"Text: {text}")
        print(f"Bounding Box (Top-left, Bottom-right): ({top_left}, {bottom_right})\n")

    y0 = y1 = None
    found_a = False
    print(filtered_result)

    for i, item in enumerate(filtered_result):
        if item["Type"] == "y0" and found_a == False:
            y0 = item["Box"][0][1] + 15
            found_a = True
        elif item["Type"] == "y1":
            y1 = item["Box"][2][1] - 20


    if y0 is None:
        print(f"Không tìm thấy ký tự 'a' hoặc từ tiếng Việt dài trước ký tự 'b' trong trang {page_number}.")

    if filtered_result:
        df_output = pd.DataFrame([item["Text"] for item in filtered_result], columns=['Filtered Text'])
        df_output.to_excel('output_filtered.xlsx', index=False)
        
        if y0 is not None and y1 is not None:
            if y0 > y1:
                y1 = None
        
        if y0 is not None or y1 is not None:
            if y1 is not None:
                if y1 < 200: return
            print("y0:")
            print(y0)
            print("y1:")
            print(y1)
            pdf_page_to_image(pdf_document, output_folder, page_number, zoom_factor, y0, y1)
        else:
            print("Không đủ thông tin để thực hiện cắt trang PDF.")
    else:
        print("Không có dữ liệu nào phù hợp với các điều kiện lọc.")




def extract_full_pdf(pdf_path=pdf_path):
    if not os.path.exists(pdf_path):
        print(f"File PDF không tồn tại: {pdf_path}")
        return
    
    pdf_document = fitz.open(pdf_path)
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    for i in range(start_page, end_page):
        image_path = os.path.join(output_folder, f"full_page_{i}.png")
        image_path1 = os.path.join(output_folder, f"cropped_page_{i}.png")
        if not os.path.exists(image_path):
            image_path = pdf_page_to_image(pdf_document, output_folder, i, zoom_factor)
        
        if image_path and not os.path.exists(image_path1):
            find_SinoNom(image_path, i, pdf_document)
    
    pdf_document.close()

extract_full_pdf()
