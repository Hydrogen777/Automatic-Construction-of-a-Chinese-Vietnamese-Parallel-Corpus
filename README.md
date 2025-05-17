# Automatic-Construction-of-a-Chinese-Vietnamese-Parallel-Corpus
Automatic Construction of a Chinese–Vietnamese Parallel Corpus
Các bước thực hiện:
1. Chạy file “extract_image.py” để tách file pdf thành ảnh hoàn chỉnh và ảnh đã crop.
2. Copy toàn bộ các ảnh cropped_page vào folder image_label
3. Sử dụng API của kandianguji và OCR chữ Hán từ ảnh đã crop. Thu được folder Json_result chứa kết quả đã OCR
4. Chạy file “Extract_file.ipynb” để OCR tiếng Việt, đọc file json kết quả OCR và tạo file xlsx
5. Chạy file “color_and_rename.py” để đổi tên cho ảnh và tô màu kết quả.
