import os
import pandas as pd
import openpyxl
from xlsxwriter import Workbook

folder_path = r'.\image_label'

for filename in os.listdir(folder_path):
    # Kiểm tra nếu tệp là ảnh (có đuôi .png)
    if filename.endswith('.png') and filename.startswith('cropped_page_'):
        # Tách lấy số trang từ tên tệp
        page_number = filename.replace('cropped_page_', '').replace('.png', '').zfill(3)
        
        # Tạo tên mới theo format yêu cầu
        new_filename = f'Sac_phong_trieu_Nguyen_tren_dia_ban_Thua_Thien_Hue_page{page_number}.png'
        
        # Đường dẫn cũ và mới
        old_path = os.path.join(folder_path, filename)
        new_path = os.path.join(folder_path, new_filename)
        
        # Đổi tên tệp
        os.rename(old_path, new_path)

print("Đổi tên thành công!")
df = pd.read_excel('result.xlsx')
list_quocngu = df['Chữ quốc ngữ'].str.lower().str.replace(',', '').str.replace(':', '').tolist()
list_ocr = list(df['SinoNom OCR'])


""
path_similar = r"SinoNom_similar_Dic.xlsx"
path_QuocNgu_SinoNom = r"QuocNgu_SinoNom_Dic.xlsx"
DICT_1 = pd.read_excel(path_QuocNgu_SinoNom)
DICT_2 = pd.read_excel(path_similar)

def compare(word, OCR): 
    result_word = DICT_1[DICT_1['QuocNgu'] == word]['SinoNom'].values
    result_OCR = DICT_2[DICT_2['Input Character'] == OCR]['Top 20 Similar Characters'].values
    
    result_word_set = set(result_word)
    result_OCR_set = set(result_OCR)
    
    if OCR in result_word_set:
        return {OCR}
    
    return result_word_set & result_OCR_set

def edit_distance(a, b):
    m, n = len(a), len(b)
    dp = [[0] * (n + 1) for _ in range(m + 1)]

    for i in range(m + 1):
        dp[i][0] = i
    for j in range(n + 1):
        dp[0][j] = j

    for i in range(1, m + 1):
        for j in range(1, n + 1):
            common_chars = compare(a[i - 1], b[j - 1])
            if common_chars:
                dp[i][j] = dp[i - 1][j - 1]
            else:
                dp[i][j] = min(
                    dp[i - 1][j] + 1,  # Deletion
                    dp[i][j - 1] + 1,  # Insertion
                    dp[i - 1][j - 1] + 1  # Substitution
                )

    distance = dp[m][n]
    a_highlighted, b_highlighted = highlight_diff(a, b, dp)

    return distance, a_highlighted, b_highlighted

def highlight_diff(a, b, dp):
    m, n = len(a), len(b)
    i, j = m, n
    a_highlighted, b_highlighted = [], []

    while i > 0 and j > 0:
        common_chars = compare(a[i - 1], b[j - 1])
        if common_chars:
            a_highlighted.append(a[i - 1])
            b_highlighted.append(b[j - 1])
            i -= 1
            j -= 1
        elif dp[i][j] == dp[i - 1][j] + 1:
            a_highlighted.append(a[i - 1])
            b_highlighted.append('*')
            i -= 1
        elif dp[i][j] == dp[i][j - 1] + 1:
            a_highlighted.append('*')
            b_highlighted.append(b[j - 1])
            j -= 1
        else:
            a_highlighted.append(a[i - 1])
            b_highlighted.append(b[j - 1])
            i -= 1
            j -= 1

    while i > 0:
        a_highlighted.append(a[i - 1])
        b_highlighted.append('*')
        i -= 1

    while j > 0:
        a_highlighted.append('*')
        b_highlighted.append(b[j - 1])
        j -= 1

    return a_highlighted[::-1], b_highlighted[::-1]




from xlsxwriter import Workbook # type: ignore
workbook = Workbook('result1.xlsx')
worksheet = workbook.add_worksheet()

red = workbook.add_format({'color': 'red'})      
blue = workbook.add_format({'color': 'blue'})  
black = workbook.add_format({'color': 'black'})  

column_widths = {
    'A': 20,
    'B' :15,
    'C': 50,  
    'D': 55,  
    'E': 100  
}

for col_letter, width in column_widths.items():
    worksheet.set_column(f'{col_letter}:{col_letter}', width)
    worksheet.write(0, 0, 'Image_name')          
    worksheet.write(0, 1, 'ID')        
    worksheet.write(0, 2, 'Image Box')            
    worksheet.write(0, 3, 'SinoNom OCR')  
    worksheet.write(0, 4, 'Chữ quốc ngữ')    

for row_num, (word, ocr) in enumerate(zip(list_quocngu, list_ocr)):
    print('line: ', row_num)
    a = word.split()
    b = list(ocr)
    distance, b_new, a_new = edit_distance(a, b)
    max_len = len(b_new)
    if max_len ==1:
        max_len +=1
        b_new.append(' ')
        a_new.append(' ')
    temp = []
    for i in range(max_len):
        if len(compare(b_new[i], a_new[i])) >1:
                temp+=(blue, a_new[i])
        elif len(compare(b_new[i], a_new[i])) == 1:
                temp+=(black, a_new[i])
        elif len(compare(b_new[i], a_new[i])) == 0:
            if a_new[i] == '*' and b_new[i] != '*':
                temp+=(red," ")
            if b_new[i] == '*' and a_new[i] != '*':
                temp+=(red, a_new[i])
            else:
                temp+=(red, a_new[i])
        

    if len(a_new) == 1 and a_new[i] != '*':  # Ensure it's not a wildcard
        temp.append(black)  # Use black for single-character
        temp.append(a_new[i])  # Add the single character

            
    print(temp)
    worksheet.write(row_num + 1, 0, df['Image_name'][row_num])
    worksheet.write(row_num + 1, 1, df['ID'][row_num])
    worksheet.write_rich_string(row_num + 1, 3, *temp)  
    worksheet.write(row_num + 1, 2, df['Image Box'][row_num])
    worksheet.write(row_num + 1, 4, df['Chữ quốc ngữ'][row_num])


workbook.close()