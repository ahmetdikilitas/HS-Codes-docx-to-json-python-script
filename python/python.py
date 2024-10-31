import json
import re
from docx import Document

def convert_word_to_json(docx_file, json_file):
    doc = Document(docx_file)
    content = []
    
    for paragraph in doc.paragraphs:
        content.append(paragraph.text.strip())

    hs_codes = {}
    current_code = None

    for line in content:
        if re.match(r'^\d{2,7}$', line):  # Sadece sayı içeren satır
            current_code = line  # HS kodunu güncelle
            hs_codes[current_code] = ""  # Açıklama başlangıcı
        elif current_code and line:  # HS kodu mevcutsa ve açıklama satırı boş değilse
            if hs_codes[current_code]:  # Eğer açıklama daha önce tanımlandıysa
                hs_codes[current_code] += " " + line  # Açıklamaya ekle
            else:
                hs_codes[current_code] = line  # İlk açıklama ataması

    # HS kodlarını listeye çevir
    hs_codes_list = [{"code": code, "description": description.strip()} for code, description in hs_codes.items()]

    with open(json_file, 'w', encoding='utf-8') as json_file:
        json.dump(hs_codes_list, json_file, ensure_ascii=False, indent=4)

# Kullanım
convert_word_to_json('HS Code.docx', 'hs_codes.json')
