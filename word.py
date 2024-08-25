import pandas as pd
import re
from docx import Document
import os


excel_file = r'E:\Desktop\python\Word\data.xlsx'
template_file = r'E:\Desktop\python\Word\base.docx'
destination_folder = r'E:\Desktop\python\Word\output'


df = pd.read_excel(excel_file)
df = df.fillna('')
os.makedirs(destination_folder, exist_ok=True)


def format_room(room):
    if re.fullmatch(r"\d+/\d+", room):
        room = re.sub(r"(\d+)/(\d+)", r"\1-\2", room)
        room = f"({room})"
    else:
        room = re.sub(r"(\d+)/(\d+)", r"\1-\2", room)
    return room


for index, row in df.iterrows():
    new_doc = Document(template_file)

    for paragraph in new_doc.paragraphs:
        if 'name' in paragraph.text:
            paragraph.text = paragraph.text.replace('name', str(row['name']))
        if 'date' in paragraph.text:
            paragraph.text = paragraph.text.replace('date', str(row['date']))
        if 'passport' in paragraph.text:
            paragraph.text = paragraph.text.replace(
                'passport', str(row['passport']))
        if 'room' in paragraph.text:
            paragraph.text = paragraph.text.replace('room', str(row['room']))
            formatted_room = format_room(str(row['room']))

    new_file_name = os.path.join(
        destination_folder, f"output_{formatted_room}.docx")
    new_doc.save(new_file_name)

print(f"'output_{formatted_room}.docx' created successfully")
