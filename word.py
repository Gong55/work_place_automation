from docx import Document
from docx.text.paragraph import Paragraph
from docx.text.run import Run
import os
import pandas as pd
import re
from datetime import datetime
import numpy as np

base_path = r'E:\Desktop\python\Word'
excel_file = os.path.join(base_path, 'data_input.xlsx')
template_file = os.path.join(base_path, 'SIC Draft Contract.docx')
destination_folder = os.path.join(base_path, 'output')
owner_contract = os.path.join(base_path, 'owner_contract.docx')

df = pd.read_excel(excel_file)
df_flipped = df.set_index('Attributes').transpose()


def month_number_to_name(month):
    months = ["January", "February", "March", "April", "May", "June",
              "July", "August", "September", "October", "November", "December"]

    return months[month - 1]


def convert_date_format(column):
    if df_flipped[column].notna().all():
        df_flipped[column] = df_flipped[column].replace(' ', np.nan)
        df_flipped[column] = pd.to_datetime(
            df_flipped[column], errors='coerce')
        df_flipped[f'{column}_day'] = df_flipped[column].dt.day
        df_flipped[f'{column}_month'] = df_flipped[column].dt.month
        df_flipped[f'{column}_year'] = df_flipped[column].dt.year
        df_flipped[f'{column}_month_en'] = df_flipped[f'{column}_month'].apply(
            month_number_to_name)
    else:
        pass


convert_date_format("owner_passport_expire_date")
convert_date_format("owner_passport_expire_date_2")


def format_room(room):
    if re.fullmatch(r"\d+/\d+", room):
        room = re.sub(r"(\d+)/(\d+)", r"\1-\2", room)
        room = f"({room})"
    else:
        room = re.sub(r"(\d+)/(\d+)", r"\1-\2", room)
    return room


def replace_text_with_format(paragraph, old_text, new_text):
    for run in paragraph.runs:
        if old_text in run.text:
            original_font = run.font
            original_text = run.text
        if not pd.isna(new_text):
            if old_text.isupper():
                new_text = new_text.upper()
            elif old_text.islower():
                new_text = new_text.lower()
            elif old_text.istitle():
                new_text = new_text.title()
        else:
            new_text = ' '

        run.text = original_text.replace(old_text, str(new_text))

        run.font.bold = original_font.bold
        run.font.italic = original_font.italic
        run.font.underline = original_font.underline
        run.font.size = original_font.size
        run.font.name = original_font.name
        if original_font.color is not None:
            run.font.color.rgb = original_font.color.rgb
        run.font.all_caps = original_font.all_caps


def replace_text_in_tables(table, old_text, new_text):
    for row in table.rows:
        for cell in row.cells:
            for paragraph in cell.paragraphs:
                replace_text_with_format(paragraph, old_text, new_text)


# for index, row in df_flipped.iterrows():
#     doc = Document(template_file)

#     for section in doc.sections:
#         footer = section.footer
#         for footer_text in footer.paragraphs:
#             replace_text_with_format(
#                 footer_text, "-owner_name-", str(row['owner_name']))
#             replace_text_with_format(
#                 footer_text, "owner_name_2", str(row['owner_name_2']))
#             replace_text_with_format(
#                 footer_text, "tenant_name", str(row['tenant_name']))
#             replace_text_with_format(
#                 footer_text, "tenant_name_2", str(row['tenant_name_2']))
#             replace_text_with_format(
#                 footer_text, "witness_name", str(row['witness_name']))

#     for paragraph in doc.paragraphs:

#         replace_text_with_format(
#             paragraph, "project_name", str(paragraph['project_name']))
#         replace_text_with_format(
#             paragraph, "-room_number-", str(row['room_number']))
#         replace_text_with_format(
#             paragraph, "floor_number", str(row['floor_number']))
#         replace_text_with_format(paragraph, "area", str(row['area']))
#         replace_text_with_format(
#             paragraph, "building_no", str(row['building_no']))
#         replace_text_with_format(paragraph, "address", str(row['address']))
#         # address TH
#         # start_date
#         # end_date
#         # start_date TH
#         # end_date TH
#         replace_text_with_format(
#             paragraph, "rent_price", str(row['rent_price']))
#         # rent_price EN
#         # rent_price TH
#         # rent_price x2
#         # rent_price x2 EN
#         # rent_price x2 TH

#         replace_text_with_format(paragraph, "late_payment_grace_period", str(
#             row['late_payment_grace_period']))
#         #
#         replace_text_with_format(
#             paragraph, "-owner_name-", str(row['owner_name']))
#         replace_text_with_format(
#             paragraph, "owner_passport", str(['owner_passport']))
#         replace_text_with_format(
#             paragraph, "owner_nationality", str(['owner_nationality']))
#         replace_text_with_format(paragraph, "owner_bank", str(['owner_bank']))
#         replace_text_with_format(
#             paragraph, "owner_bank_branch", str('owner_bank_branch'))
#         replace_text_with_format(
#             paragraph, "owner_bank_account_no", str(['owner_bank_account_no']))
#         replace_text_with_format(
#             paragraph, "owner_bank_account_name", str(['owner_bank_account_name']))

#         # owner_passport_expire_date
#         # owner_passport_expire_date TH

#     formatted_room = format_room(str(row['room_number']))
#     new_file_name = os.path.join(
#         destination_folder, f"Lease Agreement {row['project_name']} {formatted_room}.docx")
#     doc.save(new_file_name)


# for index, row in df_flipped.iterrows():
#     doc2 = Document(owner_contract)
#     for table in doc2.tables:
#         replace_text_in_tables(table, "-NAME1-", str(row['owner_name']))
#         replace_text_in_tables(table, "-NANE2-", str(row['owner_name_2']))
#         replace_text_in_tables(table, "-PASPRT1-", str(row['owner_passport']))
#         replace_text_in_tables(
#             table, "-PASPRT1-", str(row['pwner_passport_2']))

#     for paragraph2 in doc2.paragraphs:
#         replace_text_with_format(
#             paragraph2, "-UNITNAME-", str(row['project_name']))
#         replace_text_with_format(
#             paragraph2, "-UNITNUMBER-", str(row['room_number']))
#         # replace_text_with_format(paragraph2, "-SY-", )
#         # replace_text_with_format(paragraph2, "-SM-", )
#         # replace_text_with_format(paragraph2, "-SD-", )
#         # replace_text_with_format(paragraph2, "-ENY-", )
#         # replace_text_with_format(paragraph2, "-ENM-", )
#         # replace_text_with_format(paragraph2, "-EMD-", )
#         # Have to convert excel date into DD MM YY

#     formatted_room = format_room(str(row['room_number']))
#     new_file_name = os.path.join(
#         destination_folder, f"代租管合約 {row['project_name']} {formatted_room}.docx")
#     doc2.save(new_file_name)
#     print(f"'output_{formatted_room}.docx' created successfully")
