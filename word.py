from docx import Document
from docx.text.paragraph import Paragraph
from docx.text.run import Run
import os
import pandas as pd
import re
from datetime import datetime
import numpy as np
from dateutil.relativedelta import relativedelta
from pythainlp.util import bahttext
from num2words import num2words


base_path = r'E:\Desktop\python\Word'
excel_file = os.path.join(base_path, 'data_input.xlsx')
lease_agreement = os.path.join(base_path, 'SIC Draft Contract.docx')
destination_folder = os.path.join(base_path, 'output')
owner_contract = os.path.join(base_path, 'owner_contract.docx')

df = pd.read_excel(excel_file)
df_flipped = df.set_index('Attributes').transpose()

thai_month_names = {
    1: 'มกราคม',
    2: 'กุมภาพันธ์',
    3: 'มีนาคม',
    4: 'เมษายน',
    5: 'พฤษภาคม',
    6: 'มิถุนายน',
    7: 'กรกฎาคม',
    8: 'สิงหาคม',
    9: 'กันยายน',
    10: 'ตุลาคม',
    11: 'พฤศจิกายน',
    12: 'ธันวาคม'
}


def number_to_text_en(df, column):
    df[column +
        '_text_en'] = df[column].apply(lambda x: num2words(x) + ' baht only')


def calculate_two_months_deposit(df, column):
    df[column + '_times_two'] = df[column].apply(lambda x: x * 2)


def number_to_text_th(df, column):
    df[column + '_text_th'] = df[column].apply(bahttext)


def lease_period(start_date, end_date, df):

    delta_list = []
    for start, end in zip(df[start_date], df[end_date]):
        start = start + pd.DateOffset(days=1)
        delta = relativedelta(end, start)
        delta_list.append((delta.days, delta.months, delta.years))

    return delta_list


def late_payment_grace_period(start_date, df):

    answer = []

    for start in df[start_date]:
        output = start + relativedelta(days=4)
        answer.append(output.days, output.months, output.years)
    return answer


def convert_date_format(column, df):
    df[column] = df[column].replace(' ', np.nan)
    df[column] = pd.to_datetime(
        df[column], errors='coerce')
    if df[column].notna().all():
        df[f'{column}_day'] = df[column].dt.day
        df[f'{column}_month'] = df[column].dt.month
        df[f'{column}_year'] = df[column].dt.year
        df[f'{column}_month_en'] = df[column].dt.month_name()
        df[f'{column}_month_th'] = df[f'{column}_month'].map(
            thai_month_names)
        df[column + '_year_th'] = df[column + '_year'].apply(lambda x: x + 543)
    else:
        pass


def format_room(room):
    if re.fullmatch(r"\d+/\d+", room):
        room = re.sub(r"(\d+)/(\d+)", r"\1-\2", room)
        room = f"({room})"
    else:
        room = re.sub(r"(\d+)/(\d+)", r"\1-\2", room)
    return room


convert_date_format("rent_start_date", df_flipped)
convert_date_format("rent_end_date", df_flipped)
convert_date_format("owner_passport_expire_date", df_flipped)
convert_date_format("tenant_passport_expire_date", df_flipped)


def replace_text_with_format(paragraph, old_text, new_text):
    # Concatenate all runs' text to check if `old_text` exists in the paragraph
    full_text = ''.join([run.text for run in paragraph.runs])

    # If the old_text exists, proceed with replacement
    if old_text in full_text:
        print(f"Found '{old_text}' in paragraph: {full_text}")

        # Handle case transformations based on old text format
        if old_text.isupper():
            new_text = new_text.upper()
        elif old_text.islower():
            new_text = new_text.lower()
        elif old_text.istitle():
            new_text = new_text.title()

        # Replace the placeholder in the full text
        new_full_text = full_text.replace(old_text, new_text)

        # Clear all existing runs in the paragraph (to avoid splitting issues)
        for run in paragraph.runs:
            run.clear()

        # Create a new run with the updated full text
        new_run = paragraph.add_run(new_full_text)

        # Apply the font formatting (using the first run's font as an example)
        original_font = paragraph.runs[0].font if paragraph.runs else None
        if original_font:
            new_run.font.bold = original_font.bold
            new_run.font.italic = original_font.italic
            new_run.font.underline = original_font.underline
            new_run.font.size = original_font.size
            new_run.font.name = original_font.name
            if original_font.color is not None:
                new_run.font.color.rgb = original_font.color.rgb
            new_run.font.all_caps = original_font.all_caps


def replace_text_in_tables(table, old_text, new_text):
    # Loop through all rows and cells in the table
    for table_row in table.rows:
        for cell_row in table_row.cells:
            if cell_row.paragraphs:
                for cell_paragraph in cell_row.paragraphs:
                    # Debugging: Print the content of each cell before replacement
                    print(
                        f"Cell content before replacement: {cell_paragraph.text}")

                    replace_text_with_format(
                        cell_paragraph, old_text, new_text)


def replace_text_in_document(doc, old_text, new_text):
    # Replace text in paragraphs outside of tables
    for paragraph in doc.paragraphs:
        print(f"Paragraph content before replacement: {paragraph.text}")
        replace_text_with_format(paragraph, old_text, new_text)


# Main Loop
for index1, row1 in df_flipped.iterrows():
    doc = Document(owner_contract)

    # Replace text in table cells
    for table in doc.tables:
        replace_text_in_tables(table, 'NAME1HOLDER', str(
            row1['owner_name']) if pd.notna(row1['owner_name']) else '')
        replace_text_in_tables(table, 'NAME2HOLDER', str(
            row1['owner_name_2']) if pd.notna(row1['owner_name_2']) else '')
        replace_text_in_tables(table, 'PSPT1HO', str(
            row1['owner_passport']) if pd.notna(row1['owner_passport']) else '')
        replace_text_in_tables(table, '{{PSPT2}}', str(
            row1['owner_passport_2']) if pd.notna(row1['owner_passport_2']) else '')

    # Replace text in paragraphs outside of tables
    replace_text_in_document(doc, 'PROJECTNAMEHOLDER', str(
        row1['project_name']) if pd.notna(row1['project_name']) else '')
    replace_text_in_document(doc, 'UNITNUMBERHOLDER', str(
        row1['room_number']) if pd.notna(row1['room_number']) else '')

    # Debugging: Print document save confirmation
    print(f"Document for row {index1} processed.")

    # Debugging: Print document save confirmation
    print(f"Document for row {index1} processed.")

    formatted_room = format_room(str(row1['room_number']))

    new_file_name = os.path.join(
        destination_folder, f"代租管合約 {row1['project_name']} {formatted_room}.docx")
    doc.save(new_file_name)
    print(f"'output_{formatted_room}.docx' created successfully")

    # for index, row in df_flipped.iterrows():
    #     doc = Document(template_file)

    #     for section in doc.sections:
    #         footer = section.footer
    #         for footer_text in footer.paragraphs:
    #             replace_text_with_format(
    #                 footer_text, "owner_name", str(row['owner_name']))
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
