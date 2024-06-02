from pdfrw import PdfReader, PdfWriter
import pygame
import openpyxl
import tkinter as tk
from tkinter import filedialog
from PyPDF2 import PdfWriter, PdfReader
import io
import sys
import os
import re
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.pagesizes import letter

OPTIONS = [
    "OZON",
    "WILDBERRIES",
    "CDEK",
    "YANDEX"
]

X_WB_POSITION = 6
Y_WB_POSITION = 4

X_OZON_POSITION = 15
Y_OZON_POSITION = 85


def resource_path(relative):
    if hasattr(sys, "_MEIPASS"):
        return os.path.join(sys._MEIPASS, relative)
    return os.path.join(relative)


def browse_pdf_file():
    pdf_file_path.set(filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")]))
    output_pdf_file_path.set(pdf_file_path.get().split('/')[-1])


def browse_excel_file():
    excel_file_path.set(filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")]))
    output_excel_file_path.set(excel_file_path.get().split('/')[-1])


def browse_output_folder():
    output_folder = filedialog.askdirectory()
    output_folder_path.set(output_folder)


def start_processing():
    if variable.get() == OPTIONS[0]:
        start_processing_ozon()
    elif variable.get() == OPTIONS[1]:
        start_processing_wb()
    elif variable.get() == OPTIONS[2]:
        start_processing_cdek()
    else:
        start_processing_yandex()
    result.set("Calculation complete. Check the file!!!")


def start_processing_cdek():
    pdf_path = pdf_file_path.get()
    output_stream = open(f'{output_folder_path.get()}/edited_cdek.pdf', "wb")
    excel_path = excel_file_path.get()
    pdf_document = PdfReader(open(pdf_path, "rb"))
    excel_workbook = openpyxl.load_workbook(excel_path)
    excel_sheet = excel_workbook.active
    output = PdfWriter()
    for pdf_page_number, pdf_page in enumerate(pdf_document.pages):
        pdf_all_text = pdf_page.extract_text()
        pdf_text = re.search(r'\d+-\d+-\d+', pdf_all_text)[0]
        matching_values = []

        for row in excel_sheet.iter_rows(min_row=2, min_col=2, max_col=12, values_only=True):
            if any(cell is not None for cell in row):
                excel_value_to_match = row[0] if row[0] is not None else ""
                excel_value_from_col10 = row[9] if row[9] is not None else ""
                excel_value_from_col12 = row[10] if row[10] is not None else ""
                excel_value_from_col12_str = str(excel_value_from_col12)
                excel_value_from_col10_str = str(excel_value_from_col10)
                if excel_value_to_match == pdf_text:
                    matching_values.append((excel_value_from_col10_str, excel_value_from_col12_str))
        if matching_values:
            packet = io.BytesIO()
            can = canvas.Canvas(packet, pagesize=letter)
            can.setFont('FreeSans', int(entry_size.get()))
            combined_values = []
            for value_10, value_12 in matching_values:
                combined_values.append(f"{value_10}, {value_12}")
            combined_values_text = " ; ".join(combined_values)
            can.drawString(int(entry_x.get()), int(entry_y.get()), combined_values_text)
            can.save()
            new_pdf = PdfReader(packet)
            pdf_page.merge_page(new_pdf.pages[0])
            output.add_page(pdf_page)
    output.write(output_stream)
    output_stream.close()


def start_processing_ozon():
    pdf_path = pdf_file_path.get()
    output_stream = open(f'{output_folder_path.get()}/edited_ozon.pdf', "wb")
    excel_path = excel_file_path.get()
    pdf_document = PdfReader(open(pdf_path, "rb"))
    excel_workbook = openpyxl.load_workbook(excel_path)
    excel_sheet = excel_workbook.active
    output = PdfWriter()
    for pdf_page_number, pdf_page in enumerate(pdf_document.pages):
        pdf_text = pdf_page.extract_text().replace('\n', '')
        matching_values = []

        for row in excel_sheet.iter_rows(min_row=2, min_col=2, max_col=12, values_only=True):
            if any(cell is not None for cell in row):
                excel_value_to_match = row[0] if row[0] is not None else ""
                excel_value_from_col10 = row[9] if row[9] is not None else ""
                excel_value_from_col12 = row[10] if row[10] is not None else ""
                excel_value_from_col12_str = str(excel_value_from_col12)
                excel_value_from_col10_str = str(excel_value_from_col10)
                if excel_value_to_match == pdf_text.split()[3]:
                    matching_values.append((excel_value_from_col10_str, excel_value_from_col12_str))
        if matching_values:
            packet = io.BytesIO()
            can = canvas.Canvas(packet, pagesize=letter)
            can.setFont('FreeSans', int(entry_size.get()))
            combined_values = []
            for value_10, value_12 in matching_values:
                combined_values.append(f"{value_10}, {value_12}")
            combined_values_text = " ; ".join(combined_values)
            can.drawString(int(entry_x.get()), int(entry_y.get()), combined_values_text)
            can.save()
            new_pdf = PdfReader(packet)
            pdf_page.merge_page(new_pdf.pages[0])
            output.add_page(pdf_page)
    output.write(output_stream)
    output_stream.close()


def start_processing_yandex():
    pdf_path = pdf_file_path.get()
    output_stream = open(f'{output_folder_path.get()}/edited_yandex.pdf', "wb")
    excel_path = excel_file_path.get()
    pdf_document = PdfReader(open(pdf_path, "rb"))
    excel_workbook = openpyxl.load_workbook(excel_path)
    excel_sheet = excel_workbook.active
    output = PdfWriter()
    for pdf_page_number, pdf_page in enumerate(pdf_document.pages):
        pdf_text = pdf_page.extract_text().replace('\n', '').split()[0]
        matching_values = []

        for row in excel_sheet.iter_rows(min_row=2, min_col=1, max_col=3, values_only=True):
            if any(cell is not None for cell in row):
                excel_value_to_match = str(row[0]) if row[0] is not None else ""
                name = str(row[1]) if row[1] is not None else ""
                quantity = str(row[2]) if row[2] is not None else ""
                if excel_value_to_match == pdf_text:
                    matching_values.append((name, quantity))
        if matching_values:
            packet = io.BytesIO()
            can = canvas.Canvas(packet, pagesize=letter)
            can.setFont('FreeSans', int(entry_size.get()))
            combined_values = []
            for value_10, value_12 in matching_values:
                combined_values.append(f"{value_10}, {value_12}")
            combined_values_text = " ; ".join(combined_values)
            can.rotate(90)
            can.drawString(int(entry_x.get()), -int(entry_y.get()), combined_values_text)
            can.save()
            new_pdf = PdfReader(packet)
            pdf_page.merge_page(new_pdf.pages[0])
            output.add_page(pdf_page)
    output.write(output_stream)
    output_stream.close()


def start_processing_wb():
    pdf_path = pdf_file_path.get()
    output_stream = open(f'{output_folder_path.get()}/edited_wb.pdf', "wb")
    excel_path = excel_file_path.get()
    pdf_document = PdfReader(open(pdf_path, "rb"))
    excel_workbook = openpyxl.load_workbook(excel_path)
    excel_sheet = excel_workbook.active
    output = PdfWriter()
    for pdf_page_number, pdf_page in enumerate(pdf_document.pages):
        pdf_text = pdf_page.extract_text()
        pdf_text = pdf_text.replace("\n", "")
        pdf_text = re.sub(r'[a-zA-Zа-яА-Я]', '', pdf_text)
        matching_value = ""
        for row in excel_sheet.iter_rows(min_row=2, min_col=1, max_col=2, values_only=True):
            if any(cell is not None for cell in row):
                excel_value_to_match = row[1].replace(" ", '') if row[1] is not None else ""
                excel_value_from_col0 = row[0] if row[0] is not None else ""
                excel_value_from_col0_str = str(excel_value_from_col0)

                if excel_value_to_match == pdf_text:
                    matching_value = excel_value_from_col0_str
                    break
        if len(matching_value) != 0:
            packet = io.BytesIO()
            can = canvas.Canvas(packet, pagesize=letter)
            can.setFont('FreeSans', int(entry_size.get()))
            can.rotate(90)
            can.drawString(int(entry_x.get()), -int(entry_y.get()), matching_value)
            can.save()
            new_pdf = PdfReader(packet)
            pdf_page.merge_page(new_pdf.pages[0])
            output.add_page(pdf_page)
    output.write(output_stream)
    output_stream.close()


pdfmetrics.registerFont(TTFont('FreeSans', 'FreeSans.ttf', 'CP1251'))

root = tk.Tk()
root.title("Добавление значений на PDF")
root.geometry("640x480")  # Установка размера окна

result = tk.StringVar()
pdf_file_path = tk.StringVar()
excel_file_path = tk.StringVar()
output_pdf_file_path = tk.StringVar()
output_excel_file_path = tk.StringVar()
output_folder_path = tk.StringVar()  # Добавляем переменную для хранения пути к папке

option_label = tk.Label(root, text="Выберите сервис")
option_label.pack()

variable = tk.StringVar(root)
variable.set(OPTIONS[0])
w = tk.OptionMenu(root, variable, *OPTIONS)
w.pack()

pdf_label = tk.Label(root, text="Выберите PDF файл:")
pdf_label.pack()

pdf_button = tk.Button(root, text="Обзор", command=browse_pdf_file)
pdf_button.pack()

pdf_folder_label = tk.Label(root, textvariable=output_pdf_file_path)
pdf_folder_label.pack()

excel_label = tk.Label(root, text="Выберите Excel файл:")
excel_label.pack()

excel_button = tk.Button(root, text="Обзор", command=browse_excel_file)
excel_button.pack()

excel_folder_label = tk.Label(root, textvariable=output_excel_file_path)
excel_folder_label.pack()

x_label = tk.Label(root, text="X координата:")
x_label.pack()

entry_x = tk.Entry(root)
entry_x.pack()

y_label = tk.Label(root, text="Y координата:")
y_label.pack()

entry_y = tk.Entry(root)
entry_y.pack()

size_label = tk.Label(root, text="Размер шрифта: ")
size_label.pack()

entry_size = tk.Entry(root)
entry_size.pack()

pygame.init()
font = resource_path(os.path.join('Folder', 'FreeSans.ttf'))

output_folder_button = tk.Button(root, text="Выбрать папку для сохранения", command=browse_output_folder)
output_folder_button.pack()

output_folder_label = tk.Label(root, textvariable=output_folder_path)
output_folder_label.pack()

process_button = tk.Button(root, text="Начать", command=start_processing)
process_button.pack()

result_label = tk.Label(root, textvariable=result)
result_label.pack()

root.mainloop()
