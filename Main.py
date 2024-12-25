from jinja2 import Environment, FileSystemLoader
from weasyprint import HTML
import os
import pandas as pd
from PyPDF2 import PdfReader, PdfWriter
from pypdf import PdfMerger
import glob
from datetime import datetime
import pikepdf
from collections import OrderedDict
pd.options.mode.chained_assignment = None
pd.options.display.max_colwidth = 10000


def cover_page(file_path):
    project_folder = r'E:\Tejaswini - Tool Desing\Report generation'
    os.chdir(project_folder)
    env = Environment(loader=FileSystemLoader('.'))
    template = env.get_template("first_page.html")
    df = pd.read_excel(file_path)
    for index, row in df.iterrows():
        current_date = datetime.now().strftime("%d-%m-%Y")
        output_html = template.render(name=row['Beneficiary Name'],
                                      age=row['AGE'],
                                      gender=row['SEX'],
                                      beneficiary_id=row['Beneficiary ID'],
                                      ehr_id=row['EHR ID'],
                                      date=current_date)
        os.chdir(r'Z:\Tejaswini\Sleep panel\Sleep_Report_Script\Result')
        base_url = os.path.dirname(os.path.abspath(__file__))
        HTML(string=output_html, base_url=base_url).write_pdf(str(row['Barcode Id']) + "_First_page.pdf")


directory_path = "Z:/Tejaswini/Sleep panel/Sleep_Report_Script/Input"
for filename in os.listdir(directory_path):
    if filename.startswith("Sleep") and filename.endswith(".xlsx"):
        file = os.path.join(directory_path, filename)
        cover_page(file)
print("Cover page Done...")
