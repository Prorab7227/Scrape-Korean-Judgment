import os
import requests
from bs4 import BeautifulSoup
import pandas as pd
from openpyxl.styles import Font
from tqdm import tqdm
from PyPDF2 import PdfReader
import re

BASE_URL = "https://www.scourt.go.kr/supreme/info/JpBoardListAction.work?gubun=1"

def extract_pdf_link(news_url):
    response = requests.get(news_url)
    soup = BeautifulSoup(response.content, 'html.parser')

    pdf_link = None
    attachments = soup.find_all('a')
    for attachment in attachments:
        if attachment.has_attr('href') and '.pdf' in attachment['href']:
            pdf_link = attachment['href']
            break

    return pdf_link

def download_pdf(pdf_url, save_path, incident_number):
    response = requests.get(pdf_url, stream=True)
    total_size = int(response.headers.get('content-length', 0))

    with open(save_path, 'wb') as file, tqdm(
        desc=f"Downloading incident {incident_number}",
        total=total_size,
        unit='B',
        unit_scale=True,
        unit_divisor=1024,
        ncols=80,
        bar_format="{l_bar}{bar}| {n_fmt}/{total_fmt} [{elapsed}]"
    ) as bar:
        for data in response.iter_content(1024):
            file.write(data)
            bar.update(len(data))

def extract_text_from_pdf(pdf_path):
    reader = PdfReader(pdf_path)
    full_text = ""
    for page in reader.pages:
        full_text += page.extract_text()
    return full_text

def extract_decision_dates(pdf_text):
    prev_decision_date = ""
    decision_text = ""

    # Поиск текста между "원 심 판 결" или "재심대상판결" и "판 결 선 고" либо "주       문" для Previous decision date
    decision_markers = ['원 심 판 결', '재심대상판결']
    end_marker = '판 결 선 고'
    order_marker = '주       문'

    for marker in decision_markers:
        if marker in pdf_text:
            if end_marker in pdf_text:
                prev_decision_date = pdf_text.split(marker)[1].split(end_marker)[0].strip()
            elif order_marker in pdf_text:
                prev_decision_date = pdf_text.split(marker)[1].split(order_marker)[0].strip()
            break

    # Поиск текста между "주       문" и "이       유" для Decision
    if '주       문' in pdf_text and '이       유' in pdf_text:
        decision_text = pdf_text.split('주       문')[1].split('이       유')[0].strip()

        # Удаление текста после тире и числа (например, "- 2024")
        decision_text = re.split(r'\-\s*\d+', decision_text)[0].strip()

    return prev_decision_date, decision_text

def main(start_page=1, end_page=5):
    data = []
    pdf_folder = 'pdf_files'
    if not os.path.exists(pdf_folder):
        os.makedirs(pdf_folder)

    for page_index in range(start_page, end_page + 1):
        url = f"{BASE_URL}&pageIndex={page_index}"  # Формируем URL для каждой страницы
        response = requests.get(url)
        html = response.content

        soup = BeautifulSoup(html, 'html.parser')
        rows = soup.find('tbody').find_all('tr')

        for row in rows:
            cells = row.find_all('td')
            if cells:
                final_decision_date = cells[1].text.strip()
                incident_number = cells[3].text.strip()
                news_link = cells[-1].find('a')
                news_link_url = news_link['href'] if news_link else None

                pdf_link = extract_pdf_link(news_link_url) if news_link_url else None

                if pdf_link:
                    pdf_filename = f"{incident_number}.pdf"
                    pdf_path = os.path.join(pdf_folder, pdf_filename)

                    if os.path.exists(pdf_path):
                        print(f"File for incident {incident_number} already exists, skipping download.")
                    else:
                        download_pdf(pdf_link, pdf_path, incident_number)

                    # Извлечение текста из PDF
                    pdf_text = extract_text_from_pdf(pdf_path)
                    prev_decision_date, decision_text = extract_decision_dates(pdf_text)
                else:
                    prev_decision_date = decision_text = None

                data.append([final_decision_date, incident_number, pdf_link, prev_decision_date, decision_text])

    # Создаем DataFrame с новыми столбцами
    df = pd.DataFrame(data, columns=['Final decision date', 'Incident number', 'PDF link', 'Previous decision date', 'Decision'])

    # Сохраняем данные в Excel
    with pd.ExcelWriter('parsed_data.xlsx', engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Data')

        workbook = writer.book
        worksheet = writer.sheets['Data']

        # Добавляем гиперссылки на PDF
        for i in range(2, len(data) + 2):
            if data[i-2][2]:
                link = data[i-2][2]
                worksheet[f'C{i}'] = f'=HYPERLINK("{link}", "Download PDF")'
                worksheet[f'C{i}'].font = Font(color="0000FF", underline="single")

        # Делаем заголовки жирным шрифтом
        for cell in worksheet["1:1"]:
            cell.font = Font(bold=True)

    print(f"Data successfully saved to 'parsed_data.xlsx' and PDFs downloaded to '{pdf_folder}'")

if __name__ == "__main__":
    start_page = 1
    end_page = 5
    main(start_page, end_page)
