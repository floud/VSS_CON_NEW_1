import os
import re
import subprocess
import requests

# Настройки Confluence

from secrets_1 import AUTH, CONFLUENCE_CONTENT_URL, LOCAL_VSS_PATH, SPACE_KEY


HEADERS = {
    'Authorization': f'Bearer {AUTH}',
    'X-Atlassian-Token': 'nocheck'
}
PARENT_PAGE_ID = "1147052"  # ID родительской страницы в Confluence



def clean_html(html_content):
    """
    Очистка и исправление HTML-кода перед отправкой в Confluence.
    """
    # Удаляем все XML-директивы, которые могут вызвать ошибки
    cleaned_html = re.sub(r'<!DOCTYPE[^>]*>', '', html_content, flags=re.IGNORECASE)
    cleaned_html = re.sub(r'<\?xml[^>]*\?>', '', cleaned_html, flags=re.IGNORECASE)
    cleaned_html = re.sub(r'<!\[CDATA\[.*?\]\]>', '', cleaned_html, flags=re.DOTALL)

    # Исправляем некорректные атрибуты
    cleaned_html = re.sub(r'(\s\w+)=([^"\s>]+)', r'\1="\2"', cleaned_html)
    cleaned_html = re.sub(r'(\s\w+)=["\']?x([a-zA-Z0-9]+)["\']?', r'\1="\2"', cleaned_html)

    # Удаление проблемных символов, которые могут вызывать ошибки при парсинге
    cleaned_html = re.sub(r'[^\x00-\x7F]+', '', cleaned_html)  # Удаление не-ASCII символов
    cleaned_html = re.sub(r'&nbsp;', ' ', cleaned_html)  # Заменяем неразрывные пробелы

    return cleaned_html

def convert_word_to_html(word_file):
    # Генерация имени для HTML файла
    base_name = os.path.splitext(os.path.basename(word_file))[0]
    html_file = f"{base_name}.html"
    
    # Конвертация Word файла в HTML с использованием pandoc
    subprocess.run(['pandoc', word_file, '-o', html_file, '--standalone', '--embed-resources'], check=True)
    
    return html_file

def upload_html_to_confluence(html_file, parent_id):
    # Загрузка HTML файла как страницы
    with open(html_file, 'r') as f:
        html_content = f.read()
        
        # Очистка HTML-кода
        cleaned_html = clean_html(html_content)
        
        data = {
            'type': 'page',
            'title': os.path.basename(html_file),
            'ancestors': [{'id': parent_id}],
            'space': {'key': SPACE_KEY},
            'body': {
                'storage': {
                    'value': cleaned_html,
                    'representation': 'storage'
                }
            }
        }
        response = requests.post(
            f"{CONFLUENCE_CONTENT_URL}",
            json=data,
            headers=HEADERS
        )
        if response.status_code == 200:
            print(f"HTML page '{html_file}' uploaded successfully.")
            return response.json()['id']
        else:
            print(f"Failed to upload HTML page '{html_file}'. Status Code: {response.status_code}, Response: {response.text}")
            return None

def upload_attachment_to_confluence(file_path, parent_id):
    """Загрузка файла как вложения к странице Confluence"""
    file_name = os.path.basename(file_path)
    try:
        with open(file_path, 'rb') as f:
            files = {'file': (file_name, f)}
            response = requests.post(
                f"{CONFLUENCE_CONTENT_URL}{parent_id}/child/attachment",
                files=files,
                headers=HEADERS
            )
            if response.status_code == 200:
                print(f"File '{file_name}' uploaded successfully.")
            else:
                print(f"Failed to upload file '{file_name}'. Status Code: {response.status_code}, Response: {response.text}")
    except Exception as e:
        print(f"An error occurred while uploading file '{file_name}': {str(e)}")

def process_word_file(word_file, parent_id):
    html_file = convert_word_to_html(word_file)
    page_id = upload_html_to_confluence(html_file, parent_id)
    # Очистка временных файлов (если необходимо)
    if os.path.exists(html_file):
        os.remove(html_file)
    return page_id

def create_confluence_page(title, parent_id):
    """Создание новой страницы в Confluence"""
    data = {
        'type': 'page',
        'title': title,
        'ancestors': [{'id': parent_id}],
        'space': {'key': SPACE_KEY},
        'body': {
            'storage': {
                'value': '<p>This page was automatically created.</p>',
                'representation': 'storage'
            }
        }
    }
    response = requests.post(
        f"{CONFLUENCE_CONTENT_URL}",
        json=data,
        headers=HEADERS
    )
    if response.status_code == 200:
        print(f"Page '{title}' created successfully.")
        return response.json()['id']
    else:
        print(f"Failed to create page '{title}'. Status Code: {response.status_code}, Response: {response.text}")
        return None

def process_directory(directory, parent_id):
    for root, dirs, files in os.walk(directory):
        for dir_name in dirs:
            # Создание страницы для каждой директории
            dir_path = os.path.join(root, dir_name)
            page_id = create_confluence_page(dir_name, parent_id)
            if page_id:
                process_directory(dir_path, page_id)  # Рекурсивно обрабатываем вложенные директории

        for name in files:
            file_path = os.path.join(root, name)
            if name.lower().endswith(('.doc', '.docx', '.rtf')):
                print(f"Processing Word file: {file_path}")
                process_word_file(file_path, parent_id)
            else:
                # Загрузка других типов файлов как вложений
                print(f"Uploading attachment: {file_path}")
                upload_attachment_to_confluence(file_path, parent_id)
        break  # Останавливаем обработку, чтобы не заходить в поддиректории снова

# Запуск процесса обработки директории
process_directory(LOCAL_VSS_PATH, PARENT_PAGE_ID)