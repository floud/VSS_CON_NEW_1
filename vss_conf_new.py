import os
import subprocess
import requests
import shutil

# Настройки Confluence
from secrets_1 import AUTH, CONFLUENCE_CONTENT_URL, LOCAL_VSS_PATH, SPACE_KEY

HEADERS = {
    'Authorization': f'Bearer {AUTH}',
    'X-Atlassian-Token': 'nocheck'}
PARENT_PAGE_ID = "360597"  # ID родительской страницы в Confluence

def convert_word_to_html(word_file):
    # Генерация имени для HTML файла и директории media
    base_name = os.path.splitext(os.path.basename(word_file))[0]
    html_file = f"{base_name}.html"
    media_dir = f"{base_name}_media"
    
    # Конвертация Word файла в HTML с использованием pandoc
    subprocess.run(['pandoc', word_file, '-o', html_file, '--extract-media', media_dir, '--embed-resources'], check=True)
    
    return html_file, media_dir

def upload_html_and_media_to_confluence(html_file, media_dir, parent_id):
    # Загрузка HTML файла как страницы
    with open(html_file, 'r') as f:
        html_content = f.read()
        data = {
            'type': 'page',
            'title': os.path.basename(html_file),
            'ancestors': [{'id': parent_id}],
            'space': {'key': SPACE_KEY},
            'body': {
                'storage': {
                    'value': html_content,
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
            page_id = response.json()['id']
        else:
            print(f"Failed to upload HTML page '{html_file}'. Status Code: {response.status_code}, Response: {response.text}")
            return None

    # Загрузка медиафайлов (изображений)
    if os.path.exists(media_dir):
        for root, _, files in os.walk(media_dir):
            for file_name in files:
                file_path = os.path.join(root, file_name)
                upload_attachment_to_confluence(file_path, page_id)

    return page_id

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
    html_file, media_dir = convert_word_to_html(word_file)
    page_id = upload_html_and_media_to_confluence(html_file, media_dir, parent_id)
    # Очистка временных файлов
    if os.path.exists(html_file):
        os.remove(html_file)
    if os.path.exists(media_dir):
        shutil.rmtree(media_dir)
    return page_id

def process_directory(directory, parent_id):
    for root, dirs, files in os.walk(directory):
        for name in files:
            file_path = os.path.join(root, name)
            if name.lower().endswith(('.doc', '.docx', '.rtf')):
                print(f"Processing Word file: {file_path}")
                process_word_file(file_path, parent_id)
            elif name.lower()== '.ds_store':
                continue
            else:
                # Загрузка других типов файлов как вложений
                print(f"Uploading attachment: {file_path}")
                upload_attachment_to_confluence(file_path, parent_id)
        break  # Обрабатываем только верхний уровень файлов в директории

# Запуск процесса обработки директории
process_directory(LOCAL_VSS_PATH, PARENT_PAGE_ID)