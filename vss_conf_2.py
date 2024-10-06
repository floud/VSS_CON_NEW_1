import os
import subprocess
import requests
import shutil
from datetime import datetime
import random
import json
import string

# Настройки Confluence
from secrets_1 import AUTH, CONFLUENCE_CONTENT_URL, CONFLUENCE_API_URL, LOCAL_VSS_PATH, SPACE_KEY

HEADERS = {
    'Authorization': f'Bearer {AUTH}',
    'X-Atlassian-Token': 'nocheck'}
PARENT_PAGE_ID = "360468"  # ID родительской страницы в Confluence

# Функция для проверки существования страницы
def page_exists(title, space_key):
    url = f"{CONFLUENCE_API_URL}/content?title={title}&spaceKey={space_key}"
    response = requests.get(url, headers=HEADERS)
    
    if response.status_code == 200:
        data = response.json()
        if data['size'] > 0:  # Если найдена страница с таким названием
            return data['results'][0]['id']
    return None

# Функция для генерации уникального названия
def generate_unique_title(title):
    # Добавляем текущую дату и время к названию
    current_time = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    unique_suffix = ''.join(random.choices(string.ascii_lowercase + string.digits, k=5))  # Генерация случайного суффикса
    new_title = f"{title} - {current_time} - {unique_suffix}"
    return new_title
def convert_word_to_html(word_file):
    # Генерация имени для HTML файла и директории media
    base_name = os.path.splitext(os.path.basename(word_file))[0]
    html_file = f"{base_name}.html"
    media_dir = f"{base_name}_media"
    
    # Конвертация Word файла в HTML с использованием pandoc
    subprocess.run(['pandoc', word_file, '-o', html_file, 
                    # '--embed-resources', 
                    # '--standalone', 
                    '--extract-media', media_dir
                    ]
                    , check=True)
    
    return html_file, media_dir

def upload_html_and_media_to_confluence(html_file, media_dir, parent_id):
    # Проверяем, существует ли страница с таким названием
    title = os.path.basename(html_file)
    existing_page_id = page_exists(title, SPACE_KEY)
    
    if existing_page_id:
        print(f"Страница с названием '{title}' уже существует. Генерация уникального названия.")
        title = generate_unique_title(title)  # Генерация уникального названия
        
    # Загрузка HTML файла как страницы
    with open(html_file, 'r') as f:
        html_content = f.read()
        data = {
            'type': 'page',
            'title': title,
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

def create_confluence_page(title, parent_id):
    # Проверяем, существует ли страница с таким названием
    existing_page_id = page_exists(title, SPACE_KEY)
    
    if existing_page_id:
        print(f"Страница с названием '{title}' уже существует. Генерация уникального названия.")
        title = generate_unique_title(title)  # Генерация уникального названия
    
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

def process_word_file(word_file, parent_id):
    html_file, media_dir = convert_word_to_html(word_file)
    page_id = upload_html_and_media_to_confluence(html_file, media_dir, parent_id)
    # Очистка временных файлов
    if os.path.exists(html_file):
        os.remove(html_file)
    if os.path.exists(media_dir):
        shutil.rmtree(media_dir)
    return page_id

# Запуск процесса обработки директории
process_directory(LOCAL_VSS_PATH, PARENT_PAGE_ID)
