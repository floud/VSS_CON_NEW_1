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
PARENT_PAGE_ID = "1540103"  # ID родительской страницы в Confluence

def convert_word_to_html(word_file):
    base_name = os.path.splitext(os.path.basename(word_file))[0]
    html_file = f"{base_name}.html"
    
    # Конвертация Word файла в HTML с использованием pandoc
    subprocess.run(['pandoc', word_file, '-o', html_file, '--standalone', '--embed-resources'], check=True)
    
    return html_file

def tidy_html(html_file):
    """
    Проверка и исправление HTML файла с помощью Tidy.
    """
    tidy_file = f"{html_file}.tidy"
    # Удаление параметра check=True, чтобы избежать прерывания скрипта при предупреждениях
    result = subprocess.run(['tidy', '-quiet', '-indent', '--drop-empty-elements', 'no', '-asxhtml', '-utf8', '-o', tidy_file, html_file])
    
    if result.returncode != 0:
        print(f"Tidy завершился с кодом {result.returncode}. Возможны проблемы с HTML-кодом.")
    
    # Чтение и возврат отформатированного файла
    with open(tidy_file, 'r') as f:
        tidy_html_content = f.read()
    
    # Удаление промежуточного файла
    os.remove(tidy_file)
    
    return tidy_html_content

def upload_html_to_confluence(html_file, parent_id):
    # Tidy HTML для исправления кода
    cleaned_html = tidy_html(html_file)
    
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
    
    if os.path.exists(html_file):
        os.remove(html_file)
    
    return page_id

def create_confluence_page(title, parent_id):
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
            dir_path = os.path.join(root, dir_name)
            page_id = create_confluence_page(dir_name, parent_id)
            if page_id:
                process_directory(dir_path, page_id)
        
        for name in files:
            file_path = os.path.join(root, name)
            if name.lower().endswith(('.doc', '.docx', '.rtf')):
                print(f"Processing Word file: {file_path}")
                process_word_file(file_path, parent_id)
            else:
                print(f"Uploading attachment: {file_path}")
                upload_attachment_to_confluence(file_path, parent_id)
        break

# Запуск процесса обработки директории
process_directory(LOCAL_VSS_PATH, PARENT_PAGE_ID)
