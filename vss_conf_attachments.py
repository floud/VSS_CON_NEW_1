import requests
import json

from secrets_1 import CONFLUENCE_BASE_URL, AUTH, SPACE_KEY
HEADERS = {
    'Authorization': f'Bearer {AUTH}',
    'X-Atlassian-Token': 'nocheck',
    "Accept": "application/json",
    "Content-Type": "application/json"}

def get_page_content(page_id):
    url = f"{CONFLUENCE_BASE_URL}/rest/api/content/{page_id}?expand=body.storage,version"
    response = requests.get(url, headers=HEADERS)
    response.raise_for_status()
    return response.json()

def update_page_content(page_id, new_content, current_version, title):
    url = f"{CONFLUENCE_BASE_URL}/rest/api/content/{page_id}"
    data = {
        "version": {
            "number": current_version + 1
        },
        "type": "page",
        "title": title,
        "body": {
            "storage": {
                "value": new_content,
                "representation": "storage"
            }
        }
    }
    response = requests.put(url, headers=HEADERS, data=json.dumps(data))
    response.raise_for_status()
    return response.json()

def add_attachments_macro_if_needed(page_id, page_data):
    original_content = page_data['body']['storage']['value']
    title = page_data['title']
    
    # Проверяем наличие текста "This page was automatically created."
    if "This page was automatically created." in original_content:
        print(f"Adding attachments macro to page {page_id}")
        
        # Добавляем макрос вложений в конец страницы (без ac:macro-id)
        attachments_macro = '<p><ac:structured-macro ac:name="attachments" ac:schema-version="1"/></p>'
        new_content = original_content + attachments_macro
        
        # Обновляем страницу с новым содержимым
        update_page_content(page_id, new_content, page_data['version']['number'], title)
        print(f"Updated page {page_id}: Attachments macro added.")
    else:
        print(f"No changes needed for page {page_id}.")

def process_space_pages(space_key):
    pages = get_all_pages(space_key)
    
    for page in pages:
        page_id = page['id']
        page_data = get_page_content(page_id)
        
        print(f"Processing page {page_id}: {page['title']}")
        
        if 'version' not in page_data:
            print(f"Error: 'version' key not found in page {page_id}. Skipping this page.")
            continue
        
        # Добавляем макрос если нужно
        add_attachments_macro_if_needed(page_id, page_data)

def get_all_pages(space_key):
    url = f"{CONFLUENCE_BASE_URL}/rest/api/content?spaceKey={space_key}&type=page&limit=500"
    response = requests.get(url, headers=HEADERS)
    response.raise_for_status()
    return response.json()['results']

# Запуск обработки страниц пространства

process_space_pages(SPACE_KEY)
# pages = get_all_pages(SPACE_KEY)
# for page in pages:
    # print(page['id'])
