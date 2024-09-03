import requests
import re
import json

from secrets_1 import CONFLUENCE_BASE_URL, AUTH, SPACE_KEY
HEADERS = {
    'Authorization': f'Bearer {AUTH}',
    'X-Atlassian-Token': 'nocheck',
    "Accept": "application/json",
    "Content-Type": "application/json"}
def get_all_pages(space_key):
    url = f"{CONFLUENCE_BASE_URL}/rest/api/content?spaceKey={space_key}&type=page&limit=1000"
    response = requests.get(url, headers=HEADERS)
    response.raise_for_status()
    return response.json()['results']

def get_page_content(page_id):
    url = f"{CONFLUENCE_BASE_URL}/rest/api/content/{page_id}?expand=body.storage,version"
    response = requests.get(url, headers=HEADERS)
    response.raise_for_status()
    return response.json()

def get_page_attachments(page_id):
    url = f"{CONFLUENCE_BASE_URL}/rest/api/content/{page_id}/child/attachment"
    response = requests.get(url, headers=HEADERS)
    response.raise_for_status()
    return response.json()['results']

def update_page_content(title, page_id, new_content, current_version):
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
    try:
        response.raise_for_status()
    except requests.exceptions.HTTPError as e:
        print(f"Error updating page {page_id}: {e}")
        print(f"Response content: {response.text}")
        raise
    return response.json()


def find_and_replace_image_links(page_content, attachments):
    updated_content = page_content
    
    # Регулярное выражение для поиска тегов img
    image_links = re.findall(r'<img src="([^"]+)"', page_content)
    print(f"Found image links: {image_links}")
    
    for link in image_links:
        filename = link.split("/")[-1]  # Получаем только имя файла
        matching_attachment = next((att for att in attachments if att['title'] == filename), None)
        
        if matching_attachment:
            # Создание нового тега для вставки изображения из вложений
            new_tag = f'<ac:image><ri:attachment ri:filename="{matching_attachment["title"]}" /></ac:image>'
            updated_content = updated_content.replace(f'<img src="{link}"', new_tag)
            print(f"Replacing {link} with {new_tag}")
        else:
            print(f"No matching attachment found for {filename}")

    return updated_content

def process_space_pages(space_key):
    pages = get_all_pages(space_key)
    
    for page in pages:
        page_id = page['id']
        page_data = get_page_content(page_id)
        attachments = get_page_attachments(page_id)
        
        print(f"Processing page {page_id}: {page['title']}")

        if 'version' not in page_data:
            print(f"Error: 'version' key not found in page {page_id}. Skipping this page.")
            continue
        
        original_content = page_data['body']['storage']['value']
        new_content = find_and_replace_image_links(original_content, attachments)
        
        if original_content != new_content:
            update_page_content(page['title'],page_id, new_content, page_data['version']['number'])
            print(f"Updated page {page_id}: Links replaced.")
        else:
            print(f"No changes made to page {page_id}.")

# Запуск процесса
process_space_pages(SPACE_KEY)