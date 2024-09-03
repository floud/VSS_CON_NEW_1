import os
import subprocess
import requests

from secrets_1 import AUTH, CONFLUENCE_CONTENT_URL, LOCAL_VSS_PATH, SPACE_KEY

HEADERS = {
    'Authorization': f'Bearer {AUTH}',
    'X-Atlassian-Token': 'nocheck'}

def convert_doc_to_html(doc_path):
    output_html = doc_path.replace(os.path.splitext(doc_path)[1], '.html')
    media_folder = "./media"
    command = ["pandoc", doc_path, "-o", output_html, f"--extract-media={media_folder}", "--embed-resources"]
    subprocess.run(command, check=True)
    return output_html, media_folder

def upload_to_confluence(title, content, parent_id=None):
    data = {
        "type": "page",
        "title": title,
        "space": {"key": SPACE_KEY},
        "body": {"storage": {"value": content, "representation": "storage"}}
    }
    
    if parent_id:
        data["ancestors"] = [{"id": parent_id}]

    response = requests.post(CONFLUENCE_CONTENT_URL, json=data, headers=HEADERS)
    
    if response.status_code == 200:
        page_id = response.json().get("id")
        print(f"Page '{title}' created successfully with ID {page_id}.")
        return page_id
    else:
        print(f"Failed to create page '{title}'. Status Code: {response.status_code}")
        return None

def upload_attachment_to_confluence(file_path, parent_id):
    with open(file_path, 'rb') as f:
        filename = os.path.basename(file_path)
        files = {
            'file': (filename, f),
        }
        response = requests.post(
            f"{CONFLUENCE_CONTENT_URL}{parent_id}/child/attachment",
            files=files,
            headers=HEADERS
        )
        
        if response.status_code == 200:
            print(f"File '{filename}' uploaded as attachment successfully.")
        else:
            print(f"Failed to upload file '{filename}'. Status Code: {response.status_code}")

def process_directory(directory, parent_id=None):
    for item in os.listdir(directory):
        item_path = os.path.join(directory, item)
        
        if os.path.isdir(item_path):
            folder_page_id = upload_to_confluence(item, "", parent_id)
            if folder_page_id:
                process_directory(item_path, folder_page_id)
        
        elif item.endswith(('.doc', '.docx', '.rtf')):
            html_file, media_folder = convert_doc_to_html(item_path)
            with open(html_file, 'r') as f:
                content = f.read()
            page_id = upload_to_confluence(item.replace(os.path.splitext(item)[1], ''), content, parent_id)
            if page_id and os.path.exists(media_folder):
                for media_item in os.listdir(media_folder):
                    media_path = os.path.join(media_folder, media_item)
                    upload_attachment_to_confluence(media_path, page_id)

        else:
            upload_attachment_to_confluence(item_path, parent_id)

# Выполнение Checkout из VSS и обработка файлов
process_directory(LOCAL_VSS_PATH)