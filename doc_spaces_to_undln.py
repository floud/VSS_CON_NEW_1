import os

def replace_spaces_in_doc_filenames(directory):
    for root, dirs, files in os.walk(directory):
        for filename in files:
            if filename.endswith('.doc') and ' ' in filename:
                new_filename = filename.replace(' ', '_')
                old_filepath = os.path.join(root, filename)
                new_filepath = os.path.join(root, new_filename)
                os.rename(old_filepath, new_filepath)
                print(f'Renamed: "{old_filepath}" -> "{new_filepath}"')

# Укажите путь к директории, в которой нужно выполнить замену
directory_path = '/path/to/your/directory'
replace_spaces_in_doc_filenames(directory_path)