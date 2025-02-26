import os

def convert_to_crlf(file_path):
    with open(file_path, 'r', encoding='utf-8', newline='') as file:
        content = file.read()
    
    # 使用 CRLF (\r\n) 替换换行符
    content_crlf = content.replace('\r\n', '\n').replace('\n', '\r\n')

    with open(file_path, 'w', encoding='utf-8', newline='') as file:
        file.write(content_crlf)

def convert_directory_to_crlf(directory):
    for root, _, files in os.walk(directory):
        for file in files:
            file_path = os.path.join(root, file)
            try:
                convert_to_crlf(file_path)
                print(f"Converted: {file_path}")
            except Exception as e:
                print(f"Error converting {file_path}: {e}")

if __name__ == "__main__":
    target_directory = input("Enter the directory path: ")
    convert_directory_to_crlf(target_directory)
