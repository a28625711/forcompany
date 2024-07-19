# forcompany
Private

https://1drv.ms/i/c/7f34c6d1925ecc82/EZMk2qtwFEVFhNb8e3xEXH0BPkN5ClzP0JGyDR0RKbsAnQ

https://1drv.ms/i/c/7f34c6d1925ecc82/EecYfx5a9vZEnYISWN0a3UIB9uqfhQxCcwtHap8WAYqBxg


import os
import shutil

# 源文件路径
src_path = r'C:\Users\User\Documents\source_folder'

# 目标文件路径
dst_path = r'C:\Users\User\Documents\target_folder'

# 需要转移的文件类型列表
file_types = ['.pdf', '.docx', '.xlsx']

# 是否覆盖重复文件的开关
overwrite_files = True

for root, dirs, files in os.walk(src_path):
    for file in files:
        file_ext = os.path.splitext(file)[1].lower()
        if file_ext in file_types:
            src_file = os.path.join(root, file)
            dst_file = os.path.join(dst_path, file)
            
            # 检查目标文件是否已存在
            if os.path.exists(dst_file):
                if overwrite_files:
                    # 覆盖重复文件
                    os.remove(dst_file)
                    shutil.move(src_file, dst_path)
                    print(f'Overwritten: {file}')
                else:
                    # 保留重复文件
                    print(f'Skipped: {file}')
            else:
                # 移动文件
                shutil.move(src_file, dst_path)
                print(f'Moved: {file}')