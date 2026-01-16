import sys
sys.path.insert(0, '.')
from drive_uploader import DriveUploader
import os

# 初始化上传器
uploader = DriveUploader()

# 认证
if not uploader.authenticate():
    print('认证失败！')
    exit(1)

print('认证成功！')

# 查找测试图片
test_pack = r'C:\Users\txk13\Desktop\test_pack\3'
test_img = None
if os.path.exists(test_pack):
    for f in os.listdir(test_pack):
        if f.lower().endswith(('.jpg', '.png')):
            test_img = os.path.join(test_pack, f)
            break

if not test_img:
    print('未找到测试图片')
    exit(1)

print(f'测试图片: {test_img}')

# 上传到用户指定的 pic 文件夹
FOLDER_ID = '1rpmIXMFOT7Hb668lkQnjgjVBa1UQaH2J'
file_name = os.path.basename(test_img)
from googleapiclient.http import MediaFileUpload

file_metadata = {'name': file_name, 'parents': [FOLDER_ID]}
media = MediaFileUpload(test_img, resumable=True)

try:
    file = uploader.service.files().create(
        body=file_metadata, 
        media_body=media, 
        fields='id, webViewLink'
    ).execute()
    
    file_id = file.get('id')
    web_link = file.get('webViewLink')
    print(f'上传成功！')
    print(f'File ID: {file_id}')
    print(f'网页链接: {web_link}')
    
    # 设置为公开
    if uploader.make_public(file_id):
        print('已设置为公开访问')
        print(f'直接链接: https://drive.google.com/uc?id={file_id}')
    else:
        print('设置公开失败')
except Exception as e:
    print(f'上传失败: {e}')
