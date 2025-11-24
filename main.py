import pdfplumber
from flask import Flask, jsonify,send_from_directory,request
from werkzeug.utils import secure_filename
from pdf2image import convert_from_path
import os
import requests
from openai import OpenAI
from docx import Document

app = Flask(__name__)

#docx转成txt
def docx_to_txt(docx_file, txt_file):
    # 打开 docx 文件
    doc = Document(docx_file)
    # 初始化一个空字符串来存储文本内容
    full_text = []
    
    # 打开 txt 文件，准备写入
    with open(txt_file, 'w', encoding='utf-8') as txt:
        # 遍历文档中的每个段落
        for para in doc.paragraphs:
            # 写入段落文本到 txt 文件
            txt.write(para.text + '\n')
            full_text.append(para.text)
    return '\n'.join(full_text)
#删除文件
def isFileExit(workspace_path):
    if os.path.exists(workspace_path) and os.path.isdir(workspace_path):
    # 遍历目录中的所有文件和子目录
        for filename in os.listdir(workspace_path):
            file_path = os.path.join(workspace_path, filename)  # 获取文件的完整路径
            if os.path.isfile(file_path):  # 判断是否为文件
                os.remove(file_path)  # 删除文件
                print(f"已删除文件: {file_path}")
            elif os.path.isdir(file_path):  # 如果是子目录，可以选择递归删除或跳过
                print(f"跳过子目录: {file_path}")
        print("目录中的所有文件已删除。")
    else:
        print(f"路径不存在或不是一个目录: {workspace_path}")
# @app.route('/api/upload-pdf', methods=['POST'])
# def pdf_to_text():
#     global pdfDist_path
#     global pdfFilePath
#     """
#        将PDF文件转换为文本
#        :param pdf_path: PDF文件路径
#        :param output_path: 输出文本文件路径，如果为None则打印到控制台
#        """
#     #获取文件夹名称
#     folder_name = request.headers.get('X-Folder-Name')
#     pdfDist_path = folder_name
#     print(folder_name,'文件夹名称')
#     if not folder_name:
#         return jsonify({"error": "No folder name provided"}), 400
#     # 确保文件夹路径安全
#     upload_folder = secure_filename(folder_name)
#     # 如果文件夹不存在，则创建文件夹
#     if not os.path.exists(upload_folder):
#         os.makedirs(upload_folder)
#     else:
#         isFileExit(upload_folder)
#     #如果数据不存在就返回400
#     if not request.data:
#         return jsonify({"error": "No file content"}), 400

#     # 获取文件名
#     filename = secure_filename(request.headers.get('Content-Disposition', '').split('filename=')[-1].strip('"'))
#     print(filename, '文件名称==')
#     if not filename:
#         return jsonify({"error": "No filename provided"}), 400 
#     file_path = os.path.join(upload_folder, filename)
#     pdfFilePath=file_path
#     with open(file_path, 'wb') as f:
#         f.write(request.data)

#     try:
#         # 打开PDF文件
#         with pdfplumber.open(file_path) as pdf:
#             text = ""
#             # 遍历每一页
#             for page in pdf.pages:
#                 # 提取当前页的文本
#                 page_text = page.extract_text()
#                 if page_text:
#                     text += page_text + "\n\n"

#         # 创建与文件夹同名的txt文件夹
#         txt_folder = f"{upload_folder}_txt"
#         if not os.path.exists(txt_folder):
#             os.makedirs(txt_folder)

#         # 构建输出文件路径，保持原文件名但更换扩展名
#         base_filename = os.path.splitext(filename)[0]  # 获取不带扩展名的文件名
#         output_filename = f"{base_filename}.txt"
#         output_path = os.path.join(txt_folder, output_filename)

#         # 保存文本到文件
#         with open(output_path, 'w', encoding='utf-8') as f:
#             f.write(text)

#         return jsonify({"message": f"PDF converted and saved to {output_path}", "output_path": output_path,'content':text}), 200

#     except Exception as e:
#         return jsonify({"error": f"处理PDF时出错: {str(e)}"}), 500
@app.route('/api/upload-pdf', methods=['POST'])
def pdf_to_text():
    global pdfDist_path
    global pdfFilePath
    """
       将PDF文件转换为文本
       :param pdf_path: PDF文件路径
       :param output_path: 输出文本文件路径，如果为None则打印到控制台
       """
    #获取文件夹名称
    folder_name = request.headers.get('X-Folder-Name')
    pdfDist_path = folder_name
    print(folder_name,'文件夹名称')
    if not folder_name:
        return jsonify({"error": "No folder name provided"}), 400
    # 确保文件夹路径安全
    upload_folder = secure_filename(folder_name)
    # 如果文件夹不存在，则创建文件夹
    if not os.path.exists(upload_folder):
        os.makedirs(upload_folder)
    else:
        isFileExit(upload_folder)
    #如果数据不存在就返回400
    if not request.data:
        return jsonify({"error": "No file content"}), 400

    # 获取文件名
    filename = secure_filename(request.headers.get('Content-Disposition', '').split('filename=')[-1].strip('"'))
    print(filename, '文件名称==')
    if not filename:
        return jsonify({"error": "No filename provided"}), 400 
     # 检查文件扩展名
    file_extension = filename.lower().split('.')[-1]
    if file_extension not in ['pdf', 'docx']:
        return jsonify({"error": "Only .pdf or .docx files are allowed"}), 400

    # 根据文件类型选择文件夹
    if file_extension == 'pdf':
        file_folder = os.path.join(upload_folder, 'pdf_files')
    elif file_extension == 'docx':
        file_folder = os.path.join(upload_folder, 'docx_files')

    # 如果文件夹不存在，则创建文件夹
    if not os.path.exists(file_folder):
        os.makedirs(file_folder)
    file_path = os.path.join(file_folder, filename)
    pdfFilePath=file_path
    with open(file_path, 'wb') as f:
        f.write(request.data)

    try:
        # 创建与文件夹同名的txt文件夹
        txt_folder = f"{upload_folder}_txt"
        if not os.path.exists(txt_folder):
            os.makedirs(txt_folder)
        else:
            isFileExit(txt_folder)
        # 构建输出文件路径，保持原文件名但更换扩展名
        base_filename = os.path.splitext(filename)[0]  # 获取不带扩展名的文件名
        output_filename = f"{base_filename}.txt"
        output_path = os.path.join(txt_folder, output_filename)
        # 打开PDF文件
        if file_extension == 'pdf':
            with pdfplumber.open(file_path) as pdf:
                text = ""
                # 遍历每一页
                for page in pdf.pages:
                    # 提取当前页的文本
                    page_text = page.extract_text()
                    if page_text:
                        text += page_text + "\n\n"
            # 保存文本到文件
            with open(output_path, 'w', encoding='utf-8') as f:
                f.write(text)

            return jsonify({"message": f"PDF converted and saved to {output_path}", "output_path": output_path,'content':text}), 200
        elif file_extension == 'docx':
            docx_txt=docx_to_txt(file_path,output_path)
            return jsonify({"message": f"docx converted and saved to {output_path}", "output_path": output_path,'content':docx_txt}), 200

    except Exception as e:
        return jsonify({"error": f"处理PDF时出错: {str(e)}"}), 500
    

@app.route('/api/filetxt/<path:filename>', methods=['GET'])
def pdf_txt(filename):
    #获取pdf文件名称
    pdfs_list=[filename]
    print(pdfFilePath,'llllllllllllll')
    images = convert_from_path(pdfFilePath)
    print(pdfs_list,'kkkkkkk')
    file_name = filename.split(".")[1]
    # 创建与文件夹同名的img文件夹
    temp_folder = f"{pdfDist_path}_img"
    img_folder = os.path.join(temp_folder,file_name)
    # 如果文件夹不存在，则创建文件夹
    if not os.path.exists(img_folder):
        os.makedirs(img_folder)
    else:
        isFileExit(img_folder)

    image_paths = []
    for i, image in enumerate(images):
        image_path = os.path.join(img_folder, f"page_{i + 1}.jpg")
        image.save(image_path, "JPEG")
        image_paths.append(image_path)

    print(f"PDF 转换为图片成功，图片保存在 {temp_folder} 文件夹中")

    # 调用大模型（这里假设有一个函数 call_large_model）
    res = call_large_model(image_paths)
    return jsonify({"res": res}),200

'''
pdf转图片
调用多模态大模型，获取结果
'''
@app.route('/pdfToTxt', methods=['GET'])
def pdf_process():
    #获取pdf文件名称
    images = convert_from_path(pdfFilePath)
    tempname = pdfFilePath.split("/")[1]
    filename = tempname.split(".")[0]
    # 创建与文件夹同名的img文件夹
    temp_folder = f"{pdfDist_path}_img"
    img_folder = os.path.join(temp_folder,filename)
    # 如果文件夹不存在，则创建文件夹
    if not os.path.exists(img_folder):
        os.makedirs(img_folder)
    else:
        isFileExit(img_folder)

    image_paths = []
    for i, image in enumerate(images):
        image_path = os.path.join(img_folder, f"page_{i + 1}.jpg")
        image.save(image_path, "JPEG")
        image_paths.append(image_path)

    print(f"PDF 转换为图片成功，图片保存在 {temp_folder} 文件夹中")

    # 调用大模型（这里假设有一个函数 call_large_model）
    res = call_large_model(image_paths)
    return jsonify({"res": res}),200
# 定义函数：调用大模型
def call_large_model(image_paths):
    # 这里是一个示例函数，实际调用大模型的逻辑需要根据你的需求实现
    client = OpenAI(api_key='123', base_url='http://0.0.0.0:23333/v1')
    contentList = []
    for image_path in image_paths:
        print(f"调用大模型处理图片：{image_path}")
        model_name = client.models.list().data[0].id
        response = client.chat.completions.create(
            model=model_name,
            messages=[{
                'role':
                'user',
                'content': [{
                    'type': 'text',
                    'text': '展示这张图片的所有文字',
                }, {
                    'type': 'image_url',
                    'image_url': {
                        'url':
                        image_path,
                    },
                }],
            }],
            temperature=0.8,
            top_p=0.8)
        print(response)
        content = response.choices[0].message.content
        contentList.append(content)
        print(content)
    return ','.join(contentList)
    

                   
if __name__ == '__main__':
    app.run(host='0.0.0.0', port=8080, debug=False)

