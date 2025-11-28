# pdfocr

```
main.py环境安装
conda create -n lmdeploy python=3.11 -y && conda activate lmdeploy
pip install lmdeploy partial_json_parser timm
pip install flask pdfplumber pdf2image werkzeug 
pip install huggingface_hub
pip install python-docx
apt-get update
apt-get install poppler-utils
pip install modelscope
modelscope download --model OpenGVLab/InternVL3-2B-Instruct --local_dir ./OpenGVLab/InternVL3-2B-Instruct
lmdeploy serve api_server OpenGVLab/InternVL3-2B-Instruct --backend turbomind --server-port 23333 --tp 1 --chat-template internvl2_5
```


### 文本格式
```
pip install openai python-docx -i https://mirrors.aliyun.com/pypi/simple/

```