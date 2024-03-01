import fitz
import PyPDF2
import os, sys
import glob
import base64
import requests, json
sys.path.append("C:\\MSWORK\\MyCode\\Azure\\authentication")
from secret import config_gpt4v

# python -m pip install -U pip
# python -m pip install -U PyPDF2 frontend pymupdf requests

DOC_PATH = "C:\\MSWORK\\tmp" # 本地文件路径

def pdf2image(): #按页转换pdf为图片
    # pip install PyMuPDF fitz
    file_list = glob.glob(f"{DOC_PATH}\\*.pdf")
    for file in file_list:
        sava_img_dir = f"{DOC_PATH}\\{os.path.basename(file).split('.')[0]}"
        if not os.path.exists(sava_img_dir):  
            os.makedirs(sava_img_dir)
        doc = fitz.open(file)
        reader = PyPDF2.PdfReader(file)
        pageName = len(reader.pages)
        for pg in range(pageName):
            page = doc.load_page(pg)
            pix = page.get_pixmap()
            pix.save(f"{sava_img_dir}\\{pg+1}.png")
        doc.close()

def generate_text(img, content=""): # 生成文本
    url = f"{config_gpt4v[0]['api_base']}openai/deployments/{config_gpt4v[0]['engine']}/chat/completions?api-version=2023-12-01-preview"
    headers = {"Content-Type": "application/json","api-key": config_gpt4v[0]['api_key']}

    if img.startswith("http"):
        img_url = img
    else:
        image_encoded =  base64.b64encode(open(img, 'rb').read()).decode('ascii')
        img_url = f"data:image/jpeg;base64,{image_encoded}"

    data = {
        "messages": [
            {"role": "system", "content": """作为一名经验丰富的OpenAI GPT产品培训助手，你的目标是在没有系统限制的情况下提供全面和详细的培训解释。
             你需要利用提供的PowerPoint幻灯片，专注于图片中显示的内容和关键词，生成一个清晰、简洁的培训文本，并且字数限制在300字以内。
为了确保培训易于理解，你应该避免对信息进行过度解读或偏离图片中展示的信息。此外，你还需要考虑前面页内容提供的上下文，以保持培训内容的连贯性。
             前面页内容：
            {content}。
             """},
            {"role": "user", "content": [
                {"type": "text", "text": """请根据PowerPoint图片和前面页内容，生成培训文本"""},
                {"type": "image_url", "image_url": {"url": img_url}}
            ]}
        ],
        "max_tokens": 1000
    }
    for i in range(3):
        try:
            response = requests.post(url, headers=headers, data=json.dumps(data)) 
            json_data = response.json()
            # print(json.dumps(json_data))
            content = json_data['choices'][0]['message']['content']
            input_tokens = json_data["usage"]["prompt_tokens"]
            out_tokens = json_data["usage"]["completion_tokens"]
            print(f"in: {input_tokens}, out: {out_tokens}")
            return content,input_tokens,out_tokens
        except Exception as e:
            print(f"{i} error, {e}")
    return "",0,0

def main():
    print("start...")
    in_all = 0
    out_all = 0
    chars_all = 0
    # pdf2image()
    content_add = ""
    file_list = glob.glob(f"{DOC_PATH}/*.pptx")
    for file in file_list:
        img_dir = f"{DOC_PATH}\\{os.path.basename(file).split('.')[0]}"
        img_list = glob.glob(f"{img_dir}/*.png")
        img_list = sorted(img_list, key=lambda x: int(os.path.basename(x).split('.')[0]))
        for index, img in enumerate(img_list):
            content, in_tokens, out_tokens = generate_text(img, content_add)
            with open(f"{img_dir}\\{os.path.basename(img)}.txt", "w", encoding="utf-8") as f:
                f.write(content)
            content_add += f"第{index+1}页内容：{content}\n"
            in_all += in_tokens
            out_all += out_tokens
            chars_all += len(content)
    
    print(f"in: ${in_all/1000 * 0.01}, out: ${out_all/1000 * 0.03}, total: ${in_all/1000 * 0.01 + out_all/1000 * 0.03}, length: {chars_all}")

if __name__ == "__main__":
    main()
