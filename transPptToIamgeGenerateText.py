from pptx import Presentation
import os, sys
import glob
import base64
import requests, json
import zipfile  
import re
from xml.etree import ElementTree as ET  
import shutil  
sys.path.append("C:\\MSWORK\\MyCode\\Azure\\authentication")
from secret import config_gpt4v

# python -m pip install -U pip
# python -m pip install -U python-pptx pillow
# linux: apt-get install libreoffice imagemagick  

DOC_PATH = "C:\\MSWORK\\tmp" # 本地文件路径

def extract_video_from_pptx(pptx_file):
        # PPTX文件路径  
    pptx_file = "C:\\MSWORK\\tmp\\AzureOpenAISeries-GPT-4TurbowithVision.pptx"
    pptxbasename = (os.path.basename(pptx_file)).split('.')[0]
    # 解压缩后的临时目录 
    extract_dir = f"C:\\MSWORK\\tmp\\{pptxbasename}\\temp_dir"
    # 存储提取出的视频的目录 
    videos_dir = f"C:\\MSWORK\\tmp\\{pptxbasename}\\video_dir" 
    # 打开PPTX文件作为Zip文件  
    with zipfile.ZipFile(pptx_file, 'r') as zip_ref:  
        # 解压缩文件到临时目录  
        zip_ref.extractall(extract_dir)  
    # 提取并命名视频文件  
    for i, rels_file in enumerate(sorted(glob.glob(os.path.join(extract_dir, 'ppt', 'slides', '_rels', '*.rels')))):  
        tree = ET.parse(rels_file)  
        page_numbers = re.findall(r'\d+', os.path.basename(rels_file))[0]
        root = tree.getroot() 
        for element in root.findall('.//*'):
            if 'video' in element.get('Type'):
                # 创建存储视频的目录  
                if not os.path.isdir(videos_dir):  
                    os.makedirs(videos_dir) 
                video_file = element.get('Target') 
                video_filename = os.path.basename(video_file)  
                print(f'Extracting {video_filename} from slide {page_numbers}')
                video_new_path = os.path.join(videos_dir, f'slide_{page_numbers}_{video_filename}')  
                old_path = os.path.join(extract_dir, 'ppt', 'media', video_filename)
                shutil.copy(old_path, video_new_path)
    # 清理解压缩出的临时文件夹  
    if os.path.exists(extract_dir):  
        shutil.rmtree(extract_dir)  
    if os.path.exists(videos_dir):  
        print(f"视频提取完成，存放在 {videos_dir} 目录下")

def ppt2image(): #按页转换ppt为图片
    import platform
    os_name = platform.system()
    if os_name == "Windows":
        import comtypes
        from comtypes.client import CreateObject

        # 确保存储图片的目录存在  
        if not os.path.exists(DOC_PATH):  
            os.makedirs(DOC_PATH) 
        file_list = glob.glob(f"{DOC_PATH}\\*.pptx")
        for file in file_list:
            extract_video_from_pptx(file)
            sava_img_dir = f"{DOC_PATH}\\{os.path.basename(file).split('.')[0]}"
            if not os.path.exists(sava_img_dir):
                os.makedirs(sava_img_dir) 
            comtypes.CoInitialize() 
            powerpoint = CreateObject("Powerpoint.Application")
            presentation = powerpoint.Presentations.Open(file)
            for i, slide in enumerate(presentation.Slides):  
                slide.Export(f"{sava_img_dir}\\{i+1}.png", "PNG")
            presentation.Close()  
            # 退出PowerPoint  
            powerpoint.Quit()
            comtypes.CoUninitialize()
        print(f"图片装换完成，存放在 {sava_img_dir} 目录下")
    elif os_name == 'Linux':  
        print("System is Linux, need to install libreoffice and imagemagick")  
    else:  
        print("Unknown system")

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
    ppt2image()
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
