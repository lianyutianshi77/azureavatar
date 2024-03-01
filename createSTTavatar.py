#_*_ coding: utf-8 _*_
import json
import logging
import os
import sys
import time
import glob
from pathlib import Path
import requests
sys.path.append("C:\\MSWORK\\MyCode\\Azure\\authentication")
from secret import avatar_config

logging.basicConfig(stream=sys.stdout, level=logging.INFO,  # set to logging.DEBUG for verbose output
        format="[%(asctime)s] %(message)s", datefmt="%m/%d/%Y %I:%M:%S %p %Z")
logger = logging.getLogger(__name__)

SUBSCRIPTION_KEY = avatar_config[0]["key"]
SERVICE_REGION = avatar_config[0]["region"]

NAME = "Simple_avatar_synthesis"
DESCRIPTION = "Simple avatar synthesis description"

# The service host suffix.
SERVICE_HOST = "customvoice.api.speech.microsoft.com"

DOC_PATH = "C:\\MSWORK\\tmp" # 本地文件路径

def submit_synthesis(num, text: str):
    print(f"submit_synthesis: {num}")
    url = f'https://{SERVICE_REGION}.{SERVICE_HOST}/api/texttospeech/3.1-preview1/batchsynthesis/talkingavatar'
    header = {
        'Ocp-Apim-Subscription-Key': SUBSCRIPTION_KEY,
        'Content-Type': 'application/json'
    }
    payload = {
        'displayName': f"{num}.webm",
        'description': DESCRIPTION,
        "textType": "SSML",
        "inputs": [
            {"text": f"""<speak xmlns="http://www.w3.org/2001/10/synthesis" xmlns:mstts="http://www.w3.org/2001/mstts" xmlns:emo="http://www.w3.org/2009/10/emotionml" version="1.0" xml:lang="zh-CN"><voice name="zh-CN-XiaoxiaoNeural">{text}</voice></speak>"""},
        ],
        "properties": {
            "customized": False, # 是否使用自定义头像
            "talkingAvatarCharacter": "lisa",  # talking avatar character, required for prebuilt avatar, optional for custom avatar
            "talkingAvatarStyle": "technical-sitting",  # talking avatar style, required for prebuilt avatar, optional for custom avatar
            "videoFormat": "webm",  # mp4 or webm, webm is required for transparent background
            "videoCodec": "vp9",  # hevc, h264 or vp9, vp9 is required for transparent background; default is hevc
            "subtitleType": "soft_embedded",
            "backgroundColor": "transparent", # 背景为透明
        }
    }
    retries = 3
    for i in range(retries):
        try:
            response = requests.post(url, json.dumps(payload), headers=header)
            if response.status_code < 400:
                logger.info('submit_synthesis: Batch avatar synthesis job submitted successfully')
                logger.info(f'submit_synthesis: Job ID: {response.json()["id"]}')
                return response.json()["id"]
            else:
                logger.error(f'submit_synthesis: Failed to submit batch avatar synthesis job: {response}')
        except Exception as e:
            logger.error(f'submit_synthesis: Failed to submit batch avatar synthesis job: {e}')


def get_synthesis(job_id):
    print(f"get_synthesis: {job_id}")
    url = f'https://{SERVICE_REGION}.{SERVICE_HOST}/api/texttospeech/3.1-preview1/batchsynthesis/talkingavatar/{job_id}'
    header = {
        'Ocp-Apim-Subscription-Key': SUBSCRIPTION_KEY
    }
    retries = 3
    for i in range(retries):
        try:
            response = requests.get(url, headers=header)
            if response.status_code < 400:
                logger.debug('get_synthesis: Get batch synthesis job successfully')
                logger.debug(response.json())
                if response.json()['status'] == 'Succeeded':
                    logger.info(f'get_synthesis: Batch synthesis job succeeded, download URL: {response.json()["outputs"]["result"]}')
                    return response.json()['status'], response.json()["outputs"]["result"]
                return response.json()['status'], ""
        except Exception as e:
            logger.error(f'get_synthesis: Failed to get batch synthesis job: {e}')
  
  
def list_synthesis_jobs(skip: int = 0, top: int = 100):
    """List all batch synthesis jobs in the subscription"""
    url = f'https://{SERVICE_REGION}.{SERVICE_HOST}/api/texttospeech/3.1-preview1/batchsynthesis/talkingavatar?skip={skip}&top={top}'
    header = {
        'Ocp-Apim-Subscription-Key': SUBSCRIPTION_KEY
    }
    response = requests.get(url, headers=header)
    if response.status_code < 400:
        logger.info(f'List batch synthesis jobs successfully, got {len(response.json()["values"])} jobs')
        logger.info(response.json())
    else:
        logger.error(f'Failed to list batch synthesis jobs: {response.text}')

def download_synthesis_result(save_video_file, video_url):
    print(f"download_synthesis_result: {num}")
    retries = 3
    for i in range(retries):
        try:
            response = requests.get(video_url)
            if response.status_code < 400:
                with open(save_video_file, 'wb') as f:
                    f.write(response.content)
                logger.info(f"Synthesis result downloaded successfully: {save_video_file}")
                return save_video_file
            else:
                logger.error(f'Failed to download synthesis result: {response.text}')
        except Exception as e:
            logger.error(f'Failed to download synthesis result: {e}')
  
  
if __name__ == '__main__':
    file_list = glob.glob(f"{DOC_PATH}/*.pptx")
    for file in file_list:
        img_dir = f"{DOC_PATH}\\{os.path.basename(file).split('.')[0]}"
        img_list = glob.glob(f"{img_dir}/*.txt")
        img_list = sorted(img_list, key=lambda x: int(os.path.basename(x).split('.')[0]))
        for index, txt in enumerate(img_list):
            num = index + 1
            with open(txt, "r", encoding="utf-8") as f:
                content = f.read()
                job_id = submit_synthesis(num,content)
                if job_id is not None:
                    while True:
                        status, video_url = get_synthesis(job_id)
                        if status == 'Succeeded':
                            logger.info('batch avatar synthesis job succeeded')
                            save_video_file = f"{img_dir}\\{num}.webm"
                            download_synthesis_result(save_video_file, video_url)
                            break
                        elif status == 'Failed':
                            logger.error('batch avatar synthesis job failed')
                            break
                        else:
                            logger.info(f'batch avatar synthesis job is still running, status [{status}]')
                            time.sleep(5)
