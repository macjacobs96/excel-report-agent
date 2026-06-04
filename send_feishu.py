#!/usr/bin/env python3
"""通过飞书 OPENCLAW传输助手 发送文件"""

import requests, json, sys, os
from pathlib import Path

APP_ID = "cli_aaab8dfe57391cb0"
APP_SECRET = "KAYOuqXohBsnCEU2ahcF7cdqPWk2zE5E"

def get_token():
    resp = requests.post(
        "https://open.feishu.cn/open-apis/auth/v3/tenant_access_token/internal",
        json={"app_id": APP_ID, "app_secret": APP_SECRET},
        timeout=10
    )
    data = resp.json()
    if data.get("code") != 0:
        raise Exception(f"获取token失败: {data}")
    return data["tenant_access_token"]

def upload_file(token, filepath, filename=None):
    """上传文件到飞书，返回 file_key"""
    path = Path(filepath)
    if not filename:
        filename = path.name
    
    url = "https://open.feishu.cn/open-apis/im/v1/files"
    headers = {"Authorization": f"Bearer {token}"}
    
    with open(path, 'rb') as f:
        resp = requests.post(url, headers=headers, files={
            'file': (filename, f)
        }, data={'file_type': 'stream', 'file_name': filename}, timeout=30)
    
    data = resp.json()
    if data.get("code") != 0:
        raise Exception(f"上传失败: {data}")
    
    return data["data"]["file_key"]

def send_file(token, receive_id, file_key, filename, msg_text=""):
    """发送文件消息"""
    url = "https://open.feishu.cn/open-apis/im/v1/messages"
    params = {"receive_id_type": "open_id"}
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json"
    }
    
    body = {
        "receive_id": receive_id,
        "msg_type": "file",
        "content": json.dumps({"file_key": file_key})
    }
    
    resp = requests.post(url, headers=headers, params=params, json=body, timeout=15)
    return resp.json()

def send_text(token, receive_id, text):
    """发送文本消息"""
    url = "https://open.feishu.cn/open-apis/im/v1/messages"
    params = {"receive_id_type": "open_id"}
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json"
    }
    
    body = {
        "receive_id": receive_id,
        "msg_type": "text",
        "content": json.dumps({"text": text})
    }
    
    resp = requests.post(url, headers=headers, params=params, json=body, timeout=15)
    return resp.json()

if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("用法: python3 send_feishu.py <receive_id> <filepath> [message_text]")
        print("或: python3 send_feishu.py text <receive_id> <message>")
        sys.exit(1)
    
    token = get_token()
    print(f"✅ Token: {token[:20]}...")
    
    if sys.argv[1] == "text":
        result = send_text(token, sys.argv[2], sys.argv[3])
    else:
        receive_id = sys.argv[1]
        filepath = sys.argv[2]
        msg = sys.argv[3] if len(sys.argv) > 3 else ""
        
        file_key = upload_file(token, filepath)
        print(f"✅ 文件已上传: {file_key}")
        
        result = send_file(token, receive_id, file_key, Path(filepath).name, msg)
    
    print(json.dumps(result, indent=2, ensure_ascii=False))
