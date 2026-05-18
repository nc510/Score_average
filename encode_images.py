import base64
import os

def encode_image_to_base64(image_path):
    with open(image_path, 'rb') as f:
        return base64.b64encode(f.read()).decode('utf-8')

wechat_base64 = encode_image_to_base64('加好友.png')
donation_base64 = encode_image_to_base64('打赏.png')

output = f'''# -*- coding: utf-8 -*-
# 自动生成的文件，包含二维码图片的Base64编码

WECHAT_QR_CODE = """{wechat_base64}"""

DONATION_QR_CODE = """{donation_base64}"""
'''

with open('qr_codes.py', 'w', encoding='utf-8') as f:
    f.write(output)

print("已生成 qr_codes.py 文件")
print(f"微信二维码 Base64 长度: {len(wechat_base64)}")
print(f"打赏二维码 Base64 长度: {len(donation_base64)}")
