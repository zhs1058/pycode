import base64
from Cryptodome.Cipher import AES
from Cryptodome.Util.Padding import pad, unpad
from Cryptodome.Random import get_random_bytes

def encrypt_image_base64(image_path, key):
    # 读取图片并进行Base64编码
    with open(image_path, 'rb') as f:
        image_data = f.read()
        base64_encoded = base64.b64encode(image_data)

    # 使用AES加密
    cipher = AES.new(key, AES.MODE_CBC)
    encrypted_image = cipher.encrypt(pad(base64_encoded, AES.block_size))

    # 返回加密后的Base64字符串
    return base64.b64encode(encrypted_image)

def decrypt_image_base64(encrypted_base64, key):
    # 解码Base64并使用AES解密
    encrypted_image = base64.b64decode(encrypted_base64)
    cipher = AES.new(key, AES.MODE_CBC)
    decrypted_image = unpad(cipher.decrypt(encrypted_image), AES.block_size)

    # 返回解密后的Base64字符串
    return base64.b64decode(decrypted_image)

# 生成随机密钥
key = get_random_bytes(16)

# 加密图片
encrypted_base64 = encrypt_image_base64('original_image.jpg', key)

# 解密图片
decrypted_base64 = decrypt_image_base64(encrypted_base64, key)

# 保存解密后的图片
with open('decrypted_image.jpg', 'wb') as f:
    f.write(decrypted_base64)
