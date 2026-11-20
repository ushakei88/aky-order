#!/usr/bin/env python3
"""입력 파일을 DASH_PW로 암호화/복호화 (공개 저장소에 원본을 올리지 않기 위함).
  python3 vault.py enc 원본.xlsx inputs/orders.xlsx.enc
  python3 vault.py dec inputs/orders.xlsx.enc /tmp/orders.xlsx
"""
import os, sys
from cryptography.hazmat.primitives.ciphers.aead import AESGCM
from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC
from cryptography.hazmat.primitives import hashes

def key(pw, salt):
    return PBKDF2HMAC(algorithm=hashes.SHA256(), length=32, salt=salt, iterations=200000).derive(pw.encode())

def main():
    mode, src, dst = sys.argv[1:4]
    pw = os.environ["DASH_PW"]
    data = open(src, "rb").read()
    if mode == "enc":
        salt, iv = os.urandom(16), os.urandom(12)
        out = salt + iv + AESGCM(key(pw, salt)).encrypt(iv, data, None)
    else:
        salt, iv = data[:16], data[16:28]
        out = AESGCM(key(pw, salt)).decrypt(iv, data[28:], None)
    os.makedirs(os.path.dirname(dst) or ".", exist_ok=True)
    open(dst, "wb").write(out)

if __name__ == "__main__":
    main()
