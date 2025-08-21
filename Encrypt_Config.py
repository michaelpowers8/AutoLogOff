import os
import json
from cryptography.fernet import Fernet

def get_key() -> bytes:
    if os.path.exists("secret.key"):
        with open("secret.key","rb") as file:
            return file.read()
    else:
        with open("secret.key","wb") as file:
            key:bytes = Fernet.generate_key()
            file.write(key)
            return key

def encrypt_configuration():
    key:bytes = get_key()
    fernet:Fernet = Fernet(key)
    with open("Config.encrypted","wb") as encrypted_file:
        with open("Config.json","r",encoding="utf-8") as original_file:
            configuration:dict[str,str|int|bool] = json.loads(original_file.read())
            configuration_str:str = json.dumps(configuration)
            configuration_bytes:bytes = configuration_str.encode("utf-8")
            configuration_encrypted:bytes = fernet.encrypt(configuration_bytes)
            encrypted_file.write(configuration_encrypted)

def main():
    encrypt_configuration()

if __name__ == "__main__":
    main()