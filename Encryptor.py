from cryptography.fernet import Fernet

def encrypt_ini_file(file_path, key):
    with open(file_path, "rb") as file:
        file_data = file.read()
    fernet = Fernet(key)
    encrypted_data = fernet.encrypt(file_data)
    with open(file_path, "wb") as file:
        file.write(encrypted_data)

def main():
    ini_file_path = "params.ini"
    encryption_key = Fernet.generate_key()
    encrypt_ini_file(ini_file_path, encryption_key)
    print(encryption_key)

if __name__ == "__main__":
    main()
