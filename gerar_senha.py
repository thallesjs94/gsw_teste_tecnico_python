import base64
import os
import getpass
from cryptography.fernet import Fernet
from cryptography.hazmat.primitives import hashes
from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC
from cryptography.hazmat.backends import default_backend

# A chave mestra que você forneceu
CHAVE_MESTRA = "TESTERPA"
SENHA_APP = 'SENHA_APP'
SENHA_EMAIL = 'SENHA_EMAIL'


def derivar_chave(salt, chave_mestra):
    """Deriva uma chave Fernet segura a partir da chave mestra e do salt."""
    kdf = PBKDF2HMAC(
        algorithm=hashes.SHA256(),
        length=32,
        salt=salt,
        iterations=100000,
        backend=default_backend()
    )
    # Retorna a chave no formato que o Fernet espera
    return base64.urlsafe_b64encode(kdf.derive(chave_mestra.encode()))


def criptografar_senha(chave_fernet, senha):
    """Pede uma senha e a criptografa."""
    senha_plana = senha
    f = Fernet(chave_fernet)
    token = f.encrypt(senha_plana.encode())
    return token.decode()


def main():
    print("--- Gerador de Configuração Criptografada ---")

    # 1. Gerar um Salt
    # O salt é usado para tornar a derivação da chave única.
    # Deve ser o MESMO para criptografar e descriptografar.
    salt = os.urandom(16)
    salt_b64 = base64.b64encode(salt).decode()

    print("\nPASSO 1: Copie este 'salt' para o seu config.ini [Credentials]")
    print("=" * 50)
    print(f"salt = {salt_b64}")
    print("=" * 50)

    # 2. Derivar a chave Fernet
    chave_fernet = derivar_chave(salt, CHAVE_MESTRA)

    # 3. Criptografar a senha do APP
    print("\nPASSO 2: Digite a senha da APLICAÇÃO")
    senha_app_enc = criptografar_senha(chave_fernet, SENHA_APP)
    print("\nCopie a senha da APLICAÇÃO para o config.ini [Credentials]")
    print("=" * 50)
    print(f"senha_criptografada = {senha_app_enc}")
    print("=" * 50)

    # 4. Criptografar a senha do EMAIL
    print("\nPASSO 3: Digite a senha do EMAIL (rpa@seuemail.com)...")
    senha_email_enc = criptografar_senha(chave_fernet, SENHA_EMAIL)
    print("\nCopie a senha do EMAIL para o config.ini [Email]")
    print("=" * 50)
    print(f"email_senha_criptografada = {senha_email_enc}")
    print("=" * 50)

    print("\nConcluído. Seu 'config.ini' está pronto para ser preenchido.")


if __name__ == "__main__":
    main()
