import os
import base64
import logging
import configparser
import smtplib
import pandas as pd
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

# Importações de Criptografia
from cryptography.fernet import Fernet
from cryptography.hazmat.primitives import hashes
from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC
from cryptography.hazmat.backends import default_backend


def ler_configuracao(arquivo_config='config.ini'):
    """Lê o arquivo .ini e retorna um objeto de configuração."""
    try:
        config = configparser.ConfigParser()
        # Lê o arquivo de configuração (garante UTF-8 para não ter erro de acento)
        config.read(arquivo_config, encoding='utf-8')
        if not config.sections():
            logging.error(f"Arquivo 'config.ini' não encontrado ou está vazio.")
            raise FileNotFoundError("Arquivo 'config.ini' não encontrado ou vazio.")
        logging.info("Arquivo 'config.ini' lido com sucesso.")
        return config
    except Exception as e:
        logging.error(f"Erro ao ler 'config.ini': {e}")
        raise


def derivar_chave_fernet(salt, chave_mestra):
    """Deriva uma chave de criptografia segura a partir da chave mestra e do salt."""
    kdf = PBKDF2HMAC(
        algorithm=hashes.SHA256(),
        length=32,
        salt=salt,
        iterations=100000,
        backend=default_backend()
    )
    # Retorna a chave formatada para o Fernet
    return base64.urlsafe_b64encode(kdf.derive(chave_mestra.encode()))


def descriptografar(senha_criptografada, chave_mestra, salt_b64):
    """Descriptografa uma string usando a chave mestra e o salt."""
    try:
        salt = base64.urlsafe_b64decode(salt_b64)
        chave_fernet = derivar_chave_fernet(salt, chave_mestra)
        f = Fernet(chave_fernet)

        # Descriptografa e decodifica de bytes para string
        senha_bytes = f.decrypt(senha_criptografada.encode('utf-8'))
        return senha_bytes.decode('utf-8')
    except Exception as e:
        logging.error(f"Falha na descriptografia. Verifique 'chave_mestra' e 'salt' no config.ini. Erro: {e}")
        raise Exception("Falha na descriptografia. Verifique 'config.ini'.") from e


def ler_planilha(caminho_arquivo):
    """Lê a planilha Excel/XLSX e retorna uma lista de dicionários (um por linha)."""
    try:
        logging.info(f"Lendo planilha: {caminho_arquivo}")
        # Converte a planilha para um DataFrame do Pandas
        df = pd.read_excel(caminho_arquivo, engine='openpyxl')

        # Converte o DataFrame para uma lista de dicionários
        # 'records' -> [{col1: val1, col2: val2}, {col1: val3, col2: val4}]
        dados = df.to_dict('records')

        if not dados:
            logging.warning("Planilha lida, mas está vazia.")

        logging.info(f"Planilha lida. {len(dados)} registros encontrados.")
        return dados
    except FileNotFoundError:
        logging.error(f"Arquivo da planilha não encontrado em: {caminho_arquivo}")
        raise
    except Exception as e:
        logging.error(f"Erro inesperado ao ler o arquivo Excel: {e}")
        raise


def enviar_email_status(cfg_email, senha_email, sucesso=True, mensagem_erro="", registros_sucesso=0, registros_falha=0):
    """Envia um email de status (sucesso ou falha) da execução do RPA."""

    # Pega os dados do objeto de config
    remetente = cfg_email.get('email_remetente')
    destinatario = cfg_email.get('email_destinatario')
    servidor_smtp = cfg_email.get('servidor_smtp')
    porta_smtp = int(cfg_email.get('porta_smtp', 587))  # Porta 587 (TLS) como padrão

    if not all([remetente, destinatario, servidor_smtp]):
        logging.warning(
            "Email não enviado. 'remetente', 'destinatario' ou 'servidor_smtp' não configurados no config.ini.")
        return

    logging.info(f"Montando email de status para: {destinatario}")

    # --- Monta o Corpo do Email ---
    if sucesso:
        assunto = "Sucesso: RPA de Cadastro de Funcionários"
        corpo = f"""
        Olá,

        O RPA de Cadastro de Funcionários foi executado com sucesso.

        Resumo da Execução:
        - Registros cadastrados: {registros_sucesso}
        - Registros com falha: {registros_falha}
        - Total processado: {registros_sucesso + registros_falha}

        Atenciosamente,
        RPA Bot
        """
    else:
        assunto = "FALHA: RPA de Cadastro de Funcionários"
        corpo = f"""
        Olá,

        O RPA de Cadastro de Funcionários falhou durante a execução.

        Resumo da Execução:
        - Registros cadastrados antes da falha: {registros_sucesso}
        - Registros com falha: {registros_falha}

        Erro principal:
        {mensagem_erro}

        Por favor, verifique o log 'log_rpa_cadastro.log' na pasta do robô para mais detalhes.

        Atenciosamente,
        RPA Bot
        """

    try:
        msg = MIMEMultipart()
        msg['From'] = remetente
        msg['To'] = destinatario
        msg['Subject'] = assunto
        msg.attach(MIMEText(corpo, 'plain'))

        # --- Bloco de Envio (Descomentar para ativar) ---
        """
        logging.info(f"Conectando ao servidor SMTP: {servidor_smtp}:{porta_smtp}")
        server = smtplib.SMTP(servidor_smtp, porta_smtp)
        server.starttls() # Inicia conexão segura
        logging.info("Fazendo login no email...")
        server.login(remetente, senha_email)
        logging.info("Enviando email...")
        server.sendmail(remetente, destinatario, msg.as_string())
        server.quit()
        logging.info("Email de status enviado com sucesso.")
        """

        # Simulação (remover quando o bloco acima for ativado)
        logging.info("Envio de email (simulado). Para ativar, descomente o bloco em utils/helpers.py.")

    except Exception as e:
        logging.error(f"Erro ao tentar enviar o email: {e}")