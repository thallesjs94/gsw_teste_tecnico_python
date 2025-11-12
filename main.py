import os
import time
import logging
import glob
from utils import helpers as utils

# Importações do Selenium
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

# --- Configuração do Logging ---
# Define o logging para arquivo e console
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler("log_rpa_cadastro.log", mode='w', encoding='utf-8'),
        logging.StreamHandler()
    ]
)

# --- Constantes de XPaths ---
# Centraliza todos os seletores da aplicação em um único local
XPATHS = {
    # Login
    "campo_usuario": "//input[@id='username']",
    "campo_senha": "//input[@id='password']",
    "botao_entrar": "//button[contains(normalize-space(), 'Entrar')]",

    # Dashboard
    "botao_baixar_planilha": "//a[contains(normalize-space(), 'Baixar Planilha')]",
    "botao_ir_cadastro": "//a[contains(normalize-space(), 'Cadastrar')]",
    "botao_sair": "//a[contains(normalize-space(), 'Sair')]",

    # Tela de Cadastro (usando 'label' para achar o 'id')
    "campo_nome": "//*[@id=//label[contains(normalize-space(), 'Nome')]/@for]",
    "campo_sobrenome": "//*[@id=//label[contains(normalize-space(), 'Sobrenome')]/@for]",
    "campo_email": "//*[@id=//label[contains(normalize-space(), 'Email')]/@for]",
    "campo_cargo": "//*[@id=//label[contains(normalize-space(), 'Cargo')]/@for]",
    "campo_empresa": "//*[@id=//label[contains(normalize-space(), 'Empresa')]/@for]",
    "campo_endereco": "//*[@id=//label[contains(normalize-space(), 'Endereço')]/@for]",
    "campo_telefone": "//*[@id=//label[contains(normalize-space(), 'Telefone')]/@for]",
    "botao_cadastrar_funcionario": "//button[contains(normalize-space(), 'Cadastrar Funcionário')]"
}


# --- Funções do WebDriver (Lógica de Automação) ---

def setup_driver(diretorio_download):
    """Configura e inicializa o WebDriver do Chrome."""
    logging.info(f"Configurando WebDriver. Pasta de download: {diretorio_download}")

    # Garante que o diretório de download exista
    os.makedirs(diretorio_download, exist_ok=True)

    chrome_options = Options()
    # Descomentar para rodar em background (produção)
    # chrome_options.add_argument("--headless")
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_argument("--window-size=1920,1080")

    # Configurações de download
    prefs = {
        "download.default_directory": diretorio_download,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True
    }
    chrome_options.add_experimental_option("prefs", prefs)

    # Instala ou atualiza o chromedriver automaticamente
    servico = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=servico, options=chrome_options)
    logging.info("WebDriver pronto.")
    return driver


def esperar_elemento(driver, by, valor, timeout=10):
    """Função reusável para esperar por um elemento ficar clicável."""
    try:
        wait = WebDriverWait(driver, timeout)
        # Espera o elemento ser clicável (mais robusto que apenas visível)
        return wait.until(EC.element_to_be_clickable((by, valor)))
    except Exception as e:
        logging.error(f"Elemento não encontrado ou não clicável: {by}={valor} (Timeout: {timeout}s)")
        raise


def login(driver, url, usuario, senha, xpaths):
    """Acessa a URL, preenche login e senha, e clica em Entrar."""
    # Esta função SÓ tenta logar. A lógica de retentativa fica fora.
    try:
        logging.info(f"Acessando URL: {url}")
        driver.get(url)

        logging.info("Preenchendo credenciais...")
        esperar_elemento(driver, By.XPATH, xpaths["campo_usuario"]).send_keys(usuario)
        driver.find_element(By.XPATH, xpaths["campo_senha"]).send_keys(senha)
        driver.find_element(By.XPATH, xpaths["botao_entrar"]).click()

        # Espera o botão de download (confirmação de login)
        esperar_elemento(driver, By.XPATH, xpaths["botao_baixar_planilha"])
        logging.info("Login OK. Dashboard carregado.")

    except Exception as e:
        # Apenas levanta o erro para a função 'iniciar_e_logar' capturar
        logging.error(f"Falha no login. Verifique URL, credenciais ou XPaths. Erro: {e}")
        raise


def iniciar_e_logar(diretorio_rpa, url, usuario, senha, xpaths, max_tentativas=3):
    """
    Orquestrador de login: abre, loga e, se falhar, fecha e tenta de novo.
    """
    driver = None
    for tentativa in range(1, max_tentativas + 1):
        logging.info(f"--- Tentativa de Login [{tentativa}/{max_tentativas}] ---")
        try:
            driver = setup_driver(diretorio_rpa)
            login(driver, url, usuario, senha, xpaths)
            
            # Sucesso
            logging.info("Login realizado com sucesso.")
            return driver
        
        except Exception as e:
            logging.warning(f"Falha na tentativa {tentativa}: {e}")
            
            # Garante que o driver antigo seja fechado antes de tentar de novo
            if driver:
                driver.quit()
                logging.info("Navegador fechado para retentativa.")
            
            if tentativa == max_tentativas:
                logging.error("Não foi possível fazer login após 3 tentativas.")
                raise Exception("Falha fatal no login após 3 tentativas.")
            
            time.sleep(3) # Pausa antes de tentar de novo
    
    raise Exception("Loop de login concluído sem sucesso.")


def baixar_planilha(driver, diretorio_download, xpath_botao_baixar):
    """
    Limpa a pasta, baixa a planilha e espera um novo arquivo .xlsx aparecer.
    Retorna o caminho completo do arquivo encontrado.
    """
    max_tentativas = 3
    for tentativa in range(1, max_tentativas + 1):
        try:
            logging.info(f"Tentativa de download [{tentativa}/{max_tentativas}]...")

            # Limpa a pasta de arquivos .xlsx antes de baixar
            logging.warning(f"Limpando pasta de download de *.xlsx: {diretorio_download}")
            padrao_busca = os.path.join(diretorio_download, "*.xlsx")
            arquivos_antigos = glob.glob(padrao_busca)

            if not arquivos_antigos:
                logging.info("Pasta já estava limpa.")
            else:
                for f in arquivos_antigos:
                    try:
                        os.remove(f)
                        logging.info(f"Arquivo antigo removido: {f}")
                    except OSError as e:
                        logging.error(f"Não foi possível remover o arquivo antigo {f}: {e}")
                        raise

            # Clica para baixar
            esperar_elemento(driver, By.XPATH, xpath_botao_baixar).click()

            # Espera o arquivo aparecer no disco (estratégia dinâmica)
            timeout_download = 30  # segundos
            tempo_inicial = time.time()
            caminho_planilha_encontrado = None

            while not caminho_planilha_encontrado:
                time.sleep(1) 

                if time.time() - tempo_inicial > timeout_download:
                    raise Exception(f"Timeout de {timeout_download}s esperando por um novo arquivo .xlsx.")

                # Procura por QUALQUER arquivo .xlsx na pasta
                arquivos_novos = glob.glob(padrao_busca)
                if arquivos_novos:
                    caminho_planilha_encontrado = arquivos_novos[0]

            logging.info(f"Planilha baixada com sucesso: {caminho_planilha_encontrado}")
            return caminho_planilha_encontrado

        except Exception as e:
            logging.warning(f"Falha na tentativa {tentativa}: {e}")
            if tentativa < max_tentativas:
                logging.warning("Atualizando a página (via URL direta) e tentando novamente...")
                # Usar driver.get() é mais estável que refresh()
                driver.get("https://desafio-rpa-946177071851.us-central1.run.app/challenger/dashboard")
                time.sleep(2)
                esperar_elemento(driver, By.XPATH, xpath_botao_baixar, timeout=15)
            else:
                logging.error("Downloads falharam após todas as tentativas.")
                raise Exception(f"Falha crítica no download após {max_tentativas} tentativas.")


def cadastrar_funcionarios(driver, dados_planilha, xpaths):
    """Itera sobre os dados da planilha e cadastra cada funcionário DENTRO de um iframe."""
    sucessos = 0
    falhas = 0
    total = len(dados_planilha)
    logging.info(f"Iniciando cadastro de {total} funcionários...")

    ID_IFRAME = "registerIframe"
    MAX_TENTATIVAS_CADASTRO = 3

    try:
        # Navega para a tela de cadastro (só uma vez)
        esperar_elemento(driver, By.XPATH, xpaths["botao_ir_cadastro"]).click()
        logging.info("Página de cadastro carregada.")

    except Exception as e:
        logging.error(f"Não foi possível acessar a página de cadastro. Abortando. Erro: {e}")
        raise

    # Itera sobre cada linha (funcionário) lido da planilha
    for i, funcionario in enumerate(dados_planilha, start=1):
        logging.info(f"--- Processando funcionário {i}/{total} ---")
        
        nome = funcionario.get('Nome')
        sobrenome = funcionario.get('Sobrenome')
        email = funcionario.get('Email')
        cargo = funcionario.get('Cargo')
        empresa = funcionario.get('Empresa')
        endereco = funcionario.get('Endereço')
        telefone = funcionario.get('Telefone')

        # Validação básica
        if not all([nome, email, cargo]):
            logging.warning(f"Registro {i} pulado: Nome, Email ou Cargo estão faltando.")
            falhas += 1
            continue

        # --- Bloco de Retentativa por Funcionário ---
        sucesso_neste_registro = False
        for tentativa in range(1, MAX_TENTATIVAS_CADASTRO + 1):
            try:
                logging.info(f"Tentativa de cadastro [{tentativa}/{MAX_TENTATIVAS_CADASTRO}] para: {nome} {sobrenome}")
                
                # 1. ENTRAR NO IFRAME
                # (Precisamos entrar a CADA tentativa, pois o refresh nos tira dele)
                logging.debug(f"Entrando no iframe '{ID_IFRAME}'...")
                wait = WebDriverWait(driver, 10)
                wait.until(EC.frame_to_be_available_and_switch_to_it((By.ID, ID_IFRAME)))

                # 2. PREENCHER O FORMULÁRIO (DENTRO DO IFRAME)
                # (str() garante que valores None ou numéricos sejam enviados)
                esperar_elemento(driver, By.XPATH, xpaths["campo_nome"]).send_keys(str(nome or ''))
                driver.find_element(By.XPATH, xpaths["campo_sobrenome"]).send_keys(str(sobrenome or ''))
                driver.find_element(By.XPATH, xpaths["campo_email"]).send_keys(str(email or ''))
                driver.find_element(By.XPATH, xpaths["campo_cargo"]).send_keys(str(cargo or ''))
                driver.find_element(By.XPATH, xpaths["campo_empresa"]).send_keys(str(empresa or ''))
                driver.find_element(By.XPATH, xpaths["campo_endereco"]).send_keys(str(endereco or ''))
                driver.find_element(By.XPATH, xpaths["campo_telefone"]).send_keys(str(telefone or ''))

                driver.find_element(By.XPATH, xpaths["botao_cadastrar_funcionario"]).click()

                time.sleep(0.5)  # Pausa breve para o JS limpar os campos

                logging.info(f"OK: Funcionário {i} (Email: {email}) cadastrado.")
                sucessos += 1
                sucesso_neste_registro = True

                # 3. SAIR DO IFRAME (APÓS SUCESSO)
                driver.switch_to.default_content()
                
                # Se deu certo, quebra o loop de retentativas
                break 

            except Exception as e:
                logging.error(f"FALHA [Tentativa {tentativa}] ao cadastrar funcionário {i} (Email: {email}): {e}")

                # 4. SAIR DO IFRAME (APÓS FALHA)
                # Garantia de que o refresh() funcionará no contexto certo
                driver.switch_to.default_content()

                if tentativa < MAX_TENTATIVAS_CADASTRO:
                    logging.warning("Recarregando a página de cadastro (Refresh) para tentar destravar...")
                    driver.refresh()
                    time.sleep(2) # Pausa para o refresh
                    logging.info("Página recarregada. Próxima tentativa para o MESMO funcionário.")
                else:
                    logging.error(f"FALHA PERMANENTE: Funcionário {i} (Email: {email}) falhou após {MAX_TENTATIVAS_CADASTRO} tentativas.")
        
        # --- Fim do Bloco de Retentativa ---
        
        if not sucesso_neste_registro:
            falhas += 1 # Contabiliza a falha se esgotou as tentativas

    logging.info(f"Cadastro finalizado. Sucessos: {sucessos}, Falhas: {falhas}")
    return sucessos, falhas


def logout(driver, xpaths):
    """Clica no botão Sair para encerrar a sessão."""
    try:
        logging.info("Encerrando sessão (Logout)...")
        # Garante que estamos no conteúdo principal
        driver.switch_to.default_content()

        esperar_elemento(driver, By.XPATH, xpaths["botao_sair"]).click()
        # Espera o campo de usuário (tela de login)
        esperar_elemento(driver, By.XPATH, xpaths["campo_usuario"])
        logging.info("Logout OK.")
    except Exception as e:
        logging.warning(f"Problema ao fazer logout (pode já ter deslogado): {e}")


# --- Função Principal (Orquestrador) ---

def main():
    """Função principal que orquestra a execução do RPA."""
    logging.info("--- Iniciando execução do RPA de Cadastro ---")
    driver = None
    status_sucesso = False
    erro_execucao = ""
    sucessos, falhas = 0, 0

    # Guarda a config de email para usar no 'finally'
    cfg_email = None
    senha_email = None

    try:
        # --- 1. Ler Configurações ---
        config = utils.ler_configuracao('config.ini')

        cfg_geral = config['GERAL']
        cfg_creds = config['CREDENCIAS_APP']
        cfg_email = config['EMAIL']

        diretorio_rpa = cfg_geral.get('diretorio_download', 'C:\\RPA')

        # --- 2. Descriptografar Senhas ---
        logging.info("Lendo credenciais criptografadas...")
        senha_app = utils.descriptografar(
            cfg_creds['senha_criptografada'],
            cfg_creds['chave_mestra'],
            cfg_creds['salt']
        )
        senha_email = utils.descriptografar(
            cfg_email['email_senha_criptografada'],
            cfg_creds['chave_mestra'],
            cfg_creds['salt']
        )
        logging.info("Senhas carregadas na memória.")

        # --- 3. Executar RPA ---
        
        # Login (com retentativas)
        driver = iniciar_e_logar(
            diretorio_rpa,
            cfg_geral['url_login'],
            cfg_creds['usuario'],
            senha_app,
            XPATHS
        )

        # Download (com retentativas)
        caminho_planilha_real = baixar_planilha(driver, diretorio_rpa, XPATHS["botao_baixar_planilha"])

        # Leitura da planilha
        dados = utils.ler_planilha(caminho_planilha_real)

        # Cadastro (com retentativas por registro)
        sucessos, falhas = cadastrar_funcionarios(driver, dados, XPATHS)

        # Logout
        logout(driver, XPATHS)

        logging.info("Processo concluído com sucesso.")
        status_sucesso = True

    except Exception as e:
        logging.error(f"ERRO FATAL. Processo interrompido: {e}", exc_info=True)
        erro_execucao = str(e)
        status_sucesso = False

    finally:
        # --- 4. Encerrar Driver e Enviar Email ---
        if driver:
            driver.quit()
            logging.info("WebDriver encerrado.")

        # Envia email de status (mesmo se der erro)
        if cfg_email and senha_email:
            logging.info("Preparando email de status...")
            utils.enviar_email_status(
                cfg_email,
                senha_email,
                status_sucesso,
                erro_execucao,
                sucessos,
                falhas
            )
        elif cfg_email:
            logging.error("Não foi possível enviar email. Configurações de email ou senha não foram carregadas.")

        logging.info("--- Execução terminada ---")


if __name__ == "__main__":
    main()
