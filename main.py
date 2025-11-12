import os
import time
import logging
import glob  # Adicionado para busca de arquivos

# Módulo de utilitários customizado
# Esta linha agora importa o PACOTE 'utils' que está na pasta raiz
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
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler("log_rpa_cadastro.log", mode='w', encoding='utf-8'),
        logging.StreamHandler()
    ]
)

# --- Constantes de XPaths ---
XPATHS = {
    # Login
    "campo_usuario": "//input[@id='username']",
    "campo_senha": "//input[@id='password']",
    "botao_entrar": "//button[contains(normalize-space(), 'Entrar')]",

    # Dashboard
    "botao_baixar_planilha": "//a[contains(normalize-space(), 'Baixar Planilha')]",
    "botao_ir_cadastro": "//a[contains(normalize-space(), 'Cadastrar')]",
    "botao_sair": "//a[contains(normalize-space(), 'Sair')]",

    # Tela de Cadastro
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
    # chrome_options.add_argument("--headless") # Executar sem abrir janela (descomentar para produção)
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

    servico = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=servico, options=chrome_options)
    logging.info("WebDriver pronto.")
    return driver


def esperar_elemento(driver, by, valor, timeout=10):
    """Função reusável para esperar por um elemento ficar clicável."""
    try:
        wait = WebDriverWait(driver, timeout)
        return wait.until(EC.element_to_be_clickable((by, valor)))
    except Exception as e:
        logging.error(f"Elemento não encontrado ou não clicável: {by}={valor} (Timeout: {timeout}s)")
        raise


def login(driver, url, usuario, senha, xpaths):
    """Acessa a URL, preenche login e senha, e clica em Entrar."""
    # Esta função agora SÓ tenta logar. A retentativa é feita fora dela.
    try:
        logging.info(f"Acessando URL: {url}")
        driver.get(url)

        logging.info("Preenchendo credenciais...")
        esperar_elemento(driver, By.XPATH, xpaths["campo_usuario"]).send_keys(usuario)
        driver.find_element(By.XPATH, xpaths["campo_senha"]).send_keys(senha)
        driver.find_element(By.XPATH, xpaths["botao_entrar"]).click()

        # Espera o botão de download aparecer para confirmar o login
        esperar_elemento(driver, By.XPATH, xpaths["botao_baixar_planilha"])
        logging.info("Login OK. Dashboard carregado.")

    except Exception as e:
        # Apenas levanta o erro para a função de retentativa capturar
        logging.error(f"Falha no login. Verifique URL, credenciais ou XPaths. Erro: {e}")
        raise


def iniciar_e_logar(diretorio_rpa, url, usuario, senha, xpaths, max_tentativas=3):
    """
    Nova função orquestradora que lida com as retentativas de login.
    Ela abre, loga, e se falhar, fecha e tenta de novo.
    """
    driver = None
    for tentativa in range(1, max_tentativas + 1):
        logging.info(f"--- Tentativa de Login [{tentativa}/{max_tentativas}] ---")
        try:
            # 1. Abre um NOVO navegador
            driver = setup_driver(diretorio_rpa)

            # 2. Tenta fazer o login
            login(driver, url, usuario, senha, xpaths)

            # 3. Sucesso! Retorna o driver pronto para uso.
            logging.info("Login realizado com sucesso.")
            return driver

        except Exception as e:
            logging.warning(f"Falha na tentativa {tentativa}: {e}")

            # 4. Falha. Garante que o navegador seja fechado antes da próxima tentativa.
            if driver:
                driver.quit()
                logging.info("Navegador fechado para retentativa.")

            if tentativa == max_tentativas:
                logging.error("Não foi possível fazer login após 3 tentativas.")
                raise Exception("Falha fatal no login após 3 tentativas.")

            time.sleep(3)  # Pausa de 3s antes de tentar de novo

    # Se o loop terminar sem sucesso (improvável, mas por segurança)
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

            # --- Lógica de Limpeza ---
            logging.warning(f"Limpando pasta de download de *.xlsx: {diretorio_download}")
            # Padrão de busca para todos os arquivos .xlsx na pasta
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
                        raise  # Erro crítico se não puder limpar a pasta

            # --- Download ---
            esperar_elemento(driver, By.XPATH, xpath_botao_baixar).click()

            # --- Espera Dinâmica ---
            timeout_download = 30  # segundos
            tempo_inicial = time.time()
            caminho_planilha_encontrado = None

            while not caminho_planilha_encontrado:
                time.sleep(1)  # Espera 1 segundo

                # Verifica se o tempo esgotou
                if time.time() - tempo_inicial > timeout_download:
                    raise Exception(f"Timeout de {timeout_download}s esperando por um novo arquivo .xlsx.")

                # Procura por QUALQUER arquivo .xlsx na pasta
                arquivos_novos = glob.glob(padrao_busca)

                if arquivos_novos:
                    # Encontrou! Pega o primeiro da lista.
                    caminho_planilha_encontrado = arquivos_novos[0]

            logging.info(f"Planilha baixada com sucesso: {caminho_planilha_encontrado}")
            return caminho_planilha_encontrado  # RETORNA O CAMINHO REAL

        except Exception as e:
            logging.warning(f"Falha na tentativa {tentativa}: {e}")
            if tentativa < max_tentativas:
                logging.warning("Atualizando a página (via URL direta) e tentando novamente...")
                driver.get("https://desafio-rpa-946177071851.us-central1.run.app/challenger/dashboard")
                time.sleep(2)

                # Espera o botão estar pronto novamente após o refresh
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

    ID_IFRAME = "registerIframe"  # ID do Iframe
    MAX_TENTATIVAS_CADASTRO = 3  # Máximo de tentativas por funcionário

    try:
        # Navega para a tela de cadastro UMA VEZ
        esperar_elemento(driver, By.XPATH, xpaths["botao_ir_cadastro"]).click()
        logging.info("Página de cadastro carregada.")

    except Exception as e:
        logging.error(f"Não foi possível acessar a página de cadastro. Abortando. Erro: {e}")
        raise

    # Itera sobre cada linha (funcionário) lido da planilha
    for i, funcionario in enumerate(dados_planilha, start=1):
        logging.info(f"--- Processando funcionário {i}/{total} ---")

        # Pega os dados do 'dict' (linha da planilha)
        nome = funcionario.get('Nome')
        sobrenome = funcionario.get('Sobrenome')
        email = funcionario.get('Email')
        cargo = funcionario.get('Cargo')
        empresa = funcionario.get('Empresa')
        endereco = funcionario.get('Endereço')
        telefone = funcionario.get('Telefone')

        if not all([nome, email, cargo]):
            logging.warning(f"Registro {i} pulado: Nome, Email ou Cargo estão faltando.")
            falhas += 1
            continue

        # --- Início do Bloco de Retentativa por Funcionário ---
        sucesso_neste_registro = False
        for tentativa in range(1, MAX_TENTATIVAS_CADASTRO + 1):
            try:
                logging.info(f"Tentativa de cadastro [{tentativa}/{MAX_TENTATIVAS_CADASTRO}] para: {nome} {sobrenome}")

                # --- 1. ENTRAR NO IFRAME ---
                # (Precisamos entrar no iframe a CADA tentativa, pois o refresh nos tira dele)
                logging.debug(f"Entrando no iframe '{ID_IFRAME}'...")
                wait = WebDriverWait(driver, 10)
                wait.until(EC.frame_to_be_available_and_switch_to_it((By.ID, ID_IFRAME)))

                # --- 2. PREENCHER O FORMULÁRIO (DENTRO DO IFRAME) ---
                esperar_elemento(driver, By.XPATH, xpaths["campo_nome"]).send_keys(str(nome or ''))
                driver.find_element(By.XPATH, xpaths["campo_sobrenome"]).send_keys(str(sobrenome or ''))
                driver.find_element(By.XPATH, xpaths["campo_email"]).send_keys(str(email or ''))
                driver.find_element(By.XPATH, xpaths["campo_cargo"]).send_keys(str(cargo or ''))
                driver.find_element(By.XPATH, xpaths["campo_empresa"]).send_keys(str(empresa or ''))
                driver.find_element(By.XPATH, xpaths["campo_endereco"]).send_keys(str(endereco or ''))
                driver.find_element(By.XPATH, xpaths["campo_telefone"]).send_keys(str(telefone or ''))

                # Clica em Cadastrar (DENTRO DO IFRAME)
                driver.find_element(By.XPATH, xpaths["botao_cadastrar_funcionario"]).click()

                time.sleep(0.5)  # Pausa breve para o JS limpar os campos

                logging.info(f"OK: Funcionário {i} (Email: {email}) cadastrado.")
                sucessos += 1
                sucesso_neste_registro = True

                # --- 3. SAIR DO IFRAME (APÓS SUCESSO) ---
                driver.switch_to.default_content()

                # Se deu certo, sai do loop de retentativas (break)
                break

            except Exception as e:
                logging.error(f"FALHA [Tentativa {tentativa}] ao cadastrar funcionário {i} (Email: {email}): {e}")

                # --- 4. SAIR DO IFRAME (APÓS FALHA, POR GARANTIA) ---
                # Esta é a principal garantia de que o refresh() funcionará
                driver.switch_to.default_content()

                if tentativa < MAX_TENTATIVAS_CADASTRO:
                    # Se não for a última tentativa, recarrega e tenta de novo
                    logging.warning("Recarregando a página de cadastro (Refresh) para tentar destravar...")
                    driver.refresh()
                    time.sleep(2)  # Pausa para o refresh
                    logging.info("Página recarregada. Próxima tentativa para o MESMO funcionário.")
                else:
                    # Se foi a última tentativa, loga o erro e NÃO incrementa 'sucessos'
                    logging.error(
                        f"FALHA PERMANENTE: Funcionário {i} (Email: {email}) falhou após {MAX_TENTATIVAS_CADASTRO} tentativas.")

        # --- Fim do Bloco de Retentativa ---

        if not sucesso_neste_registro:
            falhas += 1  # Contabiliza a falha se esgotou as tentativas

    logging.info(f"Cadastro finalizado. Sucessos: {sucessos}, Falhas: {falhas}")
    return sucessos, falhas


def logout(driver, xpaths):
    """Clica no botão Sair para encerrar a sessão."""
    try:
        logging.info("Encerrando sessão (Logout)...")
        # Garante que estamos no conteúdo principal antes de clicar em "Sair"
        driver.switch_to.default_content()

        esperar_elemento(driver, By.XPATH, xpaths["botao_sair"]).click()
        # Espera o campo de usuário aparecer para confirmar o logout
        esperar_elemento(driver, By.XPATH, xpaths["campo_usuario"])
        logging.info("Logout OK.")
    except Exception as e:
        # Não é um erro crítico se o logout falhar, apenas avisa
        logging.warning(f"Problema ao fazer logout (pode já ter deslogado): {e}")


# --- Função Principal (Orquestrador) ---

def main():
    """Função principal que orquestra a execução do RPA."""
    logging.info("--- Iniciando execução do RPA de Cadastro ---")
    driver = None
    status_sucesso = False
    erro_execucao = ""
    sucessos, falhas = 0, 0

    cfg_email = None  # Guardado para o 'finally'
    senha_email = None  # Guardado para o 'finally'

    try:
        # --- 1. Ler Configurações ---
        config = utils.ler_configuracao('config.ini')

        # Pega os dados do 'config'
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
            cfg_creds['chave_mestra'],  # Reutiliza a mesma chave mestra
            cfg_creds['salt']  # Reutiliza o mesmo salt
        )
        logging.info("Senhas carregadas na memória.")

        # --- 3. Executar RPA ---

        # O setup_driver e o login agora são feitos dentro da função com retentativa
        driver = iniciar_e_logar(
            diretorio_rpa,
            cfg_geral['url_login'],
            cfg_creds['usuario'],
            senha_app,
            XPATHS
        )

        # Se o código chegou aqui, o login foi um sucesso e 'driver' é válido.

        # Agora a função retorna o caminho do arquivo baixado
        caminho_planilha_real = baixar_planilha(driver, diretorio_rpa, XPATHS["botao_baixar_planilha"])

        # Usamos o caminho real que a função encontrou
        dados = utils.ler_planilha(caminho_planilha_real)

        sucessos, falhas = cadastrar_funcionarios(driver, dados, XPATHS)

        logout(driver, XPATHS)

        logging.info("Processo concluído com sucesso.")
        status_sucesso = True

    except Exception as e:
        logging.error(f"ERRO FATAL. Processo interrompido: {e}", exc_info=True)
        erro_execucao = str(e)
        status_sucesso = False

    finally:
        # --- 4. Encerrar Driver e Enviar Email ---
        # Este 'finally' agora pega o 'driver' mesmo se ele falhar no login,
        # pois 'driver' é definido como None no início do try.
        if driver:
            driver.quit()
            logging.info("WebDriver encerrado.")

        # Envio de email acontece mesmo se der erro (para avisar)
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