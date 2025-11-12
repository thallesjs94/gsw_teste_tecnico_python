# 🤖 RPA de Cadastro de Funcionários

Este projeto é um **Robô de Automação de Processos (RPA)** desenvolvido em **Python** para automatizar o cadastro de funcionários em uma plataforma web. O robô é responsável por fazer login, baixar uma planilha de funcionários, ler os dados e cadastrar cada um deles no sistema, tratando erros e reportando o status final por e-mail.

---

## 🚀 Funcionalidades Principais

- **Login Robusto:** Tenta fazer login até 3 vezes, reabrindo o navegador se necessário.  
- **Download Dinâmico:** Limpa a pasta de download, baixa o arquivo e identifica a planilha pelo seu tipo (.xlsx), independentemente do nome.  
- **Cadastro Resiliente:** Tenta cadastrar cada funcionário até 3 vezes. Se falhar, recarrega a página de cadastro e tenta novamente.  
- **Manuseio de Iframe:** Entra e sai corretamente do iframe do formulário de cadastro.  
- **Segurança:** Todas as senhas (aplicação e e-mail) são armazenadas de forma criptografada no config.ini usando a biblioteca cryptography.  
- **Relatório por E-mail:** Envia um e-mail de status ao final da execução (sucesso ou falha) com estatísticas de cadastro.  
- **Logging:** Gera um arquivo log_rpa_cadastro.log detalhado para depuração.

---

## 🧱 Estrutura do Projeto

```
/gsw_teste_tecnico_python/
├── .gitignore
├── config.ini               <-- Arquivo de configuração (PRECISA SER CRIADO)
├── gerar_senha.py           <-- Ferramenta para criptografar senhas
├── main.py                  <-- Script principal (orquestrador)
├── README.md                <-- Este arquivo
├── requirements.txt         <-- Dependências do Python
├── venv/                    <-- Ambiente virtual (ignorado pelo git)
└── utils/
    ├── __init__.py
    └── helpers.py           <-- Funções (email, excel, criptografia)
```

---

## ⚙️ Guia de Instalação e Execução

Siga estes passos para configurar e rodar o projeto.

### 1️⃣ Clonar o Repositório

```bash
git clone https://github.com/thallesjs94/gsw_teste_tecnico_python.git
cd gsw_teste_tecnico_python
```

---

### 2️⃣ Criar Ambiente Virtual (Venv)

É uma boa prática isolar as dependências do projeto.

**Windows**
```bash
python -m venv venv
.
env\Scripts ctivate
```

**macOS / Linux**
```bash
python3 -m venv venv
source venv/bin/activate
```

---

### 3️⃣ Instalar Dependências

Com o venv ativado, instale as bibliotecas necessárias:

```bash
pip install -r requirements.txt
```

---

### 4️⃣ Configurar o config.ini (Passo Crítico ⚠️)

Este é o passo mais importante.

**A) Criar o arquivo config.ini:**  
Crie um arquivo chamado `config.ini` na pasta raiz e cole o seguinte template nele. Preencha com seus dados reais.

```ini
[GERAL]
url_login = https://desafio-rpa-946177071851.us-central1.run.app
diretorio_download = C:\RPA
; A linha abaixo não é usada pelo script mais recente (que usa busca dinâmica),
; mas está aqui para referência.
nome_arquivo_planilha = funcionarios.xlsx

[CREDENCIAS_APP]
usuario = usuario_app
chave_mestra = TESTERPA
; Cole o salt gerado pelo script 'gerar_senha.py' aqui
salt = COLE_O_SALT_AQUI
; Cole a senha do APP criptografada aqui
senha_criptografada = COLE_A_SENHA_DO_APP_CRIPTOGRAFADA_AQUI

[EMAIL]
; --- Configuração para GMAIL ---
; O script usa 'starttls()', então usamos a porta 587 (TLS)
smtp_server = smtp.gmail.com
smtp_port = 587
email_remetente = seu.email@gmail.com
email_destinatario = destinatario@email.com

; IMPORTANTE: Para o Gmail, você DEVE usar uma "Senha de App".
; Cole a "Senha de App" de 16 dígitos (criptografada) aqui
email_senha_criptografada = COLE_A_SENHA_DO_APP_GMAIL_CRIPTOGRAFADA_AQUI
```

---

**B) Editar o gerar_senha.py:**  
Abra o arquivo `gerar_senha.py` e preencha as variáveis no topo:

```python
SENHA_APP = "PcoqYjedlMqDPeCimMnH"
SENHA_EMAIL = "abcdefghijklmnop"
```

---

**C) Rodar o gerar_senha.py:**  
Execute o script para gerar seus dados criptografados.

```bash
python gerar_senha.py
```

---

**D) Preencher o config.ini:**  
O script acima vai imprimir no terminal o salt, a senha_criptografada e a email_senha_criptografada. Copie e cole esses valores nos campos correspondentes do seu config.ini.

---

**E) Preencher o restante do config.ini:**  
Preencha os campos restantes, como email_remetente, email_destinatario, etc.

---

### 5️⃣ Executar o Robô

Com o venv ativado e o config.ini preenchido, execute o robô:

```bash
python main.py
```

---

## 🧰 Tecnologias Utilizadas

- 🐍 Python 3  
- 🌐 Selenium  
- 📊 OpenPyXL  
- 🔒 Cryptography  
- ✉️ smtplib  

---

## 🧑‍💻 Autor

**Thalles Jerônimo**  
💼 Desenvolvedor RPA e Python  
🔗 https://github.com/thallesjs94
