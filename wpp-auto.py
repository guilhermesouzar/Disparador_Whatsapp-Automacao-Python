import pandas as pd
import pyautogui as pg
import pyperclip
import time
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
import tkinter as tk
from tkinter import filedialog

# Função para ler a planilha e retornar um DataFrame com os dados
def ler_planilha(caminho_planilha):
    try:
        df = pd.read_excel(caminho_planilha)
        required_columns = ['celularrespfinanceiro', 'celularbackup', 'localarquivo'] 
        for col in required_columns:
            if col not in df.columns:
                raise ValueError(f"A coluna '{col}' não foi encontrada na planilha.")
        return df
    except Exception as e:
        print(f"[ERRO] Não foi possível acessar a planilha: {e}")
        return None

def selecionar_planilha():
    root = tk.Tk()
    root.withdraw()
    print("Selecione a planilha:")
    root.call('wm', 'attributes', '.', '-topmost', True)
    caminho_planilha = filedialog.askopenfilename(
        title="Selecione a planilha",
        filetypes=[("Arquivos Excel", "*.xlsx;*.xls")]
    )
    return caminho_planilha

# Função para obter o telefone com base na escolha do usuário
def obter_telefone(row, escolha):
    if escolha == 1:
        telefone = row['celularrespfinanceiro']
        if pd.isna(telefone):
            telefone = row['celularbackup']
    if pd.isna(telefone):
        return None
    return telefone

# Função para escolher o caminho da mídia
def escolher_midia():
    root = tk.Tk()
    root.withdraw()
    image_path = filedialog.askopenfilename(
        title="Selecione a mídia",
        filetypes=[("Arquivos de Imagem", "*.jpg;*.jpeg;*.png;*.gif;*.bmp;*.mp4;*.mkv;*.avi")]
    )
    if not image_path:
        print("Nenhuma mídia selecionada. Encerrando o processo.")
        exit()
    return image_path

print("Aguarde a inicialização completa do navegador.\n")

service = Service(ChromeDriverManager().install())
driver = webdriver.Chrome(service=service)

driver.get('https://web.whatsapp.com/')
time.sleep(25)

escolha = int(input("Whatsapp Web foi aberto no navegador?\n1 - Sim | 2 - Não\n"))
if escolha not in [1, 2]:
    print("Escolha inválida. Encerrando o processo.")
    exit()
elif escolha == 1:
    excel_file = selecionar_planilha()
    if not excel_file:
        print("Nenhuma planilha selecionada. Encerrando o processo.")
        exit()
    df = ler_planilha(excel_file)
    if df is None:
        print("Erro ao ler a planilha. Encerrando o processo.")
        exit()
elif escolha == 2:
    print("Erro ao abrir whatsappweb. Encerrando o processo.")
    time.sleep(3)
    exit()

# Solicita ao usuário a escolha do telefone
escolha_telefone = int(input("Escolha a forma de obter o telefone:\n1 - Celular Responsável na planilha\nDigite 1: "))
if escolha_telefone not in [1]:
    print("Escolha inválida. Encerrando o processo.")
    exit()

# Solicita a mensagem do usuário
message = input("Digite a mensagem que deseja enviar: ")

# Escolhe a mídia
escolha_midia_opcao = int(input("Escolha a forma de obter a mídia:\n1 - Selecionar uma mídia salva no computador\n2 - Usar o caminho da mídia na planilha\nDigite sua escolha (1 ou 2): "))
image_path = None

if escolha_midia_opcao == 1:
    image_path = escolher_midia() 
elif escolha_midia_opcao == 2:
    print("A mídia será obtida da planilha durante o envio.")
else:
    print("Escolha inválida. Encerrando o processo.")
    exit()

escolha_tipomidia_opcao = int(input("O formato da mídia é:\n1 - Vídeo/Imagem\n2 - Documento\nDigite sua escolha (1 ou 2): "))
tipomidia = None
if escolha_tipomidia_opcao == 1:
    tipomidia = 1
elif escolha_tipomidia_opcao == 2:
    tipomidia = 0
else:
    print("Escolha inválida. Encerrando o processo.")
    exit()

input("✅ Escaneie o código e com a conta do Whatsapp aberta em seu navegador pressione Enter para continuar iniciar o envio das mensagens.")

for index, row in df.iterrows():
    phone = obter_telefone(row, escolha_telefone)
    if phone is None:
        print(f"⚠️ Não há telefone disponível para a linha {index}.")
        continue

    backup_phone = row['celularbackup'] 

    try:
        print(f"📲 Enviando mensagem para {phone}")

        # pesquisa o Contato
        def buscar_contato(phone):
            driver.find_element(By.XPATH, '//p[contains(@class,"selectable-text copyable-text")]').send_keys(Keys.ESCAPE)
            driver.find_element(By.CSS_SELECTOR, "span[data-icon='new-chat-outline']").click()
            campo_pesquisa = driver.find_element(By.XPATH, '//p[contains(@class,"selectable-text copyable-text")]')
            time.sleep(2)
            campo_pesquisa.click()
            campo_pesquisa.send_keys(phone)
            time.sleep(3)
            campo_pesquisa.send_keys(Keys.ENTER)
            time.sleep(2)
            campo_mensagem = driver.find_elements(By.XPATH, '//p[contains(@class,"selectable-text copyable-text")]')
            if len(campo_mensagem) == 3:
                print(f"⚠️ Não foi possível encontrar o contato: {phone}")
                return len(campo_mensagem)
            time.sleep(2)

        # envia a mensagem
        def enviar_mensagem(message):
            time.sleep(2)
            campo_mensagem = driver.find_elements(By.XPATH, '//p[contains(@class,"selectable-text copyable-text")]')
            if len(campo_mensagem) > 1:
                campo_mensagem[1].click()
                time.sleep(1)
                campo_mensagem[1].send_keys(message)
                campo_mensagem[1].send_keys(Keys.ENTER)
            else:
                print(f"⚠️ Não foi possível encontrar o campo de mensagem para {phone}.")

        # envia mídia como mensagem
        def enviar_midia(image_path):
            if image_path:  
                time.sleep(2)
                driver.find_element(By.CSS_SELECTOR, "span[data-icon='plus']").click()
                time.sleep(2)
                attach = driver.find_elements(By.CSS_SELECTOR, "input[type='file']")
                attach[escolha_tipomidia_opcao].send_keys(image_path)
                time.sleep(3)
                send_button = driver.find_element(By.CSS_SELECTOR, "span[data-icon='send']")
                send_button.click()

        num_mensagens = buscar_contato(phone)
        if num_mensagens == 3:  
            print(f"⚠️ Não foi possível encontrar o contato")
            continue  # Continua para a próxima iteração do loop
        time.sleep(2)

        enviar_mensagem(message) 

        # Verifica se a escolha foi 2 e obtém o caminho da mídia da coluna 'localarquivo'
        if escolha_midia_opcao == 2:
            image_path = row['localarquivo']  # Obtém o caminho da coluna 'localarquivo'
            if pd.isna(image_path):
                print(f"⚠️ Não há caminho de mídia disponível na linha {index}.")
                continue  # Se não houver caminho, pula para a próxima linha

        if image_path:  # Verifica se há uma mídia para enviar
            enviar_midia(image_path)  # Envia a mídia

    except Exception as e:
        print(f"❌ Falha ao enviar mensagem para {phone}. Erro: {e}")
        
        # Envia mensagem de erro para o celular de backup
        if not pd.isna(backup_phone):
            try:
                print(f"📲 Enviando mensagem de erro para {backup_phone}")
                buscar_contato(backup_phone)
                error_message = f"Erro ao enviar mensagem para {phone}: {e}"
                enviar_mensagem(error_message)  # Envia a mensagem de erro
            except Exception as backup_error:
                print(f"❌ Falha ao enviar mensagem de erro para BACKUP {backup_phone}. Erro: {backup_error}")

print("🏁 Todas as mensagens foram processadas!")