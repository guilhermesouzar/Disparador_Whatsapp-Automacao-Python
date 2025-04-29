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
from tkinter import filedialog, messagebox
from tkinter import ttk
import threading

# Variáveis globais
excel_file = None
message = ""
escolha_telefone = 1
escolha_midia_opcao = 1
escolha_tipomidia_opcao = 1
image_path = None
df = None
driver = None

# Função para ler a planilha
def ler_planilha(caminho_planilha):
    try:
        df = pd.read_excel(caminho_planilha)
        required_columns = ['celularrespfinanceiro', 'celularbackup', 'caminhoarquivo']
        for col in required_columns:
            if col not in df.columns:
                raise ValueError(f"A coluna '{col}' não foi encontrada na planilha.")
        return df
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao acessar a planilha: {e}")
        return None

# Função para obter o telefone baseado na escolha do usuário
def obter_telefone(row, escolha):
    telefone = row['celularrespfinanceiro'] if escolha == 1 else row['celularbackup']
    if pd.isna(telefone):
        return None
    return telefone

# Funções Selenium para WhatsApp

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
        return len(campo_mensagem)

def enviar_mensagem(message):
    campo_mensagem = driver.find_elements(By.XPATH, '//p[contains(@class,"selectable-text copyable-text")]')
    if len(campo_mensagem) > 1:
        campo_mensagem[1].click()
        time.sleep(1)
        linhas = message.splitlines()

        for linha in linhas:
            campo_mensagem[1].send_keys(linha)
            campo_mensagem[1].send_keys(Keys.SHIFT, Keys.ENTER)  # Insere quebra de linha

        campo_mensagem[1].send_keys(Keys.ENTER)

def enviar_midia(image_path):
    if image_path:
        time.sleep(2)
        driver.find_element(By.CSS_SELECTOR, "span[data-icon='plus']").click()
        time.sleep(2)
        attach = driver.find_elements(By.CSS_SELECTOR, "input[type='file']")
        attach[escolha_tipomidia_opcao - 1].send_keys(image_path)
        time.sleep(4)
        send_button = driver.find_element(By.CSS_SELECTOR, "span[data-icon='send']")
        send_button.click()


# Função principal de envio
def enviar():
    global df, image_path

    if not excel_file:
        messagebox.showwarning("Aviso", "Selecione uma planilha antes de continuar.")
        return

    df = ler_planilha(excel_file)
    if df is None:
        return

    for index, row in df.iterrows():
        phone = obter_telefone(row, escolha_telefone)
        if phone is None:
            continue

        try:
            print(f"Enviando mensagem para {phone}")
            num_mensagens = buscar_contato(phone)
            if num_mensagens == 3:
                continue

            time.sleep(2)
            enviar_mensagem(message)

            if escolha_midia_opcao == 2:
                image_path = row['caminhoarquivo']
                if pd.isna(image_path):
                    continue

            if image_path:
                enviar_midia(image_path)

        except Exception as e:
            print(f"Erro ao enviar mensagem para {phone}: {e}")

    messagebox.showinfo("Finalizado", "Todas as mensagens foram processadas!")

# Funções da interface UI
def escolher_planilha():
    global excel_file
    excel_file = filedialog.askopenfilename(title="Selecione a planilha", filetypes=[("Excel files", "*.xlsx;*.xls")])
    if excel_file:
        lbl_planilha.config(text=f"Selecionado: {excel_file.split('/')[-1]}")

def escolher_midia_manual():
    global image_path
    image_path = filedialog.askopenfilename(title="Selecione a mídia", filetypes=[("Arquivos", "*.jpg;*.png;*.gif;*.bmp;*.mp4;*.mkv;*.avi")])
    if image_path:
        lbl_midia.config(text=f"Mídia selecionada: {image_path.split('/')[-1]}")

def iniciar_envio_thread():
    threading.Thread(target=enviar).start()

# Inicia WebDriver
service = Service(ChromeDriverManager().install())
driver = webdriver.Chrome(service=service)
driver.get("https://web.whatsapp.com")

# Interface Gráfica
app = tk.Tk()
app.title("Envio de Mensagens WhatsApp")
app.geometry("900x500")

frame = tk.Frame(app)
frame.pack(pady=20)

btn_planilha = tk.Button(frame, text="Selecionar Planilha", command=escolher_planilha)
btn_planilha.pack(pady=5)

lbl_planilha = tk.Label(frame, text="Nenhuma planilha selecionada")
lbl_planilha.pack(pady=5)

lbl_msg = tk.Label(frame, text="Mensagem a ser enviada:")
lbl_msg.pack(pady=5)

entry_msg = tk.Text(frame, height=5, width=40)
entry_msg.pack(pady=5, padx=10)

lbl_tipo_tel = tk.Label(frame, text="Tipo de telefone:")
lbl_tipo_tel.pack(pady=5)

tipo_tel = ttk.Combobox(frame, values=["Responsável Financeiro"], state="readonly")
tipo_tel.current(0)
tipo_tel.pack(pady=5)

lbl_midia_opcao = tk.Label(frame, text="Fonte da Mídia:")
lbl_midia_opcao.pack(pady=5)

combo_midia_opcao = ttk.Combobox(frame, values=["Selecionar arquivo manualmente", "Usar caminho da planilha"], state="readonly")
combo_midia_opcao.current(0)
combo_midia_opcao.pack(pady=5)

btn_midia = tk.Button(frame, text="Selecionar Mídia", command=escolher_midia_manual)
btn_midia.pack(pady=5)

lbl_midia = tk.Label(frame, text="Nenhuma mídia selecionada")
lbl_midia.pack(pady=5)

lbl_tipo_midia = tk.Label(frame, text="Tipo de envio da mídia:")
lbl_tipo_midia.pack(pady=5)

combo_tipo_midia = ttk.Combobox(frame, values=["Documento", "Imagem/Vídeo"], state="readonly")
combo_tipo_midia.current(0)
combo_tipo_midia.pack(pady=5)

btn_enviar = tk.Button(app, text="Iniciar Envio", command=lambda: set_variaveis_e_iniciar(entry_msg.get("1.0", tk.END).strip(), tipo_tel.current(), combo_midia_opcao.current(), combo_tipo_midia.current()))
btn_enviar.pack(pady=20)

def set_variaveis_e_iniciar(msg, tel_opcao, midia_opcao, tipo_midia):
    global message, escolha_telefone, escolha_midia_opcao, escolha_tipomidia_opcao
    message = msg
    escolha_telefone = tel_opcao + 1
    escolha_midia_opcao = midia_opcao + 1
    escolha_tipomidia_opcao = tipo_midia + 1
    iniciar_envio_thread()

app.mainloop()
driver.quit()
