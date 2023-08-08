import os
import time
import shutil
import platform
import subprocess
import zipfile
import requests
import threading
import tkinter as tk
from tkinter import messagebox
from selenium import webdriver
from selenium.webdriver import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# Function to download and setup EdgeDriver
def download_Edgedriver():

    # Check if Edgedriver exists, and if not, download and setup it
    if not os.path.exists('./edgedriver'):
        try:

            download_label.config(text="Baixando driver do Edge...")  # Update the label text

            # Run the Windows command to retrieve Edge version from the Registry
            command = 'reg query HKCU\Software\Microsoft\Edge\BLBeacon /v version'
            output = subprocess.check_output(command, shell=True).decode('utf-8')

            # Process the output to extract the Edge version
            edge_version = output.split()[-1]

            # Construct the URL for downloading the Edge WebDriver
            download_url = f"https://msedgewebdriverstorage.blob.core.windows.net/edgewebdriver/{edge_version}/edgedriver_win64.zip"

            response = requests.get(download_url)
            response.raise_for_status()  # Raise an exception if the request was unsuccessful

            with open('edgedriver.zip', 'wb') as f:
                f.write(response.content)

            with zipfile.ZipFile('edgedriver.zip', 'r') as zip_ref:
                zip_ref.extractall('edgedriver')

            os.remove('edgedriver.zip')

            # Update the UI after download is complete
            root.after(0, update_ui_login)
        
        except Exception as e:

            messagebox.showerror("Falha ao baixar o driver do Edge", str(e))
        
    else:

        # Update the UI after download is complete
        root.after(0, update_ui_login)

def browsing():

    options = webdriver.EdgeOptions()
    options.add_argument('--disable-extensions')
        
    edge_driver_path = os.path.abspath('./edgedriver/msedgedriver.exe')
    driver = webdriver.Edge(executable_path=edge_driver_path, options=options)

    # Example: Navigate to a URL
    driver.get('https://www2.fiap.com.br/Aluno/BlackBoard')
    driver.minimize_window()
    update_status_label("OK")
        
    # Perform login
    update_status_label("\n Fazendo login com RM e senha fornecidos...")

    RM_field = driver.find_element(By.ID,'usuario')
    RM_field.send_keys(RM)

    senha_field = driver.find_element(By.ID,'senha')
    senha_field.send_keys(senha)

    conectar = driver.find_element(By.XPATH, '//*[@id="aluno"]/div[2]/form/div[3]/input')
    conectar.click()
    update_status_label("OK")

    update_status_label("\n Abrindo 'Sala de Aula Virtual'...")
    driver.get('https://www2.fiap.com.br/Aluno/BlackBoard')
    update_status_label("OK")

    time.sleep(1)

    curso = driver.find_element(By.XPATH, '//*[@id="page-local-salavirtual-conteudo-digital"]/section/div/div/div/div/div/div/div/div[1]/h3').text
    update_status_label(f"\n Contando os capítulos disponíveis do curso '{curso}'...")
    capitulos = driver.find_element(By.XPATH, '//*[@id="page-local-salavirtual-conteudo-digital"]/section/div/div/div/div/div/div/div/div[2]').get_attribute("childElementCount")
    update_status_label(f"{capitulos}")

    for contador in range(1,int(capitulos)+1):

        update_status_label(f"\n Acessando o {contador}º capítulo...")

        progresso = driver.find_element(By.XPATH, f'//*[@id="page-local-salavirtual-conteudo-digital"]/section/div/div/div/div/div/div/div/div[2]/div[{contador}]/div[2]/div[3]/div/span[4]')
        if progresso.get_attribute("className") != "is-active":
            linkCapitulo = driver.find_element(By.XPATH, f'//*[@id="page-local-salavirtual-conteudo-digital"]/section/div/div/div/div/div/div/div/div[2]/div[{contador}]/div[2]/div[3]')
            linkCapitulo.click()
            update_status_label(f"OK")

            update_status_label(f"\n Mantendo o documento aberto por 15 segundos [")
            time.sleep(3)
            update_status_label("|")
            time.sleep(3)
            update_status_label("|")
            time.sleep(3)
            update_status_label("|")
            time.sleep(3)
            update_status_label("|")
            time.sleep(3)
            update_status_label("|")
            update_status_label("]")
            driver.back()
        else:
            update_status_label(f"OK")


    update_status_label("\n Todos os capítulos concluídos, fechando o programa.")
    time.sleep(2)

    driver.quit()

    shutil.rmtree('./edgedriver')

    root.destroy()

def start_download_thread():
    download_thread = threading.Thread(target=download_Edgedriver)
    download_thread.start()

def start_browsing_thread():
    browsing_thread = threading.Thread(target=browsing)
    browsing_thread.start()

def update_ui_login():
    # Destroy the download UI elements
    download_label.destroy()

    # Create and display the login UI elements
    # (Add your login UI code here)
    global RM_label
    global RM_entry
    RM_label = tk.Label(root, text="RM:")
    RM_label.pack()
    RM_entry = tk.Entry(root)
    RM_entry.pack()

    global senha_label
    global senha_entry
    senha_label = tk.Label(root, text="Senha:")
    senha_label.pack()
    senha_entry = tk.Entry(root, show="*")
    senha_entry.pack()

    global continue_button
    continue_button = tk.Button(root, text="Continue", command=on_continue)
    continue_button.pack()

    global cancel_button
    cancel_button = tk.Button(root, text="Cancel", command=root.quit)
    cancel_button.pack()

def update_ui_browsing():
    # Destroy the login UI elements
    RM_label.destroy()
    RM_entry.destroy()
    senha_label.destroy()
    senha_entry.destroy()
    continue_button.destroy()
    cancel_button.destroy()

    # Create and display the browsing UI elements
    global status_label
    status_label = tk.Label(root, text="Iniciando browser em segundo plano...")
    status_label.pack()

def update_status_label(text):
    current_text = status_label.cget("text")
    new_text = current_text + f"{text}"
    status_label.config(text=new_text)

def on_continue():
    
    try:  

        global RM
        RM = RM_entry.get()

        global senha
        senha = senha_entry.get()

        # Update the UI after "continue" is pressed
        root.after(0, update_ui_browsing)

        # Start browsing thread
        start_browsing_thread()

    except Exception as e:
        messagebox.showerror("Error", str(e))

# Create the main UI window
root = tk.Tk()
root.title("Faz Nano Course Automático")

# Set the window size
window_width = 400  # specify the width in pixels
window_height = 300  # specify the height in pixels
root.geometry(f"{window_width}x{window_height}")

# Label to display download status
download_label = tk.Label(root, text="Baixando driver do Edge...")
download_label.pack()
start_download_thread()

root.mainloop()