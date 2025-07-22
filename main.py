from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.remote.webdriver import WebDriver
import time
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
import os
import pickle
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pyautogui
from selenium.common.exceptions import TimeoutException
import psutil

def encerrar_instancias_selenium():
    nomes_alvo = ['chromedriver', 'chrome']

    for proc in psutil.process_iter(['pid', 'name']):
        try:
            nome = proc.info['name'].lower()
            if any(alvo in nome for alvo in nomes_alvo):
                proc.kill()
                print(f"🛑 Processo encerrado: {nome} (PID {proc.pid})")
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            continue

def tirarprint():
    options = Options()
    options.add_argument("--disable-gpu")
    options.add_argument("--window-size=1920,1080")
    options.add_argument("--window-position=-32000,-32000")
    navegador = webdriver.Chrome(options=options)
    
    navegador.get("https://facto.conveniar.com.br/Fundacao/Forms/Convenio/PedidoEmpenhoCompra.aspx?CodUsuarioGestor=-1&TipoRelatorioSolicitacao=4&CodTipoPedido=16")
    navegador.find_element(By.CSS_SELECTOR,"#ctl00_ContentPlaceHolder1_ObjWucLoginCaptcha_lgUsuario_UserName").send_keys("davi.torres")
    navegador.find_element(By.CSS_SELECTOR,"#ctl00_ContentPlaceHolder1_ObjWucLoginCaptcha_lgUsuario_Password").send_keys("D1a2v3i4@")
    navegador.find_element(By.CSS_SELECTOR,"#ctl00_ContentPlaceHolder1_ObjWucLoginCaptcha_lgUsuario_btnLogin").click()
    time.sleep(2)
    print = navegador.find_element(By.CSS_SELECTOR,"#ctl00_ContentPlaceHolder1_gvPrincipal > tbody")
    print.screenshot("C:\\Users\\DaviTorresdeSousa\\Desktop\\Arquivos\\save.png")
    time.sleep(5)
    encerrar_instancias_selenium()

# URL do canal no Microsoft Teams
def enviar_teams():
    URL = "https://teams.microsoft.com/l/channel/19%3AmFbZOE7QMqj9--D0uOJ_iKxwxRJqwrVMB20tcdWWXOQ1%40thread.tacv2/teste?groupId=860ba162-eaed-45ad-a81d-4fe7c70f7082&tenantId=1332aea6-3164-47c2-a975-1ebfafd8655e"

    # Caminho para salvar os cookies
    COOKIES_FILE = "cookies_sharepoint.pkl"

    # Caminho do arquivo a ser enviado
    ARQUIVO_PATH = "C:\\Users\\DaviTorresdeSousa\\Desktop\\Arquivos\\save.png"

    # Configuração do navegador
    options = Options()

    # Modo headless moderno do Chrome
    options.add_argument("--disable-gpu")  # Recomendado junto ao headless
    options.add_argument("--window-size=1920,1080")  # Garante que o layout renderize corretamente
    options.add_argument("--window-position=-32000,-32000")
    options.add_argument("--disable-popup-blocking")
    options.add_argument("--disable-notifications")
    options.add_argument("user-data-dir=C:/Selenium/Perfil")  # Ainda pode usar seu perfil


    driver = webdriver.Chrome(options=options)
    driver.get(URL)

    # Carregar cookies, se existirem
    import builtins
    print = builtins.print  # Garante que print volte ao normal

    if os.path.exists(COOKIES_FILE):
        print("➡️ Carregando cookies salvos...")

        driver.get("https://teams.microsoft.com")  # Corrige domínio para aceitar cookies

        with open(COOKIES_FILE, "rb") as f:
            cookies = pickle.load(f)

        for cookie in cookies:
            try:
                driver.add_cookie(cookie)
            except Exception as e:
                print("Erro ao adicionar cookie:", e)

        driver.get(URL)
        
        # Ativa a janela
        janela = pyautogui.getWindowsWithTitle("Google Chrome")
        if janela:
            janela[0].activate()
        
        pyautogui.press('enter')
        # Aceita cookies se botão estiver presente
        try:
            WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div/div/div[2]/div/button[2]"))
            ).click()
        except:
            pass  # botão pode já não existir
    else:
        print("Primeira vez: faça login manualmente no navegador.")
        input("Após o login completo no Teams, pressione ENTER aqui...")

        print("💾 Salvando cookies...")
        with open(COOKIES_FILE, "wb") as f:
            pickle.dump(driver.get_cookies(), f)

        print("✅ Cookies salvos para uso futuro.")

    # Hora atual
    current_time = datetime.now().strftime("%H:%M")
    input_box = WebDriverWait(driver, 20).until(
        EC.presence_of_element_located((By.XPATH, "/html/body/div[1]/div/div/div/div[6]/div/div/div/div/div[2]/div[1]"))
    )
    time.sleep(3)
    try:
        
        WebDriverWait(driver, 7).until(
            EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div/div/div/div[9]/div/div[1]/div/div[3]/div/div/div/div/button"))
        ).click()
        
        WebDriverWait(driver, 7).until(
            EC.element_to_be_clickable((By.XPATH, "/html/body/div[9]/div[2]/div/div[2]/button[2]"))
        ).click()
        
       
        input_box = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.XPATH, "/html/body/div[1]/div/div/div/div[9]/div/div[1]/div/div[3]/div/div/div/div/div/div/div[2]/span/input"))
        )
        input_box.send_keys(current_time)

       
        clip_button = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.XPATH, "/html/body/div[1]/div/div/div/div[9]/div/div[1]/div/div[3]/div/div/div/div/div/div/div[5]/div/div[1]/div[1]/button[2]"))
        )
        driver.execute_script("arguments[0].click();", clip_button)

        
        opcao_upload = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.XPATH, "/html/body/div[10]/div/div[2]/div/div/div[2]/div/div/div[1]/ul/li[1]/div/div/div/div"))
        )
        driver.execute_script("arguments[0].click();", opcao_upload)

        
        input_file = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.XPATH, "//input[@type='file']"))
        )
        input_file.send_keys(ARQUIVO_PATH)

        
        WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div/div/div/div[9]/div/div[1]/div/div[3]/div/div[2]/div/div/div/div/div[6]/div/div[2]/div[2]"))
        ).click()
        time.sleep(5)
        driver.quit()
        print("✅ Mensagem e imagem enviadas com sucesso.")

        
    except TimeoutException:
        print("ℹ️ Botão não encontrado. Seguindo...")
        
        WebDriverWait(driver, 30).until(
            EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div/div/div/div[9]/div/div[1]/div/div[3]/div/button"))
        ).click()

        
        
        WebDriverWait(driver, 30).until(
            EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div/div/div/div[9]/div/div[1]/div/div[3]/div/div/div/div/div/div/div[2]/span/input"))
        ).send_keys(current_time)
        
        clip_button = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.XPATH, "/html/body/div[1]/div/div/div/div[9]/div/div[1]/div/div[3]/div/div/div/div/div/div/div[5]/div/div[1]/div[1]/button[2]"))
        )
        driver.execute_script("arguments[0].click();", clip_button)

        
        opcao_upload = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.XPATH, "/html/body/div[10]/div/div[2]/div/div/div[2]/div/div/div[1]/ul/li[1]/div/div/div/div"))
        )
        driver.execute_script("arguments[0].click();", opcao_upload)

        
        input_file = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.XPATH, "//input[@type='file']"))
        )
        input_file.send_keys(ARQUIVO_PATH)

        
        WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div/div/div/div[9]/div/div[1]/div/div[3]/div/div[2]/div/div/div/div/div[6]/div/div[2]/div[2]"))
        ).click()
        time.sleep(5)
        driver.quit()
        print("✅ Mensagem e imagem enviadas com sucesso.")
    

    
import threading



def tarefa_unica():
    tirarprint()
    enviar_teams()

def aguardar_e_executar(horas):
    """
    horas: lista de strings 'HH:MM' que representam os horários para rodar a tarefa
    """
    executados_hoje = set()  # para evitar executar repetido no mesmo minuto, só uma vez por horário no dia

    while True:
        agora = datetime.now().strftime("%H:%M")
        hoje = datetime.now().date()

        # Para reiniciar o set ao trocar o dia
        if hasattr(aguardar_e_executar, "dia_executado"):
            if aguardar_e_executar.dia_executado != hoje:
                executados_hoje.clear()
                aguardar_e_executar.dia_executado = hoje
        else:
            aguardar_e_executar.dia_executado = hoje

        if agora in horas and agora not in executados_hoje:
            print(f"⏳ Hora {agora} chegou! Executando tarefa...")
            thread = threading.Thread(target=tarefa_unica)
            thread.start()
            thread.join()
            print("✅ Tarefa concluída. Aguardando próximo horário...")

            executados_hoje.add(agora)
            time.sleep(60)  # espera 1 minuto para não executar duas vezes no mesmo minuto
        else:
            time.sleep(20)

if __name__ == "__main__":
    # Corrigindo a lista para elementos separados
    horarios_para_executar = ["12:50", "13:00", "13:10", "13:15", "13:20", "13:30"]
    aguardar_e_executar(horarios_para_executar)