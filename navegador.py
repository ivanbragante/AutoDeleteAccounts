import time
import subprocess
import os
import re
import pyautogui
import pyperclip
import json
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.action_chains import ActionChains
from datetime import datetime

#Este arquivo é responsável por abrir o navegador e fazer a análise dos emails

driver = webdriver.Chrome()
emailMeu = ''

# Função para carregar o email salvo
def carregar_email():
    global emailMeu
    if os.path.exists("config.json"):
        with open("config.json", "r") as f:
            data = json.load(f)
            emailMeu = data.get("email")
    return emailMeu

gmailMeu = carregar_email()

def abrir_nav():
    driver.get("https://siteparaanalisar.com")
    print("Site aberto")
    wait = WebDriverWait(driver, 10) #esperar 10 segundos pelo elemento
    time.sleep(1)


def logar():
    global gmailMeu
    print(f'Email do analista {gmailMeu}')
    driver.find_element('xpath', '//*[@id="identifierId"]').send_keys(gmailMeu) #Email para logar no site do chrome
    driver.find_element('xpath', '//*[@id="identifierNext"]/div/button/span').click() #clica em next
    time.sleep(5)
    driver.find_element('xpath', '//*[@id="initEmail"]').send_keys(gmailMeu) #Email para entrar no site da pwc
    time.sleep(0.5)
    driver.find_element('xpath', '/html/body/app-root/div/app-login-form/div[1]/div[2]/div/div/div[2]/form/div[1]/div/div[2]/button').click() # clica em next para logar
    time.sleep(15) 
    


def analisar_email(gmail): #gmail do user a ser analisado
    wait = WebDriverWait(driver, 20)
    verifData = None

    def escreve_email(gmail): #gmail
        time.sleep(20)   
        elemento = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="mat-select-value-3"]/span')))
        driver.execute_script("arguments[0].scrollIntoView();", elemento)
        ActionChains(driver).move_to_element(elemento).perform()
        time.sleep(0.5)
        wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="mat-select-value-3"]/span'))).click() #Clica no Column
        time.sleep(0.5)
        wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="mat-option-2"]/span'))).click() #seleciona email
        
        elemento = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="mat-select-value-5"]/span'))) 
        driver.execute_script("arguments[0].scrollIntoView();", elemento)
        ActionChains(driver).move_to_element(elemento).click().perform() # Clica no operators
        
        time.sleep(0.5)
        wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="mat-option-21"]/span'))).click() # Seleciona contain
        time.sleep(2)
        
        escreveEmail = '//*[@id="mat-input-0"]'
        wait.until(EC.presence_of_element_located((By.XPATH, escreveEmail)))
        driver.find_element('xpath', escreveEmail).send_keys(gmail)
        time.sleep(0.5)
    
        # Clicar no botão de pesquisar
        elemento = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="cdk-accordion-child-0"]/div/div/div[2]/div/button/mat-icon')))
        driver.execute_script("arguments[0].scrollIntoView();", elemento)
        ActionChains(driver).move_to_element(elemento).click().perform()
        time.sleep(8)

    def verifica_dataR(): #verifica se é data reviewed
        # Abaixo será um If onde ele vai parar o código e vai escrever o elemento.text do Xpath na planilha / Ex: BR/DO_NOT_DELETE/AGTech - Se legal hold e tal
        # Verificar se é Data Reviewd
        seDR = '/html/body/app-root/app-layout/div/mat-sidenav-container/mat-sidenav-content/main/app-drive-accounts/div/section/table[1]/tbody/tr/td[2]/span'
        elemento = wait.until(EC.element_to_be_clickable((By.XPATH, seDR)))
        valor_capturado = elemento.text
        if valor_capturado == '--':
            return False
        else:
            print(f"É: {valor_capturado}")
            return True
        time.sleep(1)
        driver.execute_script("arguments[0].scrollIntoView();", elemento) #faz scroll até o elemento
    
    def status_conta(): #Verifica o status da conta
        # Xpath da verificação de Legal Hold ou outros = /html/body/app-root/app-layout/div/mat-sidenav-container/mat-sidenav-content/main/app-drive-accounts/div/section/table[1]/tbody/tr/td[5]
        # Validar o status da conta
        statusConta = '/html/body/app-root/app-layout/div/mat-sidenav-container/mat-sidenav-content/main/app-drive-accounts/div/section/table[1]/tbody/tr/td[5]'
        elemento = wait.until(EC.presence_of_element_located((By.XPATH, statusConta)))
        valor_capturado = elemento.text
        print('O status da conta é: ' + valor_capturado)
        if valor_capturado in ['/BR/Generic Mailbox','/BR/VIALTO/Generic Mailbox', '/BR/DO_NOT_DELETE/AGTech']:
            return valor_capturado
        elif valor_capturado in ['/BR', '/BR/Administrators']:
            print('Status - OK')
        else:
            return 'Analisar'

    def verifica_driveF():  # Verifica se tem drive files
        seDriveF = '/html/body/app-root/app-layout/div/mat-sidenav-container/mat-sidenav-content/main/app-drive-accounts/div/section/table[1]/tbody/tr/td[7]'
        elemento = wait.until(EC.presence_of_element_located((By.XPATH, seDriveF)))
        valor_capturado = elemento.text
        if valor_capturado == '0':
            print('Não tem drive files')
            verifica_df = False  
        else:
            print('Tem drive files')
            verifica_df = True  
        return verifica_df

    def verifica_sharedD(): #Esta função tem que retornar um True, ai vai ser colocado na planilha para validação de analise de shared drive
         #Validar se tem drive files
        seSD = '/html/body/app-root/app-layout/div/mat-sidenav-container/mat-sidenav-content/main/app-drive-accounts/div/section/table[1]/tbody/tr/td[9]'
        elemento = driver.find_element('xpath' , seSD)
        valor_capturado = elemento.text
        if valor_capturado == '0':
            print('Não tem shared drive')
            return False
        else:
            print('Tem shared drive')
            return True

    def verificaDataArquivo(): #Se esta retornar True, precisa escrever na planilha que tem arquivos com menos de 1 ano
        #Ir no xpath para verificar se o ultimo arquivo foi modificado a menos de um ano
        seData = '/html/body/app-root/app-layout/div/mat-sidenav-container/mat-sidenav-content/main/app-drive-tab/div/div[4]/app-my-drive/section/table[1]/tbody/tr[1]/td[4]'
        time.sleep(5)
        # elemento = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, seData)))
        elemento = driver.find_element('xpath', seData)
        # elemento = wait.until(EC.presence_of_element_located((By.XPATH, seData)))
        data_cap = elemento.text

        data_formatada = datetime.strptime(data_cap, "%d %b %Y")
        # Obter a data atual
        data_atual = datetime.now()
        # Calcular a diferença em anos
        diferenca = data_atual - data_formatada
        # Verificar se a diferença é inferior a 1 ano (365 dias)
        if diferenca.days < 365:
            print("Tem arquivos e foram modificados a MENOS de um ano.")
            return True
        else:
            print("Tem arquivos mas foram modificados a MAIS de um ano.") #Se estiver mais de um ano, voltar uma página e selecionar data reviewd, na planilha colocar
            return False

    def verifica_GoogleS():
        #Verificar se tem google sites
        time.sleep(5)
        driver.execute_script("window.scrollBy(0, 1000);")
        driver.find_element('xpath', '//*[@id="mat-select-value-11"]/span').click()#clica em column
        time.sleep(0.3)
        driver.find_element('xpath', '//*[@id="mat-option-26"]/span').click()#clica em type
        time.sleep(0.5)
        driver.find_element('xpath', '//*[@id="mat-select-value-13"]/span').click()#clica em operator
        time.sleep(0.3)
        driver.find_element('xpath', '//*[@id="mat-option-62"]/span').click()#clica em contain
        time.sleep(0.3)
        driver.find_element('xpath', '//*[@id="mat-input-1"]').send_keys('application/vnd.google-apps.site') #coloca o valor pra pesquisa do google site
        time.sleep(0.3)
        # driver.find_element('xpath', '//*[@id="cdk-accordion-child-5"]/div/div/div[2]/button[2]/span[1]/mat-icon').click() #clica em pesquisar
        driver.find_element('xpath', '//*[@id="cdk-accordion-child-5"]/div/div/div[2]/div/button/mat-icon').click() #clica em pesquisar

        time.sleep(10.5)

        seGS = '/html/body/app-root/app-layout/div/mat-sidenav-container/mat-sidenav-content/main/app-drive-tab/div/div[4]/app-my-drive/section/table[1]/thead/tr[1]/th/div/div[1]/b'
        elemento = driver.find_element('xpath', seGS)
        # elemento = wait.until(EC.presence_of_element_located((By.XPATH, seGS)))
        verifica_gs = elemento.text
        numero = int(re.search(r'\((\d+)\)', verifica_gs).group(1))
        # Verificar se o número é diferente de zero
        if numero != 0:
            print("Tem Google Site.")
            return True
        else:
            print("Não tem Google Site.")
            return False

    #Funções ativando
    escreve_email(gmail)

    if verifica_dataR():
        return "Data Reviewed"
    if status_conta(): #Se for DoNotDelete stop, Generic MailBox stop, etc...
         return status_conta()

    verifica_sd = verifica_sharedD()
        # return "Tem shared drive"
    verifica_df = verifica_driveF() #3 Verifica se tem drive files, retorna na variavel se True or False

    #Desce até a conta do google e clica
    time.sleep(1)
    gConta = wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/app-root/app-layout/div/mat-sidenav-container/mat-sidenav-content/main/app-drive-accounts/div/section/table[1]/tbody/tr/td[3]/a'))) #Coloquei wait until ******************

    # Faz scroll até o elemento aparecer
    driver.execute_script("arguments[0].scrollIntoView();", gConta)
    time.sleep(2)
    gConta.click()  # Clica no elemento após o scroll
    time.sleep(8)

    if verifica_df:
        verifData = verificaDataArquivo()
    
    verifica_gs = verifica_GoogleS()
    
    print(f'verifica_df: {verifica_df}')
    print(f'verifica_sd: {verifica_sd}')
    print(f'verifData: {verifData}')
    print(f'verifica_gs: {verifica_gs}')

    def marcar_DR():#função que vai marcar como DR, se cumprir todos os requisitos de analise
        if (verifica_df == False and verifData == None and verifica_sd == False and verifica_gs == False) or (verifica_df == True and verifData == False and verifica_sd == False and verifica_gs == False):
            driver.get("https://pg-gx-p-app-362124.appspot.com/drive/accounts")
            time.sleep(8)
            escreve_email(gmail)
            time.sleep(1)
            elemento = wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/app-root/app-layout/div/mat-sidenav-container/mat-sidenav-content/main/app-drive-accounts/div/section/table[1]/tbody/tr/td[3]/a')))
            # Faz scroll até o elemento aparecer
            driver.execute_script("arguments[0].scrollIntoView();", elemento)
            time.sleep(5)
            # Seleciona a conta
            driver.find_element('xpath', '//*[@id="mat-mdc-checkbox-52-input"]').click()
            time.sleep(1.5)
            # Clica em "Mark Data Reviewed"
            driver.find_element('xpath', '/html/body/app-root/app-layout/div/mat-sidenav-container/mat-sidenav-content/main/app-drive-accounts/div/section/table[1]/thead/tr[1]/th/div/div[5]/button/span[2]').click()
            time.sleep(3)
            # Tenta clicar no botão "Yes" e verifica se ele existe
            try:
                driver.find_element('xpath', '//*[@id="mat-mdc-dialog-0"]/div/div/app-confirm-dialog/mat-dialog-actions/button[2]/span[2]').click()
                time.sleep(10)
                return True
            except NoSuchElementException:
                return 'Botão Data Reviewed desabilitado'
        else:
            return False


    if marcar_DR():
            return 'Marcado como DR no Google'

    if verifica_df: 
        if verifica_df and verifData and verifica_sd and verifica_gs:
            return 'Tem SD, arquivos com menos de 1 ano e google sites' 
        elif verifica_df and verifData and verifica_sd:
            return 'Tem SD e arquivos a menos de 1 ano'
        elif verifica_df and verifData and verifica_gs:
            return 'Tem GSites e arquivos a menos de 1 ano'
        elif verifica_sd and verifica_gs:
            return 'Tem SD e tem GSites'
        elif verifica_df and verifData:
            return 'Arquivo com menos de um ano'
        elif verifica_sd:
            return 'Tem shared drive'
        elif verifica_gs:
            return 'Tem google site'

    return False

    
    #função que leva pra home page e da um refresh na página uma vez a mais para evitar problemas de loading
def home_page():
    driver.get("https://siteparaanalisar.com")
    time.sleep(3)
    for i in range (1):
        driver.refresh()
        time.sleep(3)
        print('Dando refresh')