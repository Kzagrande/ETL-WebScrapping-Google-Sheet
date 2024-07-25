# API GOOGLE
from __future__ import print_function

import os.path

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# If modifying these scopes, delete the file token.json.
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

# The ID and range of a sample spreadsheet.

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from functions.time_interval import get_current_and_last_hour
import re
from datetime import datetime,timedelta
from functions.slack_report import log_error_and_notify_slack
import os


options = webdriver.EdgeOptions()
options.use_chromium = True  # Usa o modo Chromium do Edge
options.add_argument("--start-maximized")  # Maximiza a janela do navegador
browser = webdriver.Edge(options=options)

# URLs
login_url = "urlexample"
wms_url = "urlexample"
dash_url = "urlexample"
sub_package_url = "urlexample"
# Dados de login
username = "xxx"
password = "xxx"

def run_script():
    def wait_for_element(by, value, timeout=10):
        return WebDriverWait(browser, timeout).until(
            EC.presence_of_element_located((by, value))
        )
    def login():
        browser.get(login_url)
        username_input = wait_for_element(By.CSS_SELECTOR, 'input[name="name"]')
        username_input.clear()
        username_input.send_keys(username)

        password_input = wait_for_element(By.CSS_SELECTOR, 'input[name="password"]')
        password_input.clear()
        password_input.send_keys(password)
        password_input.submit()
    def navigate_to_wms():
        div_wmsAncor = wait_for_element(
            By.XPATH,
            '//*[@id="container"]/section/section/main/div/div/div[2]/div/div/div/div[2]/div/div[2]/div[2]/div[1]/div',
        )
        div_wmsAncor.click()
        browser.get(wms_url)
        
    # def check_user():
    #     try:
    #         valid_user =  wait_for_element(
    #             By.XPATH,
    #             '/html/body/div[6]/div/div/div[2]/button',
    #         )
    #         valid_user.click()
    #     except:
    #         pass    

    def select_warehouse():  
        time.sleep(2)
        # check_user()
        browser.switch_to.window(browser.window_handles[0])
        warerhouse_menu = wait_for_element(
            By.XPATH,
            '//html/body/div[2]/div/div[1]/div/div/div[1]/div[1]/div[2]/div[1]/div[2]/div/div[1]/a',
        )
        warerhouse_menu.click()
        time.sleep(2)
        warerhouse_select = wait_for_element(
            By.XPATH,
            '/html/body/div[2]/div/div[1]/div/div/div[1]/div[1]/div[2]/div[1]/div[2]/div/div[2]/div/div[2]/div/a[2]',
        )
        warerhouse_select.click()
        time.sleep(2)
        btn_ok = wait_for_element(
            By.XPATH,
            "/html/body/div[8]/div/div/div[2]/button[2]",
        )
        time.sleep(2)
        btn_ok.click()
        

        #entra em subpack seta os filtros e busca o valor do througput
    def extract_througput():
        browser.get(sub_package_url)
        time.sleep(5)

        select_time = wait_for_element(By.XPATH, '/html/body/div[2]/div/div[1]/div/div/div[2]/section[3]/section/div[1]/div/div/div/div/div/form/div[16]/div/div[2]/div/div')
        select_time.click()
        time.sleep(1)

        packed_time = wait_for_element(By.XPATH, '/html/body/div[2]/div/div[1]/div/div/div[2]/section[3]/section/div[1]/div/div/div/div/div/form/div[16]/div/div[2]/div/div/div[1]/div[2]/div/div[2]/div/a[3]')
        packed_time.click()
        time.sleep(2)

        time_element = wait_for_element(By.XPATH, '/html/body/div[2]/div/div[1]/div/div/div[2]/section[3]/section/div[1]/div/div/div/div/div/form/div[16]/div/div[2]/div/label/div')
        time_element.click()

        click_calendar = wait_for_element(By.XPATH,'/html/body/div[2]/div/div[1]/div/div/div[2]/section[3]/section/div[1]/div/div/div/div/div/form/div[16]/div/div[2]/div/label/div/div[2]/div/div[1]/div[1]/div/div[3]/div[25]')
        click_calendar.click()
        actions = ActionChains(browser)
        actions.double_click(click_calendar).perform()


        first_time = wait_for_element(By.XPATH, '/html/body/div[2]/div/div[1]/div/div/div[2]/section[3]/section/div[1]/div/div/div/div/div/form/div[16]/div/div[2]/div/label/div/div[1]/div/div[2]/span[1]')
        second_time = wait_for_element(By.XPATH, '/html/body/div[2]/div/div[1]/div/div/div[2]/section[3]/section/div[1]/div/div/div/div/div/form/div[16]/div/div[2]/div/label/div/div[1]/div/div[2]/span[3]')
        first_time.click()

        hours = get_current_and_last_hour()

        browser.execute_script(f"arguments[0].textContent = '{hours['last_hour']}'",first_time)
        time.sleep(2)

        browser.execute_script(f"arguments[0].textContent = '{hours['last_second']}'",second_time)
        time.sleep(2)
        
        second_time.send_keys(Keys.ENTER)
        time.sleep(1)

        btn_search = wait_for_element(By.XPATH, '/html/body/div[2]/div/div[1]/div/div/div[2]/section[3]/section/div[1]/div/div/div/div/div/form/div[17]/div/button[1]')
        btn_search.click()
        time.sleep(1)

        througput = wait_for_element(By.XPATH,('/html/body/div[2]/div/div[1]/div/div/div[2]/section[3]/section/div[1]/div/div/section[2]/div/div[2]/div[1]/span'))
        througput_value = througput.text
        print(througput_value)
    
        througput_number = int(re.sub(r'\D', '', througput_value))

        return througput_number
    def extract_backlog():
        # operation_monitoring = wait_for_element(
        #     By.XPATH, '//*[@id="app"]/section/aside/div/div/div[2]/div/div/div/ul/li[1]/a'
        # )
        # operation_monitoring.click()

        # overseas_warehouse = wait_for_element(
        #     By.XPATH,
        #     '//*[@id="app"]/section/aside/div/div/div[2]/div/div/div/ul/li[1]/ul/li/a',
        # )
        # overseas_warehouse.click()
        # time.sleep(2)
        browser.get(dash_url)
        browser.refresh()
        time.sleep(2)
        # check_user()

        to_be_packed = wait_for_element(
            By.XPATH,'/html/body/div[2]/section[3]/section/div[1]/div/div/div[2]/div[2]/section/div[2]/div[2]/div[2]/div[7]/p[2]/span[1]'
        )
        to_be_packed_text = to_be_packed.text
        return to_be_packed_text


    def get_sheet_name():
        # Pega a data e hora atual
        now = datetime.now()

        # Verifica se está entre meia-noite e meia-noite e 59 segundos
        if now.hour == 0:
            # Calcula a data do dia anterior
            yesterday = now - timedelta(days=1)

            # Nome dos dias da semana
            week_days = {
                0: "Monday",
                1: "Tuesday",
                2: "Wednesday",
                3: "Thursday",
                4: "Friday",
                5: "Saturday",
                6: "Sunday"
            }

            # Pega o nome do dia anterior
            day_name = week_days[yesterday.weekday()]
            # Formata para o nome da planilha
            formatted_date = f"{day_name} {yesterday.strftime('%m.%d')}"
            print('É meia noite -->  ',formatted_date )
        else:
            # Nome dos dias da semana
            week_days = {
                0: "Monday",
                1: "Tuesday",
                2: "Wednesday",
                3: "Thursday",
                4: "Friday",
                5: "Saturday",
                6: "Sunday"
            }

            # Pega o nome do dia atual
            day_name = week_days[now.weekday()]

            # Formata para o nome da planilha
            formatted_date = f"{day_name} {now.strftime('%m.%d')}"
            print(' não é meia noite -->',formatted_date )
        return formatted_date


    def get_hour_column():
            # Obtenha a hora atual
            current_hour = datetime.now()
            one_hour = timedelta(hours=1)
            recurring_time = current_hour - one_hour
            recurring_hour = recurring_time.hour

            # Mapear a hora 22 para a coluna AA e hora 23 para a coluna AB
            if recurring_hour == 22:
                hour_column = 'AA'
            elif recurring_hour == 23:
                hour_column = 'AB'
            else:
                # Defina o mapeamento das horas para as colunas do Google Sheets
                # A hora 0 corresponde à coluna E, a hora 1 corresponde à coluna F e assim por diante
                initial_column = ord('E')
                hour_column = chr(initial_column + recurring_hour)

            return hour_column 
    def get_google_sheets_service():
        # Função para obter o serviço da API do Google Sheets
        creds = None

        if os.path.exists('token.json'):
            creds = Credentials.from_authorized_user_file('token.json', SCOPES)
        # If there are no (valid) credentials available, let the user log in.
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file(
                    'client_secret.json', SCOPES)
                creds = flow.run_local_server(port=0)
            # Save the credentials for the next run
            with open('token.json', 'w') as token:
                token.write(creds.to_json())

        service = build('sheets', 'v4', credentials=creds)
        return service   
    def update_sheet(service, backlog_result, througput_result):
        try:
            print(backlog_result)
            print(througput_result)
            
            # Call the Sheets API
            sheet = service.spreadsheets()
            result = sheet.values().get(spreadsheetId='sheetid',
                                        range=f"{get_sheet_name()}!B18:AB36").execute()
            values = result.get('values', [])
            
            sheet_backlog = [
                [backlog_result]
            ]
            
            sheet_througput = [
                [througput_result]
            ]
            
            result = sheet.values().update(spreadsheetId='sheetid',
                                            range=f'{get_sheet_name()}!{get_hour_column()}120', valueInputOption="USER_ENTERED", body={"values":sheet_backlog}).execute()
            
            result = sheet.values().update(spreadsheetId='sheetid',
                                            range=f'{get_sheet_name()}!{get_hour_column()}131', valueInputOption="USER_ENTERED", body={"values":sheet_througput}).execute()
        
        except HttpError as e:
            log_http_error = str(e)
            current_path = os.path.abspath(__file__)
            arquive_name = os.path.basename(current_path)

            mensagem_de_erro = (
                f"Arquivo: {arquive_name} \n\n "
                f"Hora atual: {datetime.now().time()} \n\n"
                f"Erro na atualização do Google Sheets: {log_http_error}"
            )

            log_error_and_notify_slack(mensagem_de_erro)

   
    try:      
        func_status = 'login'
        login()

        func_status = 'navigate_to_wms'
        navigate_to_wms()

        func_status = 'select_warehouse'
        select_warehouse()

        func_status = 'througput_result'
        througput_result = extract_througput()

        func_status = 'backlog_result'
        backlog_result = extract_backlog()

        browser.quit()

        service = get_google_sheets_service()

        update_sheet(service, backlog_result, througput_result)

        return True

    except Exception as e:

        log_http_error = str(e)
        current_path = os.path.abspath(__file__) 
        arquive_name = os.path.basename(current_path)

        
        mensagem_de_erro = (
                f"Arquivo: {arquive_name} \n\n "
                f"Hora atual: {datetime.now().time()} \n\n"
                f"Função: {func_status},\n\n "
                f"Erro: {log_http_error}"
            )
        log_error_and_notify_slack(mensagem_de_erro)
        return False
if __name__ == '__main__':
    # Chamamos a função principal e verificamos seu retorno
    if run_script():
        print('Código executado com sucesso.')
    else:
        print('Ocorreu um erro durante a execução.')