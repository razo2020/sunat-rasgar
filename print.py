from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium import webdriver
from selenium.webdriver.remote.webelement import WebElement
from selenium.common.exceptions import TimeoutException
from webdriver_manager.chrome import ChromeDriverManager
import re
import time
import os
import pyautogui
from pywinauto.application import Application
from pywinauto.backend import element_class

import click
class shadowDOM :
    def __init__(self, driver:webdriver.Chrome) :
        self.driver = driver
        self.root:WebElement
    
    #función recursiva
    def CSS(self, css:str):
        self.root = self.driver.execute_async_script("""
            const traverse = e => {
                let el
                if(el = e.querySelector('""" + css + """')){
                    arguments[0](el)
                }
                [...e.querySelectorAll('*')].filter(e => e.shadowRoot).map(e => traverse(e.shadowRoot))
            }
            [...document.querySelectorAll('*')].filter(e => e.shadowRoot).map(e => traverse(e.shadowRoot))
            arguments[0](null)""")

    def Tag(self, element:str):
        self.root = self.driver.execute_script('return arguments[0].shadowRoot',self.driver.find_element(By.TAG_NAME, element))

    #dando click al Element
    def CSSck(self):
        self.root.click()
        text = self.root.text
        print(f"{text} has echo click")

def imprimir(driver:webdriver.Chrome, src:str):
    src = re.sub("[()]*", "", src, flags=re.IGNORECASE)
    src = src.replace(" ","_")
    print(src)
    if os.path.isfile(src):
        os.remove(src)
    #ventana de imprimir 
    original_window = driver.current_window_handle
    pyautogui.hotkey('ctrl', 'p')
    time.sleep(1)
    driver.switch_to.window(driver.window_handles[-1])

    _extracted_from_imprimir(
        driver
        , '#destinationSettings'
        , '#destinationSelect'
        , 'print-preview-settings-section:nth-child(9) > div > select > option:nth-child(3)[value="Save as PDF/local/"]'
    )
    time.sleep(1)
    _extracted_from_imprimir(
        driver
        , 'print-preview-button-strip'
        , 'div > cr-button.action-button'
    )
    time.sleep(1)

    app = Application().connect(class_name="#32770", title='Guardar como')
    SaveAsWin = app.window()
    SaveAsWin.Edit.type_keys(src)
    app.Guardar.Button1.click()
    driver.switch_to.window(original_window)


# TODO Rename this here and in `imprimir`
def _extracted_from_imprimir(driver, *arg):
    result = shadowDOM(driver)
    result.Tag('print-preview-app')
    result.CSS('#sidebar')
    for a in arg:
        result.CSS(a)
    result.CSSck()

def buscar(ruc, destino, prefijo):
    url = "https://e-consultaruc.sunat.gob.pe/cl-ti-itmrconsruc/FrameCriterioBusquedaWeb.jsp"
    chrome_service = Service(ChromeDriverManager().install())
    #chrome_service = Service(executable_path=r"C:/Users/Lenovo/Downloads/recuperar/web/chromedriver-win64/chromedriver.exe")
    options = webdriver.ChromeOptions()
    options.add_argument('--ignore-certificate-errors')
    options.add_argument('--ignore-ssl-errors')
    #options.add_argument('--headless=new')
    driver = webdriver.Chrome(service=chrome_service,options=options)
    driver.implicitly_wait(20)
    driver.get(url)
    #driver.maximize_window()
    
    if str.find(destino,":") == -1 :
        actual = os.path.dirname(__file__)
        destino = os.path.join(actual,destino)
    if not os.path.exists(destino):
        os.makedirs(destino)
    
    elem = driver.find_element(By.ID, 'txtRuc')
    elem.clear()
    elem.send_keys(ruc)
    wiat = WebDriverWait(driver, 25)
    elem = wiat.until(EC.presence_of_element_located((By.ID, 'btnAceptar')))
    elem.click()
    destino = destino+prefijo+ruc+"_"
    
    try:
        #Consulta RUC
        elem = wiat.until(EC.presence_of_element_located((By.XPATH, '/html/body/div/div[2]/div/div[1]/h1')))
        imprimir(driver,destino+elem.text+".pdf")

        #Información Histórica
        elem = wiat.until(EC.presence_of_element_located((By.XPATH, '/html/body/div/div[2]/div/div[5]/div[1]/div[1]/form/button')))
        src=destino+elem.text+".pdf"
        elem.click()
        imprimir(driver,src)
        elem = wiat.until(EC.presence_of_element_located((By.CLASS_NAME, 'btnNuevaConsulta')))
        elem.click()

        #Cantidad de Trabajadores y/o Prestadores de Servicio
        elem = wiat.until(EC.presence_of_element_located((By.XPATH, '/html/body/div/div[2]/div/div[5]/div[2]/div[1]/form/button')))
        src=destino+elem.text.replace("/","_")+".pdf"
        elem.click()
        imprimir(driver,src)
        elem = wiat.until(EC.presence_of_element_located((By.CLASS_NAME, 'btnNuevaConsulta')))
        elem.click()

        #Representante(s) Legal(es)
        elem = wiat.until(EC.presence_of_element_located((By.XPATH, '/html/body/div/div[2]/div/div[5]/div[3]/div[3]/form/button')))
        src=destino+elem.text+".pdf"
        elem.click()
        imprimir(driver,src)
        elem = wiat.until(EC.presence_of_element_located((By.CLASS_NAME, 'btnNuevaConsulta')))
        elem.click()

        #Establecimiento(s) Anexo(s)
        elem = wiat.until(EC.presence_of_element_located((By.XPATH, '/html/body/div/div[2]/div/div[5]/div[4]/div/form/button')))
        src=destino+elem.text+".pdf"
        elem.click()
        imprimir(driver,src)
        elem = wiat.until(EC.presence_of_element_located((By.CLASS_NAME, 'btnNuevaConsulta')))
        elem.click()

    except TimeoutException as ex:
        driver.quit
        driver.close

@click.command(name='print')
@click.option('--destino', '-d', default='dest\\empresa10\\', help='Destination for final pdf file')
@click.option('--ruc', '-r', default='20546202845', help='Ruc buscar en SUNAT')
@click.option('--prefijo', '-p', default='RUC_', help='inicio del nombre de documento')
@click.argument('file', type=click.Path() , default='')
def main(file, ruc, destino, prefijo):
    if file != "":
        with open(file) as archivo:
            for linea in archivo:
                arg = linea.split("|")
                print(arg[1])
                buscar(arg[0],arg[1], prefijo)
    else:
        buscar(ruc, destino, prefijo)

if __name__ == "__main__":
    main() # pylint: disable=no-value-for-parameter