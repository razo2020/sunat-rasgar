from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium import webdriver
from selenium.webdriver.remote.webelement import WebElement
from selenium.common.exceptions import TimeoutException, NoSuchElementException, NoAlertPresentException
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

def imprimir(driver:webdriver.Chrome, src:str, strxpath:str, atras=True):
    try:
        wait = WebDriverWait(driver,3)
        elem = wait.until(EC.presence_of_element_located((By.XPATH, strxpath)))
    except TimeoutException as ex:
        print(strxpath)
        print(ex)
        return
    src=src+elem.text.replace("/","_")+".pdf"
    elem.click()
    src = re.sub("[()]*", "", src, flags=re.IGNORECASE)
    src = src.replace(" ","_")
    print(src)
    if os.path.isfile(src):
        os.remove(src)
    original_window = driver.current_window_handle
    pyautogui.hotkey('ctrl', 'p')
    #ventana de imprimir 
    driver.switch_to.window(driver.window_handles[-1])
    time.sleep(1)
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
    if atras:
        driver.find_element(By.CLASS_NAME, 'btnNuevaConsulta').click()

# TODO Rename this here and in `imprimir`
def _extracted_from_imprimir(driver, *arg):
    result = shadowDOM(driver)
    result.Tag('print-preview-app')
    result.CSS('#sidebar')
    for a in arg:
        result.CSS(a)
    result.CSSck()

def init_url():
    url = "https://e-consultaruc.sunat.gob.pe/cl-ti-itmrconsruc/jcrS00Alias"
    #url = "https://e-consultaruc.sunat.gob.pe/cl-ti-itmrconsruc/FrameCriterioBusquedaWeb.jsp"
    chrome_service = Service(ChromeDriverManager().install())
    #chrome_service = Service(executable_path=r"C:/Users/Lenovo/Downloads/recuperar/web/chromedriver-win64/chromedriver.exe")
    options = Options()
    options.add_argument('--ignore-certificate-errors')
    options.add_argument('--ignore-ssl-errors')
    #options.add_argument('--headless=new')
    driver = webdriver.Chrome(service=chrome_service,options=options)
    #driver.implicitly_wait(15)
    driver.get(url)
    #driver.maximize_window()
    return driver

def buscar(driver:webdriver.Chrome, ruc:str, destino:str, prefijo:str):
    if str.find(destino,":") == -1 :
        actual = os.path.dirname(__file__)
        destino = os.path.join(actual,destino)
    if not os.path.exists(destino):
        os.makedirs(destino)
    n = 0
    while n < 5:
        try:
            elem = driver.find_element(By.ID,'txtRuc')
            elem.clear()
            elem.send_keys(ruc)
            driver.find_element(By.ID,'btnAceptar').click()
            break
        except NoSuchElementException as ex:
            n += 1
            print(ex)
            driver.refresh()
            time.sleep(2)

    wait = WebDriverWait(driver,2)
    try:
        if wait.until(EC.alert_is_present()):
            driver.switch_to.alert.accept()
            return
    except TimeoutException:
        print("sin alerta")
    destino = destino+prefijo+ruc+"_"
    #Consulta RUC
    imprimir(driver,destino,'/html/body/div/div[2]/div/div[1]/h1',False)
    #Información Histórica
    imprimir(driver,destino,'/html/body/div/div[2]/div/div[5]/div[1]/div[1]/form/button')
    #Cantidad de Trabajadores y/o Prestadores de Servicio
    imprimir(driver,destino,'/html/body/div/div[2]/div/div[5]/div[2]/div[1]/form/button')
    #Representante(s) Legal(es)
    imprimir(driver,destino,'/html/body/div/div[2]/div/div[5]/div[3]/div[3]/form/button')
    #Establecimiento(s) Anexo(s)
    imprimir(driver,destino,'/html/body/div/div[2]/div/div[5]/div[4]/div/form/button')

@click.command(name='print')
@click.option('--destino', '-d', default='dest\\empresa10\\', help='Destination for final pdf file')
@click.option('--ruc', '-r', default='20546202845', help='Ruc buscar en SUNAT')
@click.option('--prefijo', '-p', default='RUC_', help='inicio del nombre de documento')
@click.argument('file', type=click.Path() , default='')
def main(file, ruc, destino, prefijo):
    driver=init_url()
    if file != "":
        with open(file) as archivo:
            for linea in archivo:
                arg = linea.split("|")
                print(arg[1])
                buscar(driver, arg[0], arg[1], prefijo)
                driver.refresh()
                time.sleep(2)
    else:
        buscar(driver, ruc, destino, prefijo)
    driver.quit
    driver.close
if __name__ == "__main__":
    main() # pylint: disable=no-value-for-parameter