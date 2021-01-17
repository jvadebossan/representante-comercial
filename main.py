import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import pyautogui
from time import sleep as wait

login = ''
senha = ''

df = pd.read_excel("pedido.xlsx")

cnpj = input('digite o cnpj do cliente: ')
transp = input('digite a transportadora do cliente: ')

global driver
driver = webdriver.Chrome()
driver.get("https://www.pontodasfestas.com.br/PortalFestas/index.php")


def pegar_itens():
    codigos = []
    quantidades = []
    for i in range(0, len(df)):
        cods = str(df['item'] [i])
        quants = str(df['quant'] [i])
        codigos.append(str(cods))
        quantidades.append(str(quants.replace('.0', '')))

    print(f'{codigos}\n\n\n {quantidades}')

    index = 0
    while index < len(codigos):
        print('ITEM: ', index, 'ADICIONADO')
        
        elem = driver.find_element_by_id('descricaoProduto')
        elem.clear()
        elem.send_keys(codigos[index])
        wait(0.1)
        pyautogui.moveTo(140, 938, 1)
        pyautogui.click()
        wait(1)

        elem = driver.find_element_by_id('quantidadeProduto')
        elem.clear()
        elem.send_keys(quantidades[index])
        wait(1)

        elem = driver.find_element_by_id('btnAdicionarProduto')
        elem.click()
        index +=1
        wait(1)

        pyautogui.moveTo(1268, 360, 1)
        pyautogui.click()
        wait(1)
    print('todos os itens foram adicionados ;) ')

def login(cnpj, transp):
    elem = driver.find_element_by_id('txtUsuario')
    elem.send_keys(login)

    elem = driver.find_element_by_id('txtSenha')
    elem.clear()
    elem.send_keys(senha)

    elem = driver.find_element_by_id('logIn')
    elem.click()

    pyautogui.moveTo(567, 194, 20)
    pyautogui.click()

    elem = driver.find_element_by_id('cnpjCliente')
    elem.clear()
    elem.send_keys(cnpj)
    elem = driver.find_element_by_id('btnPesquisarCliente')
    elem.click()

    elem = driver.find_element_by_id('transportadora')
    elem.clear()
    elem.send_keys(transp)
    pegar_itens()

login(cnpj, transp)
