# Imports =============================================
import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.keys import Keys


# =====================================================

def buscar_contato(contato, driver):
    # Buscar contatos
    campo_pesquisa = driver.find_element_by_xpath('//div[contains(@class, "copyable-text selectable-text")]')
    time.sleep(3)
    campo_pesquisa.click()
    time.sleep(1)
    campo_pesquisa.send_keys(contato)
    time.sleep(1)
    campo_pesquisa.send_keys(Keys.ENTER)
    time.sleep(1)


def enviar_mensagem(mensagem, driver):
    # Enviar mensagem para contato


    campo_mensagem = driver.find_elements_by_xpath('//div[contains(@class, "copyable-text selectable-text")]')
    time.sleep(2)
    campo_mensagem[1].click()
    campo_mensagem[1].send_keys(mensagem)
    time.sleep(1)
    campo_mensagem[1].send_keys(Keys.ENTER)
    time.sleep(2)


def abrir_whatsapp(df_planilha):
    # Navegar até o whatsapp web
    driver = webdriver.Edge()
    driver.get('https://web.whatsapp.com')

    # Pressionar após fazer a leitura do QR Code
    print('Após fazer o login, aperta alguma tecla!')
    input()

    # Itera de acordo com o tamanho do DataFrame
    for index, contato in df_planilha.iterrows():
        # Chamada das funções e passa o contato, mensagem e driver

        buscar_contato(contato[0], driver)
        enviar_mensagem(contato[1], driver)

    driver.close()


def main():
    # Nome das colunas
    col_nome = 'Nome'
    col_mensagem = 'Mensagem'

    # Leitura da planilha
    planilha = pd.read_excel('teste.xlsx')

    # Nomeia as colunas
    planilha1 = (planilha[[col_nome, col_mensagem]])

    # Cria DataFrame e elimina Na's
    df_planilha = pd.DataFrame(planilha1)
    df_planilha = df_planilha.dropna()

    # Chamada da função de web scraping
    abrir_whatsapp(df_planilha)


main()