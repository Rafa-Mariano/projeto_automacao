# Criar um navegador
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
import pandas as pd

servico = Service(ChromeDriverManager().install())
navegador = webdriver.Chrome(service=servico)

# Importando arquivo tabela produtos
tabela_produtos = pd.read_excel("buscas.xlsx")
print(tabela_produtos)


import time

def verificar_tem_termos_banidos(lista_termos_banidos, nome):
# analisar se ele não tem nenhum termo banido
    tem_termos_banidos = False
    for palavra in lista_termos_banidos:
        if palavra in nome:
            tem_termos_banidos = True
    return tem_termos_banidos

def verificar_tem_todos_termos_produto(listas_termos_nome_produto, nome):
# analisar se ele tem TODOS os produtos do nome do produto
    tem_todos_termos_produtos = True
    for palavra in listas_termos_nome_produto:
        if palavra not in nome:
            tem_todos_termos_produtos = False
    return tem_todos_termos_produtos

def busca_google_shopping(navegador, produto, termos_banidos, preco_min, preco_max):
    
    produto = produto.casefold()
    termos_banidos = termos_banidos.casefold()
    lista_termos_banidos = termos_banidos.split(" ")
    listas_termos_nome_produto = produto.split(" ")
    lista_ofertas = []
    preco_min = float(preco_min)
    preco_max = float(preco_max)

    # entrar no google
    navegador.get("https://www.google.com.br/?hl=pt-BR")
    navegador.find_element('xpath' , '//*[@id="APjFqb"]').send_keys(produto)
    navegador.find_element('xpath' , '//*[@id="APjFqb"]').send_keys(Keys.ENTER)
    # entrar na aba shopping
    # PRIMEIRA FORMA DE CLICAR NA ABA SHOPPING
    #navegador.find_element('xpath' , '//*[@id="hdtb-msb"]/div[1]/div/div[2]/a').click()

    # SEGUNDA FORMA DE CLICAR NA ABA SHOPPING
    elementos = navegador.find_elements('class name' , "hdtb-mitem")
    for item in elementos:
        if "Shopping" in item.text:
            item.click()
            break

    # pegar as informações do produto 
    lista_resultados = navegador.find_elements('class name' , 'i0X6df')

    for resultado in lista_resultados:

        nome = resultado.find_element('class name' , 'tAxDx').text
        nome = nome.casefold()

        # analisar se ele não tem nenhum termo banido
        tem_termos_banidos = verificar_tem_termos_banidos(lista_termos_banidos, nome)
        # analisar se ele tem TODOS os produtos do nome do produto
        tem_todos_termos_produtos = verificar_tem_todos_termos_produto(listas_termos_nome_produto, nome)
        # selecionar só os elementos que tem_termos_banidos = Fale e ao mesmo tempo tem_todos_termos_produtos = True
        try:
            if not tem_termos_banidos and tem_todos_termos_produtos:
                preco = resultado.find_element('class name' , 'a8Pemb').text
                preco = preco.replace("R$", "").replace(" ", "").replace(".","").replace("," , ".")
                preco = float(preco)
                # se o preco ta entre o preco_min e preco_max
                if preco_min <= preco <= preco_max:
                    elemento_referencia = resultado.find_element('class name', 'bONr3b')
                    elemento_pai = elemento_referencia.find_element('xpath' , '..')
                    link = elemento_pai.get_attribute('href')
                    lista_ofertas.append((nome,preco,link))
        except:
            continue

    return lista_ofertas


def busca_buscape(navegador, produto, termos_banidos, preco_min, preco_max):

    # tratar os valores
    produto = produto.casefold()
    termos_banidos = termos_banidos.casefold()
    lista_termos_banidos = termos_banidos.split(" ")
    listas_termos_nome_produto = produto.split(" ")
    lista_ofertas = []
    preco_min = float(preco_min)
    preco_max = float(preco_max)
    # buscar no buscape
    navegador.get("https://www.buscape.com.br/")
    navegador.find_element('xpath' , '//*[@id="new-header"]/div[1]/div/div/div[3]/div/div/div[2]/div/div[1]/input').send_keys(produto)
    navegador.find_element('xpath' , '//*[@id="new-header"]/div[1]/div/div/div[3]/div/div/div[2]/div/div[1]/input').send_keys(Keys.ENTER)
    # pegar os resultados
    while len(navegador.find_elements('class name' , 'Select_Select__1S7HV')) < 1: #codigo para verificar se na pagina contém item que comprove que esta na pagina certa
        time.sleep(3)
    lista_resultados = navegador.find_elements('class name' , 'ProductCard_ProductCard_Inner__tsD4M')

    for resultado in lista_resultados:
        preco = resultado.find_element('class name', 'Text_MobileHeadingS__Zxam2').text
        nome = resultado.find_element('class name' , 'ProductCard_ProductCard_Name__LT7hv').text
        nome = nome.casefold()
        link = resultado.get_attribute('href')

    # analista se o resultado tem termo banido e tem todos os termos do produtos
        tem_termos_banidos = verificar_tem_termos_banidos(lista_termos_banidos, nome)
        tem_todos_termos_produtos = verificar_tem_todos_termos_produto(listas_termos_nome_produto, nome)

    # analisar se o preco está entre o preco minimo e preco maximo
        if not tem_termos_banidos and tem_todos_termos_produtos:
            preco = preco.replace("R$", "").replace(" ", "").replace(".","").replace("," , ".").replace("impostos" , "").replace("+", "")
            preco = float(preco)
            if preco_min <= preco <= preco_max:
                lista_ofertas.append((nome, preco, link))
    # retornar alista de ofertas do buscape

    return lista_ofertas


tabela_ofertas = pd.DataFrame()

for linha in tabela_produtos.index:
    # pesquisar sobre o produto
    produto = tabela_produtos.loc[linha, 'Nome']
    termos_banidos = tabela_produtos.loc[linha, 'Termos banidos']
    preco_min = tabela_produtos.loc[linha, 'Preço mínimo']
    preco_max = tabela_produtos.loc[linha, 'Preço máximo']

    lista_ofertas_google_shopping = busca_google_shopping(navegador, produto, termos_banidos, preco_min, preco_max)
    if lista_ofertas_google_shopping:
        tabela_google_shopping = pd.DataFrame(lista_ofertas_google_shopping, columns=['produto','preco','link'])
        tabela_ofertas = pd.concat([tabela_ofertas, tabela_google_shopping])
    else:
        tabela_google_shopping = None
    lista_ofertas_buscape = busca_buscape(navegador, produto, termos_banidos, preco_min, preco_max) 
    if lista_ofertas_buscape:
        tabela_buscape = pd.DataFrame(lista_ofertas_buscape, columns=['produto','preco','link'])
        tabela_ofertas = pd.concat([tabela_ofertas, tabela_buscape])
    else:
        tabela_buscape = None
    
print(tabela_ofertas)


tabela_ofertas.to_excel('Ofertas.xlsx', index=False)


import win32com.client as win32
import smtplib
import email.message

def enviar_email():
    corpo_email = f"""<p> Prezados,</p>
    <p> encontramos alguns produtos em oferta dentro da faixa de preço desejada</p>
    {tabela_ofertas.to_html(index=False)}
    <p>Att,</p>
    <p>Rafael Mariano</p>
    """

    msg = email.message.Message()
    msg['Subject'] = 'Ofertas'
    msg['From'] ='rafaelmonitor.ect@gmail.com'
    msg['To'] = 'rafael_ferreira.mariano@outlook.com'
    password = 'rfxgmtwihvgcrhkq'
    msg.add_header('Content-Type' , 'text/html')
    msg.set_payload(corpo_email)

    s = smtplib.SMTP('smtp.gmail.com: 587')
    s.starttls()

    s.login(msg['From'], password)
    s.sendmail(msg['From'], [msg['To']], msg.as_string().encode('utf-8'))
    print('Email enviado')


enviar_email()