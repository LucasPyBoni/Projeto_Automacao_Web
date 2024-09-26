#!/usr/bin/env python
# coding: utf-8

# # Projeto Automação Web - Busca de Preços
# 
# ### Objetivo: treinar um projeto em que a gente tenha que usar automações web com Selenium para buscar as informações que precisamos
# 
# - Já fizemos um projeto com esse objetivo no Módulo de Python e Web e em gravações de encontros ao vivo, mas não custa nada treinar mais um pouco.
# 
# ### Como vai funcionar:
# 
# - Imagina que você trabalha na área de compras de uma empresa e precisa fazer uma comparação de fornecedores para os seus insumos/produtos.
# 
# - Nessa hora, você vai constantemente buscar nos sites desses fornecedores os produtos disponíveis e o preço, afinal, cada um deles pode fazer promoção em momentos diferentes e com valores diferentes.
# 
# - Seu objetivo: Se o valor dos produtos for abaixo de um preço limite definido por você, você vai descobrir os produtos mais baratos e atualizar isso em uma planilha.
# - Em seguida, vai enviar um e-mail com a lista dos produtos abaixo do seu preço máximo de compra.
# 
# - No nosso caso, vamos fazer com produtos comuns em sites como Google Shopping e Buscapé, mas a ideia é a mesma para outros sites.
# 
# ### Outra opção:
# 
# - APIs
# 
# ### O que temos disponível?
# 
# - Planilha de Produtos, com os nomes dos produtos, o preço máximo, o preço mínimo (para evitar produtos "errados" ou "baratos de mais para ser verdade" e os termos que vamos querer evitar nas nossas buscas.
# 
# ### O que devemos fazer:
# 
# - Procurar cada produto no Google Shopping e pegar todos os resultados que tenham preço dentro da faixa e sejam os produtos corretos
# - O mesmo para o Buscapé
# - Enviar um e-mail para o seu e-mail (no caso da empresa seria para a área de compras por exemplo) com a notificação e a tabela com os itens e preços encontrados, junto com o link de compra. (Vou usar o e-mail pythonimpressionador@gmail.com. Use um e-mail seu para fazer os testes para ver se a mensagem está chegando)

# In[1]:


# criar um navegador
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
import time
import pandas as pd
import pprint
import win32com.client as win32
nav = webdriver.Chrome()

tabela_produtos = pd.read_excel('buscas.xlsx')

def buscador_google(nav, nome, preco_min, preco_max, termos_banidos):
    nome = nome.lower()
    termos_banidos = termos_banidos.lower()
    
    lista_termos_nome = nome.split(' ')
    lista_termos_banidos = termos_banidos.split(' ')


    preco_min = float(preco_min)
    preco_max = float(preco_max)

    # Abrir navegador, pesquisar e pressionar enter
    nav.get('https://www.google.com.br/')
    nav.find_element(By.XPATH, '//*[@id="APjFqb"]').send_keys(nome, Keys.ENTER)
    time.sleep(1)

    # Clicar em shopping
    nav.find_element(By.XPATH, '//*[@id="hdtb-sc"]/div/div/div[1]/div/div[2]/a/div').click()
    time.sleep(1)

    # Aqui você inicializa as listas e dicionários fora do loop de resultados
    lista_infos = []

    # Obtenção dos resultados
    resultados = nav.find_elements(By.CLASS_NAME, 'i0X6df')
    for resultado in resultados:
        produto = resultado.find_element(By.CLASS_NAME, 'tAxDx').text
        produto = produto.lower()

        # Verificação do nome - se no nome tem algum termo banido
        tem_termos_banidos = False
        for palavra in lista_termos_banidos:
            if palavra in produto:
                tem_termos_banidos = True

        # Verificar se no nome tem todos os termos do nome do produto
        tem_todos_termos_produto = True
        for palavra in lista_termos_nome:
            if palavra not in produto:
                tem_todos_termos_produto = False

        # Se passar nas verificações
        if not tem_termos_banidos and tem_todos_termos_produto:
            preco = resultado.find_element(By.CLASS_NAME, 'a8Pemb').text
            preco = preco.replace("R$", "").replace(" ", "").replace(".", "").replace(",", ".").replace("+impostos", '')
            preco = float(preco)

            # Verificando se o preço está dentro do mínimo e máximo
            if preco_min <= preco <= preco_max:
                link = resultado.find_element(By.CLASS_NAME, 'Lq5OHe').get_attribute('href')

                # Adiciona os valores ao dicionário e à lista
                dic_infos = {'produto': produto, 'preco': preco, 'link': link}

                # Adiciona o dicionário à lista
                lista_infos.append([produto, preco, link])

    # Agora 'lista_infos' contém todos os produtos que passaram pelas verificações
    return lista_infos

def buscador_buscape(nav, nome, preco_min, preco_max, termos_banidos):
    nome = nome.lower()
    termos_banidos = termos_banidos.lower()
    
    lista_termos_nome = nome.split(' ')
    lista_termos_banidos = termos_banidos.split(' ')


    preco_min = float(preco_min)
    preco_max = float(preco_max)

    # Abrir navegador, pesquisar e pressionar enter
    nav.get('https://www.buscape.com.br/')
    time.sleep(1)
    nav.find_element(By.XPATH, '//*[@id="new-header"]/div[1]/div/div/div[3]/div/div/div[2]/div/div[1]/input').send_keys(nome, Keys.ENTER)
    time.sleep(1)

    # Aqui você inicializa as listas e dicionários fora do loop de resultados
    lista_infos = []
    
    # Obtenção dos resultados
    resultados = nav.find_elements(By.CLASS_NAME, 'Hits_ProductCard__Bonl_')
    for resultado in resultados:
        produto = resultado.find_element(By.CLASS_NAME, 'ProductCard_ProductCard_Name__U_mUQ').text
        produto = produto.lower()

        # Verificação do nome - se no nome tem algum termo banido
        tem_termos_banidos = False
        for palavra in lista_termos_banidos:
            if palavra in produto:
                tem_termos_banidos = True

        # Verificar se no nome tem todos os termos do nome do produto
        tem_todos_termos_produto = True
        for palavra in lista_termos_nome:
            if palavra not in produto:
                tem_todos_termos_produto = False

        # Se passar nas verificações
        if not tem_termos_banidos and tem_todos_termos_produto:
            preco = resultado.find_element(By.CLASS_NAME, 'Text_MobileHeadingS__HEz7L').text
            preco = preco.replace("R$", "").replace(" ", "").replace(".", "").replace(",", ".").replace("+impostos", '')
            preco = float(preco)

            # Verificando se o preço está dentro do mínimo e máximo
            if preco_min <= preco <= preco_max:
                link = resultado.find_element(By.CLASS_NAME, 'ProductCard_ProductCard_Inner__gapsh').get_attribute('href')

                # Adiciona os valores ao dicionário e à lista
                dic_infos = {'produto': produto, 'preco': preco, 'link': link}

                # Adiciona o dicionário à lista
                lista_infos.append([produto, preco, link])

    # Agora 'lista_infos' contém todos os produtos que passaram pelas verificações
    return lista_infos

df_ofertas = pd.DataFrame()
for linha in tabela_produtos.index:
    # Tratamentos
    nome = tabela_produtos.loc[linha, 'Nome']
    termos_banidos = tabela_produtos.loc[linha, 'Termos banidos']
    preco_min = tabela_produtos.loc[linha, 'Preço mínimo']
    preco_max = tabela_produtos.loc[linha, 'Preço máximo']

    busca_buscape = buscador_buscape(nav, nome, preco_min, preco_max, termos_banidos)
    if busca_buscape:
        df_busca_buscape = pd.DataFrame(busca_buscape, columns=['Produto','Preço','Link'])
        df_ofertas = pd.concat([df_ofertas, df_busca_buscape], ignore_index=True)
    
    busca_google = buscador_google(nav, nome, preco_min, preco_max, termos_banidos)
    if busca_google:
        df_busca_google = pd.DataFrame(busca_google, columns=['Produto','Preço','Link'])
        df_ofertas = pd.concat([df_ofertas, df_busca_google], ignore_index=True)
    
# criando df
df_ofertas.to_excel("Ofertas.xlsx",index=False)

if len(df_ofertas.index) > 0:
    # vou enviar email
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = 'lordtsyrinxbr@gmail.com'
    mail.Subject = 'Produto(s) Encontrado(s) na faixa de preço desejada'
    mail.HTMLBody = f"""
    <p>Prezados,</p>
    <p>Encontramos alguns produtos em oferta dentro da faixa de preço desejada. Segue tabela com detalhes</p>
    {df_ofertas.to_html(index=False)}
    <p>Qualquer dúvida estou à disposição</p>
    <p>Att., Lukas Maciel</p>
    """
    
    mail.Send()

nav.quit()  
    

