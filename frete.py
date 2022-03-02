from calendar import calendar
from datetime import datetime
import tkinter
from numpy import append
import pandas as pd
import sys, os
from pandas.core.accessor import register_dataframe_accessor
from pandas.core.frame import DataFrame
import requests
from requests import auth
from requests.api import get
from requests.models import ContentDecodingError
import csv, json
import math
import time
import pyodbc
from requests.auth import HTTPBasicAuth
import base64
from sys import exit
from selenium import webdriver
import selenium
from selenium.webdriver.chrome import options
from selenium.webdriver.chrome.options import Options
from time import sleep
import tkinter as tk
from tkinter import *
from tkinter import ttk 
from tkinter import WORD
from tkinter import messagebox
from tkinter import scrolledtext
from pandas import json_normalize
from datetime import date
from datetime import datetime, timedelta
import threading
import time
from tkinter import Text, filedialog
from tkcalendar import Calendar

server = '192.168.1.15' 
database = 'PROTHEUS_ZANOTTI_PRODUCAO' 
username = 'totvs' 
password = 'totvsip' 
conn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER='+server+';DATABASE='+database+';UID='+username+';PWD='+ password)
cursor = conn.cursor()
x="S"
happening =1
erro = 0

def sol_cotacao():
    var_mensagem_cf.set('Um instante...')
    sku= var_cod_prod_cf.get()
    sku = sku.upper()
    query_produtos="SELECT TRIM(SB1.B1_COD) AS SKU,\
                    SB1.B1_ZDESRDZ AS TITULO,\
                    CASE WHEN SB1.B1_ZMKTPLC = 'S' THEN 'enable' WHEN SB1.B1_ZMKTPLC = 'N' THEN 'disable' ELSE '' END AS STATUS,\
                    ISNULL((SELECT SUM(B2_QATU - B2_RESERVA)\
                    FROM SB2010 WHERE D_E_L_E_T_ = '' AND B2_LOCAL = '01' AND B2_COD = SB1.B1_COD),0) AS SALDO,\
                    SB1.B1_CUSTD AS CUSTO,\
                    SB1.B1_PESBRU AS PESO,\
                    SB5.B5_ALTURA AS ALTURA,\
                    SB5.B5_LARG AS LARGURA,\
                    SB5.B5_COMPR AS COMPRIMENTO,\
                    TRIM(SA2.A2_NREDUZ) AS MARCA,\
                    TRIM(SB1.B1_POSIPI) AS NCM,\
                    ZB2.ZB2_CODFAM AS COD_CATEGORIA,\
                    ZB2.ZB2_DESFAM AS CATEGORIA,\
                    DA1.DA1_PRCVEN,\
                    'TENSAO' AS COD_SPEC,\
                    TRIM(SB1.B1_ZTENSAO) AS SPEC \
                    FROM SB1010 SB1\
                    LEFT JOIN SB5010 SB5 ON SB5.D_E_L_E_T_ = '' AND SB5.B5_COD = SB1.B1_COD \
                    LEFT JOIN SA2010 SA2 ON SA2.D_E_L_E_T_ = '' AND SA2.A2_COD = SB1.B1_PROC AND SA2.A2_LOJA = SB1.B1_LOJPROC \
                    LEFT JOIN ZB2010 ZB2 ON ZB2.D_E_L_E_T_ = '' AND ZB2.ZB2_CODFAM = SB1.B1_ZCODFAM \
                    LEFT JOIN DA1010 DA1 ON DA1.D_E_L_E_T_ = '' AND DA1.DA1_CODTAB = 'LJA' AND DA1.DA1_CODPRO = SB1.B1_COD\
                    WHERE 1=1 AND SB1.D_E_L_E_T_ = '' AND SUBSTRING(SB1.B1_GRUPO,1,2) = 'MR' AND SB1.B1_COD='"+sku+"'"
    
    df_produto = pd.read_sql(query_produtos,conn)

    if df_produto.empty:
         var_mensagem_cf.set('Produto não encontrado')
    else:
        var_mensagem_cf.set('Um instante...')
        produto = str(df_produto.iloc[0,1]).title()
        peso = df_produto.iloc[0,5]
        altura = df_produto.iloc[0,6]
        largura = df_produto.iloc[0,7]
        profundidade = df_produto.iloc[0,8]
        stk = df_produto.iloc[0,3]
        valor_bd = df_produto.iloc[0,13]

        if altura == 0 or largura == 0 or profundidade == 0 or peso == 0:
            var_mensagem_cf.set('Produto com cadastro incompleto, verifique e tente novamente.')
        else:            
            var_alt_prod_cf.set(altura)
            var_peso_bruto_prod_cf.set(peso)
            var_larg_prod_cf.set(largura)
            var_prof_prod_cf.set(profundidade)
            var_cod_prod_bd_cf.set(sku)
            var_produto_bd_cf.set(produto)
            var_stk_prod_cf.set(stk)
            cnpjdestinatario = var_cpf_cnpj_cf.get()
            cnpjdestinatario = str(cnpjdestinatario).replace("-","")
            cnpjdestinatario = (cnpjdestinatario).replace(".","")
            cnpjdestinatario = (cnpjdestinatario).replace("/","")
            valor = var_valor_nf_cf.get()
            if valor == '':
                valor = valor_bd
            valor = str(valor).replace(",",".")
            valor = (valor).replace("R$","")
            valor = round(float(valor),2)
            peso = str(peso)
            peso = float(peso)
            diametro= math.sqrt((largura*largura)+(profundidade*profundidade))
            diametro = str(diametro)
            m3= (altura*largura*profundidade)
            altura= str(altura)
            largura= str(largura)
            profundidade= str(profundidade)
            m3=str(m3)
            data = var_dt_exped_nf_cf.get()
            data = str(data).replace("/","")
            dia = data[0:2]
            mes = data[2:4]
            ano = data[4:8]
            print(dia+"/"+mes+"/"+ano)
            cepdestino = var_cep_cf.get()
            cepdestino = str(cepdestino).replace("-","")
            cep = str(cepdestino).replace("-","")
            servico = "04510" #pac--sedex "04014"
            peso = float(peso)
            altura=float(altura)
            largura=float(largura)
            profundidade=float(profundidade)           

            url_cep = "https://viacep.com.br/ws/"+cepdestino+"/json/"
            request_endereco_cep = requests.request("GET",url_cep)
            request_endereco_cep = json.loads(request_endereco_cep.content)
            teste_cep = pd.DataFrame.from_dict(json_normalize(request_endereco_cep),orient = 'columns')
            try:
                rua = request_endereco_cep['logradouro'] 
                bairro = request_endereco_cep['bairro'] 
                cidade = request_endereco_cep['localidade'] 
                estado = request_endereco_cep['uf']
                var_label_consulta_frete.set(rua+', '+bairro+' - '+cidade+", "+estado)
            except:
                messagebox.askretrycancel('Atenção','CEP não encontrado.')
                #limpar()
            if (peso <= 30) and (altura <=  1.05) and (largura <= 1.05) and (profundidade <= 1.05):
                peso = str(peso)
                altura=str(altura)
                largura=str(largura)
                profundidade=str(profundidade)
                json_correios= '{"identificador":"0","cep_origem":"06530075","cep_destino":"'+cepdestino+'","formato":"1","peso":"'+peso+'","Comprimento":"'+profundidade+'","altura":"'+altura+'","largura":"'+largura+'","diametro":"'+diametro+'","mao_propria":"N","aviso_recebimento":"S","valor_declarado":"'+str(valor)+'","servicos":["41211"]}'
                headers_correios= {'Content-Type':'application/json'}
                url_correios = 'https://www.sgpweb.com.br/novo/api/consulta-precos-prazos?chave_integracao=ccfd6937b513db320221aac29b54f93c'
                response_correios = requests.request("POST", url_correios, headers = headers_correios, data = json_correios)
                valor_correios = "["+response_correios.text+"]"
                valor_correios = pd.read_json(valor_correios)
                valor_correios = str(valor_correios.iloc[0,12])
                valor_correios = valor_correios.replace("'",'"')
                valor_correios = pd.read_json(valor_correios)
                valor_correios2 = str(valor_correios.iloc[6,0])
                prazo_correios = str(valor_correios.iloc[5,0])
                valor_correios2 = "R$ "+valor_correios2.replace(".",",")
                transportadora_correios = "Correios"
                cotacao_correios = [transportadora_correios,valor_correios2,prazo_correios]
                acompanhamento_cf.insert('','end',values = cotacao_correios)
            else:
                transportadora_correios = "Correios"
                prazo_correios = "-"
                valor_correios2 = "-"
                peso = str(peso)
                peso = str(peso)
                altura=str(altura)
                largura=str(largura)
                profundidade=str(profundidade)
                cotacao_correios = [transportadora_correios,valor_correios2,prazo_correios]
                acompanhamento_cf.insert('','end',values = cotacao_correios)
        
            headers= {'Content-Type':'application/json'}
            url_autenticacao= "https://01wapi.rte.com.br/token"
            parametros= {'auth_type':'dev', 'grant_type':'password','username':'ZANOTTI', 'password':'7EPC6I8Z'}
            autenticacao = requests.request("POST", url_autenticacao, headers=headers, data=parametros)
            autenticacao=(autenticacao.text)
            autenticacao = "["+autenticacao+"]"
            autenticacao = pd.read_json(autenticacao)
            token = str(autenticacao.iloc[0,0])
            
            altura = float(altura)
            largura = float(largura)
            profundidade = float(profundidade)

            url_cidade= "https://01wapi.rte.com.br/api/v1/busca-por-cep?zipCode="+cepdestino
            headers_rte= {"Authorization": "Bearer "+token, "Content-Type":"application/json" }
            parametros_cidade = {}
            cidade_consulta = requests.request("GET", url_cidade, headers=headers_rte, data=parametros_cidade)
            IdCidade = json.loads(cidade_consulta.content)
            print(IdCidade)
            message = pd.DataFrame(IdCidade)
            message = message.iloc[0,1]

            if 'RTE' in message:                
                transportadora_rte = 'Rodonaves'
                prazo_rte = '-'
                valor_rte = '-'
            else:
                IdCidade=str(IdCidade['CityId'])
                url_simulacotacao = "https://01wapi.rte.com.br/api/v1/simula-cotacao"
                data_cotacao = '{"OriginZipCode": "06530075" ,"OriginCityId": 9560, "DestinationZipCode":"'+str(cepdestino)+'","DestinationCityId":"'+str(IdCidade)+'", "TotalWeight":"'+str(peso)+'", "EletronicInvoiceValue":"'+str(valor)+'", "CustomerTaxIdRegistration": "55963771000176", "Packs": [{"AmountPackages": 1,"Weight":"'+str(peso)+'","Length":"'+str(profundidade)+'","Height":"'+str(altura)+'","Width":"'+str(largura)+'"}]}'
                response_rte = requests.request("POST", url_simulacotacao, headers=headers_rte, data=data_cotacao)
                valor_rte = json.loads(response_rte.content)
                prazo_rte = str(valor_rte['DeliveryTime'])
                valor_rte = "R$ "+str(valor_rte['Value']).replace(".",",")
                transportadora_rte = 'Rodonaves'
            cotacao_rte = [transportadora_rte,valor_rte,prazo_rte]
            acompanhamento_cf.insert('','end',values=cotacao_rte)

            peso=float(peso)
            altura=float(altura)
            largura=float(largura)
            profundidade=float(profundidade)
            
            dados_happening = {'Cidades': ['AGUA VERMELHA','ALTINOPOLIS','AMERICANA','AMERICO BRASILIENSE','ANALANDIA','ARARAQUARA','ARARAS','ARUJÁ','ATIBAIA','BARRETOS','BARRINHA','BARUERI','BATATAIS','BEBEDOURO','BONFIM PAULISTA','BRODOWSKI','BUENO DE ANDRADE','BURITIZAL','CABREUVA','CACHOEIRA DE EMAS','CAIEIRAS','CAJAMAR','CAJOBI','CAJURU','CAMPINAS','CANDIA','CANDIDO RODRIGUES','CAPIVARI','CARAPICUÍBA','CASA BRANCA','CASSIA DOS COQUEIROS','COLINA','CORDEIROPOLIS','CÓRREGO RICO','COSMOPOLIS','COTIA','CRAVINHOS','CRISTAIS PAULISTA','CRUZ DAS POSSES','DESCALVADO','DIADEMA','DOBRADA','DUMONT','EMBU','FERNANDO PRESTES','FERRAZ DE VASCONCELOS','FRANCA','FRANCISCO MORATO','FRANCO DA ROCHA','GUAIRA','GUARA','GUARIBA','GUARULHOS','GUATAPARA','HORTOLANDIA','IBATE','IBITIUVA','INDAIATUBA','IPUA','ITAPECERICA DA SERRA','ITAPEVI','ITAQUAQUECETUBA','ITIRAPINA','ITIRAPUA','ITUPEVA','ITUVERAVA','JABORANDI','JABOTICABAL','JANDIRA','JARDINOPOLIS','JUND','JURUCE','LAGOA BRANCA','LEME','LIMEIRA','LUIS ANTONIO','LUSITÂNIA','MAIRIPORÃ','MARCONDESIA','MATAO','MAUÁ','MOCOCA','MOGI DAS CRUZES','MONTE ALTO','MONTE AZUL PAULISTA','MONTE MOR','MORRO AGUDO','MOTUCA','NOVA EUROPA','NOVA ODESSA','NUPORANGA','ORLANDIA','OSASCO','PATROCINIO PAULISTA','PAULINIA','PEDREGULHO','PIRACICABA','PIRAPORA DO BOM JESUS','PIRASSUNUNGA','PITANGUEIRAS','POÁ','PONTAL','PORTO FERREIRA','PRADOPOLIS','RESTINGA','RIBEIRÃO PIRES','RIBEIRAO PRETO','RINCAO','RIO CLARO','RIO GRANDE DA SERRA','SALES OLIVEIRA','SALTO','SANTA BARBARA D OESTE','SANTA CRUZ DA ESPERANCA','SANTA CRUZ DAS PALMEIRAS','SANTA ERNESTINA','SANTA GERTRUDES','SANTA ISABEL','SANTA LUCIA ','SANTA RITA DO PASSA QUATRO','SANTA ROSA DE VITERBO','SANTANA DE PARNAÍBA','SANTO ANDRÉ','SANTO ANTONIO DA ALEGRIA','SÃO BERNARDO DO CAMPO','SÃO CAETANO DO SUL','SAO CARLOS','SAO JOAQUIM DA BARRA','SAO JOSE DA BELA VISTA','SAO JOSE DO RIO PARDO','SÃO PAULO','SAO SIMAO','SERRA AZUL','SERRANA','SERTAOZINHO','SEVERINIA','SOROCABA','SUMARE','SUZANO','TABOÃO DA SERRA','TAMBAU','TAQUARITINGA','TERRA ROXA','VALINHOS','VARGEM GRANDE DO SUL','VINHEDO','VIRADOURO','VISTA ALEGRE DO ALTO'],'Prazos':['1','3','2','1','2','1','2','2','2','1','3','3','3','3','2','3','1','4','2','4','2','2','2','3','2','3','1','4','1','3','4','1','2','1','2','2','2','4','2','3','1','1','3','2','4','2','3','2','2','4','4','3','1','4','2','1','4','2','4','2','2','2','3','4','2','3','1','3','2','2','2','2','2','2','2','4','1','2','4','1','2','3','2','3','2','2','4','2','2','2','4','3','1','4','2','4','2','2','3','3','2','3','3','3','3','2','1','1','2','2','4','2','2','3','3','4','2','2','1','1','3','2','1','4','1','1','1','3','4','3','1','3','3','2','1','2','3','3','2','2','3','3','4','3','2','2','4','4']}
            dados_happening = pd.DataFrame(dados_happening, columns=['Cidades', 'Prazos'])
            
            try:
                cidade_happening = dados_happening.loc[dados_happening['Cidades'] == cidade]
                prazo_happening = str(cidade_happening.iloc[0,1])
                cidade_happening = str(cidade_happening.iloc[0,1])
                happening = 1
                transportadora_happening = 'Happening'
            except IndexError:
                transportadora_happening = 'Happening'
                prazo_happening = '-'
                valor_happening = '-'
                happening = 0
            
            if happening == 1:
                if valor <= 1846.14:
                    valor_happening = 120
                if 10000 >= valor >= 1846.15:
                    valor_happening = valor * 0.065
                if 10000.01 <= valor <= 15000:
                    valor_happening = valor *0.055
                if 15000.01 <= valor <= 25000:
                    valor_happening = valor *0.045
                if 25000.01 <= valor <= 30000:
                    valor_happening = valor *0.035
                if 30000.00 < valor:
                    valor_happening = valor *0.03
                valor_happening = round(valor_happening,2)
                valor_happening = "R$ "+str(valor_happening).replace(".",",")      
            cotacao_happening = [transportadora_happening,valor_happening,prazo_happening]
            acompanhamento_cf.insert('','end',values = cotacao_happening)
            
            peso=str(peso)
            altura=str(altura)
            largura=str(largura)
            profundidade=str(profundidade)
            
            if var_cot_zan.get() == 1:
                options = webdriver.ChromeOptions()
                options.add_argument("--headless")
                navegador = webdriver.Chrome('//192.168.1.16/02 - Público/09 - Marketing/Marketplaces/Ferramentas/chromedriver.exe', chrome_options=options)
                navegador.get('https://rotasbrasil.com.br/')
                #navegador.delete_all_cookies()
                sleep(1)
                navegador.find_element_by_xpath('//*[@id="painelVeiculos"]/div[2]').click()
                sleep(1)
                input_cidade_origem = navegador.find_element_by_xpath('//*[@id="txtEnderecoPartida"]')
                cidade_origem = 'Ribeirão Preto'
                input_cidade_origem.send_keys(cidade_origem)
                sleep(1)
                navegador.find_element_by_xpath('//*[@id="ui-id-1"]/li[1]').click()
                input_cidade_destino = navegador.find_element_by_xpath('//*[@id="txtEnderecoChegada"]')
                input_cidade_destino.send_keys(cidade)
                sleep(1)
                navegador.find_element_by_xpath('//*[@id="ui-id-2"]/li[1]').click()
                sleep(1)
                navegador.find_element_by_xpath('//*[@id="btnSubmit"]').click()

                sleep(1)

                pedagio = navegador.find_element_by_xpath('//*[@id="infoRoute0"]/span[2]/span').text
                distancia = navegador.find_element_by_xpath('//*[@id="painelResults"]/div/div[3]').text
                prazo_zanotti = navegador.find_element_by_xpath('//*[@id="painelResults"]/div/div[1]/div[1]').text
                transportadora_zanotti = 'Zanotti'
                pedagio = str(pedagio).replace(",",".")
                pedagio = float(pedagio)
                distancia = str(distancia).replace(" km","")
                distancia = (distancia).replace(".","")
                distancia = (distancia).replace(",",".")
                distancia = float(distancia)
                valor_zanotti = round(((distancia*2)*2)+(pedagio*2),2)
                valor_zanotti = "R$ "+ str(valor_zanotti).replace(".",",")
            
                cotacao_zanotti = [transportadora_zanotti,valor_zanotti,prazo_zanotti]
                acompanhamento_cf.insert('','end',values = cotacao_zanotti)
            '''          
            if (float(peso) <= 90) and (float(altura) <= 2) and (float(largura)<=2) and (float(profundidade)<=2):
                peso = str(peso)
                altura = str(altura)
                largura = str(largura)
                profundidade = str(profundidade)
                url_jadlog =  "https://www.jadlog.com.br/embarcador/api/frete/valor"
                json_jadlog = '{"cepori":"14080350","cepdes":"'+cepdestino+'","frap":"","peso":"'+peso+'","cnpj":"55963771000176","conta":"99999","contrato":"99999","modalidade":"4", "tpentrega":"D","tpseguro":"N","vldeclarado":"'+valor+'","vlcoleta":""}'
                headers_jadlog = {'Content-Type':'application/json', 'Authorization':'eyJ0eXAiOiJKV1QiLCJhbGciOiJIUzI1NiJ9.eyJqdGkiOjEwMjc5NCwiZHQiOiIyMDIxMDMxMiJ9.NfjRmgRnQ_JrTmhPO09EelxWmME5jtSdrf3ZFPaxPDM'}
                response_jadlog = requests.request("POST", url = url_jadlog, headers = headers_jadlog, data = json_jadlog)
                valor_jadlog = response_jadlog.text
                valor_jadlog = "["+valor_jadlog+"]"
                valor_jadlog = pd.read_json(valor_jadlog)
                print(valor_jadlog)
                valor_jadlog2 = valor_jadlog.iloc[0,2]
                prazo_jadlog = (valor_jadlog.iloc[0,1])
                msg_jadlog = "Jadlog: R$", valor_jadlog2, "prazo de", prazo_jadlog,"dia(s)"
            else:
                msg_jadlog = "Jadlog: Produto não atendido pela Jadlog"
            '''
            url_jamef = 'http://www.jamef.com.br/frete/rest/v1/1/55963771000176/ribeiraopreto/SP/000004/'+str(peso)+'/'+str(valor)+'/'+str(m3)+'/'+str(cepdestino)+'/18/'+str(dia)+'/'+str(mes)+'/'+str(ano)+'/T911416'
            response_jamef = requests.get(url_jamef)
            valor_jamef = response_jamef.text
            valor_jamef = "["+valor_jamef+"]"
            valor_jamef = pd.read_json(valor_jamef)

            if valor_jamef.empty:
                transportadora_jamef = 'Jamef'
                valor_jamef = "-"
                prazo_jamef = "-"
            else:    
                valor_jamef = float(valor_jamef.iloc[0,0])
                valor_jamef = round(valor_jamef,2)
                valor_jamef = "R$ "+str(valor_jamef).replace(".",",")
                transportadora_jamef = 'Jamef'
                prazo_jamef= '-'

            cotacao_jamef= [transportadora_jamef,valor_jamef,prazo_jamef]
            acompanhamento_cf.insert('','end',values = cotacao_jamef)

            peso=float(peso)
            altura=float(altura)
            largura=float(largura)
            profundidade=float(profundidade)

            if (peso <= 70) and (altura <= 1.2) and (largura<=1.2) and (profundidade<=1.2):
                peso = str(peso)
                altura = str(altura)
                largura = str(largura)
                profundidade = str(profundidade)
                url_braspress =  "https://hml-api.braspress.com/v1/cotacao/calcular/json"
                json_braspress= '{"cnpjRemetente":"55963771000176","cnpjDestinatario":"'+cnpjdestinatario+'","modal":"R","tipoFrete":"1","cepOrigem":"06530075","cepDestino":"'+cepdestino+'","vlrMercadoria":"'+str(valor)+'","peso":"'+peso+'","volumes":"1","cubagem":[{"comprimento":"'+profundidade+'","largura":"'+largura+'","altura":"'+altura+'","volumes":"1"}]}'
                headers_braspress = {'Content-Type':'application/json'}
                response_braspress = requests.request("POST", url=url_braspress, auth=HTTPBasicAuth('ZANOTTI_HML', '#3&Jw23rZ7Y%$w&O'),headers=headers_braspress,data=json_braspress)
                valor_braspress = response_braspress.text
                valor_braspress = "["+valor_braspress+"]"
                valor_braspress = pd.read_json(valor_braspress)
                valor_braspress2 = "R$ "+str(valor_braspress.iloc[0,2]).replace(".",",")
                prazo_braspress = (valor_braspress.iloc[0,1])
                transportadora_braspress = 'Braspress'
                cotacao_braspress= [transportadora_braspress,valor_braspress2,prazo_braspress]
                acompanhamento_cf.insert('','end',values = cotacao_braspress)
            else:
                transportadora_braspress = 'Braspress'
                valor_braspress2 = "-"
                prazo_braspress = "-"
                cotacao_braspress= [transportadora_braspress,valor_braspress2,prazo_braspress]
                acompanhamento_cf.insert('','end',values = cotacao_braspress)
                
            peso=float(peso)
            altura=float(altura)
            largura=float(largura)
            profundidade=float(profundidade)
            
            cidades_transexpress = pd.read_excel('//192.168.1.16/02 - Público/09 - Marketing/Marketplaces/Tabelas Frete/transexpress.xlsx')
            
            try:
                cidade_transexpress = cidades_transexpress.loc[cidades_transexpress['cidade'] == str(cidade).upper()]
                praca_transexpress =  str(cidade_transexpress.iloc[0,1])
                prazo_transexpress = '-'                
                valor_transexpress = float(valor)*0.0185
                if praca_transexpress == 'RP':
                    if valor_transexpress < 35:
                        valor_transexpress = 35
                if praca_transexpress == 'SJRP':
                    if valor_transexpress < 38:
                        valor_transexpress = 38
                
                valor_transexpress = round(valor_transexpress,2)
                valor_transexpress = "R$ "+str(valor_transexpress).replace('.',',')
                transportadora_transexpress = 'Trans Express'
                cotacao_transexpress= [transportadora_transexpress,valor_transexpress,prazo_transexpress]
                acompanhamento_cf.insert('','end',values = cotacao_transexpress)
            except IndexError:
                transportadora_transexpress = 'Trans Express'
                prazo_transexpress = '-'
                valor_transexpress = '-'           
                cotacao_transexpress= [transportadora_transexpress,valor_transexpress,prazo_transexpress]
                acompanhamento_cf.insert('','end',values = cotacao_transexpress)
           
            ###Leofran###
            m3 = largura*altura*profundidade
            peso_cubado = round(m3*300,2)
            
            if peso < peso_cubado:
                peso_leofran = peso_cubado
            else:
                peso_leofran = peso
            cpp_leofran = pd.read_excel('//192.168.1.16/02 - Público/09 - Marketing/Marketplaces/Tabelas Frete/leofran.xlsx')
            
            try:
                cidade_leofran = cpp_leofran.loc[cpp_leofran['cidade'] == str(cidade).upper()]
                
                prazo_leofran = cidade_leofran.iloc[0,1]
                praca_leofran = str(cidade_leofran.iloc[0,2])
                
                fracao_pedagio_valor = 3
                seguro_porc = 0.2
                gris_porc = 0.2
                excedente_valor = 0

                if (praca_leofran == 'RAO') or (praca_leofran == 'FRA'):
                    excedente_valor_por_kg= 0.25 
                    
                    if peso_leofran <= 20:
                        fixo_valor =  20

                    if peso_leofran > 20 and peso_leofran <= 50:
                        fixo_valor =  23

                    if peso_leofran > 50 and peso_leofran <= 70:
                        fixo_valor =  28

                    if peso_leofran > 70 and peso_leofran <= 100:
                        fixo_valor =  32
                    if peso_leofran > 100:
                        fixo_valor = 32

                if (praca_leofran == 'CAM') or (praca_leofran == 'SAO'):
                    
                    excedente_valor_por_kg= 0.30
                    
                    if peso_leofran <= 20:
                        fixo_valor =  23

                    if peso_leofran > 20 and peso_leofran <= 50:
                        fixo_valor =  32

                    if peso_leofran > 50 and peso_leofran <= 70:
                        fixo_valor =  38

                    if peso_leofran > 70 and peso_leofran <= 100:
                        fixo_valor =  43

                    if peso_leofran > 100:
                        fixo_valor = 32
                    
                if praca_leofran == 'SJC':
                    excedente_valor_por_kg = 0.40
                    if peso_leofran <= 20:
                        fixo_valor =  32

                    if peso_leofran > 20 and peso_leofran <= 50:
                        fixo_valor =  40

                    if peso_leofran > 50 and peso_leofran <= 70:
                        fixo_valor =  48

                    if peso_leofran > 70 and peso_leofran <= 100:
                        fixo_valor =  55

                    if peso_leofran > 100:
                        fixo_valor = 32
                
                if peso_leofran > 100:
                    
                    excedente_valor = (peso_leofran-100)*excedente_valor_por_kg
                    pedagio = (peso_leofran//100)*fracao_pedagio_valor
                        
                    if peso_leofran % 100 != 0:
                        pedagio = pedagio + fracao_pedagio_valor
                        
                else:
                    pedagio = fracao_pedagio_valor  
                
                valor_leofran = fixo_valor + excedente_valor + pedagio 
                valor_leofran = float((valor_leofran + (valor_leofran*seguro_porc)+(valor_leofran*gris_porc))*1.12)
                valor_leofran = round(valor_leofran,2)
                valor_leofran = str(valor_leofran)
                valor_leofran = "R$ " + str(valor_leofran).replace(".",",")
                transportadora_leofran = 'Leofran'
                cotacao_leofran= [transportadora_leofran,valor_leofran,prazo_leofran]
                acompanhamento_cf.insert('','end',values = cotacao_leofran)
            except:
                transportadora_leofran = 'Leofran'
                prazo_leofran = '-'
                valor_leofran = '-'
                cotacao_leofran= [transportadora_leofran,valor_leofran,prazo_leofran]
                acompanhamento_cf.insert('','end',values = cotacao_leofran)

            var_mensagem_cf.set('Consulta de frete realizada com sucesso!')
       
def limpar():
    var_cod_prod_cf.set('')
    var_cpf_cnpj_cf.set('')
    var_cep_cf.set('')
    var_valor_nf_cf.set('')
    var_dt_exped_nf_cf.set('')
    var_cot_zan.set(0)
    var_cod_prod_bd_cf.set('')
    var_produto_bd_cf.set('')
    var_alt_prod_cf.set('')
    var_larg_prod_cf.set('')
    var_prof_prod_cf.set('')
    var_peso_bruto_prod_cf.set('')
    var_stk_prod_cf.set('')
    var_mensagem_cf.set('Faça sua cotação')
    var_label_consulta_frete.set('')
        
    for item in acompanhamento_cf.get_children():
        acompanhamento_cf.delete(item)

janela = tk.Tk()
canvas = tk.Canvas(janela)
janela.title("Cotação de Frete")
janela.geometry('435x500+300+60')
janela.resizable(width=0,height=0)
janela.iconphoto(False,tk.PhotoImage(file='//192.168.1.16/02 - Público/09 - Marketing/Marketplaces/Inserções Manuais/Imagens/icone.png'))

label_cod_prod_cf = tk.Label(janela, text="Cod.Prod", fg="black").place(x=18,y=10)
var_cod_prod_cf = tk.StringVar()
ent_cod_prod_cf = tk.Entry(janela,fg="black", width=10, textvariable = var_cod_prod_cf).place(x=20,y=30)

label_cpf_cnpj_cf = tk.Label(janela, text="CPF/CNPJ", fg="black").place(x=88,y=10)
var_cpf_cnpj_cf = tk.StringVar()
entrada_cpf_cnpj_cf = tk.Entry(janela, fg="black", width=29, textvariable = var_cpf_cnpj_cf).place(x=90,y=30)

label_cep_cf = tk.Label(janela, text="CEP", fg="black").place(x=18,y=50)
var_cep_cf = tk.StringVar()
ent_cep_cf = tk.Entry(janela,fg="black", width=15, textvariable = var_cep_cf).place(x=20,y=70)

label_valor_nf_cf = tk.Label(janela, text="Valor da NF", fg="black").place(x=118,y=50)
var_valor_nf_cf = tk.StringVar()
ent_valor_nf_cf = tk.Entry(janela,fg="black", width = 10, textvariable = var_valor_nf_cf).place(x=120,y=70)

label_dt_exped_nf_cf = tk.Label(janela, text="Data Expedição", fg="black").place(x=188,y=50)
var_dt_exped_nf_cf = tk.StringVar()
ent_dt_exped_nf_cf = tk.Entry(janela,fg="black", width = 12, textvariable = var_dt_exped_nf_cf).place(x=190,y=70)

#cal = Calendar(janela, selectmode='day').place(x=190,y=70)

var_cot_zan = tk.IntVar()
cb_cot_zan = tk.Radiobutton(janela,text="Cotar pela Zanotti", variable=var_cot_zan, value=1).place(x=280,y=40)
cb_nao_cot_zan = tk.Radiobutton(janela,text="Não cotar pela Zanotti", variable=var_cot_zan, value=0).place(x=280,y=60)

botao_ok_cf = tk.Button(janela, text='Consultar', width=10,command=sol_cotacao).place(x=115, y=110)
botao_limpar_cf = tk.Button(janela, text='Limpar', width=10, command=limpar ).place(x=240, y=110)

lbl_cod_prod_cf = tk.Label(janela, text="Cod.Produto", fg="black").place(x=23,y=150)
var_cod_prod_bd_cf = StringVar()
txt_cod_prod_cf = tk.Entry(janela, bg= "#eff0f1", textvariable = var_cod_prod_bd_cf, width = 10, state=DISABLED).place(x=25,y=170)

lbl_produto_cf = tk.Label(janela, text="Produto", fg="black").place(x=103,y=150)
var_produto_bd_cf = StringVar()
txt_produto_cf = tk.Entry(janela, bg= "#eff0f1", textvariable = var_produto_bd_cf, width = 50, state=DISABLED).place(x=105,y=170)

lbl_alt_prod_cf = tk.Label(janela, text="Altura", fg="black").place(x=23,y=190)
var_alt_prod_cf = StringVar()
txt_alt_prod_cf = tk.Entry(janela, bg= "#eff0f1", textvariable = var_alt_prod_cf, width = 10, state=DISABLED).place(x=25,y=210)

lbl_larg_prod_cf = tk.Label(janela, text="Largura", fg="black").place(x=103,y=190)
var_larg_prod_cf = StringVar()
txt_larg_prod_cf = tk.Entry(janela, bg= "#eff0f1", textvariable = var_larg_prod_cf, width = 10, state=DISABLED).place(x=105,y=210)

lbl_prof_prod_cf = tk.Label(janela, text="Profundidade", fg="black").place(x=183,y=190)
var_prof_prod_cf = StringVar()
txt_prof_prod_cf = tk.Entry(janela, bg= "#eff0f1", textvariable = var_prof_prod_cf, width = 10, state=DISABLED).place(x=185,y=210) 

lbl_peso_bruto_prod_cf = tk.Label(janela, text="Peso Bruto", fg="black").place(x=263,y=190)
var_peso_bruto_prod_cf = StringVar()
txt_peso_bruto_prod_cf = tk.Entry(janela, bg= "#eff0f1", textvariable = var_peso_bruto_prod_cf, width = 10, state=DISABLED).place(x=265,y=210)

lbl_stk_prod_cf = tk.Label(janela, text="Estoque", fg="black").place(x=343,y=190)
var_stk_prod_cf = StringVar()
txt_stk_prod_cf = tk.Entry(janela, bg= "#eff0f1", textvariable = var_stk_prod_cf, width = 10, state=DISABLED).place(x=345,y=210)

acompanhamento_cf = ttk.Treeview(janela, selectmode="browse", column=('Transportadora','Preço','Prazo'),show ='headings')
acompanhamento_cf.column('Transportadora',width=180, minwidth=50, stretch=NO)
acompanhamento_cf.heading('#1', text='Transportadora',anchor=CENTER)
acompanhamento_cf.column('Preço',width=110, minwidth=50, stretch=NO,anchor=CENTER)
acompanhamento_cf.heading('#2', text='Preço',anchor=CENTER)
acompanhamento_cf.column('Prazo',width=108, minwidth=50, stretch=NO,anchor=CENTER)
acompanhamento_cf.heading('#3', text='Prazo',anchor=CENTER)
acompanhamento_cf.place(height=200, width=400,x=18, y=290)

var_mensagem_cf = StringVar()
label_mensagem_cf = tk.Label(janela, textvariable=var_mensagem_cf).place(x=23,y=240)
var_mensagem_cf.set('Faça sua cotação')

var_label_consulta_frete = StringVar()
label_atualizacao_consulta_frete = tk.Label(janela, textvariable=var_label_consulta_frete).place(x=23,y=260)
var_label_consulta_frete.set('')

janela.mainloop()
