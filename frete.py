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

server = '192.168.1.15' 
database = 'PROTHEUS_ZANOTTI_PRODUCAO' 
username = 'totvs' 
password = 'totvsip' 
conn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER='+server+';DATABASE='+database+';UID='+username+';PWD='+ password)
cursor = conn.cursor()
x="S"
happening =1
erro = 0
'''
janela = Tk()
janela.title("Simulação de Frete")
janela.geometry('351x345+400+150')
janela.resizable(width=0,height=0)
cod_produto = tkinter.Label(janela, text="Sku", fg="black").place(x=8,y=10)
#ent_sku=tkinter.Entry(fg="black",bg= "#eff0f1", width=10).place(x=10,y=30)
label_valor = tkinter.Label(janela, text="Preço", fg="black").place(x=81,y=10)
ent_valor=tkinter.Entry(fg="black",bg= "#eff0f1", width=20, ).place(x=83,y=30)
label_cnpj = tkinter.Label(janela, text="CPF\CNPJ", fg="black").place(x=214,y=10)
ent_cnpj=tkinter.Entry(fg="black",bg= "#eff0f1", width=20).place(x=216,y=30)
label_cep = tkinter.Label(janela, text="CEP", fg="black").place(x=7,y=50)
ent_cep=tkinter.Entry(fg="black",bg= "#eff0f1", width=10).place(x=10,y=70)
label_cidade = tkinter.Label(janela, text="Cidade", fg="black").place(x=81,y=50)
ent_cidade=tkinter.Entry(fg="black",bg= "#d4d4d4", width=30).place(x=83,y=70)
label_data = tkinter.Label(janela, text="Data", fg="black").place(x=271,y=50)
ent_data = tkinter.Entry(fg="black",bg= "#eff0f1", width=10).place(x=275,y=70)
#botao_consulta =Button(janela, text="Consultar", fg="black", width=10).place(x=60,y=300)
botao_novaconsulta =Button(janela, text="Nova", fg="black", width=10).place(x=200,y=300)
resultado_consulta =  Listbox(bg="#eff0f1",width=55, fg="black").place(x=8, y=120)


def clicar():
   sku= ent_sku.get()
   print("clique")

botao_consulta =Button(janela, text="Consultar", fg="black", width=10, command=clicar()).place(x=60,y=300)
ent_sku=tkinter.Entry(fg="black",bg= "#eff0f1", width=10).place(x=10,y=30)

janela.mainloop()

'''
while x=="S":
    while erro == 0:
        sku= input("Digite o código do produto: ")
        sku = sku.upper()
        query_produtos="SELECT TRIM(SB1.B1_COD) AS SKU, TRIM(SB5.B5_ECTITU) AS TITULO, SB5.B5_ECDESCR AS DESCRIÇÃO, CASE WHEN SB1.B1_ZMKTPLC = 'S' THEN 'enable' WHEN SB1.B1_ZMKTPLC = 'N' THEN 'disable' ELSE '' END AS STATUS, ISNULL((SELECT SUM(B2_QATU - B2_RESERVA) FROM SB2010 WHERE D_E_L_E_T_ = '' AND B2_LOCAL = '01' AND B2_COD = SB1.B1_COD),0) AS SALDO, SB1.B1_CUSTD	AS CUSTO, SB1.B1_PESBRU	AS PESO, SB5.B5_ALTURA AS ALTURA, SB5.B5_LARG AS LARGURA, SB5.B5_COMPR AS COMPRIMENTO, TRIM(SA2.A2_NREDUZ) AS MARCA, TRIM(SB1.B1_CODBAR) AS CODBAR, TRIM(SB1.B1_POSIPI) AS NCM, ZB2.ZB2_CODFAM AS COD_CATEGORIA, ZB2.ZB2_DESFAM AS CATEGORIA, TRIM(SB5.B5_ECIMGFI) AS IMAGEM,'TENSAO' AS COD_SPEC, TRIM(SB1.B1_ZTENSAO) AS SPEC FROM SB1010 SB1 LEFT JOIN SB5010 SB5 ON SB5.D_E_L_E_T_ = '' AND SB5.B5_COD = SB1.B1_COD LEFT JOIN SA2010 SA2 ON SA2.D_E_L_E_T_ = '' AND SA2.A2_COD = SB1.B1_PROC AND SA2.A2_LOJA = SB1.B1_LOJPROC LEFT JOIN ZB2010 ZB2 ON ZB2.D_E_L_E_T_ = '' AND ZB2.ZB2_CODFAM = SB1.B1_ZCODFAM WHERE 1=1 AND SB1.D_E_L_E_T_ = '' AND SUBSTRING(SB1.B1_GRUPO,1,2) = 'MR' AND SB1.B1_COD='"+sku+"'"
        df_produto=pd.read_sql(query_produtos,conn)

        if df_produto.empty:
            erro=0
        else:
            erro=1

    peso = df_produto.iloc[0,6]
    altura = df_produto.iloc[0,7]
    largura = df_produto.iloc[0,8]
    profundidade = df_produto.iloc[0,9]
    print("Peso:", peso,"kg -", "Altura:", altura, "metros -", "Largura:", largura,"metros -", "Profundidade:",profundidade,"metros" )
                      
    cnpjdestinatario = input("Digite o CNPJ ou CPF do cliente: ")
    valor= input("Digite o valor do produto: R$ ")
    valor =float(valor.replace(',','.'))
    peso = str(peso)
    peso = float(peso)
    diametro= math.sqrt((largura*largura)+(profundidade*profundidade))
    diametro = str(diametro)
    m3= (altura*largura*profundidade)
    altura= str(altura)
    largura= str(largura)
    profundidade= str(profundidade)
    m3=str(m3)
    data = input("Digite a data que será expedido (dd/mm/aaaa): ")
    dia = data[0:2]
    mes = data[3:5]
    ano = data[6:10]
    cepdestino = input("Digite o CEP de destino: ")
    servico = "04510" #pac--sedex "04014"

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
    cidade_consulta=(cidade_consulta.text)
    IdCidade="["+cidade_consulta+"]"

    if IdCidade == '[[{"PropertyName":"","Message":"RTE não atende CEP (62500025) informado"}]]':
        IdCidade = 0
    else:
        print(IdCidade)
        IdCidade2= pd.read_json(IdCidade)
        IdCidade=str(IdCidade2.iloc[0,0])
        dados_happening = {'Cidades': ['AGUA VERMELHA','ALTINOPOLIS','AMERICANA','AMERICO BRASILIENSE','ANALANDIA','ARARAQUARA','ARARAS','ARUJÁ','ATIBAIA','BARRETOS','BARRINHA','BARUERI','BATATAIS','BEBEDOURO','BONFIM PAULISTA','BRODOWSKI','BUENO DE ANDRADE','BURITIZAL','CABREUVA','CACHOEIRA DE EMAS','CAIEIRAS','CAJAMAR','CAJOBI','CAJURU','CAMPINAS','CANDIA','CANDIDO RODRIGUES','CAPIVARI','CARAPICUÍBA','CASA BRANCA','CASSIA DOS COQUEIROS','COLINA','CORDEIROPOLIS','CÓRREGO RICO','COSMOPOLIS','COTIA','CRAVINHOS','CRISTAIS PAULISTA','CRUZ DAS POSSES','DESCALVADO','DIADEMA','DOBRADA','DUMONT','EMBU','FERNANDO PRESTES','FERRAZ DE VASCONCELOS','FRANCA','FRANCISCO MORATO','FRANCO DA ROCHA','GUAIRA','GUARA','GUARIBA','GUARULHOS','GUATAPARA','HORTOLANDIA','IBATE','IBITIUVA','INDAIATUBA','IPUA','ITAPECERICA DA SERRA','ITAPEVI','ITAQUAQUECETUBA','ITIRAPINA','ITIRAPUA','ITUPEVA','ITUVERAVA','JABORANDI','JABOTICABAL','JANDIRA','JARDINOPOLIS','JUND','JURUCE','LAGOA BRANCA','LEME','LIMEIRA','LUIS ANTONIO','LUSITÂNIA','MAIRIPORÃ','MARCONDESIA','MATAO','MAUÁ','MOCOCA','MOGI DAS CRUZES','MONTE ALTO','MONTE AZUL PAULISTA','MONTE MOR','MORRO AGUDO','MOTUCA','NOVA EUROPA','NOVA ODESSA','NUPORANGA','ORLANDIA','OSASCO','PATROCINIO PAULISTA','PAULINIA','PEDREGULHO','PIRACICABA','PIRAPORA DO BOM JESUS','PIRASSUNUNGA','PITANGUEIRAS','POÁ','PONTAL','PORTO FERREIRA','PRADOPOLIS','RESTINGA','RIBEIRÃO PIRES','RIBEIRAO PRETO','RINCAO','RIO CLARO','RIO GRANDE DA SERRA','SALES OLIVEIRA','SALTO','SANTA BARBARA D OESTE','SANTA CRUZ DA ESPERANCA','SANTA CRUZ DAS PALMEIRAS','SANTA ERNESTINA','SANTA GERTRUDES','SANTA ISABEL','SANTA LUCIA ','SANTA RITA DO PASSA QUATRO','SANTA ROSA DE VITERBO','SANTANA DE PARNAÍBA','SANTO ANDRÉ','SANTO ANTONIO DA ALEGRIA','SÃO BERNARDO DO CAMPO','SÃO CAETANO DO SUL','SAO CARLOS','SAO JOAQUIM DA BARRA','SAO JOSE DA BELA VISTA','SAO JOSE DO RIO PARDO','SÃO PAULO','SAO SIMAO','SERRA AZUL','SERRANA','SERTAOZINHO','SEVERINIA','SOROCABA','SUMARE','SUZANO','TABOÃO DA SERRA','TAMBAU','TAQUARITINGA','TERRA ROXA','VALINHOS','VARGEM GRANDE DO SUL','VINHEDO','VIRADOURO','VISTA ALEGRE DO ALTO'],'Prazos':['1','3','2','1','2','1','2','2','2','1','3','3','3','3','2','3','1','4','2','4','2','2','2','3','2','3','1','4','1','3','4','1','2','1','2','2','2','4','2','3','1','1','3','2','4','2','3','2','2','4','4','3','1','4','2','1','4','2','4','2','2','2','3','4','2','3','1','3','2','2','2','2','2','2','2','4','1','2','4','1','2','3','2','3','2','2','4','2','2','2','4','3','1','4','2','4','2','2','3','3','2','3','3','3','3','2','1','1','2','2','4','2','2','3','3','4','2','2','1','1','3','2','1','4','1','1','1','3','4','3','1','3','3','2','1','2','3','3','2','2','3','3','4','3','2','2','4','4']}
        dados_happening = pd.DataFrame(dados_happening, columns=['Cidades', 'Prazos'])
        cidade = str(IdCidade2.iloc[0,1])
    
        try:
            cidade_happening = dados_happening.loc[dados_happening['Cidades'] == cidade]
            prazo_happening = str(cidade_happening.iloc[0,1])
            cidade_happening = str(cidade_happening.iloc[0,1])
            happening = 1
        except IndexError:
            happening = 0
    
        if happening == 1:
            if valor <= 1846.14:
                preco_happening = 120
            if 10000 >= valor >= 1846.15:
                preco_happening = valor * 0.065
            if 10000.01 <= valor <= 15000:
                preco_happening = valor *0.055
            if 15000.01 <= valor <= 25000:
                preco_happening = valor *0.045
            if 25000.01 <= valor <= 30000:
                preco_happening = valor *0.035
            if 30000.00 < valor:
                preco_happening = valor *0.03
            preco_happening = round(preco_happening,2)         
            msg_happening="Happening: R$",preco_happening, "prazo de", prazo_happening,"dia(s)"
        else:
            msg_happening="Happening: Cidade não atendida"
        
        msg_happening = str(msg_happening)
        msg_happening = msg_happening.replace("'","")
        msg_happening = msg_happening.replace("(","",1)
        msg_happening = msg_happening.replace(")","",1)
        msg_happening = msg_happening.replace(",","")

    valor = str(valor)
    peso=str(peso)
    altura=str(altura)
    largura=str(largura)
    profundidade=str(profundidade)

    navegador = webdriver.Chrome('C:/Users/Thales.Bartolomeu/Desktop/MKTPLACE/chromedriver.exe')
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
    prazo_zan = navegador.find_element_by_xpath('//*[@id="painelResults"]/div/div[1]/div[1]').text

    
    if IdCidade != 0:
        url_simulacotacao = "https://01wapi.rte.com.br/api/v1/simula-cotacao"
        data_cotacao = '{"OriginZipCode": "06530075" ,"OriginCityId": 9560, "DestinationZipCode":"'+cepdestino+'","DestinationCityId":"'+IdCidade+'", "TotalWeight":"'+peso+'", "EletronicInvoiceValue":"'+valor+'", "CustomerTaxIdRegistration": "55963771000176", "Packs": [{"AmountPackages": 1,"Weight":"'+peso+'","Length":"'+profundidade+'","Height":"'+altura+'","Width":"'+largura+'"}]}'
        response_rte = requests.request("POST", url_simulacotacao, headers=headers_rte, data=data_cotacao)
        valor_rte = response_rte.text
        valor_rte = "["+valor_rte+"]"
        valor_rte = pd.read_json(valor_rte)
        #prazo_rte = str(valor_rte.iloc[0,1])
        valor_rte = str(valor_rte.iloc[0,0])
    peso=float(peso)
    altura=float(altura)
    largura=float(largura)
    profundidade=float(profundidade)

    
    """
    if (peso < 90) and (altura <= 2) and (largura<=2) and (profundidade<=2):
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

    msg_jadlog = str(msg_jadlog)
    mmsg_jadlog = msg_jadlog.replace("'","")
    msg_jadlog = msg_jadlog.replace("(","",1)
    msg_jadlog = msg_jadlog.replace(")","",1)
    msg_jadlog = msg_jadlog.replace(",","")
    """
    if (peso < 30) and (altura <=  0.9) and (largura <= 0.9) and (profundidade <= 0.9):
        peso = str(peso)
        altura=str(altura)
        largura=str(largura)
        profundidade=str(profundidade)
        valor=str(valor)

        json_correios= '{"identificador":"0","cep_origem":"06530075","cep_destino":"'+cepdestino+'","formato":"1","peso":"'+peso+'","Comprimento":"'+profundidade+'","altura":"'+altura+'","largura":"'+largura+'","diametro":"'+diametro+'","mao_propria":"N","aviso_recebimento":"S","valor_declarado":"'+valor+'","servicos":["41211"]}'
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
        valor_correios2 = valor_correios2.replace(",",".")
        msg_correios = "Correios: R$",valor_correios2, "prazo de", prazo_correios,"dia(s)"
    else:
        msg_correios = "Correios: Produto com mais de 30kg, não atendido pelos Correios"
        peso = str(peso)
        peso = str(peso)
        altura=str(altura)
        largura=str(largura)
        profundidade=str(profundidade)
    msg_correios = str(msg_correios)
    msg_correios = msg_correios.replace("'","")
    msg_correios = msg_correios.replace("(","",1)
    msg_correios = msg_correios.replace(")","",1)
    msg_correios = msg_correios.replace(",","")

    url_jamef = 'http://www.jamef.com.br/frete/rest/v1/1/55963771000176/ribeiraopreto/SP/000004/'+peso+'/'+valor+'/'+m3+'/'+cepdestino+'/18/'+dia+'/'+mes+'/'+ano+'/T911416'
    response_jamef = requests.get(url_jamef)
    valor_jamef = response_jamef.text
    valor_jamef = "["+valor_jamef+"]"
    valor_jamef = pd.read_json(valor_jamef)
    valor_jamef = float(valor_jamef.iloc[0,0])
    valor_jamef = round(valor_jamef,2)

    peso=float(peso)
    altura=float(altura)
    largura=float(largura)
    profundidade=float(profundidade)

    if (peso < 50) and (altura <= 1) and (largura<=1) and (profundidade<=1):
        peso = str(peso)
        altura = str(altura)
        largura = str(largura)
        profundidade = str(profundidade)
        url_braspress =  "https://hml-api.braspress.com/v1/cotacao/calcular/json"
        json_braspress= '{"cnpjRemetente":"55963771000176","cnpjDestinatario":"'+cnpjdestinatario+'","modal":"R","tipoFrete":"1","cepOrigem":"06530075","cepDestino":"'+cepdestino+'","vlrMercadoria":"'+valor+'","peso":"'+peso+'","volumes":"1","cubagem":[{"comprimento":"'+profundidade+'","largura":"'+largura+'","altura":"'+altura+'","volumes":"1"}]}'
        headers_braspress = {'Content-Type':'application/json'}
        response_braspress = requests.request("POST", url=url_braspress, auth=HTTPBasicAuth('ZANOTTI_HML', '#3&Jw23rZ7Y%$w&O'),headers=headers_braspress,data=json_braspress)
        valor_braspress = response_braspress.text
        valor_braspress = "["+valor_braspress+"]"
        valor_braspress = pd.read_json(valor_braspress)
        valor_braspress2 = valor_braspress.iloc[0,2]
        prazo_braspress = (valor_braspress.iloc[0,1])
        msg_braspress = "Braspress: R$", valor_braspress2, "prazo de", prazo_braspress,"dia(s)"
    else:
        msg_braspress = "Braspress: Produto não atendido pela Braspress"

    msg_braspress = str(msg_braspress)
    msg_braspress = msg_braspress.replace("'","")
    msg_braspress = msg_braspress.replace("(","",1)
    msg_braspress = msg_braspress.replace(")","",1)
    msg_braspress = msg_braspress.replace(",","")

    peso=float(peso)
    altura=float(altura)
    largura=float(largura)
    profundidade=float(profundidade)
    
    
    if IdCidade != 0:
        print("Frete para a cidade de "+cidade)
        print("")
        print("Rodonaves: R$", valor_rte)
        print(msg_happening)
    print("Jamef: R$", valor_jamef)
    print(msg_braspress)
    print(msg_correios)

    distancia = str(distancia)
    distancia = distancia.replace(" km","")
    distancia = distancia.replace(",",".")
    distancia = float(distancia)
    pedagio = str(pedagio)
    pedagio = pedagio.replace(",",".")
    pedagio=float(pedagio)
    print("Zanotti: R$",round((distancia*2)*1.7)+(pedagio*2))
    #print(msg_jadlog)
    print("")
    nova_consulta= input("Deseja fazer uma nova consulta? (S/N): ")
    nova_consulta=nova_consulta.upper()
    if nova_consulta=="S":
       x="S"
       erro=0
    else:
       x="N"
       exit

