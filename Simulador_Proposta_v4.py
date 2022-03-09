from cProfile import label
from email.mime.application import MIMEApplication
from email.mime.base import MIMEBase
from email.mime.image import MIMEImage
from email.utils import collapse_rfc2231_value
import pstats
from multiprocessing.sharedctypes import Value
from typing import Optional
from click import launch
import streamlit as st
import numpy as np
import pandas as pd
import pathlib
import time
#import win32com.client as win32
#from oauth2client.service_account import ServiceAccountCredentials
#from email.mime.multipart import MIMEMultipart
#from email.mime.text import MIMEText
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import dataframe_image as dfi
from fpdf import FPDF 
from datetime import datetime
from datetime import date


df = pd.read_excel("Base_Preços.xlsx")

imagem_Logo = "https://github.com/cbaracho2/SIMULADOR_PROPOSTA_V4/blob/main/logo7lm.png?raw=true"
Imagem_data = "https://github.com/cbaracho2/SIMULADOR_PROPOSTA_V4/blob/main/Solicitar_Aprovacao.png?raw=true"

image = "https://grupoimerge.com.br/wp-content/uploads/2020/06/logo-7lm.svg"
st.sidebar.image(image,use_column_width=False, width=None)
st.title("Simulador do Plano de Pgto | 7LM")
st.subheader('Construção do plano de pagamento:')
botao_003 = st.button("Solicitar Liberação da Proposta")

   
def enviar_email_002(cidade,empreendimento,bloco,unidade):
    try:
        username = "solicitarliberacaodaproposta@gmail.com"
        password = "G@fisa2014"
        mail_from = "comercial@7lm.com.br"
        mail_to = "comercial@7lm.com.br"
        mail_subject = f"7LM | SOLICITAÇÃO DE APROVAÇÃO | {cidade} | {empreendimento} | {bloco} | {unidade}"
        mail_body = "Segue proposta para liberação."
        arquivo ="Solicitação_Proposta.pdf"

        mimemsg = MIMEMultipart()
        mimemsg['From']=mail_from
        mimemsg['To']=mail_to
        mimemsg['Subject']=mail_subject
        mimemsg.attach(MIMEText(mail_body, 'plain'))
        fp = open(arquivo, 'rb')
        anexo = MIMEApplication(fp.read(), _subtype="pdf")
        fp.close()
        anexo.add_header('Content-Disposition', 'attachment', filename=arquivo)
        mimemsg.attach(anexo)
        
        connection = smtplib.SMTP(host='smtp.gmail.com', port=587)
        connection.starttls()
        connection.login(username,password)
        connection.send_message(mimemsg)
        connection.quit()
        st.write("E-mail enviado com sucesso.")
    except:
        st.write("Opa! Algo saiu errado.")
        
    
def add_image(caminho_da_imagem,logo,EMPREEND,TORRE,UNID, situacao):
    y=190
    x=125
    ASSINATURA = "CIENTE DIRETORIA [   ]"
    pdf = FPDF() 
    pdf.add_page() 
    pdf.set_font('Arial','',18)
    pdf.image(logo, x = 10 , y = 5 , w = 40 ) 
    pdf.cell(195,40,'7LM | LIBERAÇÃO DE PROPOSTA | ALÇADA DIRETORIA' ,0,1,'C')#fill=True
    pdf.set_font ( "Arial" , size=12) 
    pdf.set_text_color(0,102,153)
    pdf.text(10 , 40 , f"EMPREENDIMENTO: {EMPREEND}")
    pdf.text(10 , 45 ,  f"BLOCO: {TORRE}")
    pdf.text(10 , 50 ,  f"UNIDADE: {UNID}") 
    pdf.image(caminho_da_imagem, x = 20 , y = 90 , w = 165 )
    if situacao == "APROVADO":
        pdf.set_font ( "Arial" , size=14) 
        pdf.set_text_color(26,188,41)
        pdf.text(x, y , f"PARECER: {situacao}")
        pdf.set_text_color(0,0,0)
        pdf.text(125, 205, f"{ASSINATURA}")
        pdf.set_font ( "Arial" , size=10) 
        pdf.text(10, 250,f"Arquivo gerado automaticamente | Inteligência Grupo Imerge | {datetime.today().strftime('%Y-%m-%d %H:%M')}")  
        pdf.output("Solicitação_Proposta.pdf" )
    else:
        pdf.set_font ( "Arial" , size=14) 
        pdf.set_text_color(234,50,0)
        pdf.text(x, y,  f"PARECER: {situacao}")
        pdf.set_text_color(0,0,0)
        pdf.text(125, 205, f"{ASSINATURA}")
        pdf.set_font ( "Arial" , size=10) 
        pdf.text(10, 250,f"Arquivo gerado automaticamente | Inteligência Grupo Imerge | {datetime.today().strftime('%Y-%m-%d %H:%M')}") 
        pdf.output("Solicitação_Proposta.pdf")
        
        
        
def resposta_proposta():
    if  (RESULTADO.loc["PREÇO TOTAL","% VARIAVEL"] + DESCONTO_EMP[LISTA_EMPREENDIMENTOS])  < 1:
        return f'REPROVADO!'
    elif VALOR_GARANTIDO_01  < OBJ_GAR1:
        return f'REVISAR'
    elif (VALOR_GARANTIDO_01 + CH_MORADIA) > VALOR_DO_LAUDO and POS_CHAVE != 0 and PRE_CHAVE != VALOR_GARANTIDO_01:
        return f'REVISAR LAUDO'
    else:
        return f'PRÉ APROVADO'


login = st.sidebar.text_input(label="Login",type="password")
if login == "":
    st.sidebar.write("Inserir a senha de acesso!")
elif login != "7lm2022":
    st.sidebar.write("Senha errada!")
else:
    st.sidebar.caption(f"Bem vindo {login}")
    
    with st.sidebar.expander("CONSULTA CPF::", expanded=False):
        acesso=st.text_input(label="Acesso Secreto",type="password")
        if acesso != "020386":
            st.caption(f"Acesso Negado")
        else:    
            cpf = st.number_input(label="CPF:",  format="%.0f")
            botao_002 = st.button("Consulta do CPF")

    with st.form(key="one"):
        with st.sidebar.expander("CIDADE/EMP/BLOCO/UNID::", expanded=False):
            df=pd.read_excel("Base_Preços.xlsx")
            DATA_HOJE = pd.to_datetime(date.today(),errors="coerce")
            DT_ENTREGA = {"VILA DO SOL":"2023-05-01","VILA DAS AGUAS":"2022-07-01","VILA AZALEIA":"2023-05-01", "VILA DAS ORQUÍDEAS":"2023-06-01","VILA DAS TULIPAS":"2023-12-01"}
            CIDADES = df.groupby("CIDADE")["UF"].count()
            CIDADES = pd.DataFrame(CIDADES)
            CIDADES = CIDADES.reset_index()
            CIDADES = list(CIDADES["CIDADE"].values[0:3].astype(str))
            CIDADES = st.selectbox("Escolha uma Cidade:", options=CIDADES)
            
            
            EMPREENDIMENTOS = df.groupby(["CIDADE","EMP"])["CIDADE"].count() 
            EMPREENDIMENTOS = pd.DataFrame(data=EMPREENDIMENTOS)
            EMPREENDIMENTOS.rename(columns={"CIDADE":"CID"}, inplace=True)
            EMPREENDIMENTOS = EMPREENDIMENTOS.reset_index(level='EMP', col_level=2)
            EMPREENDIMENTOS = EMPREENDIMENTOS.reset_index()
            EMPREENDIMENTOS = list(EMPREENDIMENTOS.loc[EMPREENDIMENTOS.CIDADE == CIDADES,"EMP"])
            LISTA_EMPREENDIMENTOS = st.selectbox("Escolha Empreendimento:", options=EMPREENDIMENTOS)
            
            
            BLOCO = df.groupby(["CIDADE","EMP","BLOCO"])["DESCONTO"].count()
            BLOCO = pd.DataFrame(BLOCO)
            BLOCO = BLOCO.reset_index()
            BLOCO = list(BLOCO.loc[(BLOCO.CIDADE == CIDADES) & (BLOCO.EMP == LISTA_EMPREENDIMENTOS),"BLOCO"])
            LISTA_BLOCOS = st.selectbox("Escolha o Bloco:", options=BLOCO)
                      
            UNIDADE = df.groupby(["CIDADE","EMP","BLOCO","UNIDADE"])["DESCONTO"].count()
            UNIDADE = pd.DataFrame(UNIDADE)
            UNIDADE = UNIDADE.reset_index()
            UNIDADE = list(UNIDADE.loc[(UNIDADE.CIDADE == CIDADES) & (UNIDADE.EMP == LISTA_EMPREENDIMENTOS) & (UNIDADE.BLOCO == LISTA_BLOCOS),"UNIDADE"])
            LISTA_UNIDADES = st.selectbox("Escolha a unidade:", options=UNIDADE)
            
            LAUDO = df.groupby(["CIDADE","EMP","BLOCO","UNIDADE","VALOR DO LAUDO"])["DESCONTO"].count()
            LAUDO = pd.DataFrame(LAUDO)
            LAUDO = LAUDO.reset_index()
            LAUDO = LAUDO.loc[(LAUDO.CIDADE == CIDADES) & (LAUDO.EMP == LISTA_EMPREENDIMENTOS) & (LAUDO.BLOCO == LISTA_BLOCOS) & (LAUDO.UNIDADE == LISTA_UNIDADES),"VALOR DO LAUDO"]

            VGV = df.groupby(["CIDADE","EMP","BLOCO","UNIDADE","VALOR DE VENDA"])["DESCONTO"].count()
            VGV = pd.DataFrame(VGV)
            VGV = VGV.reset_index()
            VGV = VGV.loc[(VGV.CIDADE == CIDADES) & (VGV.EMP == LISTA_EMPREENDIMENTOS) & (VGV.BLOCO == LISTA_BLOCOS) & (VGV.UNIDADE == LISTA_UNIDADES),"VALOR DE VENDA"]

            SITUACAO = df.groupby(["CIDADE","EMP","BLOCO","UNIDADE","VALOR DE VENDA","SITUAÇÃO"])["DESCONTO"].count()
            SITUACAO = pd.DataFrame(SITUACAO)
            SITUACAO = SITUACAO.reset_index()
            SITUACAO_001 = SITUACAO.loc[(SITUACAO.CIDADE == CIDADES) & (SITUACAO.EMP == LISTA_EMPREENDIMENTOS) & (SITUACAO.BLOCO == LISTA_BLOCOS) & (SITUACAO.UNIDADE == LISTA_UNIDADES),"SITUAÇÃO"]
 
            
            
        
        with st.sidebar.expander("SINAL E PRINCÍPIO DE PGTO", expanded=False):
            #st.subheader("PROPOSTA DE PGTO | TOTAL:")
            sinal = st.number_input(label="Valor da Parcela SINAL:",format="%.2f")
            dt_sinal = st.date_input(label="Vencimento da parcela do SINAL:")
            
        
        with st.sidebar.expander("PARCELAS MENSAIS 1", expanded=False):
            #st.subheader("CONDIÇÃO DE PGTO | PRÉ:")
            QTD_MENSAIS_001 = st.slider('Quantidade de parcelas MENSAIS 001', 0, 60, 1,format="%.0f")
            mensais1 = st.number_input(label="Valor das parcelas MENSAIS 001:",format="%.2f")
            data_mensais1 = st.date_input(label="Vencimento das parcelas MENSAIS 001:")
            DIF_001 = pd.to_datetime(DT_ENTREGA[LISTA_EMPREENDIMENTOS], errors="coerce" )  
            #st.subheader("//////////////////////////////////////////////////")
        
          
        with st.sidebar.expander("PARCELAS MENSAIS 2", expanded=False):
            QTD_MENSAIS_002 = st.slider('Quantidade de parcelas MENSAIS 002', 0, 60, 1,format="%.0f")
            mensais2 = st.number_input(label="Valor das parcelas MENSAIS 002:",format="%.2f")
            data_mensais2 = st.date_input(label="Vencimento das percelas MENSAIS 002:")
            #st.subheader("////////////////////////////////////////")
             
        with st.sidebar.expander("PARCELAS INTERMEDIÁRIAS 1", expanded=False):    
            QTD_INTER_001 = st.slider('Quantidade de Intermediárias 1:', 0, 6, 1,format="%.0f")
            TIPO_INTER_001 = st.selectbox('Qual o tipo da Intermediária 1:', options=["--","SEMESTRAIS", "ANUAIS"])
            INTER_001 = st.number_input(label=f"Valor das {TIPO_INTER_001} 1:",format="%.2f")
            DATA_INTER_001 = st.date_input(label=f"Vencimento das parcelas {TIPO_INTER_001} 1:")
            #st.subheader("//////////////////////////////////////////////////")
        
        with st.sidebar.expander("PARCELAS INTERMEDIÁRIAS 2", expanded=False):     
            QTD_INTER_002 = st.slider('Quantidade de Intermediárias 2:', 0, 6, 1,format="%.0f")
            TIPO_INTER_002 = st.selectbox('Qual o tipo da Intermediária 2:', options=["--","SEMESTRAIS", "ANUAIS"])
            INTER_002 = st.number_input(label=f"Valor das {TIPO_INTER_002} 2:",format="%.2f")
            DATA_INTER_002 = st.date_input(label=f"Vencimento das parcelas {TIPO_INTER_002} 2:")
            #st.subheader("CONDIÇÃO DE PGTO | PÓS:")
            
        with st.sidebar.expander("PARCELAS MENSAIS PÓS", expanded=False):    
            QTD_MENSAIS_001_POS = st.slider('Quantidade de parcelas MENSAIS PÓS', 0, 60, 1,format="%.0f")
            mensais1_POS = st.number_input(label="Valor das parcelas MENSAIS PÓS:",format="%.2f")
            data_mensais1_POS = st.date_input(label="Vencimento das parcelas MENSAIS PÓS:")
            #st.subheader("VALOR GARANTIDO:")
        
        with st.sidebar.expander("FIN/SUB/FGTS/CHEQUE", expanded=False):     
            FIN = st.number_input(label="VALOR FINANCIADO:",format="%.2f")
            SUBSIDIO = st.number_input(label="VALOR DO SUBSÍDIO:",format="%.2f")
            FGTS = st.number_input(label="VALOR DO FGTS:",format="%.2f")
            CH_MORADIA = st.number_input(label="VALOR DO CHEQUE MORADIA:",format="%.2f")
            
            
        VALOR_TOTAL_PROPOSTA = ((FIN + SUBSIDIO + FGTS + CH_MORADIA + sinal) + (QTD_MENSAIS_001 * mensais1) + (QTD_MENSAIS_002 * mensais2) + (INTER_001 * QTD_INTER_001) + (INTER_002*QTD_INTER_002) + (QTD_MENSAIS_001_POS * mensais1_POS))
        GAR1_EMP = {"VILA DO SOL":112500,"VILA DAS AGUAS":112500,"VILA AZALEIA":118000, "VILA DAS ORQUÍDEAS":115000,"VILA DAS TULIPAS":105000}
        GAR2_EMP = {"VILA DO SOL":120000,"VILA DAS AGUAS":114600,"VILA AZALEIA":125000, "VILA DAS ORQUÍDEAS":121000,"VILA DAS TULIPAS":115000}
        
        DT_ENTREGA = {"VILA DO SOL":"2023-05-01","VILA DAS AGUAS":"2022-07-01","VILA AZALEIA":"2023-05-01", "VILA DAS ORQUÍDEAS":"2023-06-01","VILA DAS TULIPAS":"2023-12-01"}
        
        
        BASE_CALCULO = pd.DataFrame(columns=["DATA","SINAL","MENSAIS","MENSAIS_2","MENSAIS_POS","SEMESTRAIS","ANUAIS","FINANC","EV_OBRA","EV_OBRA_1","TOTAL"], index=range(80))
        
        BASE_CALCULO.iloc[0,1] = sinal
        BASE_CALCULO.iloc[1,7] = FIN + SUBSIDIO + FGTS
         
         
        #TRATAMENTO DAS MENSAIS 001
                  
        DATA_MENSAIS_PRÉ = pd.to_datetime(data_mensais1,errors="coerce")
        DATA_TRAT_001 = pd.date_range(DATA_MENSAIS_PRÉ, periods=QTD_MENSAIS_001, freq="M")
        LISTA_DATA_TRAT_001 = pd.Series(range(len(DATA_TRAT_001)), index=DATA_TRAT_001)
        DATA_TRAT_002 = pd.DataFrame(LISTA_DATA_TRAT_001[:]).reset_index()
        DATA_TRAT_002.drop(columns=[0],inplace=True)
        DATA_TRAT_002.rename(columns={"index":"Data_Mensal"}, inplace=True)
        
        
        #TRATAMENTO DAS MENSAIS_002
        DATA_MENSAIS_PRÉ2 = pd.to_datetime(data_mensais2,errors="coerce")
        DATA_TRAT_100 = pd.date_range(DATA_MENSAIS_PRÉ2, periods=QTD_MENSAIS_002, freq="M")
        LISTA_DATA_TRAT_002 = pd.Series(range(len(DATA_TRAT_100)), index=DATA_TRAT_100)
        DATA_TRAT_003 = pd.DataFrame(LISTA_DATA_TRAT_002[:]).reset_index()
        DATA_TRAT_003.drop(columns=[0],inplace=True)
        DATA_TRAT_003.rename(columns={"index":"Data_Mensal"}, inplace=True)
        
        
        #TRATAMENTO DAS MENSAIS_PÓS
        DATA_MENSAIS_POS = pd.to_datetime(data_mensais1_POS,errors="coerce")
        DATA_TRAT_200 = pd.date_range(DATA_MENSAIS_POS, periods=QTD_MENSAIS_001_POS, freq="M")
        LISTA_DATA_TRAT_200 = pd.Series(range(len(DATA_TRAT_200)), index=DATA_TRAT_200)
        DATA_TRAT_200 = pd.DataFrame(LISTA_DATA_TRAT_200[:]).reset_index()
        DATA_TRAT_200.drop(columns=[0],inplace=True)
        DATA_TRAT_200.rename(columns={"index":"Data_Mensal"}, inplace=True)
        
              
        
        #TRATAMENTO DAS ANUAIS_001
        DATA_INTERMEDIADIA_001 = pd.to_datetime(DATA_INTER_001,errors="coerce")
        DATA_TRAT_300 = pd.date_range(DATA_INTERMEDIADIA_001, periods=QTD_INTER_001, freq="12M")
        LISTA_DATA_TRAT_300 = pd.Series(range(len(DATA_TRAT_300)), index=DATA_TRAT_300)
        DATA_TRAT_300 = pd.DataFrame(LISTA_DATA_TRAT_300[:]).reset_index()
        DATA_TRAT_300.drop(columns=[0],inplace=True)
        DATA_TRAT_300.rename(columns={"index":"Data_Mensal"}, inplace=True)
        
        #TRATAMENTO DAS SEMESTRAIS_001
        DATA_SEMESTRAIS_001 = pd.to_datetime(DATA_INTER_001,errors="coerce")
        DATA_TRAT_400 = pd.date_range(DATA_SEMESTRAIS_001, periods=QTD_INTER_001, freq="6M")
        LISTA_DATA_TRAT_400 = pd.Series(range(len(DATA_TRAT_400)), index=DATA_TRAT_400)
        DATA_TRAT_400 = pd.DataFrame(LISTA_DATA_TRAT_400[:]).reset_index()
        DATA_TRAT_400.drop(columns=[0],inplace=True)
        DATA_TRAT_400.rename(columns={"index":"Data_Mensal"}, inplace=True)
        
        
        
        #TRATAMENTO DAS ANUAIS_002
        DATA_INTERMEDIADIA_002 = pd.to_datetime(DATA_INTER_002,errors="coerce")
        DATA_TRAT_301 = pd.date_range(DATA_INTERMEDIADIA_002, periods=QTD_INTER_002, freq="12M")
        LISTA_DATA_TRAT_301 = pd.Series(range(len(DATA_TRAT_301)), index=DATA_TRAT_301)
        DATA_TRAT_301 = pd.DataFrame(LISTA_DATA_TRAT_301[:]).reset_index()
        DATA_TRAT_301.drop(columns=[0],inplace=True)
        DATA_TRAT_301.rename(columns={"index":"Data_Mensal"}, inplace=True)
        
        #TRATAMENTO DAS SEMESTRAIS_002
        DATA_SEMESTRAIS_002 = pd.to_datetime(DATA_INTER_002,errors="coerce")
        DATA_TRAT_401 = pd.date_range(DATA_SEMESTRAIS_002, periods=QTD_INTER_002, freq="6M")
        LISTA_DATA_TRAT_401 = pd.Series(range(len(DATA_TRAT_401)), index=DATA_TRAT_401)
        DATA_TRAT_401 = pd.DataFrame(LISTA_DATA_TRAT_401[:]).reset_index()
        DATA_TRAT_401.drop(columns=[0],inplace=True)
        DATA_TRAT_401.rename(columns={"index":"Data_Mensal"}, inplace=True)
        
        
        DATA_TRAT = pd.date_range("2022-03-31", periods=100, freq="M")
        LISTA_DATA_TRAT = pd.Series(range(len(DATA_TRAT)), index=DATA_TRAT)
        DATA_TRAT_001 = pd.DataFrame(LISTA_DATA_TRAT[:]).reset_index()
        DATA_TRAT_001.drop(columns=[0],inplace=True)
        DATA_TRAT_001["index"] = pd.to_datetime(DATA_TRAT_001["index"], errors="coerce")

        cont = 0
        limite=99
        while cont <= limite:
            BASE_CALCULO.loc[cont:cont,"DATA"] = DATA_TRAT_001["index"][cont]
            cont=cont+1
            BASE_CALCULO["DATA"] = pd.to_datetime(BASE_CALCULO["DATA"], errors="coerce")
            
        for i in BASE_CALCULO["DATA"]:
            if QTD_MENSAIS_001 > 0:  
                if i >= DATA_TRAT_002["Data_Mensal"][0] and i <= DATA_TRAT_002["Data_Mensal"][QTD_MENSAIS_001-1]:
                    BASE_CALCULO.loc[(BASE_CALCULO.DATA >= DATA_TRAT_002["Data_Mensal"][0]) & (BASE_CALCULO.DATA <= DATA_TRAT_002["Data_Mensal"][QTD_MENSAIS_001-1]) ,"MENSAIS"] = mensais1
                    
                    
        for i in BASE_CALCULO["DATA"]:
            if QTD_MENSAIS_002 > 0:  
                if i >= DATA_TRAT_003["Data_Mensal"][0] and i <= DATA_TRAT_003["Data_Mensal"][QTD_MENSAIS_002-1]:
                    BASE_CALCULO.loc[(BASE_CALCULO.DATA >= DATA_TRAT_003["Data_Mensal"][0]) & (BASE_CALCULO.DATA <= DATA_TRAT_003["Data_Mensal"][QTD_MENSAIS_002-1]) ,"MENSAIS_2"] = mensais2
                                
        
        for i in BASE_CALCULO["DATA"]:
            if QTD_MENSAIS_001_POS > 0:  
                if i >= DATA_TRAT_200["Data_Mensal"][0] and i <= DATA_TRAT_200["Data_Mensal"][QTD_MENSAIS_001_POS-1]:
                    BASE_CALCULO.loc[(BASE_CALCULO.DATA >= DATA_TRAT_200["Data_Mensal"][0]) & (BASE_CALCULO.DATA <= DATA_TRAT_200["Data_Mensal"][QTD_MENSAIS_001_POS-1]) ,"MENSAIS_POS"] = mensais1_POS
                                
        
        # INTERMEDIÁRIAS ANUAIS --------------------------------------------------------------------------------------------------------
        for i in BASE_CALCULO["DATA"]:
            if QTD_INTER_001 >= 1 and TIPO_INTER_001 == "ANUAIS":  
                if i >= DATA_TRAT_300["Data_Mensal"][0] and i <= DATA_TRAT_300["Data_Mensal"][0]:
                    BASE_CALCULO.loc[(BASE_CALCULO.DATA >= DATA_TRAT_300["Data_Mensal"][0]) & (BASE_CALCULO.DATA <= DATA_TRAT_300["Data_Mensal"][0]) ,"ANUAIS"] = INTER_001
                
        for i in BASE_CALCULO["DATA"]:
            if QTD_INTER_001 >= 2 and TIPO_INTER_001 == "ANUAIS":  
                if i >= DATA_TRAT_300["Data_Mensal"][1] and i <= DATA_TRAT_300["Data_Mensal"][1]:
                    BASE_CALCULO.loc[(BASE_CALCULO.DATA >= DATA_TRAT_300["Data_Mensal"][1]) & (BASE_CALCULO.DATA <= DATA_TRAT_300["Data_Mensal"][1]) ,"ANUAIS"] = INTER_001        
        
        for i in BASE_CALCULO["DATA"]:
            if QTD_INTER_001 >= 3 and TIPO_INTER_001 == "ANUAIS":  
                if i >= DATA_TRAT_300["Data_Mensal"][2] and i <= DATA_TRAT_300["Data_Mensal"][2]:
                    BASE_CALCULO.loc[(BASE_CALCULO.DATA >= DATA_TRAT_300["Data_Mensal"][2]) & (BASE_CALCULO.DATA <= DATA_TRAT_300["Data_Mensal"][2]) ,"ANUAIS"] = INTER_001        

        for i in BASE_CALCULO["DATA"]:
            if QTD_INTER_001 >= 4 and TIPO_INTER_001 == "ANUAIS":  
                if i >= DATA_TRAT_300["Data_Mensal"][3] and i <= DATA_TRAT_300["Data_Mensal"][3]:
                    BASE_CALCULO.loc[(BASE_CALCULO.DATA >= DATA_TRAT_300["Data_Mensal"][3]) & (BASE_CALCULO.DATA <= DATA_TRAT_300["Data_Mensal"][3]) ,"ANUAIS"] = INTER_001        


        # INTERMEDIÁRIAS SEMESTRAIS --------------------------------------------------------------------------------------------------------
        for i in BASE_CALCULO["DATA"]:
            if QTD_INTER_001 >= 1 and TIPO_INTER_001 == "SEMESTRAIS":  
                if i >= DATA_TRAT_300["Data_Mensal"][0] and i <= DATA_TRAT_300["Data_Mensal"][0]:
                    BASE_CALCULO.loc[(BASE_CALCULO.DATA >= DATA_TRAT_300["Data_Mensal"][0]) & (BASE_CALCULO.DATA <= DATA_TRAT_300["Data_Mensal"][0]) ,"SEMESTRAIS"] = INTER_001
                
        for i in BASE_CALCULO["DATA"]:
            if QTD_INTER_001 >= 2 and TIPO_INTER_001 == "SEMESTRAIS":  
                if i >= DATA_TRAT_300["Data_Mensal"][1] and i <= DATA_TRAT_300["Data_Mensal"][1]:
                    BASE_CALCULO.loc[(BASE_CALCULO.DATA >= DATA_TRAT_300["Data_Mensal"][1]) & (BASE_CALCULO.DATA <= DATA_TRAT_300["Data_Mensal"][1]) ,"SEMESTRAIS"] = INTER_001        
        
        for i in BASE_CALCULO["DATA"]:
            if QTD_INTER_001 >= 3 and TIPO_INTER_001 == "SEMESTRAIS":  
                if i >= DATA_TRAT_300["Data_Mensal"][2] and i <= DATA_TRAT_300["Data_Mensal"][2]:
                    BASE_CALCULO.loc[(BASE_CALCULO.DATA >= DATA_TRAT_300["Data_Mensal"][2]) & (BASE_CALCULO.DATA <= DATA_TRAT_300["Data_Mensal"][2]) ,"SEMESTRAIS"] = INTER_001        

        for i in BASE_CALCULO["DATA"]:
            if QTD_INTER_001 >= 4 and TIPO_INTER_001 == "SEMESTRAIS":  
                if i >= DATA_TRAT_300["Data_Mensal"][3] and i <= DATA_TRAT_300["Data_Mensal"][3]:
                    BASE_CALCULO.loc[(BASE_CALCULO.DATA >= DATA_TRAT_300["Data_Mensal"][3]) & (BASE_CALCULO.DATA <= DATA_TRAT_300["Data_Mensal"][3]) ,"SEMESTRAIS"] = INTER_001        


        for i in BASE_CALCULO["DATA"]:
            if QTD_INTER_001 == 5 and TIPO_INTER_001 == "ANUAIS":  
                if i >= DATA_TRAT_300["Data_Mensal"][4] and i <= DATA_TRAT_300["Data_Mensal"][4]:
                    BASE_CALCULO.loc[(BASE_CALCULO.DATA >= DATA_TRAT_300["Data_Mensal"][4]) & (BASE_CALCULO.DATA <= DATA_TRAT_300["Data_Mensal"][4]) ,"ANUAIS"] = INTER_001        


        
        # INTERMEDIÁRIAS_002 ANUAIS --------------------------------------------------------------------------------------------------------
        for i in BASE_CALCULO["DATA"]:
            if QTD_INTER_002 >= 1 and QTD_INTER_002 == "ANUAIS":  
                if i >= DATA_TRAT_301["Data_Mensal"][0] and i <= DATA_TRAT_301["Data_Mensal"][0]:
                    BASE_CALCULO.loc[(BASE_CALCULO.DATA >= DATA_TRAT_301["Data_Mensal"][0]) & (BASE_CALCULO.DATA <= DATA_TRAT_301["Data_Mensal"][0]) ,"ANUAIS"] = INTER_002
                
        for i in BASE_CALCULO["DATA"]:
            if QTD_INTER_002 >= 2 and TIPO_INTER_002 == "ANUAIS":  
                if i >= DATA_TRAT_301["Data_Mensal"][1] and i <= DATA_TRAT_301["Data_Mensal"][1]:
                    BASE_CALCULO.loc[(BASE_CALCULO.DATA >= DATA_TRAT_301["Data_Mensal"][1]) & (BASE_CALCULO.DATA <= DATA_TRAT_301["Data_Mensal"][1]) ,"ANUAIS"] = INTER_002       
        
        for i in BASE_CALCULO["DATA"]:
            if QTD_INTER_002 >= 3 and TIPO_INTER_002 == "ANUAIS":  
                if i >= DATA_TRAT_301["Data_Mensal"][2] and i <= DATA_TRAT_301["Data_Mensal"][2]:
                    BASE_CALCULO.loc[(BASE_CALCULO.DATA >= DATA_TRAT_301["Data_Mensal"][2]) & (BASE_CALCULO.DATA <= DATA_TRAT_301["Data_Mensal"][2]) ,"ANUAIS"] = INTER_002        

        for i in BASE_CALCULO["DATA"]:
            if QTD_INTER_002 >= 4 and TIPO_INTER_002 == "ANUAIS":  
                if i >= DATA_TRAT_301["Data_Mensal"][3] and i <= DATA_TRAT_301["Data_Mensal"][3]:
                    BASE_CALCULO.loc[(BASE_CALCULO.DATA >= DATA_TRAT_301["Data_Mensal"][3]) & (BASE_CALCULO.DATA <= DATA_TRAT_301["Data_Mensal"][3]) ,"ANUAIS"] = INTER_002        


        # INTERMEDIÁRIAS_002 SEMESTRAIS --------------------------------------------------------------------------------------------------------
        for i in BASE_CALCULO["DATA"]:
            if QTD_INTER_002 >= 1 and TIPO_INTER_002 == "SEMESTRAIS":  
                if i >= DATA_TRAT_401["Data_Mensal"][0] and i <= DATA_TRAT_401["Data_Mensal"][0]:
                    BASE_CALCULO.loc[(BASE_CALCULO.DATA >= DATA_TRAT_401["Data_Mensal"][0]) & (BASE_CALCULO.DATA <= DATA_TRAT_401["Data_Mensal"][0]) ,"SEMESTRAIS"] = INTER_002
                
        for i in BASE_CALCULO["DATA"]:
            if QTD_INTER_002 >= 2 and TIPO_INTER_002 == "SEMESTRAIS":  
                if i >= DATA_TRAT_401["Data_Mensal"][1] and i <= DATA_TRAT_401["Data_Mensal"][1]:
                    BASE_CALCULO.loc[(BASE_CALCULO.DATA >= DATA_TRAT_401["Data_Mensal"][1]) & (BASE_CALCULO.DATA <= DATA_TRAT_401["Data_Mensal"][1]) ,"SEMESTRAIS"] = INTER_002        
        
        for i in BASE_CALCULO["DATA"]:
            if QTD_INTER_002 >= 3 and TIPO_INTER_002 == "SEMESTRAIS":  
                if i >= DATA_TRAT_401["Data_Mensal"][2] and i <= DATA_TRAT_401["Data_Mensal"][2]:
                    BASE_CALCULO.loc[(BASE_CALCULO.DATA >= DATA_TRAT_401["Data_Mensal"][2]) & (BASE_CALCULO.DATA <= DATA_TRAT_401["Data_Mensal"][2]) ,"SEMESTRAIS"] = INTER_002        

        for i in BASE_CALCULO["DATA"]:
            if QTD_INTER_002 >= 4 and TIPO_INTER_002 == "SEMESTRAIS":  
                if i >= DATA_TRAT_401["Data_Mensal"][3] and i <= DATA_TRAT_401["Data_Mensal"][3]:
                    BASE_CALCULO.loc[(BASE_CALCULO.DATA >= DATA_TRAT_401["Data_Mensal"][3]) & (BASE_CALCULO.DATA <= DATA_TRAT_401["Data_Mensal"][3]) ,"SEMESTRAIS"] = INTER_002        


        for i in BASE_CALCULO["DATA"]:
            if QTD_INTER_002 == 5 and TIPO_INTER_002 == "ANUAIS":  
                if i >= DATA_TRAT_401["Data_Mensal"][4] and i <= DATA_TRAT_401["Data_Mensal"][4]:
                    BASE_CALCULO.loc[(BASE_CALCULO.DATA >= DATA_TRAT_401["Data_Mensal"][4]) & (BASE_CALCULO.DATA <= DATA_TRAT_401["Data_Mensal"][4]) ,"ANUAIS"] = INTER_002        

        
      
                
        DT_ENTREGA_TRAT =  pd.DataFrame(columns=["DATA_ENTREGA"],index=[pd.to_datetime(DT_ENTREGA[LISTA_EMPREENDIMENTOS], errors="coerce")]).reset_index()
        DT_ENTREGA_TRAT.drop(columns=["DATA_ENTREGA"], inplace=True)
        DT_ENTREGA_TRAT.rename(columns={"index":"Data_Entrega"}, inplace=True)
                      
       
        BASE_CALCULO["TIPO"] = "PÓS"
        BASE_CALCULO.loc[(BASE_CALCULO["DATA"] < DT_ENTREGA_TRAT["Data_Entrega"][0]),"TIPO"] = "PRÉ"
                
            
        BASE_CALCULO["TOTAL"] = BASE_CALCULO.loc[0:99,["SINAL","MENSAIS","MENSAIS_2","MENSAIS_POS","SEMESTRAIS","ANUAIS", "FINANC"]].sum(axis = 1)
        BASE_CALCULO.loc[BASE_CALCULO.FINANC >0,"TIPO"] = "FIN"
        BASE_CALCULO["DATA"] = pd.to_datetime(BASE_CALCULO["DATA"], errors="coerce").dt.strftime('%m-%Y')
        
     

        
        botao_001 = st.form_submit_button("SIMULAR")

    
    
    if botao_001:
        
        RESULTADO = pd.DataFrame(columns=["OBJETIVO", "REALIZADO", "VARIAVEL","% VARIAVEL"], index=["PREÇO TOTAL","TOTAL PRÉ","TOTAL PÓS","CHEQUE MORADIA","GARANTIDO 1"])

        col1,col2,col3 = st.columns(3)
        col1.metric(label="EMPREENDIMENTO", value=LISTA_EMPREENDIMENTOS)
        col2.metric(label="BLOCO", value=LISTA_BLOCOS)
        col3.metric(label="UNIDADE", value=LISTA_UNIDADES)
        
        PRE_EMP = {"VILA DO SOL":0.03,"VILA DAS AGUAS":0.01,"VILA AZALEIA":0.05, "VILA DAS ORQUÍDEAS":0.04,"VILA DAS TULIPAS":0.04}
        POS_EMP = {"VILA DO SOL":0.07,"VILA DAS AGUAS":0.09,"VILA AZALEIA":0.10, "VILA DAS ORQUÍDEAS":0.08,"VILA DAS TULIPAS":0.06}

        VALOR_DO_LAUDO = float(LAUDO)
        VALOR_MAX_CHEQUE = float(42000)
        VALOR_GARANTIDO_01 = float(FIN + FGTS + SUBSIDIO + sinal)
        PRE_CHAVE = BASE_CALCULO.loc[BASE_CALCULO["TIPO"] == "PRÉ","TOTAL"].sum(axis = 0)
        POS_CHAVE = BASE_CALCULO.loc[BASE_CALCULO["TIPO"] == "PÓS","TOTAL"].sum(axis = 0)
        DIF_LAUDO_X_CHEQUE = (VALOR_DO_LAUDO - VALOR_GARANTIDO_01)
        VALOR_TOTAL_AJUSTADO = float(VGV) + (VALOR_DO_LAUDO*PRE_EMP[LISTA_EMPREENDIMENTOS]) + (VALOR_DO_LAUDO*POS_EMP[LISTA_EMPREENDIMENTOS])
        OBJ_PRE = VALOR_DO_LAUDO*PRE_EMP[LISTA_EMPREENDIMENTOS]
        OBJ_POS = VALOR_DO_LAUDO*POS_EMP[LISTA_EMPREENDIMENTOS]
        
       
        if DIF_LAUDO_X_CHEQUE < 0:
            CHEQUE_MORADIA_TOTAL = 0.0
        elif DIF_LAUDO_X_CHEQUE <= VALOR_MAX_CHEQUE:
            CHEQUE_MORADIA_TOTAL = DIF_LAUDO_X_CHEQUE
        else:
            CHEQUE_MORADIA_TOTAL = VALOR_MAX_CHEQUE
            
        if CHEQUE_MORADIA_TOTAL == 0:
            RESULTADO_DO_CHEQUE = CHEQUE_MORADIA_TOTAL  
        else:
            RESULTADO_DO_CHEQUE = (CH_MORADIA / CHEQUE_MORADIA_TOTAL)
            
            
            
        
        OBJ_GAR1 = VALOR_DO_LAUDO - CHEQUE_MORADIA_TOTAL
                    
        RESULTADO.loc["PREÇO TOTAL","OBJETIVO"] = VALOR_TOTAL_AJUSTADO
        RESULTADO.loc["PREÇO TOTAL","REALIZADO"] = VALOR_TOTAL_PROPOSTA
        RESULTADO.loc["PREÇO TOTAL","VARIAVEL"] = VALOR_TOTAL_PROPOSTA - VALOR_TOTAL_AJUSTADO
        RESULTADO.loc["PREÇO TOTAL","% VARIAVEL"] = VALOR_TOTAL_PROPOSTA / VALOR_TOTAL_AJUSTADO
    
        
        RESULTADO.loc["TOTAL PRÉ","OBJETIVO"] = OBJ_PRE
        RESULTADO.loc["TOTAL PRÉ","REALIZADO"] = PRE_CHAVE
        RESULTADO.loc["TOTAL PRÉ","VARIAVEL"] = PRE_CHAVE - OBJ_PRE
        RESULTADO.loc["TOTAL PRÉ","% VARIAVEL"] = (PRE_CHAVE / OBJ_PRE)
        
        RESULTADO.loc["TOTAL PÓS","OBJETIVO"] = OBJ_POS
        RESULTADO.loc["TOTAL PÓS","REALIZADO"] = POS_CHAVE
        RESULTADO.loc["TOTAL PÓS","VARIAVEL"] =  POS_CHAVE - OBJ_POS
        RESULTADO.loc["TOTAL PÓS","% VARIAVEL"] = (POS_CHAVE / OBJ_POS)
        
        RESULTADO.loc["GARANTIDO 1","OBJETIVO"] = OBJ_GAR1
        RESULTADO.loc["GARANTIDO 1","REALIZADO"] = VALOR_GARANTIDO_01
        RESULTADO.loc["GARANTIDO 1","VARIAVEL"] = VALOR_GARANTIDO_01 - OBJ_GAR1
        RESULTADO.loc["GARANTIDO 1","% VARIAVEL"] = (VALOR_GARANTIDO_01 / OBJ_GAR1)
        
        RESULTADO.loc["CHEQUE MORADIA","OBJETIVO"] = CHEQUE_MORADIA_TOTAL
        RESULTADO.loc["CHEQUE MORADIA","REALIZADO"] = CH_MORADIA
        RESULTADO.loc["CHEQUE MORADIA","VARIAVEL"] = CH_MORADIA - CHEQUE_MORADIA_TOTAL
        RESULTADO.loc["CHEQUE MORADIA","% VARIAVEL"] = RESULTADO_DO_CHEQUE
        
        
        RESULTADO.fillna(value=0, inplace=True)
        pd.options.display.latex.repr = True
        
        st.dataframe(RESULTADO.style.format(subset=['OBJETIVO',"REALIZADO","VARIAVEL","% VARIAVEL"], formatter="{:.2f}"))
        
        DESCONTO_EMP = {"VILA DO SOL":0.01,"VILA DAS AGUAS":0.03,"VILA AZALEIA":0.00, "VILA DAS ORQUÍDEAS":0.02,"VILA DAS TULIPAS":0.00}
        
        def resposta_proposta():
            if  (RESULTADO.loc["PREÇO TOTAL","% VARIAVEL"] + DESCONTO_EMP[LISTA_EMPREENDIMENTOS])  < 1:
                return f'REPROVADO!'
            elif VALOR_GARANTIDO_01  < OBJ_GAR1:
                return f'REVISAR'
            elif (VALOR_GARANTIDO_01 + CH_MORADIA) > VALOR_DO_LAUDO and POS_CHAVE != 0 and PRE_CHAVE != VALOR_GARANTIDO_01:
                return f'REVISAR LAUDO'
            else:
                return f'PRÉ APROVADO'
        
        col1, col2, col3 = st.columns(3)
        col1.metric(label="STATUS DE APROVAÇÃO:", value=resposta_proposta())
        col2.metric(label="STATUS UNIDADE:", value=str(SITUACAO_001)[5:18])
        col3.metric(label="$_VAR:", value=np.round(VALOR_TOTAL_PROPOSTA - VALOR_TOTAL_AJUSTADO,2), delta=np.round(-(VALOR_TOTAL_PROPOSTA / VALOR_TOTAL_AJUSTADO),2))

        BASE_CALCULO.fillna(value=0, inplace=True)
        st.dataframe(BASE_CALCULO.style.format(subset=["SINAL","MENSAIS","MENSAIS_2","MENSAIS_POS","SEMESTRAIS","ANUAIS", "FINANC","EV_OBRA","EV_OBRA_1","TOTAL"], formatter="{:.2f}"))
        
        add_image(Imagem_data,imagem_Logo,LISTA_EMPREENDIMENTOS,LISTA_BLOCOS,LISTA_UNIDADES,resposta_proposta())  
        
    if botao_003:
        enviar_email_002(CIDADES, LISTA_EMPREENDIMENTOS,LISTA_BLOCOS,LISTA_UNIDADES)
        st.write("Arquivo Gerado!")
        #add_image(Imagem_data,imagem_Logo,LISTA_EMPREENDIMENTOS,LISTA_BLOCOS,LISTA_UNIDADES,resposta_proposta())   

        
        

    if acesso == "020386" and botao_002:
        if cpf !="":
            acesso = "https://caixaaqui.caixa.gov.br/caixaaqui/CaixaAquiController/index"
            chrome_options = webdriver.ChromeOptions()
            #chrome_options.add_argument('--headless')
            driver = webdriver.Chrome(r'C:\Users\carlos.baracho\.conda\chromedriver.exe')#,chrome_options=chrome_options)
            
            driver.get(acesso)
            st.write("ATENÇÃO!!!! NÃO MEXA NO MOUSE E TECLADO!!!")
            st.write("Aguarde!!!")
            
            time.sleep(1)
            driver.find_element_by_id('convenio').click()
            time.sleep(1)
            driver.find_element_by_id('convenio').send_keys('000317314')
            time.sleep(1)
            driver.find_element_by_id('login').click()
            time.sleep(1)
            driver.find_element_by_id('login').send_keys('Michelle')
            time.sleep(1)
            driver.find_element_by_id('password').click()
            time.sleep(1)
            driver.find_element_by_id('password').send_keys('Prime028')
            time.sleep(3)
            driver.find_element_by_id('btLogin').click()
            time.sleep(3)
            st.write("1. Passo: Logado!")
            driver.find_element_by_link_text('Serviços ao Cliente').click()
            time.sleep(3)
            driver.find_element_by_link_text('Negócios').click()
            time.sleep(3)
            driver.find_element_by_link_text('Pesquisar Clientes').click()
            time.sleep(3)
            driver.find_element_by_id('dataCpf').send_keys(cpf)
            time.sleep(3)
            driver.find_element_by_link_text('Consultar').click()
            st.write("2. Passo: Consulta Feita!")
            time.sleep(3)
            webdriver.ActionChains(driver).key_down(Keys.CONTROL).send_keys("a").perform()
            time.sleep(3)
            copiar=webdriver.ActionChains(driver).key_down(Keys.CONTROL).send_keys("c").perform()
            time.sleep(3)
            texto = clipboard.paste()
            clipboard.copy("")
            Consulta = texto.split("\n")[9][0:6]
            if Consulta == "SERASA":
                st.write("Cliente Negativado")
                driver.close()
            else:
                st.write("Nada Consta")
                driver.close()

        else:
            st.write("PREENCHA O CAMPO CPF!!!")
                                  

     
