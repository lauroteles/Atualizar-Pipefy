import pandas as pd
import streamlit as st
import openpyxl as op
import os
class Pipefy():

    def __init__(self,caminho_dados):
        os.chdir(caminho_dados)

    def compilando_planilha_controle(self):

        lista_de_paginas = ['BTG', 'Guide', 'Genial', 'Ágora', 'Órama']
        arquivo_controle = []

        for pag in lista_de_paginas:
            pagina_x = pd.read_excel(r'Z:\Mesa de Operações\Pipefy\Controle de Contratos - Atualizado Janeiro de 2024.xlsx', pag,skiprows=1)
            pagina_x = pagina_x.iloc[:-5,:]
            arquivo_controle.append(pagina_x)
        controle = pd.concat(arquivo_controle)
        return controle

    def atualizando_operadortes(self,df_controle):
        
        pipefy_id = pd.read_excel(r'Z:\Mesa de Operações\Pipefy\new_report_07-02-2024.xlsx')
        df_controle = df_controle.iloc[:,[2,6]].reset_index()
        pipefy_id = pipefy_id.iloc[:,[2,21]]

        pipefy_id['Conta'] = pipefy_id['Conta'].astype(str)
        df_controle['Conta'] = df_controle['Conta'].astype(str).str.replace('.0','')

        pipe_operador_atualizado = pd.merge(pipefy_id,df_controle,on='Conta',how='inner').drop(columns='index')
        return pipe_operador_atualizado


if __name__ =="__main__":

    pipefy = Pipefy(caminho_dados=r'Z:\Mesa de Operações\Pipefy\Operadores_mes_atual')
    
    controle = pipefy.compilando_planilha_controle()
    pipe_atualizar = pipefy.atualizando_operadortes(df_controle=controle)
    pipe_atualizar.to_excel('Operadores_07/02/2024.xlsx')
    










# pipefy_id = pd.read_excel('new_report_07-02-2024.xlsx')
# controle = controle.iloc[:,[2,6]].reset_index()
# pipefy_id = pipefy_id.iloc[:,[2,21]]

# pipefy_id['Conta'] = pipefy_id['Conta'].astype(str)
# controle['Conta'] = controle['Conta'].astype(str).str.replace('.0','')

# pipe_novo_operador = pd.merge(pipefy_id,controle,on='Conta',how='inner').drop(columns='index')



# st.dataframe(pipe_novo_operador)
# st.dataframe(controle)
# st.dataframe(pipefy_id)







