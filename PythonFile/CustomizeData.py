'''Desenvolver um Script em python que ao final do processo será chamado
dentro da sua automação para ler a planilha, transformar o texto das
descrições para minúsculo e remover acentos das colunas de texto e remover
tudo que for diferente de número das colunas de códigos exceto da coluna
Código Secção'''

import pandas as pd

def CustomizeData(File):

    try:
        #Lendo arquivo gerado pelo robo

        df = pd.read_excel(File, engine = 'openpyxl')

        #Mudando todas as colunas para letras minusculas

        df = df.apply(lambda x: x.astype(str).str.lower())

        #Retirando a acentuação dos campos

        cols = df.select_dtypes(include=[object]).columns
        df[cols] = df[cols].apply(lambda x: x.str.normalize('NFKD').str.encode('ascii', errors='ignore').str.decode('utf-8'))

        #Retirando tudo que é diferente de número dos campos 'Códigos', exceto 'Código Secção'

        df['Código Divisão'] = df['Código Divisão'].str.replace(r'[^0-9]+','', regex=True)
        df['Código Grupo'] = df['Código Grupo'].str.replace(r'[^0-9]+','', regex=True)
        df['Código Classe'] = df['Código Classe'].str.replace(r'[^0-9]+','', regex=True)
        df['Código SubClasse'] = df['Código SubClasse'].str.replace(r'[^0-9]+','', regex=True)

        #Salvando arquivo

        df.to_excel(excel_writer=str(File), index=False)

        return ('Sucesso')

    except Exception as e: return str(e)


