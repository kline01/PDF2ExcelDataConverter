import pdfquery 
import pandas as pd
import os
from openpyxl import load_workbook

NotaFiscal = {"Cnpj prestador": [],
            "prestador": [],
            "Número Nota": [],
            "Data de Emissão": [],
            "Tomador": [],
            "Cnpj Tomador": [],
            "Valor": [] ,
            "Base": []}

colunas = ["Cnpj prestador",
           "prestador",
            "Número Nota",
            "Data de Emissão",
            "Tomador",
            "Cnpj Tomador",
            "Valor",
            "Base"]

def listar_itens_pasta(caminho):
    if os.path.isdir(caminho):
        arquivos = [f for f in os.listdir(caminho) if os.path.isfile(os.path.join(caminho, f))]
        return arquivos

def write_to_excel(df):

        #endereço da planilha que cola as informações
        caminho = r"C:\Users\User\Downloads\notas\NotasFiscais.xlsx" 
        #ler os dados atuais da planilha
                
        #adiciona o dicionário da variável pdata
        pdata = pd.DataFrame.from_dict(NotaFiscal)
        #coloca os valores no dicionário
        pdata = pdata[colunas]
        #inicia o Escritor na planilha 'NotasFiscais'
        writer = pd.ExcelWriter(caminho, engine='openpyxl', mode='a', if_sheet_exists='replace')
        #Escrever os itens na pasta 'Main', sem index começando na próxima linha vazia
        pdata.to_excel(writer,'Main' ,index=False,header=True)
        writer.close()

def lerPDF_SALVADOR(arquivo):
    pdf = pdfquery.PDFQuery(r"C:\Users\User\Downloads\{}".format(arquivo))
    pdf.load()
    text_elements = pdf.pq('LTTextLineHorizontal')
    if len(text_elements) >= 1:
        cnpj_prestador = pdf.pq('LTTextLineHorizontal:in_bbox("12.72, 633.865, 93.685, 642.848")').text()
        prestador = pdf.pq('LTTextLineHorizontal:in_bbox("12.72, 606.025, 167.032, 615.008")').text()
        nf = text_elements[3].text
        data = text_elements[5].text
        Tomador = text_elements[23].text
        cnpj_Tomador = text_elements[25].text
        #valor = text_elements[43].text
        base = pdf.pq('LTTextLineHorizontal:in_bbox("223.56, 161.545, 263.755, 170.528")').text()
        valor = pdf.pq('LTTextLineHorizontal:in_bbox("228.6, 161.545, 263.799, 170.528")').text()
        
    else:
        print("Element at index 1 not found.")


    NotaFiscal["Cnpj prestador"].append(cnpj_prestador)
    NotaFiscal["prestador"].append(prestador)
    NotaFiscal["Número Nota"].append(nf)
    NotaFiscal["Data de Emissão"].append(data)
    NotaFiscal["Tomador"].append(Tomador)
    NotaFiscal["Cnpj Tomador"].append(cnpj_Tomador)
    NotaFiscal["Valor"].append(valor)
    NotaFiscal["Base"].append(base)


pasta_salvador = r"C:\Users\User\Downloads\notas"


lista_arquivos = listar_itens_pasta(pasta_salvador)
print(len(lista_arquivos))

for i in range(0, len(lista_arquivos)):
    lerPDF_SALVADOR(lista_arquivos[i])
    print(NotaFiscal)

write_to_excel(NotaFiscal)