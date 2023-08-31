import pdfquery 
import pandas as pd
import os
from openpyxl import load_workbook

NotaFiscal = {"Cnpj Unidade": [],
            "Unidade": [],
            "Número Nota": [],
            "Data de Emissão": [],
            "Operadora": [],
            "Cnpj Operadora": [],
            "Valor": [] ,
            "Base": []}

colunas = ["Cnpj Unidade",
           "Unidade",
            "Número Nota",
            "Data de Emissão",
            "Operadora",
            "Cnpj Operadora",
            "Valor",
            "Base"]

def listar_itens_pasta(caminho):
    if os.path.isdir(caminho):
        arquivos = [f for f in os.listdir(caminho) if os.path.isfile(os.path.join(caminho, f))]
        return arquivos

def write_to_excel(df):

        #endereço da planilha que cola as informações
        caminho = r"C:\\Users\\kevin.soares\\Desktop\\Projetos python\\NotasFiscais.xlsx" 
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
    pdf = pdfquery.PDFQuery(r"C:\Users\kevin.soares\Downloads\Cárdio Pulmonar = Nota Fiscal convênio Asfeb vencimento_ 15_09_2023 (1)\{}".format(arquivo))
    pdf.load()
    text_elements = pdf.pq('LTTextLineHorizontal')
    if len(text_elements) >= 1:
        cnpj_unidade = pdf.pq('LTTextLineHorizontal:in_bbox("12.72, 633.865, 93.685, 642.848")').text()
        unidade = pdf.pq('LTTextLineHorizontal:in_bbox("12.72, 606.025, 167.032, 615.008")').text()
        nf = text_elements[3].text
        data = text_elements[5].text
        operadora = text_elements[23].text
        cnpj_operadora = text_elements[25].text
        #valor = text_elements[43].text
        base = pdf.pq('LTTextLineHorizontal:in_bbox("223.56, 161.545, 263.755, 170.528")').text()
        valor = pdf.pq('LTTextLineHorizontal:in_bbox("228.6, 161.545, 263.799, 170.528")').text()
        
    else:
        print("Element at index 1 not found.")


    NotaFiscal["Cnpj Unidade"].append(cnpj_unidade)
    NotaFiscal["Unidade"].append(unidade)
    NotaFiscal["Número Nota"].append(nf)
    NotaFiscal["Data de Emissão"].append(data)
    NotaFiscal["Operadora"].append(operadora)
    NotaFiscal["Cnpj Operadora"].append(cnpj_operadora)
    NotaFiscal["Valor"].append(valor)
    NotaFiscal["Base"].append(base)


pasta_salvador = r"C:\Users\kevin.soares\Downloads\Cárdio Pulmonar = Nota Fiscal convênio Asfeb vencimento_ 15_09_2023 (1)"


lista_arquivos = listar_itens_pasta(pasta_salvador)
print(len(lista_arquivos))

for i in range(0, len(lista_arquivos)):
    lerPDF_SALVADOR(lista_arquivos[i])
    print(NotaFiscal)

write_to_excel(NotaFiscal)