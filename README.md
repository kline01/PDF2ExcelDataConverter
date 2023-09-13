# Extração de dados NFs-e de Salvador para Excel
O código extrai dados da nota fiscal eletrônica do municipio de Salvador para Excel. Esse código é útil para quando temos que analisar dados de diversas notas fiscais.

## ETL

Para extração de dados é utilizado a biblioteca [pdfquery](https://pypi.org/project/pdfquery/) que utiliza a posição ``` bbox ``` para localizar os itens.
```
pip install pdfquery
```

A biblioteca ``` os ``` utilizada para a possibilidade de listar todas as notas fiscais dentro da pasta.

```
pip install os-sys
```

Para transformação, os dados são colocados em um dict.
```
NotaFiscal["Cnpj prestador"].append(cnpj_prestador)
    NotaFiscal["prestador"].append(prestador)
    NotaFiscal["Número Nota"].append(nf)
    NotaFiscal["Data de Emissão"].append(data)
    NotaFiscal["Tomador"].append(Tomador)
    NotaFiscal["Cnpj Tomador"].append(cnpj_Tomador)
    NotaFiscal["Valor"].append(valor)
    NotaFiscal["Base"].append(base)
```
E então, carregados para o excel com a biblioteca pandas.
```
pip install pandas
```
Os dados do dict ```NotaFiscal``` são colocados em um ```DataFrame``` para então serem carregados com a função ExcelWriter que tem o motor ```openpyxl```:
```
writer = pd.ExcelWriter(caminho, engine='openpyxl', mode='a', if_sheet_exists='replace')
DataFrame.to_excel(writer,'Main' ,index=False,header=True) 
```

Por fim, os dados sairão em uma planilha no seguinte formato:

| Prestador  | Nr. da Nota | Dr. Emissão | Tomador | Cnpj Tomador | Valor | Base |
| ------------- | ------------- | ------------- |  ------------- |  ------------- |  ------------- |  ------------- |
| Prestador1  | 1  | 01/01/2023 | Tomador 1  | 11.111.111/0001-10  | 1,00  | 1,00  |
| Prestador2  | 2  | 02/01/2023 | Tomador 2  | 22.222.222/0002-20  | 2,00  | 2,00  |