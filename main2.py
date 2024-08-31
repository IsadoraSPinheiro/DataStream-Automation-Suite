# Grupo2: Arquivos que precisam de iteração nas linhas

import pandas as pd
#import tabula
import os
import re

# Função para converter arquivos .xls em .xlsx
def xls_to_xlsx(pasta, arquivo):
    if arquivo.endswith(".xls"):
        xls_file = os.path.join(pasta, arquivo)
        name, ext = os.path.splitext(xls_file)
        xlsx_file = f"{name}.xlsx"

        df = pd.read_excel(xls_file)
        df.to_excel(xlsx_file, index=False)

        if os.path.exists(xls_file):
            os.remove(xls_file)

# Função para converter arquivos .pdf em .xlsx
def pdf_to_excel(arquivo):
    pdf_file = arquivo  
    dt = tabula.read_pdf(pdf_file, pages="all")
    name = os.path.splitext(pdf_file)[0]
    excel_file = f"{name}.xlsx"
    excel_writer = pd.ExcelWriter(excel_file)
    for i, df in enumerate(dt):
        df.to_excel(excel_writer, sheet_name=f'Tabela_{i}', index=False)
    excel_writer.save()
    if os.path.exists(arquivo):
        os.remove(arquivo)

# Função para localizar as colunas no arquivo
def localizar(Padrao, pasta, arquivo):
    df = None  
    if arquivo.endswith('.xlsx'):
        df = pd.read_excel(os.path.join(pasta, arquivo))
        selecionar = {}
        for chave_df, valor_df in df.iterrows():
            categorias = []
            for chave_padrao, valor_padrao in Padrao.items():
                for valor in valor_df:
                    if str(valor).strip() in valor_padrao:
                        categorias.append(chave_padrao)
            if categorias:
                selecionar[chave_df] = categorias

        linha_cabecalho = None
        for indice, categorias in selecionar.items():
            if 'Valor' in categorias:
                linha_cabecalho = indice
                break

        df = df.iloc[linha_cabecalho:]
        df.reset_index(drop=True, inplace=True)
        df.columns = df.iloc[0]
        df = df[1:]
    return df

# Função para selecionar as colunas de interesse
def variantes(df, colunas):
    coluna_selecionada = []
    for coluna in colunas:
        if coluna in df.columns:
            coluna_selecionada.append(coluna)
    return coluna_selecionada

# Padrões de colunas
Padrao = {
    'Loja': ["NomSuc", "Filial", "DESCRIÇÃO LOJA", "Nome Empresa", "Loja", "Lojas", "loja", "lojas", "LOJA", "LOJAS", "Nome da Porta no Sistema do Cliente", "FabricanteChave", "DESC_Empresa", "EMPRESA", "Empresa", "NOME DA LOJA"],
    'Marca': ["Marca Produto", "Marca", "marca", "MARCA", "Cód./Grife", "cod./grife", "LOR GRIFE"],
    'Valor': ["Vda Bruta","Vl Total", "Preço Total"," Valor Vendido "," TOTAL R$ ","ValorVendas", "Soma Total", "VALOR VENDA", "Valor Vendido no Mês", "Total Sales Amount", "Valor Total Item", "Valor Faturado", " Valor Vendido", "Valor Vendido", "Valor", "Venda", "vendido", "VENDA", "VENDIDO", "VI Total ", "Pr.Venda (Líq.)", "Venda líquida", "venda liquida", "VENDA LIQUIDA", "valor total", "vendas ($)", "VALOR TOTAL", "Tot. Líquido", "Vendas ($)", "P.Vendido Total","Valor Total"],
    'Ean': ["C. Fornec","Código Auxiliar","C. Fornecedor","Cod.Barras", "Código de Barra", "Cd Barra", "CODIGO DE BARRAS","Referência","EAN", "ean", "Código do cliente", "Código de Barras do Produto", "Cód. barras", "Código Produto", "CodBarras", "Código", "COD DE BARRAS", "Cod. Barra"]
}

# Função para remover linhas em branco do DataFrame
def excluir_blank(df):
    df_remover = df.loc[df.isnull().all(axis=1)]
    df = df.drop(df_remover.index)
    df = df.reset_index(drop=True)
    return df

# Função para remover linhas que contenham "tot" no texto
def excluir_total(df):
    padrao = re.compile(r'tot[a-z]*', re.IGNORECASE)
    linhas_com_padrao = df.apply(lambda row: row.astype(str).str.contains(padrao).any(), axis=1)
    df = df.loc[~linhas_com_padrao]
    df = df.reset_index(drop=True)
    return df

# Função para remoção da soma dos valores
def excluir_soma(df):
    df = df.dropna(subset=['Loja'])
    df = df.reset_index(drop=True)
    return df

# Função principal para processar arquivos
def processamento(arquivo, Padrao):
    df_arquivo = localizar(Padrao, pasta, arquivo)
    df_arquivo = excluir_total(df_arquivo)


    coluna_selecionada = {
        'Loja': variantes(df_arquivo, Padrao['Loja']),
        'Marca': variantes(df_arquivo, Padrao['Marca']),
        'EAN': variantes(df_arquivo, Padrao['Ean']),
        'Valor': variantes(df_arquivo, Padrao['Valor'])
    }

    df_arquivo_selecionado = pd.DataFrame(columns=['Loja', 'Marca', 'Ean', 'Valor'])

    if coluna_selecionada['Loja']:
        df_arquivo_selecionado['Loja'] = df_arquivo[coluna_selecionada['Loja']]
    if coluna_selecionada['Valor']:
        df_arquivo_selecionado['Valor'] = df_arquivo[coluna_selecionada['Valor']]
    if coluna_selecionada['Marca'] or coluna_selecionada['EAN']:
        if coluna_selecionada['EAN']:
            df_arquivo_selecionado['Ean'] = df_arquivo[coluna_selecionada['EAN']]
        if coluna_selecionada['Marca']:
            df_arquivo_selecionado['Marca'] = df_arquivo[coluna_selecionada['Marca']]

    df_arquivo_selecionado = excluir_blank(df_arquivo_selecionado)
    df_arquivo_selecionado = excluir_soma(df_arquivo_selecionado)
    nome = os.path.splitext(arquivo)[0]
    nome_arquivo = f"{nome}"
    df_arquivo_selecionado['Fonte'] = nome_arquivo
    return df_arquivo_selecionado

# Pasta de origem dos arquivos
pasta = "c:\\Users\\isadora.pinheiro\\OneDrive - L'Oréal\\Desktop\\Automatização\\Grupo 2\\2.0"
arquivos = os.listdir(pasta)

# DataFrame para consolidar os dados
df_consolidado = pd.DataFrame(columns=['Loja', 'Marca', 'Ean', 'Valor', 'Fonte'])
for arquivo in arquivos:
    if arquivo.endswith(('.xls', '.xlsx', '.pdf')):
        if arquivo.endswith('.pdf'):
            pdf_to_excel(arquivo)
        elif arquivo.endswith('.xls'):
            xls_to_xlsx(pasta, arquivo)

        df_arquivo = processamento(arquivo,Padrao)
        if df_arquivo is not None:
            df_consolidado = pd.concat([df_consolidado, df_arquivo], ignore_index=True)

# Salvar o DataFrame consolidado
df_consolidado.to_excel("c:\\Users\\isadora.pinheiro\\OneDrive - L'Oréal\\Desktop\\Automatização\\Grupo 2\\Consolidado2.xlsx", index=False)
