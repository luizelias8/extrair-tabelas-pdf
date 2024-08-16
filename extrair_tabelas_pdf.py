import PySimpleGUI as sg
import camelot
import pandas as pd
import os

# Função para extrair tabelas e salvar no arquivo Excel
def extrair_tabelas(caminho_pdf, caminho_saida, intervalo_paginas, metodo):
    try:
        # Extrair tabelas das páginas especificadas
        tabelas = camelot.read_pdf(caminho_pdf, pages=intervalo_paginas, flavor=metodo)

        # Lista para armazenar os DataFrames das tabelas
        lista_tabelas = []

        # Iterar sobre as tabelas extraídas
        for tabela in tabelas:
            # Converter cada tabela para um DataFrame do pandas
            df = tabela.df
            lista_tabelas.append(df)

        # Combinar todas as tabelas em um único DataFrame
        df_combinado = pd.concat(lista_tabelas, ignore_index=True)

        # Exportar para um arquivo Excel
        df_combinado.to_excel(caminho_saida, index=False, header=False)

        sg.popup('Exportação concluída com sucesso!')
    except Exception as e:
        sg.popup_error(f'Erro ao extrair tabelas: {e}')

# Layout da janela
layout = [
    [sg.Text('Selecione o arquivo PDF')],
    [sg.Input(key='caminho_pdf'), sg.FileBrowse(button_text='Selecionar Arquivo', file_types=(('PDF Files', '*.pdf'),))],
    [sg.Text('Selecione o local para salvar o arquivo Excel (opcional)')],
    [sg.Input(key='caminho_saida'), sg.FileSaveAs(button_text='Selecionar Pasta', file_types=(('Arquivos Excel', '*.xlsx'),))],
    [sg.Text('Intervalo de páginas (ex: 1-3,5,7-10)')],
    [sg.Input(key='intervalo_paginas')],
    [sg.Text('Método de extração')],
    [sg.Combo(['lattice', 'stream'], default_value='lattice', key='metodo')],
    [sg.Button('Extrair Tabelas'), sg.Button('Sair')]
]

# Criar a janela
janela = sg.Window('Extrator de Tabelas PDF', layout)

while True:
    evento, valores = janela.read()
    if evento == sg.WIN_CLOSED or evento == 'Sair':
        break
    elif evento == 'Extrair Tabelas':
        caminho_pdf = valores['caminho_pdf']
        caminho_saida = valores['caminho_saida']
        intervalos_paginas = valores['intervalo_paginas']
        metodo = valores['metodo']

        # Se não for fornecido o caminho de saída, criar um caminho padrão
        if not caminho_saida:
            # Obter o diretório do arquivo PDF e definir o nome do arquivo de saída
            diretorio_pdf = os.path.dirname(caminho_pdf)
            caminho_saida = os.path.join(diretorio_pdf, 'tabelas_extraidas.xlsx')

        # Se o intervalo de páginas estiver vazio, definir como 'all'
        if not intervalos_paginas:
            intervalos_paginas = 'all'

        # Verificar se o caminho do arquivo PDF está preenchido
        if caminho_pdf:
            extrair_tabelas(caminho_pdf, caminho_saida, intervalos_paginas, metodo)
        else:
            sg.popup_error('Por favor, selecione um arquivo PDF.')

# Fechar a janela
janela.close()