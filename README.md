# Extrator de Tabelas de PDFs

Este projeto é uma ferramenta simples para extrair tabelas de arquivos PDF e exportá-las para um arquivo Excel. Ele utiliza as bibliotecas Camelot para a extração das tabelas e PySimpleGUI para criar uma interface gráfica de fácil utilização.

## Funcionalidades

- Extração de Tabelas: Extrai tabelas de arquivos PDF e as exporta para um arquivo Excel (.xlsx).
- Intervalo de Páginas: Permite selecionar um intervalo de páginas específico para extrair as tabelas.
- Método de Extração: Oferece a opção de escolher entre dois métodos de extração (lattice e stream).
- Caminho de Saída: Salva o arquivo Excel no caminho especificado pelo usuário. Se não fornecido, o arquivo é salvo no mesmo diretório do PDF de origem com o nome tabelas_extraidas.xlsx.
- Interface Gráfica: Interface amigável que facilita a utilização da ferramenta.

## Requisitos

- Python 3.x
- [Camelot](https://camelot-py.readthedocs.io/en/master/) (para extração de tabelas)
- [PySimpleGUI](https://pysimplegui.readthedocs.io/en/latest/) (para a interface gráfica)
- [pandas](https://pandas.pydata.org/) (para manipulação e exportação de tabelas)

## Instalação

1. Clone o repositório para sua máquina local:
```
git clone https://github.com/luizelias8/extrair-tabelas-pdf.git
```

2. Instale as dependências:
```
pip install -r requirements.txt
```

## Uso

1. Execute o script principal:
```
python extrair_tabelas_pdf.py
```

2. Na interface gráfica:

- Selecione o arquivo PDF.
- Informe o intervalo de páginas desejado ou deixe em branco para processar todas as páginas.
- Escolha o método de extração (lattice ou stream).
- Informe o caminho de saída ou deixe em branco para salvar o arquivo Excel no mesmo diretório do PDF.

3. Clique em "Extrair Tabelas" para iniciar o processo. Uma mensagem será exibida ao final informando sobre a conclusão da exportação.

## Contribuição

Contribuições são bem-vindas!

## Autor

- [Luiz Elias](https://github.com/luizelias8)