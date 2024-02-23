# Notas Fiscais em XML para Excel

Este projeto é uma aplicação em Python que lê informações de notas fiscais em formato XML e as salva em um arquivo Excel (.xlsx).

## Tecnologias Utilizadas

- **Python**: Linguagem de programação principal do projeto.
- **Pandas**: Biblioteca Python para manipulação e análise de dados, utilizada para criar e manipular o DataFrame que será salvo no Excel.
- **xmltodict**: Biblioteca Python para converter XML em dicionários, utilizada para analisar os arquivos XML das notas fiscais.

## Instalação

1. Certifique-se de ter o Python instalado. Você pode baixá-lo em [python.org](https://www.python.org/downloads/).
2. Clone este repositório para o seu ambiente local:

git clone https://github.com/seu_usuario/nf-xml-to-excel.git

3. Instale as dependências utilizando pip:

pip install -r requirements.txt

## Utilização

1. Coloque seus arquivos XML de notas fiscais na pasta `tests`.
2. Execute o script Python `main.py`:

python main.py


3. Após a execução, um arquivo Excel chamado `NotasFiscais.xlsx` será gerado na pasta raiz do projeto, contendo as informações das notas fiscais em formato de tabela.

## Exemplo de Funcionamento

Suponha que temos o seguinte arquivo XML de nota fiscal:

```xml
<notaFiscal>
    <numero>001</numero>
    <data>2024-02-23</data>
    <cliente>
        <nome>Fulano de Tal</nome>
        <cpf>123.456.789-00</cpf>
        <endereco>Rua Exemplo, 123</endereco>
        <cidade>Cidade Exemplo</cidade>
        <estado>Estado Exemplo</estado>
    </cliente>
    <itens>
        <item>
            <descricao>Produto 1</descricao>
            <quantidade>2</quantidade>
            <valorUnitario>50.00</valorUnitario>
            <total>100.00</total>
        </item>
        <item>
            <descricao>Produto 2</descricao>
            <quantidade>1</quantidade>
            <valorUnitario>75.00</valorUnitario>
            <total>75.00</total>
        </item>
    </itens>
    <total>175.00</total>
</notaFiscal>
Este arquivo XML será lido pelo script Python e as informações serão armazenadas em um arquivo Excel, com cada item da nota fiscal em uma linha separada.
