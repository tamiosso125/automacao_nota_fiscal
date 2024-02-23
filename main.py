import xmltodict
import os
import pandas as pd

all_tests = os.listdir("tests")

def get_test(test):
    with open(f"tests/{test}", 'rb') as test_file:
            info_test = xmltodict.parse(test_file)
            number = info_test['notaFiscal']['numero']
            date = info_test['notaFiscal']['data']
            name = info_test['notaFiscal']['cliente']['nome']
            if "cnpj" in info_test['notaFiscal']['cliente']:
                cnpj = info_test['notaFiscal']['cliente']['cnpj']
            else:
                cnpj = 'Não Informado'
            if "cpf" in info_test['notaFiscal']['cliente']:
                cpf = info_test['notaFiscal']['cliente']['cpf']
            else:
                cpf = 'Não Informado'

            adress =  info_test['notaFiscal']['cliente']['endereco']
            city = info_test['notaFiscal']['cliente']['cidade']
            state = info_test['notaFiscal']['cliente']['estado']
            itens = info_test['notaFiscal']['itens']['item']
            total = info_test['notaFiscal']['total']
            values.append([number, date, name, cpf, cnpj, adress, city, state, itens, total])


colums = ['numero_nota', 'data_emissao', 'nome_cliente', 'cpf', 'cnpj', 'endereco', 'cidade', 'estado', 'item', 'total']
values = []

for test in all_tests:
    get_test(test)
    
tabela = pd.DataFrame(columns=colums, data=values)
tabela.to_excel("NotasFiscais.xlsx", index=False)