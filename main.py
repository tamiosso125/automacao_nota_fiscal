import xmltodict
import os
import pandas as pd

all_tests = os.listdir("tests")

def formatar_item(item):
    descricao = item['descricao']
    quantidade = int(item['quantidade'])
    valor_unitario = float(item['valorUnitario'])
    total = float(item['total'])
    
    return f"{quantidade}x {descricao} (R${valor_unitario:.2f} cada) = R${total:.2f}"

def get_test(test, max_itens):
    with open(f"tests/{test}", 'rb') as test_file:
        info_test = xmltodict.parse(test_file)['notaFiscal']
        
        cliente = info_test['cliente']
        try:
            cnpj = cliente['cnpj']
        except KeyError:
            cnpj = 'Não Informado'
        try:
            cpf = cliente['cpf']
        except KeyError:
            cpf = 'Não Informado'

        itens = {f'item_{i+1}': formatar_item(item) for i, item in enumerate(info_test['itens']['item'])}
        
        # Preencher com None para garantir que todas as linhas tenham o mesmo número de colunas
        for i in range(len(itens), max_itens):
            itens[f'item_{i+1}'] = None
        
        total = info_test['total']
        
        values.append({
            'numero_nota': info_test['numero'],
            'data_emissao': info_test['data'],
            'nome_cliente': cliente['nome'],
            'cpf': cpf,
            'cnpj': cnpj,
            'endereco': cliente['endereco'],
            'cidade': cliente['cidade'],
            'estado': cliente['estado'],
            **itens,
            'total': f'R${total}'
        })

# Encontrar o número máximo de itens em todas as notas fiscais
max_itens = max(len(xmltodict.parse(open(f"tests/{test}", 'rb'))['notaFiscal']['itens']['item']) for test in all_tests)

values = []
for test in all_tests:
    get_test(test, max_itens)

tabela = pd.DataFrame(values)
if os.path.exists('NotasFiscais.xlsx'):
    # Carregar o arquivo existente para atualização
    tabela_existente = pd.read_excel('NotasFiscais.xlsx')
    
    # Verificar se as colunas da tabela existente correspondem às colunas da nova tabela
    if set(tabela.columns) == set(tabela_existente.columns):
        # Atualizar a tabela existente
        tabela_existente = pd.concat([tabela_existente, tabela], ignore_index=True)
        
        # Salvar a tabela atualizada
        tabela_existente.to_excel("NotasFiscais.xlsx", index=False)
        
        print("Tabela existente atualizada com sucesso!")
    else:
        print("As colunas da tabela existente não correspondem às colunas da nova tabela. Não é possível atualizar.")
else:
    # Salvar a nova tabela, pois não existe uma tabela existente
    tabela.to_excel("NotasFiscais.xlsx", index=False)
    print("Nova tabela criada com sucesso!")
