import pandas as pd
import pyautogui 
import time
import pyperclip

pyautogui.PAUSE=2

base_dados = pd.read_excel(r"C:\Users\loganmiranda\Desktop\vendas\Vendas - Dez.xlsx")

base_dados = base_dados.drop('Código Venda', axis=1)

lojas = set(base_dados['ID Loja'])
lojas = list(lojas)

produtos = set(base_dados['Produto'])
produtos = list(produtos)


pd.set_option('display.max_rows', None)
pd.set_option('display.max_columns', None)

for loja in colunas:
    nao_vendidos = []
    
    loja_especifica = base_dados[base_dados['ID Loja'] == loja]
    
    faturamento = loja_especifica['Valor Final'].sum()
    qtd_vendida = loja_especifica['Quantidade'].sum()
    produto_info = loja_especifica.groupby('Produto')[['Quantidade', 'Valor Final']].sum()
    nome_produto = produto_info.index.tolist()
    
    for produto in produtos:
        if produto not in nome_produto:
            nao_vendidos.append(produto)
    
    
    qtd_total_produto = produto_info['Quantidade'].tolist()
    valor_total_produto = produto_info['Valor Final'].tolist()
    relatorio_produtos = list(zip(qtd_total_produto, valor_total_produto,nome_produto))
    relatorio_produtos.sort(reverse=True)
    
    mensagem = [f"""{info[2]} teve {info[0]} unidades vendidas \n""" for info in relatorio_produtos[:10]]
    print(mensagem)
    pyautogui.hotkey('ctrl','t')
    pyautogui.write('gmail.com')
    pyautogui.press('enter')
    time.sleep(3)
    pyautogui.click(x=2005, y=220)
    pyautogui.write('loganmiranda@mast.br')
    pyautogui.press('tab')
    pyautogui.press('tab')
    mensagem = f"Relatório de vendas em Dezembro - {loja}"
    pyperclip.copy(mensagem)
    pyautogui.hotkey('ctrl', 'v')
    pyautogui.press('tab')
    
    mensagem = f"""Boa tarde, senhores
    
    Em dezembro, no {loja}:
    Vendemos {qtd_vendida} produtos
    O FATURAMENTO TOTAL foi de R${faturamento}
    
    Os 10 produtos mais vendidos foram:
    {relatorio_produtos[:10]}
    
    Por fim, os produtos que NÃO foram vendidos no mês de Dezembro:
    {nao_vendidos}"""
    
    pyperclip.copy(mensagem)
    pyautogui.hotkey('ctrl', 'v')
    time.sleep(1)
    pyautogui.hotkey('ctrl', 'enter')
    time.sleep(2)
    pyautogui.hotkey('ctrl', 'w')

