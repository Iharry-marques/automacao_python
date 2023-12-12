import openpyxl
import pyperclip
import pyautogui
from time import sleep

workbook = openpyxl.load_workbook('produtos_ficticios.xlsx')
sheet_produtos = workbook['Produtos']
for linha in sheet_produtos.iter_rows(min_row=2):
    nome_produto = linha[0].value
    pyperclip.copy(nome_produto)
    pyautogui.click(161,298, duration=1)
    pyautogui.hotkey('ctrl','v')
    
    descricao = linha[1].value
    pyperclip.copy(descricao)
    pyautogui.click(160,406, duration=1)
    pyautogui.hotkey('ctrl','v')
    
    categoria = linha[2].value
    pyperclip.copy(categoria)
    pyautogui.click(168,569, duration=1)
    pyautogui.hotkey('ctrl','v')
    
    codigo_produto = linha[3].value
    pyperclip.copy(codigo_produto)
    pyautogui.click(162,679, duration=1)
    pyautogui.hotkey('ctrl','v')
    
    peso = linha[4].value
    pyperclip.copy(peso)
    pyautogui.click(169,781, duration=1)
    pyautogui.hotkey('ctrl','v')
    
    dimensoes = linha[5].value
    pyperclip.copy(dimensoes)
    pyautogui.click(161,891, duration=1)
    pyautogui.hotkey('ctrl','v')
    
    pyautogui.click(185,976,duration=1)
    sleep(3)
    
    preco = linha[6].value
    pyperclip.copy(preco)
    pyautogui.click(180,330, duration=1)
    pyautogui.hotkey('ctrl','v')
    
    quantidade_em_estoque = linha[7].value
    pyperclip.copy(quantidade_em_estoque)
    pyautogui.click(179,437, duration=1)
    pyautogui.hotkey('ctrl','v')
    
    data_de_validade = linha[8].value
    pyperclip.copy(data_de_validade)
    pyautogui.click(174,542, duration=1)
    pyautogui.hotkey('ctrl','v')
    
    cor = linha[9].value
    pyperclip.copy(cor)
    pyautogui.click(178,652, duration=1)
    pyautogui.hotkey('ctrl','v')
    
    pyautogui.click(186,757, duration=1)
    
    tamanho = linha[10].value
    if tamanho == 'Pequeno':
        pyautogui.click(195,798, duration=1)
    elif tamanho == 'Médio':
        pyautogui.click(196,829, duration=1)
    else:
        pyautogui.click(198,852, duration=1)
        
    material = linha[11].value
    pyperclip.copy(material)
    pyautogui.click(186,859, duration=1)
    pyautogui.hotkey('ctrl','v')
    
    pyautogui.click(192,940, duration=1)
    
    fabricante = linha[12].value
    pyperclip.copy(fabricante)
    pyautogui.click(183,362, duration=1)
    pyautogui.hotkey('ctrl','v')
    
    pais_origem = linha[13].value
    pyperclip.copy(pais_origem)
    pyautogui.click(184,468, duration=1)
    pyautogui.hotkey('ctrl','v')
    
    observacoes = linha[14].value
    pyperclip.copy(observacoes)
    pyautogui.click(179,573, duration=1)
    pyautogui.hotkey('ctrl','v')
    
    codigo_de_barras = linha[15].value
    pyperclip.copy(codigo_de_barras)
    pyautogui.click(176,735, duration=1)
    pyautogui.hotkey('ctrl','v')
    
    localizacao_armazem = linha[16].value
    pyperclip.copy(localizacao_armazem)
    pyautogui.click(172,837, duration=1)
    pyautogui.hotkey('ctrl','v')
    
    pyautogui.click(191,917, duration=1)
    pyautogui.click(699,221, duration=1)
    pyautogui.click(469,629, duration=1)