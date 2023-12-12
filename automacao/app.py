import openpyxl
import pyperclip
import pyautogui
from time import sleep

workbook = openpyxl.load_workbook('produtos_ficticios.xlsx')
sheet_produtos = workbook['Produtos']
for linha in sheet_produtos.iter_rows(min_row=2):
    nome_produto = linha[0].value
    pyperclip.copy(nome_produto)
    pyautogui.click(1995,259, duration=1)
    pyautogui.hotkey('ctrl','v')
    
    descricao = linha[1].value
    pyperclip.copy(descricao)
    pyautogui.click(2001,349, duration=1)
    pyautogui.hotkey('ctrl','v')
    
    categoria = linha[2].value
    pyperclip.copy(categoria)
    pyautogui.click(2000,480, duration=1)
    pyautogui.hotkey('ctrl','v')
    
    codigo_produto = linha[3].value
    pyperclip.copy(codigo_produto)
    pyautogui.click(1991,564, duration=1)
    pyautogui.hotkey('ctrl','v')
    
    peso = linha[4].value
    pyperclip.copy(peso)
    pyautogui.click(1996,655, duration=1)
    pyautogui.hotkey('ctrl','v')
    
    dimensoes = linha[5].value
    pyperclip.copy(dimensoes)
    pyautogui.click(1997,742, duration=1)
    pyautogui.hotkey('ctrl','v')
    
    pyautogui.click(2014,795, duration=1)
    sleep(3)
    
    preco = linha[6].value
    pyperclip.copy(preco)
    pyautogui.click(2258,283, duration=1)
    pyautogui.hotkey('ctrl','v')
    
    quantidade_em_estoque = linha[7].value
    pyperclip.copy(quantidade_em_estoque)
    pyautogui.click(2252,366, duration=1)
    pyautogui.hotkey('ctrl','v')
    
    data_de_validade = linha[8].value
    pyperclip.copy(data_de_validade)
    pyautogui.click(2251,454, duration=1)
    pyautogui.hotkey('ctrl','v')
    
    cor = linha[9].value
    pyperclip.copy(cor)
    pyautogui.click(2247,539, duration=1)
    pyautogui.hotkey('ctrl','v')
    
    
    tamanho = linha[10].value
    if tamanho == 'Pequeno':
        pyautogui.click(2257,656, duration=1)
    elif tamanho == 'Médio':
        pyautogui.click(2254,678, duration=1)
    else:
        pyautogui.click(2268,698, duration=1)
    material = linha[11].value
    fabricante = linha[12].value
    pais_origem = linha[13]