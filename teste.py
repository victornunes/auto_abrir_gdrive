#!/usr/bin/env python
# coding: utf-8

# In[17]:


import pyautogui
import time
import pyperclip
import pandas as pd

pyautogui.PAUSE = 1
#Abrir Navegador
pyautogui.press("winleft")
pyautogui.write("chrome")
pyautogui.press("enter")
pyautogui.alert("Atenção o programa vai começar a rodar, aperte OK e não mexa em nada!")
pyautogui.hotkey('ctrl','t')

#Abrir Drive
#Criando variavel para armazenar informação do link
link = "https://drive.google.com/file/d/1obXbboXP-LYVa546JLf_hl7fSIYrm0es/view?usp=sharing"
pyperclip.copy(link)
pyautogui.hotkey('ctrl','v')
pyautogui.press("enter")
time.sleep(10)
#Selecionar e baixar a imagem
pyautogui.click(x=1157, y=369, clicks=2)
pyautogui.click(x=1260, y=135)
pyautogui.click(x=1043, y=345)
time.sleep(10)


# In[21]:


import pyautogui
import time
import pyperclip
import pandas as pd

df = pd.read_excel(r'C:\Users\victo\OneDrive\Documentos\Exportar\Vendas - Dez.xlsx')
faturamento = df['Valor Final'].sum()
qtde_produtos = df['Quantidade'].sum()

#Abrindo o Gmail
pyautogui.press("winleft")
pyautogui.write("chrome")
pyautogui.press("enter")
pyautogui.alert("Atenção o programa vai começar a rodar, aperte OK e não mexa em nada!")
pyautogui.hotkey('ctrl','t')

link = "https://mail.google.com/mail/u/0/#inbox"
pyperclip.copy(link)
pyautogui.hotkey('ctrl','v')
pyautogui.press("enter")
time.sleep(5)

#ESCREVENDO O EMAIL
pyautogui.click(x=194, y=208)
pyautogui.write('victortrabalhoti@gmail.com')
pyautogui.press("tab")
pyautogui.press("tab")
assunto = "Relatorio de Vendas de Ontem"
pyperclip.copy(assunto)
pyautogui.hotkey('ctrl','v')
pyautogui.press("tab")
texto = f"""
Prezados, bom dia!

O Faturamento de ontem foi de: R${faturamento:,.2f}
A quantidade de produtos foi de: {qtde_produtos:,}

Cordialmente,
Victor Nunes Pinho"""
pyperclip.copy(texto)
pyautogui.hotkey('ctrl','v')
pyautogui.hotkey('ctrl','enter')


# In[20]:


pyautogui.position()

