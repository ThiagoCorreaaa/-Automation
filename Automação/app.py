# Ler dados da planilha 
# Inserir os dados da planilha no sistema
import openpyxl
import pyautogui
workbook=openpyxl.load_workbook('produtos.xlsx')
cadastro_produtos=workbook['Worksheet']

for linha in cadastro_produtos.iter_rows(min_row=2):
    #["MOLO000251","Logitech","Mouse","1,25.95","6.50",]
    #Essa parte vai depender do sistema do Cliente mas vamos usar pyautogui para automação
    pyautogui.click(46,261,duration=1.5)
    pyautogui.write(linha[0].value)
    pyautogui.click(113,266,duration=1)
    pyautogui.write(linha[1].value)
    pyautogui.click(189,267,duration=1)
    pyautogui.write(linha[2].value)
    pyautogui.click(250,266,duration=1)
    pyautogui.write(str(linha[3].value))
    pyautogui.click(312,265,duration=1)
    pyautogui.write(str(linha[4].value))
    pyautogui.click(38,290,duration=2)

    print(linha[0].value)
    print(linha[1].value)
    print(linha[2].value)
    print(linha[3].value)
    print(linha[4].value)
    

    