import openpyxl
import pyautogui
import time
from fuzzywuzzy import process
from setoresBcoDados import setores
import tkinter as tk
import os


def executePlanilha():
    os.startfile('planilha com horário automatizado.xlsm')


def executeAutomacao():
    caminho_planilha_helpdesk = 'planilha com horário automatizado.xlsm'

    def encontrar_setor_codigo(setor_nome1):
        setor_nome_mais_proximo = process.extractOne(setor_nome1, setores.keys())[0]
        return setores[setor_nome_mais_proximo]

    def preencher_MV2000(ocorrencia, codigo_setor, solucao, horario):
        pyautogui.click(997, 159)
        time.sleep(0.1)
        pyautogui.write(str(horario))
        time.sleep(0.1)
        pyautogui.press('space')
        pyautogui.write('-')
        pyautogui.press('space')
        time.sleep(0.1)
        pyautogui.write(str(ocorrencia))
        time.sleep(0.1)
        pyautogui.press('tab')
        time.sleep(0.1)
        pyautogui.write(str(codigo_setor))
        time.sleep(0.1)
        pyautogui.press('tab')
        time.sleep(0.1)
        pyautogui.write(str('XX - Oficina Helpdesk'))
        time.sleep(0.1)
        pyautogui.press('tab')
        time.sleep(0.1)
        pyautogui.write(str(ramal))
        pyautogui.press('tab')
        time.sleep(0.1)
        pyautogui.write(str(solucao))
        time.sleep(0.1)
        pyautogui.click(741, 442)

    planilha_helpdesk = openpyxl.load_workbook(caminho_planilha_helpdesk)
    folha_helpdesk = planilha_helpdesk.active

    for row in folha_helpdesk.iter_rows(min_row=2, min_col=1, max_col=7, values_only=True):
        nome, setor_nome, ocorrencia, ip, ramal, solucao, horario = row

        if not setor_nome:
            continue
        codigo_setor = encontrar_setor_codigo(setor_nome)
        preencher_MV2000(ocorrencia, codigo_setor, solucao, horario)

    planilha_helpdesk.close()


root = tk.Tk()
root.title("Automação HelpDesk")
root.geometry('250x50')
abrir_planilha_button = tk.Button(root, text="Abrir Planilha", command=executePlanilha)
abrir_planilha_button.pack()
cd_iniciar = tk.Button(root, text="Executar Automação", command=executeAutomacao)
cd_iniciar.pack()
root.mainloop()

