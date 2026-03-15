import customtkinter as ctk
import time
from pynput.keyboard import Controller, Key, Listener
import win32gui
import win32con
import winsound
#import pyautogui windows nao permita baixar
#______________________________________ variaveis
kb = Controller()
parar = False

#______________________________________ funcoes
def on_press(key):
     global parar
     if key == Key.esc:
      parar = True
      try:
            label_result.configure(text="AÇÃO INTERROMPIDA", **amarelo, font=fonte3)
            janela.update()   
            winsound.PlaySound("SystemExclamation", winsound.SND_ALIAS)
      except:
            pass    
      return #False # encerra listener
#____________________melhor lugar para Listener
listener = Listener(on_press=on_press)
listener.start()

def focar_janela(trecho_titulo):
    alvo = None
    def enum_handler(hwnd, _):
        nonlocal alvo
        if win32gui.IsWindowVisible(hwnd):
            titulo = win32gui.GetWindowText(hwnd)
            if trecho_titulo.lower() in titulo.lower():
                alvo = hwnd
    win32gui.EnumWindows(enum_handler, None)

    if not alvo:
        return False
    # restaura se estiver minimizada

    if trecho_titulo == "Tabela de Apoio ":
        win32gui.ShowWindow(alvo, win32con.SW_MAXIMIZE)
        # time.sleep(0.2)
    else:
        win32gui.ShowWindow(alvo, win32con.SW_RESTORE)
        # time.sleep(0.2)
    time.sleep(0.3)
    # força foco corretamente
    try:
        win32gui.SetForegroundWindow(alvo)
    except:
        pass
    return True

def testeabas():
    janela.iconify()  # minimiza meu app
    focar_janela("Tabela de Apoio ")
    if not focar_janela("Tabela de Apoio "):
        print("janela não encontrada do excel")
    janela.update()
    time.sleep(2)
    focar_janela("análise ")
    if not focar_janela("análise "):
        print("janela não encontrada do access")

def segura(modifier, key):
    kb.press(modifier)
    kb.press(key)
    kb.release(key)
    kb.release(modifier)

def coluna_esquerda():
    focar_janela("análise ")
    for _ in range(3):
        segura(Key.shift_l, Key.tab)
        time.sleep(0.2)  # pula para acessou acessou help
    focar_janela("Tabela de Apoio ")
    if checar_parada():
        return
    for _ in range(4):
        kb.press(Key.down)
        time.sleep(0.2)  # pula para acessou help
    repeticao(1)
    for _ in range(5):
        kb.press(Key.up)
        time.sleep(0.2)  # pula para acessou nº manifestacao
    focar_janela("análise ")
    for _ in range(2):
        kb.press(Key.tab)
        time.sleep(0.2)  # pula para acessou nº manifestacao
    #--------------------------
    focar_janela("Tabela de Apoio ")
    repeticao(2)
    time.sleep(0.2)
    focar_janela("análise ")
    if checar_parada():
        return
    for _ in range(3):  # levar para n manifestacao
        kb.press(Key.tab)
        time.sleep(0.2)
    focar_janela("Tabela de Apoio ")

    repeticao(2)
    time.sleep(0.1)
    kb.press(Key.down)
    time.sleep(0.1)
    kb.press(Key.down)
    if checar_parada():
        return
    time.sleep(0.1)  # pula valor
    kb.press(Key.down)  # pula causa
    time.sleep(0.2)
    repeticao(3)
    time.sleep(0.1)
    teste_subida()
def teste_subida():
    kb.press(Key.right)
    time.sleep(0.2)
    if checar_parada():
        return
    segura(Key.ctrl, Key.up)
    time.sleep(0.2)
    kb.press(Key.right)

def teste_descida():
    time.sleep(0.2)
    kb.press(Key.left)
    time.sleep(0.2)
    if checar_parada():
        return
    segura(Key.ctrl, Key.down)
    time.sleep(0.2)
    kb.press(Key.up)
    time.sleep(0.2)
    kb.press(Key.left)

def incluir_psg():
    kb.press(Key.tab)
    time.sleep(0.2)
    if checar_parada():
        return
    kb.press(Key.tab)
    time.sleep(0.2)
    kb.press(Key.enter)
    time.sleep(0.2)
    kb.press(Key.enter)
    time.sleep(0.2)
    kb.press(Key.tab)
    for _ in range(4):
        kb.press(Key.tab)
        time.sleep(0.2)
    focar_janela("Tabela de Apoio ")

def coluna_direita():
    repeticao(10)
    if checar_parada():
        return
    time.sleep(0.2)
    teste_descida()
    time.sleep(0.2)
    focar_janela("análise ")
    if checar_parada():
        return
    kb.press(key='s')
    time.sleep(0.2)
    incluir_psg()

def executar_tudo(): #funcao principal para rodar tudo, ela chama as outras funcoes, e tem a contagem regressiva, validacao de numero digitado e mensagem de sucesso ou erro
    global parar
    parar = False
    try:
        qtd = int(entry.get())  # aqui valida se usuario digitou zero ou acima de 10, pra nao rodar
        if qtd <= 0 or qtd > 10:
            raise ValueError
    except ValueError:
        janela.update()
        label_result.configure(text="Digite um número válido de 1 à 10", font=fonte)
        janela.update()
        winsound.PlaySound("SystemQuestion", winsound.SND_ALIAS)
        return

    for n in range(5, 0, -1):  # aqui serve uma contagem antes de tudo iniciar
        if checar_parada():
            return
        label_result.configure(text=f"Em {n}...", font=fonte3, **vermelho)
        janela.update()
        time.sleep(1)

    label_result.configure(text=f"Executando {qtd} vezes...", font=fonte3, **amarelo)
    janela.update()

    if not focar_janela("tabela de apoio"):  # para validar se o excel ta aberto antes de iniciar
        label_result.configure(text="Excel não encontrado aberto")
        janela.deiconify()
        return

    time.sleep(0.2)
    for _ in range(qtd):  # aqui roda as quantidades definidas
        if checar_parada():
            return
        coluna_esquerda()
        if checar_parada():
            return
        coluna_direita()

    label_result.configure(text="---Concluído---", font=fonte3, **verde)  # mensagem de sucesso quando terminar
    winsound.MessageBeep(winsound.MB_ICONHAND)  # janela.update()

def analise():
    global parar
    parar = False

    for n in range(5, 0, -1):  # aqui serve uma contagem antes de tudo iniciar
        if checar_parada():
            return
        label_result.configure(text=f"Em {n}...", **vermelho, font=fonte3)
        janela.update()
        time.sleep(1)
    label_result.configure(text=f"Executando...", **amarelo, font=fonte3)
    janela.update()  # atualiza a tela com a informacao acima
    time.sleep(0.2)
    if checar_parada():
        return
    focar_janela("Tabela de Apoio ")
    time.sleep(0.2)
    if checar_parada():
        return
    repeticao(9)
    
  
    if checar_parada():
        return
    focar_janela("Tabela de Apoio ")  # volta para excel

    kb.press(Key.right)  # vai para celula ao lado
    time.sleep(0.2)
    segura(Key.ctrl, Key.up)  # sobe ate avaliador
    time.sleep(0.2)
    kb.press(Key.right)  # vai para data siscap
    time.sleep(0.2)
    if checar_parada():
        return
    repeticao(7)
    
   
    if check_var.get() is True:
        kb.press(key="s")  # MOTIVO ACionamento sem passagem
        time.sleep(0.1)
        kb.press(Key.tab)
        time.sleep(0.1)
        focar_janela("Tabela de Apoio ")
        kb.press(Key.down)
        time.sleep(0.1)
        repeticao(1)
    else:
        if checar_parada():
            return
        focar_janela("Tabela de Apoio ")
        repeticao(3)
    
    if checar_parada():
        return
    
    
    label_result.configure(text="---Concluído---", **verde, font=fonte3)
    janela.update()
    winsound.MessageBeep(winsound.MB_ICONHAND)  # mensagem de sucesso quando terminar

def checar_parada():
    if parar:
        return True
    janela.update()
    return False

def repeticao(vezes): #mini funcao usada para evitar repetir codigo, ela recebe a quantidade de vezes que deve repetir o processo definido nela, e tem a checagem de parada para cada etapa
    for _ in range(vezes):
        if checar_parada():
            return
        focar_janela("Tabela de Apoio ")
        segura(Key.ctrl, 'c')
        time.sleep(0.2)
        if checar_parada():
            return
        focar_janela("análise ")
        segura(Key.ctrl, 'v')
        kb.press(Key.tab)
        time.sleep(0.2)
        if checar_parada():
            return
        focar_janela("Tabela de Apoio ")
        kb.press(Key.down)
        time.sleep(0.2)

#def teste():
#    pyautogui.rightClick(x=200, y=300)

#-----------------------interface----------------

verde = {"text_color":"#47F20D"}
amarelo= {"text_color":"#FFEE00"}
vermelho= {"text_color":"#FA0000"}

janela = ctk.CTk()
ctk.set_appearance_mode("dark")
ctk.set_widget_scaling(1.0)
fonte= ctk.CTkFont(family="MarkPro",size=17,weight="bold")
fonte2= ctk.CTkFont(family="MarkPro",size=12,slant="italic")
fonte3= ctk.CTkFont(family="MarkPro",size=24,slant="roman")
janela.title("Automacen")
janela.geometry("350x400")
janela.resizable(True,True)
janela.attributes("-topmost", True)
check_var = ctk.BooleanVar() # checkbox true ou falsa
conta = ctk.CTkLabel(janela,text="Inicia 5 segs após clicar",font=fonte2,text_color="#A9A9AF") 
entry = ctk.CTkEntry(janela, width=120, placeholder_text="Max: 10")
label_result = ctk.CTkLabel(janela, text="",text_color="yellow",font=fonte3,height=25)

check = ctk.CTkCheckBox(janela, text="Sem passagem", variable=check_var, checkbox_height=18, checkbox_width=18)
botao = ctk.CTkButton(janela, text="  Enviar Dados  ", command=analise,width=35, height=8,fg_color="#050791",text_color="white",font=fonte)
botao2 = ctk.CTkButton(janela, text="Enviar passagens", command=executar_tudo,width=35,hover_color="#e692f1", height=8,fg_color="#89019b",text_color="white",font=fonte)
#-----------------------interface dentro da  janela -------------------------------
conta.pack(pady=10)
label_result.pack(pady=(5,8))
ctk.CTkLabel(janela, text="___________________________________ \nDados Principais e Análise ",font=fonte2,height=20).pack(pady=(5,4))
check.pack(pady=(1,1))
botao.pack(pady=(4,20))

ctk.CTkLabel(janela, text="___________________________________ \nProtocolos Relacionados ",font=fonte2,height=20).pack(pady=(5))
ctk.CTkLabel(janela, text="Quantas passagens ",font=fonte2,height=20).pack(pady=(4,4))
entry.pack(pady=4)
botao2.pack(pady=(4,20))

#ctk.CTkButton(janela, text="teste", command=teste,width=20, height=8).pack(pady=(5,5))
# texto_formatado = f"""
#     resumo de apoio
#     """
# textbox = ctk.CTkTextbox(janela, width=550, height=400)
# textbox.pack(pady=10)
# textbox.insert("end", texto_formatado)
janela.update()
janela.mainloop()
