import pyautogui as ptg
from selenium import webdriver
from time import sleep
from email.message import EmailMessage
import smtplib
import pyscreeze


# Módulo fornecedor de funções utilitárias para os demais módulos do programa.


FAILSAFE = True

def checar_failsafe():
    z, f = ptg.position()
    if z == 0 and f == 0:
        raise ptg.FailSafeException
    

def encontrar_imagem(imagem):
    cont = 0
    while True:
        try:
            encontrou = ptg.locateOnScreen(imagem, grayscale=True, confidence = 0.8)
            checar_failsafe()
            return encontrou
        except:
            sleep(0.8)
            cont += 1
            if cont == 3:
                break
            checar_failsafe()
            pass
            

def encontrar_centro_imagem(imagem):
    cont = 0
    while True:
        try:
            x, y = ptg.locateCenterOnScreen(imagem, grayscale=True, confidence=0.92)
            checar_failsafe()      
            return (x, y)
        except:
            sleep(0.8)
            cont += 1
            if cont == 3:
                break
            checar_failsafe()
            pass


def formatador(variavel, casas_decimais="{:.2f}"):
    variavel = float(variavel)
    variavel = casas_decimais.format(variavel)
    variavel = variavel.replace(".", ",")
    return variavel


def formatador2(variavel):
    variavel = float(variavel)
    variavel = "{:.2f}".format(variavel)
    return variavel


def formatador3(variavel):
    variavel = variavel.replace(",", ".")
    variavel = float(variavel)
    return variavel


def formatador4(variavel):
    variavel = variavel.replace(".", "")
    variavel = formatador3(variavel)
    return variavel


def descer_copiar():
    ptg.press("down", interval=0.1)
    ptg.hotkey("ctrl", "c", interval=0.1)
    checar_failsafe()


def clicar_microsiga():
    try:
        x, y = encontrar_centro_imagem(r'Imagens\IconeMicrosiga.png')
        ptg.click(x, y)
    except:
        x, y = encontrar_centro_imagem(r'Imagens\IconeMicrosigaWin11.png')
        ptg.click(x, y)
    checar_failsafe()


def voltar_descer(passos=1):
    ptg.hotkey(["shift", "tab"]*passos, interval=0.15)
    ptg.press("down")
    checar_failsafe()


def cancelar_lancamento():
    checar_failsafe()
    while True:
        cancelar_lancamento_click = encontrar_centro_imagem(r'Imagens\BotaoCancelarLancamento.png')
        try:
            x, y = cancelar_lancamento_click
            ptg.click(x,y, clicks=3, interval=0.1)
            checar_failsafe()
            break
        except:
            pass
    aguarde = encontrar_centro_imagem(r'Imagens\ReferenciaAguarde.png') 
    while type(aguarde) == tuple:
        aguarde = encontrar_imagem(r'Imagens\ReferenciaAguarde.png') 
        sleep(1)
    checar_failsafe()


def cancelar1():
    sleep(0.8)
    voltar_descer()
    sleep(0.5)
    clicar_microsiga()
    checar_failsafe()


def cancelar2():
    sleep(0.5)
    cancelar_lancamento()
    voltar_descer()
    sleep(0.3)
    clicar_microsiga()
    checar_failsafe()


def cancelar3():
    sleep(0.5)
    cancelar_lancamento()
    sleep(2)
    ptg.press("esc", interval=1)
    ptg.press("enter")
    cancelar1()
    checar_failsafe()


def erro_no_portal():
    sleep(0.3)
    ptg.hotkey("alt", "tab", interval=0.1)
    ptg.hotkey("ctrl", "w")
    sleep(0.2)
    clicar_microsiga()
    voltar_descer(passos=3)
    sleep(0.2)
    checar_failsafe()


def cancelar_mudar():
    cancelar_lancamento()
    mudar_a_selecao = encontrar_centro_imagem(imagem=r'Imagens\ClicarMudarSelecao.png')
    x, y = mudar_a_selecao
    ptg.click(x,y, clicks=4, interval=0.4)
    sleep(1)
    checar_failsafe()


def escrever_natureza(natureza):
    ptg.press("enter")
    ptg.write(natureza)
    ptg.press("enter")
    ptg.press("left")
    checar_failsafe()


def insistir_clique(imagem, cliques=2):
    while True:
        try:
            clicar_microsiga()
            sleep(1.5)
            checar_failsafe()
            try:
                ptg.click(250, 150)
                elemento = encontrar_centro_imagem(imagem)
                a, b = elemento
                sleep(0.5)
                checar_failsafe()
                ptg.click(a,b, clicks=cliques, interval=0.1)
                sleep(0.5)
                break
            except:
                sleep(0.3)
        except:
            ptg.moveTo(100, 150)
            sleep(0.3)
    checar_failsafe()


def clicar_dados_da_nota(): 
    encontrar = encontrar_centro_imagem(r'Imagens\BotaoDadosDaNota.png')
    if type(encontrar) != tuple:            
        insistir_clique(r'Imagens\BotaoDadosDaNota.png')
        sleep(0.5)
        checar_failsafe()
    else:
        x, y = encontrar
        ptg.doubleClick(x, y)
    try:
        aparece_enter = encontrar_imagem(r'Imagens\NCMIgnorar.png')
        if type(aparece_enter) == pyscreeze.Box:
            sleep(0.5)
            ptg.press("enter")
            checar_failsafe()
    finally:
        ptg.write("408")
    checar_failsafe()


def clicar_valor_parcela():
    valor_parcela = encontrar_centro_imagem(r'Imagens\ClicarParcela.png')
    while type(valor_parcela) != tuple:
        ptg.moveTo(180, 200)
        aba_duplicatas = encontrar_centro_imagem(r'Imagens\BotaoAbaDuplicatas.png')
        x, y =  aba_duplicatas
        checar_failsafe()
        ptg.click(x,y, clicks=4, interval=0.1)
        valor_parcela = encontrar_centro_imagem(r'Imagens\ClicarParcela.png')
        sleep(0.4)
    x, y = valor_parcela
    ptg.click(x,y)
    checar_failsafe()


def clicar_natureza_duplicata():
    while True:
        natureza_duplicata_clique = encontrar_centro_imagem(r'Imagens\ClicarNaturezaDuplicata.png')
        checar_failsafe()
        if type(natureza_duplicata_clique) != tuple:
            ptg.moveTo(150, 250)
            ptg.click(x,y, clicks=4, interval=0.1)
            sleep(0.3)
        else:
            break
    x, y = natureza_duplicata_clique
    ptg.click(x,y)
    checar_failsafe()


def enviar_email(corpo):
    mensagem = EmailMessage()
    mensagem.set_content(corpo)
    mensagem['Subject'] = "DANFE PARA LANÇAR"
    mensagem['From'] = "bot.contabil@eqseng.com.br"
    mensagem['To'] = "entrada.doc@eqsengenharia.com.br"
 
    try:
        with smtplib.SMTP_SSL('mail.eqseng.com.br', 465) as servidor:
            servidor.login("bot.contabil@eqseng.com.br", "EQSeng852@")
            servidor.send_message(mensagem)
    except Exception as e:
        pass
 
 
def acrescer_lista(lista, lista2, link, variavel):
    try:
        verificador = lista.index(link)
    except:
        lista.append(link)
    try:
        verificador = lista2.index(link)
    except:
        lista2.append(link)
        corpo = f"""
        Olá, colaborador!

 
        Não consegui lançar o processo abaixo, pode me ajudar?

        {link}
        
        Situação: {variavel}

 
        Atenciosamente,
        Bot.Contabil
        """
        enviar_email(corpo)


def tratar_xml_ilegivel(XML_ilegivel, nao_lancadas, link, mensagem_xi, aux=False):
    clicar_microsiga()
    if aux == True:
        ptg.press(["tab"]*3, interval=0.1)
    else:
        ptg.hotkey(["shift","tab"]*3, interval=0.1)
    ptg.press("down")
    sleep(0.5)
    clicar_microsiga()
    acrescer_lista(XML_ilegivel, nao_lancadas, link, mensagem_xi)
    checar_failsafe()


def tratar_lista(lista1, lista2):
    lista_unica = lista1 + lista2
    lista_unica = list(set(lista_unica))
    return lista_unica


def abrir_link_selenium(lista):
    options = webdriver.ChromeOptions()
    options.add_argument(r'user-data-dir=C:\Users\Usuário\AppData\Local\Google\Chrome\User Data\Default')
    driver = webdriver.Chrome(options=options)
    if len(lista) > 1:
        try:
            driver.get(lista[0])
            for link in lista[1:]:
                sleep(0.2)
                driver.execute_script("window.open('');")
                driver.switch_to.window(driver.window_handles[-1])
                sleep(0.5)
                driver.get(link)
            while True:
                sleep(1)
                if not driver.window_handles:
                    break
        except IndexError:
            pass
    else: 
        try:
            driver.get(lista[0])
            while True:
                sleep(1)
                if not driver.window_handles:
                    break
        except IndexError:
            driver.quit()
    
      