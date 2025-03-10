from pathlib import Path
from PIL import ImageTk, Image
from tkinter import Tk, Canvas, Button, PhotoImage, Label, IntVar
from tigrinho import robozinho
from time import sleep
from utils import abrir_link_selenium, tratar_lista, checar_failsafe
from inicializadorUsuario import inicializar_usuario
import mensagens
import threading
 

continuar_loop = False
abortar = False
resetar = False
lancadas = 0
sem_boleto = []
processo_bloqueado = []
processo_errado = []
XML_ilegivel = []
nao_lancadas = []
 
 
def ativar_robozinho():
    """
    Executa o robô principal (robozinho) para processar as DANFEs.
    Atualiza as variáveis globais com os resultados do processamento.
    Se o robô for resetado, limpa as listas de resultados antes de executar.
    """
    global continuar_loop, lancadas, qtd_lancadas, qtd_sem_boleto, qtd_processo_bloqueado, qtd_processo_errado, qtd_XML_ilegivel, qtd_nao_lancadas
    global sem_boleto, processo_bloqueado, processo_errado, XML_ilegivel, nao_lancadas, resetar
 
    if resetar == True:
        s_boleto, proc_bloqueado, proc_errado, xml_ilegivel, n_lancadas, abortar = robozinho(resetar)
        resetar = False
    else:
        s_boleto, proc_bloqueado, proc_errado, xml_ilegivel, n_lancadas, abortar = robozinho()
        
    if abortar == False:
        lancadas += 1
        qtd_lancadas.set(lancadas)
        continuar_loop = True
    else:
        continuar_loop = False

    sem_boleto = tratar_lista(sem_boleto, s_boleto)
    processo_bloqueado = tratar_lista(processo_bloqueado, proc_bloqueado)
    processo_errado = tratar_lista(processo_errado, proc_errado)
    XML_ilegivel = tratar_lista(XML_ilegivel, xml_ilegivel)
    nao_lancadas = tratar_lista(nao_lancadas, n_lancadas)

    qtd_sem_boleto.set(len(sem_boleto))
    qtd_processo_bloqueado.set(len(processo_bloqueado))
    qtd_processo_errado.set(len(processo_errado))
    qtd_XML_ilegivel.set(len(XML_ilegivel))
    qtd_nao_lancadas.set(len(nao_lancadas))
    checar_failsafe()
   
 
def resetar_bot():
    """
    Reseta as listas de resultados e as variáveis de contagem.
    """
    global resetar
    sem_boleto.clear()
    processo_bloqueado.clear()
    processo_errado.clear()
    XML_ilegivel.clear()
    nao_lancadas.clear()
    qtd_sem_boleto.set(0)
    qtd_processo_bloqueado.set(0)
    qtd_processo_errado.set(0)
    qtd_XML_ilegivel.set(0)
    qtd_nao_lancadas.set(0)
    resetar = True
 
 
def abrir_gui():
    global qtd_lancadas, qtd_sem_boleto, qtd_processo_bloqueado, qtd_processo_errado, qtd_XML_ilegivel, qtd_nao_lancadas, continuar_loop
 
 
    def rodar_robozinho(): 
        ativar_robozinho()
        if continuar_loop:
            window.after(0.2, rodar_robozinho())
            checar_failsafe()
    
 
    def responder_clique(funcao):
        """
        Função interna que responde ao clique do usuário, executando a função passada como argumento.
        Minimiza a janela e executa a função em uma thread separada.
        """
        sleep(1)
        window.iconify()
        threading.Thread(target=funcao, daemon=True).start()
        checar_failsafe()
        
 
 
    OUTPUT_PATH = Path(__file__).parent
    ASSETS_PATH = OUTPUT_PATH / Path(r"Imagens\InterfaceGrafica")
 

    def relative_to_assets(path: str) -> Path:
        """
        Retorna o caminho absoluto para um arquivo na pasta de assets.
        """
        return ASSETS_PATH / Path(path)
    

 
    cor_fundo = "#FFFFFF"
    window = Tk()
 
    bot = mensagens.Mensagens(window)
 
    bot.mostrar_info(bot.info, bot.texto)
    bot.mostrar_info(bot.info2, bot.texto2)
    bot.mostrar_erro(bot.texto3)
    bot.mostrar_info(bot.info4, bot.texto4)
    bot.mostrar_aviso(bot.info5, bot.texto5)
 
    window.deiconify()
 
    window.iconbitmap(relative_to_assets("robozinho.ico"))
    window.geometry("788x478+390+110")
    window.title("Automação Entrada de DANFE")
    window.configure(bg = cor_fundo)
 
    qtd_sem_boleto = IntVar()
    qtd_processo_bloqueado = IntVar()
    qtd_processo_errado = IntVar()
    qtd_XML_ilegivel = IntVar()
    qtd_nao_lancadas = IntVar()
    qtd_lancadas = IntVar()
 
    canvas = Canvas(
        window,
        bg = "#FFFFFF",
        height = 478,
        width = 788,
        bd = 0,
        highlightthickness = 0,
        relief="solid",
    )
    canvas.place(x = 0, y = 0)
 
    canvas = Canvas(
        window,
        bg = "#E2E2E2",
        height = 292,
        width = 788,
        bd = 0,
        highlightthickness = 0,
        relief="solid",
    )
    canvas.place(x = 0, y = 186)
 
    label_1_imagem = PhotoImage(
        file=relative_to_assets("imagem_lancadas.png"))
    label_1 = Label(
        window,
        image=label_1_imagem,
        )
    label_1.place(
        x=45,
        y=250,
        width=138.0,
        height=55.0
    )
 
    label_sub_1 = Label(
        window,
        textvariable=qtd_lancadas,
        font=("Malgun Gothic", 17, "bold"),
        fg="#207C00",
        anchor="center",
        justify="center",
        bg="#ffffff",
        relief="groove"
        )
   
    label_sub_1.place(
        x=45,
        y=306,
        width=138.0,
        height=40.0
    )
 
    label_2_imagem = PhotoImage(
        file=relative_to_assets("imagem_nao_lancadas.png"))
    label_2 = Label(
        window,
        image=label_2_imagem,
        )
    label_2.place(
        x=183,
        y=250,
        width=138.0,
        height=55.0
    )
 
    label_sub_2 = Label(
        window,
        textvariable=qtd_nao_lancadas,
        font=("Malgun Gothic", 17, "bold"),
        fg="#D30000",
        anchor="center",
        justify="center",
        bg="#ffffff",
        relief="groove"
        )
    label_sub_2.place(
        x=183,
        y=306,
        width=138.0,
        height=40.0
    )
 
    label_3_imagem = PhotoImage(
        file=relative_to_assets("imagem_sem_boleto.png"))
    label_3 = Label(
        window,
        image=label_3_imagem,
        )
    label_3.place(
        x=321,
        y=250,
        width=138.0,
        height=55.0
    )
 
    label_sub_3 = Label(
        window,
        textvariable=qtd_sem_boleto,
        font=("Malgun Gothic", 17, "bold"),
        fg="#000000",
        anchor="center",
        justify="center",
        bg="#ffffff",
        relief="groove"
        )
    label_sub_3.place(
        x=321,
        y=306,
        width=138.0,
        height=40.0
    )
 
    label_4_imagem = PhotoImage(
        file=relative_to_assets("imagem_processo_bloqueado.png"))
    label_4 = Label(
        window,
        image=label_4_imagem,
        )
    label_4.place(
        x=459,
        y=250,
        width=138.0,
        height=55.0
    )
 
    label_sub_4 = Label(
        window,
        textvariable=qtd_processo_bloqueado,
        font=("Malgun Gothic", 17, "bold"),
        fg="#000000",
        anchor="center",
        justify="center",
        bg="#ffffff",
        relief="groove"
        )
    label_sub_4.place(
        x=459,
        y=306,
        width=138.0,
        height=40.0
    )
 
    label_5_imagem = PhotoImage(
        file=relative_to_assets("imagem_processo_errado.png"))
    label_5 = Label(
        window,
        image=label_5_imagem,
        )
    label_5.place(
        x=597,
        y=250,
        width=146.0,
        height=55.0
    )
 
    label_sub_5 = Label(
        window,
        textvariable=qtd_processo_errado,
        font=("Malgun Gothic", 17, "bold"),
        fg="#000000",
        anchor="center",
        justify="center",
        bg="#ffffff",
        relief="groove"
        )
    label_sub_5.place(
        x=597,
        y=306,
        width=147.0,
        height=40.0
    )
 
    button_image_1 = PhotoImage(
        file=relative_to_assets("BotaoInicializarUsuario.png"))
    button_1 = Button(
        image=button_image_1,
        borderwidth=5,
        highlightthickness=0,
        command=lambda: responder_clique(inicializar_usuario),
        relief="raised",
        cursor="hand2"
    )
    button_1.place(
        x=50.0,
        y=70.0,
        width=261.0,
        height=41.0
    )
 
    button_image_2 = PhotoImage(
        file=relative_to_assets("BotaoPlay.png"))
    button_2 = Button(
        image=button_image_2,
        borderwidth=3,
        highlightthickness=0,
        bd=3,
        command=lambda: responder_clique(rodar_robozinho),
        relief="solid",
        cursor="hand2"
    )
    button_2.place(
        x=101.0,
        y=155.0,
        width=590,
        height=60
    )
 
    button_image_3 = PhotoImage(
    file=relative_to_assets("BotaoSemBoleto.png"))
    button_3 = Button(
        image=button_image_3,
        borderwidth=5,
        highlightthickness=0,
        command=lambda: threading.Thread(target=abrir_link_selenium, args=(sem_boleto,), daemon=True).start(),
        relief="groove",
        cursor="hand2"
    )
    button_3.place(
        x=28,
        y=397.99999999999994,
        width=145.0,
        height=33.0
    )
 
    button_image_4 = PhotoImage(
        file=relative_to_assets("BotaoProcessoBloqueado.png"))
    button_4 = Button(
        image=button_image_4,
        borderwidth=5,
        highlightthickness=0,
        command=lambda: threading.Thread(target=abrir_link_selenium, args=(processo_bloqueado,), daemon=True).start(),
        relief="groove",
        cursor="hand2"
    )
    button_4.place(
        x=203,
        y=397.99999999999994,
        width=189.0,
        height=33.0
    )
 
    button_image_5 = PhotoImage(
        file=relative_to_assets("BotaoXMLIndecifravel.png"))
    button_5 = Button(
        image=button_image_5,
        borderwidth=5,
        highlightthickness=0,
        command=lambda: threading.Thread(target=abrir_link_selenium, args=(XML_ilegivel,), daemon=True).start(),
        relief="groove",
        cursor="hand2"
    )
    button_5.place(
        x=422,
        y=397.99999999999994,
        width=165.0,
        height=33.0
    )
 
    button_image_6 = PhotoImage(
        file=relative_to_assets("BotaoProcessoErrado.png"))
    button_6 = Button(
        image=button_image_6,
        borderwidth=5,
        highlightthickness=0,
        command=lambda: threading.Thread(target=abrir_link_selenium, args=(processo_errado,), daemon=True).start(),
        relief="groove",
        cursor="hand2"
    )
    button_6.place(
        x=617,
        y=397.99999999999994,
        width=143.0,
        height=33.0
    )
 
    button_image_7 = PhotoImage(
        file=relative_to_assets("BotaoEQS.png"))
    button_7 = Button(
        image=button_image_7,
        borderwidth=2,
        highlightthickness=0,
        command=lambda: resetar_bot(),
        relief="groove",
        cursor="hand2"
    )
    button_7.place(
        x=590,
        y=20,
        width=166.0,
        height=100.0
    )
 
    canvas.create_rectangle(
        14.999999999999886,
        21.000000000000057,
        770.9999999999999,
        406.00000000000006,
        fill="#ffffff",
        outline="")
 
 
    segunda_logo = relative_to_assets("logoBratec.png")
    imagem_logo_esquerda = ImageTk.PhotoImage(Image.open(segunda_logo))
    label_logo_esquerda = Label(window, image=imagem_logo_esquerda, bg=cor_fundo)
    label_logo_esquerda.image = imagem_logo_esquerda
    label_logo_esquerda.place(x=265, y=12)
 
    window.resizable(False, False)
    window.mainloop()
 
 
