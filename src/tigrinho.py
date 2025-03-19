from selenium.webdriver.common.by import By
from selenium import webdriver                         
from pyperclip import paste, copy     
from time import sleep
import utils
import extratorXML
import tratamentoItem
import operadoresLancamento
import pyautogui as ptg
import xmltodict   
import pyscreeze

ptg.FAILSAFE = True
sem_boleto = []
processo_bloqueado = []
processo_errado = []
XML_ilegivel = []
nao_lancadas = []
processos_ja_vistos = []
mensagem_sb = "Processo sem boleto."
mensagem_pb = "Processo Bloqueado."
mensagem_pe = "Processo com algum erro impeditivo de lançamento."
mensagem_xi = "Processo com um XML que não consigo ler."


def robozinho(resetar=False):
    """
    Função principal. Nela está o fluxo da tarefa que esse programa realiza.
    É uma função recursiva devido a necessidade de reinicialização do processo
    que alguma circunstância indesejada pode provocar.
    """

    try:
        ver_documento = r'Imagens\BotaoVerDocumentos.png'
        utils.insistir_clique(ver_documento, cliques=1)
        sleep(0.4)
        utils.checar_failsafe()
        insistir_no_clique = utils.encontrar_imagem(ver_documento)
        if type(insistir_no_clique) == pyscreeze.Box:
            while True:
                utils.insistir_clique(ver_documento, cliques=1)
                insistir_no_clique = utils.encontrar_imagem(ver_documento)
                if type(insistir_no_clique) != pyscreeze.Box:
                    break
        ptg.hotkey("alt", "d", interval=0.1)
        sleep(0.5)
        ptg.hotkey("ctrl", "c", interval=0.5)
        link = paste()
        options = webdriver.ChromeOptions()
        options.add_argument(r'user-data-dir=C:\Users\Usuário\AppData\Local\Google\Chrome\User Data\Default')
        driver = webdriver.Chrome(options=options)
        sleep(0.5)
        driver.get(link)
        sleep(2)
        tempo_max = 0

        if resetar == True:
            lista_reset = [sem_boleto, processo_bloqueado, processo_errado, XML_ilegivel, nao_lancadas, processos_ja_vistos]
            for lista in lista_reset:
                lista.clear()


        while True:
            try:
                elemento1 = driver.find_element(By.XPATH, '/html/body/app-root/app-main/div/app-processo-pagamento-nota-manutencao/po-page-default/po-page/div/po-page-content/div/div[2]/po-tabs/div[1]/div/div/po-tab-button[2]/div[1]/span')
                sleep(0.3)
                elemento1.click()
                
                if elemento1 != '':
                    try:
                        elemento2 = driver.find_element(By.XPATH, '/html/body/app-root/app-main/div/app-processo-pagamento-nota-manutencao/po-page-default/po-page/div/po-page-content/div/div[1]/div[1]/po-widget/div/po-container/div/div/po-info[5]/div/div[2]/span')
                        chave_de_acesso = elemento2.text
                        try:
                            verificador = processos_ja_vistos.index(chave_de_acesso)
                            utils.reposicionar_robo()
                            driver.quit()
                            utils.checar_failsafe()
                            sleep(0.2)
                            return robozinho() 

                        except:
                            processos_ja_vistos.append(chave_de_acesso)
                            
                        elemento3 = driver.find_element(By.XPATH, '/html/body/app-root/app-main/div/app-processo-pagamento-nota-manutencao/po-page-default/po-page/div/po-page-content/div/div[2]/po-tabs/div[2]/po-tab[2]/div[1]/po-select/po-field-container/div/div[2]/select')
                        valor = elemento3.get_attribute("value")

                        if valor == 'B':
                            try:

                                # Trecho "removido" da automação devido a necessidade de otimizar o fluxo do processo
                                # (O setor compras não segue rigorosamente a regra de anexar os boletos
                                # no portal, o que implicava em muitos processos serem "pulados" pela automação).
                                # Caso um dia essa regra passe a ser seguida assiduamente, a funcionalidade de verificação
                                # tornará a ser ativada no código.

                                
                                #elemento4 = driver.find_element(By.XPATH, '/html/body/app-root/app-main/div/app-processo-pagamento-nota-manutencao/po-page-default/po-page/div/po-page-content/div/div[2]/po-tabs/div[2]/po-tab[2]/div[2]/po-table/po-container/div/div/div/div/div/table/tbody[1]/tr/td[4]/div/span/div[3]/po-input/po-field-container/div/div[2]/input')
                                #boleto = elemento4.get_attribute("value")
                                #if len(boleto) == 0:
                                #    utils.reposicionar_robo()
                                #    driver.quit()
                                #    sleep(0.2)
                                #    utils.acrescer_lista(sem_boleto, nao_lancadas, link, mensagem_sb)
                                #    return robozinho()
                                #else:

                                driver.quit() 
                                break

                            except Exception:
                                tempo_max += 1 
                                pass  

                        else:
                            driver.quit() 
                            break                                           

                    except Exception:
                        tempo_max += 1 
                        pass  

            except Exception:
                tempo_max += 1 
                pass
            
            if tempo_max == 15:
                ptg.press("enter")
            if tempo_max == 40:
                utils.reposicionar_robo()
                driver.quit()
                sleep(0.2)
                utils.checar_failsafe()

                # Circunstância indesejada:
                # Instabilidade no portal compras.
                
                return robozinho()  
            
        sleep(0.2)
        ptg.hotkey("ctrl", "w")
        sleep(0.3)


        # <DETALHES DO TRECHO>

        # O Código abaixo se dá desta forma devido as muitas possibilidades de estruturação 
        # que pode ter um XML. Através da captura de exceções o programa toma caminhos diferentes para tentar ler o arquivo.

        caminho = "C:\\Users\\Usuário\\Desktop\\xmlFiscalio\\" + chave_de_acesso + ".xml"

        aux = False
        while True:
            try:
                with open(caminho) as fd:
                    doc = xmltodict.parse(fd.read())
                    break
            except UnicodeDecodeError:
                with open(caminho, encoding='utf-8') as fd:
                    doc = xmltodict.parse(fd.read())
                    break
            except FileNotFoundError:
                while True:
                    utils.clicar_microsiga()
                    sleep(1)
                    x, y = utils.encontrar_centro_imagem(r'Imagens\BotaoExportarXML.png')
                    ptg.doubleClick(x, y)
                    sleep(2)
                    caixa_de_texto = utils.encontrar_centro_imagem(r'Imagens\ClicarInputServidor.png')
                    if type(caixa_de_texto) != tuple:  
                        ptg.doubleClick(x, y)
                        sleep(2)
                        utils.checar_failsafe()
                        caixa_de_texto = utils.encontrar_centro_imagem(r'Imagens\ClicarInputServidor.png')
                    if type(caixa_de_texto) == tuple:
                        break
                sleep(2)
                x, y = utils.encontrar_centro_imagem(r'Imagens\ClicarInputServidor.png')
                ptg.click(x,y, clicks=3, interval=0.07)
                copy("C:\\Users\\Usuário\\Desktop\\xmlFiscalio\\")
                ptg.hotkey("ctrl", "v")
                sleep(1)
                ptg.press(["tab"]*6, interval=0.5)
                ptg.press("enter", interval=0.8)
                utils.checar_failsafe()
                caixa_de_texto = utils.encontrar_centro_imagem(r'Imagens\ClicarInputServidor.png')
                if type(caixa_de_texto) == tuple:
                    botao_salvar = utils.encontrar_centro_imagem(r'Imagens\BotaoSalvar1.png')
                    x, y = botao_salvar
                    ptg.doubleClick(x, y)
                while True:
                    aparece_enter = utils.encontrar_centro_imagem(r'Imagens\XMLEnter.png')
                    if type(aparece_enter) == tuple:
                        ptg.press("enter", interval=0.5)
                    aparece_enter2 = utils.encontrar_centro_imagem(r'Imagens\XMLEnter2.png')
                    if type(aparece_enter2) == tuple:
                        while type(aparece_enter2) == tuple:
                            sleep(0.5)
                            ptg.press("enter", interval=0.5)
                            aparece_enter2 = utils.encontrar_centro_imagem(r'Imagens\XMLEnter2.png')
                        break
                caminho = "C:\\Users\\Usuário\\Desktop\\xmlFiscalio\\" + chave_de_acesso + ".xml"
                aux = True
            except:
                try:
                    with open(caminho, encoding='utf-8') as fd:
                        doc = xmltodict.parse(fd.read(), attr_prefix="@", cdata_key="#text")
                        break
                except xmltodict.expat.ExpatError as e:
                    if aux == True:
                        utils.tratar_xml_ilegivel(XML_ilegivel, nao_lancadas, link, mensagem_xi, aux)
                        
                        # Circunstância indesejada:
                        # XML ilegível
                        
                        return robozinho()
                    else:
                        utils.tratar_xml_ilegivel(XML_ilegivel, nao_lancadas, link, mensagem_xi)
                        return robozinho()
                
        # <DETALHES DO TRECHO/>
                


        # <DETALHES DO TRECHO>

        # O Código abaixo pode parecer um tanto confuso, mas ele é dessa maneira devido as muitas possibilidades de árvores XML
        # que podemos encontrar. O XML nada mais é do que um conjunto de chaves e valores, e chaves que
        # comportam mais chaves. Ele pode ter mais de um item, e tendo mais de um item,
        # a abordagem para extrair seus dados é uma (que segue a linha [const_item], que nada mais é do que acessar item por item
        # através de seu indice. Ex: 0, 1, 2...), enquanto quando há apenas um item no XML a abordagem é outra.

        # O XML da NF é transformado em um objeto, um dicionario Python, e como todo objeto, os valores são acessados através da chave.
        
        processador = extratorXML.ProcessadorXML(doc)
        valor_total_da_nf, filial_xml = processador.processar_totais_nota_fiscal()

        const_item = 0
        while True:
            try:
                coletor_xml = doc["nfeProc"]["NFe"]["infNFe"]["det"]["prod"]
                impostos_xml = doc["nfeProc"]["NFe"]["infNFe"]["det"]["imposto"]
                valores_do_item = processador.coletar_dados_XML(coletor_xml, impostos_xml)
                break
            except KeyError:
                try:
                    coletor_xml = doc["enviNFe"]["NFe"]["infNFe"]["det"]["prod"]
                    impostos_xml = doc["enviNFe"]["NFe"]["infNFe"]["det"]["imposto"]
                    valores_do_item = processador.coletar_dados_XML(coletor_xml, impostos_xml)
                    break
                except KeyError:
                    try:
                        coletor_xml = doc["NFe"]["infNFe"]["det"]["prod"]
                        impostos_xml = doc["NFe"]["infNFe"]["det"]["imposto"]
                        valores_do_item = processador.coletar_dados_XML(coletor_xml, impostos_xml)
                        break
                    except TypeError:
                        try:
                            coletor_xml = doc["NFe"]["infNFe"]["det"][const_item]["prod"]
                            impostos_xml = doc["NFe"]["infNFe"]["det"][const_item]["imposto"]
                            valores_do_item = processador.coletar_dados_XML(coletor_xml, impostos_xml)
                            const_item += 1
                        except IndexError:
                            break
                except TypeError:
                    try:
                        coletor_xml = doc["enviNFe"]["NFe"]["infNFe"]["det"][const_item]["prod"]
                        impostos_xml = doc["enviNFe"]["NFe"]["infNFe"]["det"][const_item]["imposto"]
                        valores_do_item = processador.coletar_dados_XML(coletor_xml, impostos_xml)
                        const_item += 1
                    except IndexError:
                        break
            except TypeError:
                try:
                    coletor_xml = doc["nfeProc"]["NFe"]["infNFe"]["det"][const_item]["prod"]
                    impostos_xml = doc["nfeProc"]["NFe"]["infNFe"]["det"][const_item]["imposto"]
                    valores_do_item = processador.coletar_dados_XML(coletor_xml, impostos_xml)
                    const_item += 1
                except IndexError:
                    break

        itens, indices_e_impostos = processador.trabalhar_dados_XML(valores_do_item)

        # <DETALHES DO TRECHO/>


        while True:
            utils.checar_failsafe()
            utils.clicar_dados_da_nota()
            sleep(1)
            abriu_a_tela = utils.encontrar_centro_imagem(r'Imagens\ReferenciaAbriuDadosDaNota.png')
            if type(abriu_a_tela) == tuple:
                break
        while True:
            ptg.press("tab", interval=0.7)
            ptg.hotkey("ctrl", "c")
            filial_pedido = paste()
            if filial_pedido == filial_xml:
                ptg.press("tab", interval=0.5)
                ptg.press("enter", interval=1)
                utils.checar_failsafe()
                clicar_confirmar = utils.encontrar_centro_imagem(r'Imagens\BotaoConfirmar.png')
                if type(clicar_confirmar) == tuple:
                    cont = 0
                    while cont < 5:
                        ptg.moveTo(150, 100)
                        x, y = clicar_confirmar
                        ptg.click(x,y, clicks=2, interval=0.07)
                        cont+=1
                break
            else:
                try:
                    utils.checar_failsafe()
                    clicar_cancelar = utils.encontrar_centro_imagem(r'Imagens\BotaoCancelarDadosNF.png')
                    x, y = clicar_cancelar
                    ptg.click(x,y, clicks=2, interval=0.07)
                    sleep(1)
                    clicar_cancelar = utils.encontrar_centro_imagem(r'Imagens\BotaoCancelarDadosNF.png')
                    if type(clicar_cancelar) == tuple:
                        while type(clicar_cancelar) == tuple:
                            ptg.moveTo(150, 100)
                            x, y = clicar_cancelar
                            ptg.click(x,y, clicks=2, interval=0.07)
                            clicar_cancelar = utils.encontrar_centro_imagem(r'Imagens\BotaoCancelarDadosNF.png')  
                    utils.cancelar1()
                    utils.acrescer_lista(processo_errado, nao_lancadas, link, mensagem_pe)

                    # Circunstância indesejada:
                    # Filial do pedido não corresponde a filial de XML
                    
                    return robozinho()
                except TypeError:
                    utils.clicar_dados_da_nota()


        try:
            sleep(0.5)
            utils.checar_failsafe()
            aparece_enter = utils.encontrar_imagem(r'Imagens\AtencaoEstoque.png')
            if type(aparece_enter) == pyscreeze.Box:
                sleep(0.3)
                ptg.press("enter")
            aparece_enter2 = utils.encontrar_imagem(r'Imagens\TES102.png')
            if type(aparece_enter2) == pyscreeze.Box:
                sleep(0.2)
                ptg.press("enter", interval=0.2)
                ptg.press(["tab"]*2)
                sleep(0.2)
                ptg.write("102")
                sleep(0.2)
                ptg.press(["tab"]*2, interval=0.5)
                sleep(0.2)
                ptg.press("enter") 
                utils.checar_failsafe()
        finally:
            pass

        
        tela_de_lancamento = utils.encontrar_imagem(r'Imagens\ReferenciaAbriuOProcesso.png')
        cont = 0
        while type(tela_de_lancamento) != pyscreeze.Box:
            cont +=1

            tela_de_lancamento = utils.encontrar_imagem(r'Imagens\ReferenciaAbriuOProcesso.png')
            lancamento_retroativo = utils.encontrar_imagem(r'Imagens\LancamentoRetroativo.png')
            nota_ja_lancada = utils.encontrar_imagem(r'Imagens\ProcessoJaLancado.png')
            fornecedor_bloqueado = utils.encontrar_imagem(r'Imagens\FornecedorBloqueado.png')
            utils.checar_failsafe()
            if type(lancamento_retroativo) == pyscreeze.Box or type(nota_ja_lancada) == pyscreeze.Box or type(fornecedor_bloqueado) == pyscreeze.Box:
                sleep(1)
                ptg.press("enter", interval=1)
                if type(fornecedor_bloqueado) == pyscreeze.Box:
                    utils.acrescer_lista(processo_bloqueado, nao_lancadas, link, mensagem_pb)
                cont = 0

            tela_de_lancamento = utils.encontrar_imagem(r'Imagens\ReferenciaAbriuOProcesso.png')
            erro_esquisito = utils.encontrar_imagem(r'Imagens\ErroEsquisito2.png')
            if type(erro_esquisito) == pyscreeze.Box:
                sleep(1)
                utils.checar_failsafe()
                ptg.press("enter")
                utils.cancelar1()

                # Circunstância indesejada:
                # Erro sistêmico na abertura do processo para lançamento.
                
                return robozinho()
            
            tela_de_lancamento = utils.encontrar_imagem(r'Imagens\ReferenciaAbriuOProcesso.png')
            erro_generico = utils.encontrar_imagem(r'Imagens\ErroGenerico.png')
            if type(erro_generico) == pyscreeze.Box:
                sleep(1)
                ptg.press("enter", interval=2) 
                ptg.press("esc", interval=2) 
                ptg.press("enter", interval=2)    
                utils.cancelar1()
                utils.checar_failsafe()
                utils.acrescer_lista(processo_bloqueado, nao_lancadas, link, mensagem_pb)
                return robozinho()
            
            tela_de_lancamento = utils.encontrar_imagem(r'Imagens\ReferenciaAbriuOProcesso.png')
            chave_nao_encontrada = utils.encontrar_imagem(r'Imagens\chaveNaoEncontradaNoSefaz.png')
            nf_cancelada = utils.encontrar_imagem(r'Imagens\NFCancelada.png')
            natureza_bloq = utils.encontrar_imagem(r'Imagens\NaturezaBloq.png')
            if type(chave_nao_encontrada) == pyscreeze.Box or type(natureza_bloq) == pyscreeze.Box or type(nf_cancelada) == pyscreeze.Box:
                sleep(1)
                ptg.press("enter")
                utils.cancelar3()
                utils.acrescer_lista(processo_bloqueado, nao_lancadas, link, mensagem_pb)
                return robozinho()
            if cont == 15:
                ptg.press("enter")
                cont = 0


        # Trecho diferente do Bot da matriz
        # Em vez de apertar "tab" 10x, como é no da matriz, o da Bratec anda apenas 9x

        sleep(0.5)
        ptg.press(["tab"]*9)
        sleep(0.8)
        ptg.press(["right"]*8)

        # Trecho diferente do Bot da matriz



        for i, ctrl_imposto in enumerate(indices_e_impostos):

            verificador, item_fracionado = operadoresLancamento.verificar_valor_item(itens, i)
            if verificador == True:
                utils.acrescer_lista(processo_errado, nao_lancadas, link, mensagem_pe)
                utils.checar_failsafe()
                
                # Circunstância indesejada:
                # Valor total do item não confere nem é passível de correção
                
                return robozinho()
            tratamento_item = tratamentoItem.TratadorItem(item_fracionado, itens, i, ctrl_imposto)
            item = tratamento_item.tratar_item()
            cont = 0

            match ctrl_imposto:
                case "Nenhum imposto":
                    for lista in item:
                        desc_no_item, frete_no_item, seg_no_item, desp_no_item, icms_no_item, icmsST_no_item, ipi_no_item = lista
                        natureza = operadoresLancamento.copiar_natureza()
                        codigo = operadoresLancamento.selecionar_caso(natureza)
                        tes = operadoresLancamento.definir_TES(codigo, ctrl_imposto)
                        if tes == True:
                            utils.acrescer_lista(processo_errado, nao_lancadas, link, mensagem_pe)

                            # Circunstância indesejada:
                            # Natureza não mapeada pela automação
                            
                            return robozinho()
                        operadoresLancamento.escrever_TES(tes, natureza)
                        operadoresLancamento.inserir_desconto(desc_no_item)
                        operadoresLancamento.inserir_frete(frete_no_item)
                        operadoresLancamento.inserir_seguro(seg_no_item)
                        operadoresLancamento.inserir_despesa(desp_no_item)
                        if tes in ["102", "405", "408"]:
                            operadoresLancamento.zerar_imposto()
                        elif tes in ["406", "421", "423"]:
                            operadoresLancamento.zerar_imposto()
                            operadoresLancamento.zerar_imposto(passos_ida=12, passos_volta=13)
                        ptg.press("down")
                        cont+=1
                        operadoresLancamento.corrigir_passos_horizontal(cont, item)
                    ptg.press("up")     
                                        
                case "Apenas o ICMS":
                    for lista in item:
                        desc_no_item, frete_no_item, seg_no_item, desp_no_item, icms_no_item, bc_icms, aliq_icms, icmsST_no_item, ipi_no_item = lista
                        natureza = operadoresLancamento.copiar_natureza()
                        codigo = operadoresLancamento.selecionar_caso(natureza)
                        tes = operadoresLancamento.definir_TES(codigo, ctrl_imposto)
                        if tes == True:
                            utils.acrescer_lista(processo_errado, nao_lancadas, link, mensagem_pe)
                            return robozinho()
                        operadoresLancamento.escrever_TES(tes, natureza)
                        operadoresLancamento.inserir_desconto(desc_no_item)
                        operadoresLancamento.inserir_frete(frete_no_item)
                        operadoresLancamento.inserir_seguro(seg_no_item)
                        operadoresLancamento.inserir_despesa(desp_no_item)
                        operadoresLancamento.inserir_ICMS(icms_no_item, bc_icms, aliq_icms)
                        ptg.press(["left"]*9)
                        if tes in ["406", "421", "423"]:
                            operadoresLancamento.zerar_imposto(passos_ida=12, passos_volta=13)
                        ptg.press("down")
                        cont+=1
                        operadoresLancamento.corrigir_passos_horizontal(cont, item)
                    ptg.press("up")
                                            
                case "Apenas o ICMSST":
                    for lista in item:
                        desc_no_item, frete_no_item, seg_no_item, desp_no_item, icms_no_item, icmsST_no_item, base_icms_ST, aliq_icms_ST, ipi_no_item = lista
                        natureza = operadoresLancamento.copiar_natureza()
                        codigo = operadoresLancamento.selecionar_caso(natureza)
                        tes = operadoresLancamento.definir_TES(codigo, ctrl_imposto)
                        if tes == True:
                            utils.acrescer_lista(processo_errado, nao_lancadas, link, mensagem_pe)
                            return robozinho()
                        operadoresLancamento.escrever_TES(tes, natureza)
                        operadoresLancamento.inserir_desconto(desc_no_item)
                        operadoresLancamento.inserir_frete(frete_no_item)
                        operadoresLancamento.inserir_seguro(seg_no_item)
                        operadoresLancamento.inserir_despesa(desp_no_item)
                        if tes in ["102", "405", "408"]:
                            operadoresLancamento.zerar_imposto()
                        elif tes in ["406", "421", "423"]:
                            operadoresLancamento.zerar_imposto()
                            operadoresLancamento.zerar_imposto(passos_ida=12, passos_volta=13)
                        operadoresLancamento.inserir_ICMSST(icmsST_no_item, base_icms_ST, aliq_icms_ST)
                        ptg.press("down")
                        cont+=1
                        operadoresLancamento.corrigir_passos_horizontal(cont, item)
                    ptg.press("up")
                                            
                case "Apenas o IPI":
                    for lista in item:
                        desc_no_item, frete_no_item, seg_no_item, desp_no_item, icms_no_item, icmsST_no_item, ipi_no_item, base_ipi, aliq_ipi = lista
                        natureza = operadoresLancamento.copiar_natureza()
                        codigo = operadoresLancamento.selecionar_caso(natureza)
                        tes = operadoresLancamento.definir_TES(codigo, ctrl_imposto)
                        if tes == True:
                            utils.acrescer_lista(processo_errado, nao_lancadas, link, mensagem_pe)
                            return robozinho()
                        operadoresLancamento.escrever_TES(tes, natureza)
                        operadoresLancamento.inserir_desconto(desc_no_item)
                        operadoresLancamento.inserir_frete(frete_no_item)
                        operadoresLancamento.inserir_seguro(seg_no_item)
                        operadoresLancamento.inserir_despesa(desp_no_item)
                        operadoresLancamento.inserir_IPI(ipi_no_item, base_ipi, aliq_ipi)
                        if tes in ["406", "421", "423", "102", "403", "411"]:
                            operadoresLancamento.zerar_imposto()
                        ptg.press("down")
                        cont+=1
                        operadoresLancamento.corrigir_passos_horizontal(cont, item)
                    ptg.press("up")
                                           
                case "Apenas ICMSST e IPI":
                    for lista in item:
                        desc_no_item, frete_no_item, seg_no_item, desp_no_item, icms_no_item, icmsST_no_item, base_icms_ST, aliq_icms_ST, ipi_no_item, base_ipi, aliq_ipi = lista
                        natureza = operadoresLancamento.copiar_natureza()
                        codigo = operadoresLancamento.selecionar_caso(natureza)
                        tes = operadoresLancamento.definir_TES(codigo, ctrl_imposto)
                        if tes == True:
                            utils.acrescer_lista(processo_errado, nao_lancadas, link, mensagem_pe)
                            return robozinho()
                        operadoresLancamento.escrever_TES(tes, natureza)
                        operadoresLancamento.inserir_desconto(desc_no_item)
                        operadoresLancamento.inserir_frete(frete_no_item)
                        operadoresLancamento.inserir_seguro(seg_no_item)
                        operadoresLancamento.inserir_despesa(desp_no_item)
                        if tes in ["406", "421", "423", "102", "411"]:
                            operadoresLancamento.zerar_imposto()
                        operadoresLancamento.inserir_ICMSST(icmsST_no_item, base_icms_ST, aliq_icms_ST)
                        operadoresLancamento.inserir_IPI(ipi_no_item, base_ipi, aliq_ipi, passosIPI=0)
                        ptg.press("down")
                        cont+=1
                        operadoresLancamento.corrigir_passos_horizontal(cont, item)
                    ptg.press("up")
                                            
                case "Apenas ICMS e IPI":
                    for lista in item:
                        desc_no_item, frete_no_item, seg_no_item, desp_no_item, icms_no_item, base_icms, aliq_icms, icmsST_no_item, ipi_no_item, base_ipi, aliq_ipi = lista
                        natureza = operadoresLancamento.copiar_natureza()
                        codigo = operadoresLancamento.selecionar_caso(natureza)
                        tes = operadoresLancamento.definir_TES(codigo, ctrl_imposto)
                        if tes == True:
                            utils.acrescer_lista(processo_errado, nao_lancadas, link, mensagem_pe)
                            return robozinho()
                        operadoresLancamento.escrever_TES(tes, natureza)
                        operadoresLancamento.inserir_desconto(desc_no_item)
                        operadoresLancamento.inserir_frete(frete_no_item)
                        operadoresLancamento.inserir_seguro(seg_no_item)
                        operadoresLancamento.inserir_despesa(desp_no_item)
                        operadoresLancamento.inserir_ICMS(icms_no_item, base_icms, aliq_icms)
                        operadoresLancamento.inserir_IPI(ipi_no_item, base_ipi, aliq_ipi, passosIPI=3)
                        ptg.press("down")
                        cont+=1
                        operadoresLancamento.corrigir_passos_horizontal(cont, item)
                    ptg.press("up")
                                           
                case "Apenas ICMS e ICMSST":
                    for lista in item:
                        desc_no_item, frete_no_item, seg_no_item, desp_no_item, icms_no_item, base_icms, aliq_icms, icmsST_no_item, base_icms_ST, aliq_icms_ST, ipi_no_item = lista
                        natureza = operadoresLancamento.copiar_natureza()
                        codigo = operadoresLancamento.selecionar_caso(natureza)
                        tes = operadoresLancamento.definir_TES(codigo, ctrl_imposto)
                        if tes == True:
                            utils.acrescer_lista(processo_errado, nao_lancadas, link, mensagem_pe)
                            return robozinho()
                        operadoresLancamento.escrever_TES(tes, natureza)
                        operadoresLancamento.inserir_desconto(desc_no_item)
                        operadoresLancamento.inserir_frete(frete_no_item)
                        operadoresLancamento.inserir_seguro(seg_no_item)
                        operadoresLancamento.inserir_despesa(desp_no_item)
                        operadoresLancamento.inserir_ICMS(icms_no_item, base_icms, aliq_icms)
                        operadoresLancamento.inserir_ICMSST(icmsST_no_item, base_icms_ST, aliq_icms_ST, passosST=0)
                        ptg.press("down")
                        cont+=1
                        operadoresLancamento.corrigir_passos_horizontal(cont, item)
                    ptg.press("up")
                                            
                case "Todos os impostos":
                    for lista in item:
                        desc_no_item, frete_no_item, seg_no_item, desp_no_item, icms_no_item, base_icms, aliq_icms, icmsST_no_item, base_icms_ST, aliq_icms_ST, ipi_no_item, base_ipi, aliq_ipi = lista
                        natureza = operadoresLancamento.copiar_natureza()
                        codigo = operadoresLancamento.selecionar_caso(natureza)
                        tes = operadoresLancamento.definir_TES(codigo, ctrl_imposto)
                        if tes == True:
                            utils.acrescer_lista(processo_errado, nao_lancadas, link, mensagem_pe)
                            return robozinho()
                        operadoresLancamento.escrever_TES(tes, natureza)
                        operadoresLancamento.inserir_desconto(desc_no_item)
                        operadoresLancamento.inserir_frete(frete_no_item)
                        operadoresLancamento.inserir_seguro(seg_no_item)
                        operadoresLancamento.inserir_despesa(desp_no_item)
                        operadoresLancamento.inserir_ICMS(icms_no_item, base_icms, aliq_icms)
                        operadoresLancamento.inserir_ICMSST(icmsST_no_item, base_icms_ST, aliq_icms_ST, passosST=0)
                        operadoresLancamento.inserir_IPI(ipi_no_item, base_ipi, aliq_ipi, passosIPI=12)
                        ptg.press("down")
                        cont+=1
                        operadoresLancamento.corrigir_passos_horizontal(cont, item)
                    ptg.press("up")
                                            

            if len(indices_e_impostos) > 1:
                ptg.press("down")
            if i+1 == len(indices_e_impostos):
                ptg.press("up")
            sleep(1.5)


        aba_duplicatas = utils.encontrar_centro_imagem(r'Imagens\BotaoAbaDuplicatas.png')
        x, y =  aba_duplicatas
        ptg.click(x,y, clicks=4, interval=0.1)
        utils.checar_failsafe()
        sleep(0.6)
        lista_parc = []
        utils.clicar_valor_parcela()
        sleep(0.5)
        ptg.hotkey("ctrl", "c", interval=0.2)
        valor_parcela = paste()
        valor_parcela = utils.formatador4(valor_parcela)
        utils.checar_failsafe()


        # <DETALHES DO TRECHO>

        # Esse trecho faz a validação do valor das parcelas, aplicando correção caso necessário. 
        # Ele verifica se o valor da primeira parcela corresponde ao valor total da NF, 
        # em caso de correspondência ele dá sequência no roteiro de lançamento acionando a função utils.clicar_natureza_duplicata(),
        # caso contrário, ele tenta seguir um primeiro aparelho lógico para tratamento do caso. 
        # Esse primeiro aparelho consiste em verificar se o valor da parcela copiada é maior ou menor que o valor total da NF, 
        # e aplica um pequeno roteiro para cada caso. - Entenda que há muitas possibilidades de erro de informação no Microsiga
        # nesta etapa. O sistema tende a puxar a informação do pedido, mas às vezes este dado está errado, 
        # ao mesmo tempo em que ele tenta puxar a informação do momento do lançamento. Sendo franco, é bem confuso a forma como 
        # o sistema trata essa informação, por isso eu desenvolvi a sequencia lógica abaixo, para tratar essas incertezas - 
        # Perceba que a todo momento está sendo verificado se o "ErroParcela" surge na tela. Dependendo do momento em que ele aparece,
        # ele servirá como um gatilho para executar o segundo aparelho lógico. Esse segundo aparelho consiste em buscar no Siga 
        # a quantidade de parcelas para aquele processo, e então fazer a divisão por igual do valor total da NF. Esse aparelho é pouco
        # acionado, a maioria dos casos conseguem ser tratado pelo primeiro, mas, quando primeiro caso não resolve tem essa segunda
        # possibilidade. E se nenhum dos dois resolver ele cancela o lançamento e trata o caso como um "Processo Errado".

        if valor_parcela < valor_total_da_nf:
            lista_parc.append(valor_parcela)
            while round(sum(lista_parc),2) < valor_total_da_nf:
                utils.descer_copiar()
                erro_parcela = utils.encontrar_imagem(r'Imagens\ErroParcela.png')
                if type(erro_parcela) == pyscreeze.Box:
                    ptg.press("enter", interval=0.7)
                    ptg.press("enter", interval=0.7)
                    lista_parc = lista_parc[:-1]
                    valor_parc = valor_total_da_nf - round(sum(lista_parc),2)
                    valor_parc = utils.formatador2(valor_parc)
                    ptg.write(valor_parc, interval=0.03)
                    ptg.press("left")
                    lista_parc.append(float(valor_parc))
                else:
                    valor_parcela = paste()
                    valor_parcela = utils.formatador4(valor_parcela)
                    lista_parc.append(valor_parcela)
            somatoria = utils.formatador2(sum(lista_parc))
            somatoria = float(somatoria)
            parcela_errada = lista_parc[-1]
            if somatoria != valor_total_da_nf:
                if lista_parc[-1] == lista_parc[-2]:
                    parcela_errada = lista_parc.pop()
                    somatoria = utils.formatador2(sum(lista_parc))
                    somatoria = float(somatoria)
                diferenca_NF_siga = valor_total_da_nf - somatoria 
                ultima_parcela = parcela_errada + diferenca_NF_siga
                ultima_parcela = "{:.2f}".format(ultima_parcela)  
                ptg.click(x,y)
                descida = len(lista_parc) - 1
                ptg.press(["down"]*descida)
                sleep(0.7)
                ptg.write(ultima_parcela, interval=0.03)
                utils.checar_failsafe()
            sleep(1)

        elif valor_parcela > valor_total_da_nf:
            valor_total_da_nf = utils.formatador2(valor_total_da_nf)
            ptg.write(valor_total_da_nf)
            sleep(1)

        utils.clicar_natureza_duplicata()
        sleep(1)

        erro_parcela = utils.encontrar_imagem(r'Imagens\ErroParcela.png')
        if type(erro_parcela) == pyscreeze.Box:
            ptg.press("enter")
            utils.clicar_valor_parcela()
            utils.checar_failsafe()
            ptg.press(["left"]*2)
            sleep(0.3)
            ptg.hotkey("ctrl", "c", interval=0.1)
            primeira_parc = paste()
            ordem_parc = []
            ordem_parc.append(primeira_parc)

            if primeira_parc == '001':
                utils.descer_copiar()
                proxima_parcela = paste()
                ordem_parc.append(proxima_parcela)
                utils.checar_failsafe()
                if ordem_parc[-2] != ordem_parc[-1]:
                    while ordem_parc[-2] != ordem_parc[-1]:
                        utils.descer_copiar()
                        proxima_parcela = paste()
                        ordem_parc.append(proxima_parcela)
                    ordem_parc.pop()
                    valor_parcela = valor_total_da_nf / len(ordem_parc)
                    valor_parcela = "{:.2f}".format(valor_parcela)
                    utils.clicar_valor_parcela()
                    for vezes in range(len(ordem_parc)):
                        ptg.write(valor_parcela, interval=0.08)
                        ptg.press("left")
                        ptg.press("down", interval=0.8)
                    valor_parcela = utils.formatador3(valor_parcela)
                    valor_atingido = valor_parcela * len(ordem_parc)
                    utils.checar_failsafe()
                    sleep(2)
                    if valor_atingido != valor_total_da_nf:
                        diferenca_NF_siga = valor_atingido - valor_total_da_nf
                        valor_ultima_parcela = valor_parcela - diferenca_NF_siga
                        valor_ultima_parcela = "{:.2f}".format(valor_ultima_parcela)
                        ptg.write(valor_ultima_parcela, interval=0.08)
                        sleep(2)

        # <DETALHES DO TRECHO/>


            utils.clicar_natureza_duplicata()
            sleep(0.6)

            erro_parcela = utils.encontrar_imagem(r'Imagens\ErroParcela.png')
            if type(erro_parcela) == pyscreeze.Box:
                ptg.press("enter")
                utils.cancelar2()
                utils.acrescer_lista(processo_errado, nao_lancadas, link, mensagem_pe)
                utils.checar_failsafe()

                # Circunstância indesejada:
                # Não foi possível corrigir o valor dos títulos
                # com a lógica estabelecida nessa automação
                
                return robozinho()
            
        ptg.hotkey("ctrl", "c", interval=0.2)

        natureza_perc = paste() 
        if natureza_perc != "0,00":
            lista_perc = []
            while round(sum(lista_perc),2) < 100.0:
                natureza_perc = utils.formatador3(natureza_perc)
                lista_perc.append(natureza_perc)
                utils.descer_copiar()
                natureza_perc = paste() 
            maior_perc = max(lista_perc)
            natureza_duplicata_clique = utils.encontrar_centro_imagem(r'Imagens\ClicarNaturezaDuplicata.png')
            x, y = natureza_duplicata_clique
            ptg.click(x,y)
            ptg.press("up", interval=0.2)
            ptg.hotkey("ctrl", "c", interval=0.1)
            perc_majoritario = paste()
            utils.checar_failsafe()
            perc_majoritario = utils.formatador3(perc_majoritario)
            while perc_majoritario != maior_perc:
                utils.descer_copiar()
                perc_majoritario = paste()
                perc_majoritario = utils.formatador3(perc_majoritario)
            ptg.press("left")
            ptg.hotkey("ctrl", "c", interval=0.1)
            natureza_duplicata = paste()
            ptg.hotkey(["shift", "tab"]*5, interval=0.2)
            ptg.write(natureza_duplicata)
            ptg.press("tab", interval=1)
            utils.checar_failsafe()


        salvar = utils.encontrar_centro_imagem(r'Imagens\BotaoSalvarLancamento.png')
        salvarx, salvary = salvar
        sleep(0.7)
        ptg.click(salvarx,salvary, clicks=2, interval=0.1)
        sleep(2)
        utils.checar_failsafe()
        cont = 0
        while True:
            salvar = utils.encontrar_centro_imagem(r'Imagens\BotaoSalvarLancamento.png')
            if type(salvar) == tuple:
                ptg.click(salvarx,salvary, clicks=2, interval=0.1)
                cont += 1
                sleep(1)
                if cont == 2:
                    break
            else:
                break

        erro_de_serie = utils.encontrar_imagem(r'Imagens\ErroDeSerie.png')
        erro_de_modelo = utils.encontrar_imagem(r'Imagens\ErroDeModulo.png')
        if type(erro_de_serie) == pyscreeze.Box or type(erro_de_modelo) == pyscreeze.Box:
            ptg.press("enter", interval=0.2) 
            espec_doc = utils.encontrar_centro_imagem(r'Imagens\CampoESPEC.png')
            x, y = espec_doc
            sleep(0.5)
            ptg.click(x,y, clicks=2)
            ptg.write("NF", interval=0.1)
            ptg.press("enter", interval=0.5)
            utils.checar_failsafe()
            ptg.doubleClick(salvarx, salvary)
        erro_esquisito = utils.encontrar_imagem(r'Imagens\ErroEsquisito.png')
        if type(erro_esquisito) == pyscreeze.Box:
            ptg.press("esc")
            quit()
        erro_quantidade = utils.encontrar_imagem(r'Imagens\ErroDeQuantidade.png')
        if type(erro_quantidade) == pyscreeze.Box:
            ptg.press("enter")
            utils.cancelar_lancamento()
            mudar_a_selecao = utils.encontrar_centro_imagem(imagem=r'Imagens\ClicarMudarSelecao.png')
            x, y = mudar_a_selecao
            ptg.doubleClick(x, y)
            sleep(0.3)
            utils.clicar_microsiga()
            utils.acrescer_lista(processo_errado, nao_lancadas, link, mensagem_pe)

            # Circunstância indesejada:
            # Erro de quantidade no estoque da empresa
            # (Esse erro só pode ser ajustado pelos colaboradores do almoxarifado)
            
            return robozinho()


        cont = 0
        etapa_final = utils.encontrar_imagem(r'Imagens\ReferenciaEtapaFinal.png')
        while type(etapa_final) != pyscreeze.Box:
            sleep(0.2)
            etapa_final = utils.encontrar_imagem(r'Imagens\ReferenciaEtapaFinal.png')
        ptg.press(["tab"]*3, interval=0.9)
        ptg.press("enter", interval=1.5)
        ultimo_enter = utils.encontrar_imagem(r'Imagens\ReferenciaFinalizarLancamento.png')
        if type(ultimo_enter) != pyscreeze.Box:
            while type(ultimo_enter) != pyscreeze.Box:
                sleep(0.2)
                ultimo_enter = utils.encontrar_imagem(r'Imagens\ReferenciaFinalizarLancamento.png')
                cont +=1
                if cont == 6:
                    ptg.press("enter")
                    cont = 0
        ptg.press("tab", interval=0.9)
        ptg.press("enter")
        utils.checar_failsafe()
        aux = False
        cont2 = 0
        while True:
            ultima_tela = utils.encontrar_imagem(r'Imagens\ReferenciaAguarde.png')
            if type(ultima_tela) == pyscreeze.Box:
                aux = True
                while type(ultima_tela) == pyscreeze.Box:
                    ultima_tela = utils.encontrar_imagem(r'Imagens\ReferenciaAguarde.png')
                    sleep(0.2)
            if aux == True:
                break
            ultimo_enter = utils.encontrar_imagem(r'Imagens\ReferenciaFinalizarLancamento.png')
            if type(ultimo_enter) == pyscreeze.Box:
                cont +=1
                if cont == 4:
                    aux = True
                    ultima_tentativa = utils.encontrar_centro_imagem(imagem=r'Imagens\BotaoConfirma2.png')
                    x, y = ultima_tentativa
                    ptg.doubleClick(x, y)
                    ptg.moveTo(150, 100)
            if cont2 == 5:
                break
            cont2 +=1
        

        ptg.hotkey("win", "d", interval=0.2)


        abortar = False
        return sem_boleto, processo_bloqueado, processo_errado, XML_ilegivel, nao_lancadas, abortar
    except ptg.FailSafeException:
        abortar = True
        return sem_boleto, processo_bloqueado, processo_errado, XML_ilegivel, nao_lancadas, abortar    


   
