import time
import pyautogui as p
import keyboard as k
import acesso_bd as bd

# Constantes
p.PAUSE = 0.8
tempo_de_execucao = 0.3
sgp_geral = bd.sgp_geral
acoes_abertas = bd.acoes_abertas
acoes_fechadas = bd.acoes_fechadas
dia_inicial = bd.dia_inicial
mes_inicial = bd.mes_inicial
ano_inicial = bd.ano_inicial
botao_cadastros = [30, 28]
botao_relatorios = [90, 31]
botao_obras = [52, 160]
botao_relatorios_acoes = [133, 75]


def localizar_e_clicar(imagem):
    elemento = None
    while not elemento:
        elemento = p.locateCenterOnScreen(imagem)
    p.click(elemento[0], elemento[1])


def localizar_e_rolar(coordenadas):
    p.moveTo(coordenadas[0], coordenadas[1])
    p.mouseDown()
    p.moveTo(coordenadas[2], coordenadas[3])
    p.mouseUp()


def esperar_tela_aparecer(imagem):
    # while not p.pixelMatchesColor(coordenadas[0], coordenadas[1], cor_esperada):
    while not p.locateOnScreen(imagem):
        pass


def imprimir():
    elemento = None
    while not elemento:
        elemento = p.locateOnScreen("imagens/of2.png")
    impressora = p.locateCenterOnScreen("imagens/impressora.png")
    p.click(impressora[0], impressora[1])


def localizar_e_preencher(imagem, texto):
    elemento = None
    while not elemento:
        elemento = p.locateCenterOnScreen(imagem)
    p.click(elemento[0], elemento[1], clicks=2)
    k.write(texto)


def aguardar_e_fechar():
    elemento = p.locateCenterOnScreen("imagens/cancel.png")
    while elemento:
        elemento = p.locateCenterOnScreen("imagens/cancel.png")
    close = p.locateCenterOnScreen("imagens/close.png")
    p.click(close[0], close[1])


def preencher_datas(coordenadas, texto):
    p.click(coordenadas[0], coordenadas[1])
    p.write(texto)


def localizar_e_clicar_coordenada(coordenada):
    p.click(coordenada[0], coordenada[1])


def Relatorio_sgp_geral():
    # localizar_e_clicar('imagens/cadastros.png')
    localizar_e_clicar_coordenada(botao_cadastros)
    # localizar_e_clicar('imagens/Obras.png')
    localizar_e_clicar_coordenada(botao_obras)
    localizar_e_clicar('imagens/filtrar_por.png')
    localizar_e_clicar('imagens/todas_ns.png')
    localizar_e_clicar('imagens/empresa.png')
    localizar_e_clicar('imagens/cro.png')
    localizar_e_clicar('imagens/filtrar.png')
    time.sleep(2)
    p.moveTo(1382, 1063)
    p.click()
    esperar_tela_aparecer('imagens/verificacao.png')
    time.sleep(1)
    # Manda imprimir o relatorio
    localizar_e_clicar('imagens/botao_imprimir.png')
    # Solicita impressão do relatorio SGP Geral
    time.sleep(10)
    imprimir()
    # Preparar impressão
    localizar_e_clicar('imagens/print_to_file.png')
    localizar_e_clicar('imagens/type.png')
    coordenadas = [1080, 614, 1086, 656]
    localizar_e_rolar(coordenadas)
    localizar_e_clicar('imagens/xlsx_date.png')
    localizar_e_clicar('imagens/....png')
    # preencher endereço e nome e salva no servidor
    endereco = sgp_geral
    localizar_e_preencher('imagens/local_nome.png', endereco)
    localizar_e_clicar('imagens/salvar.png')
    localizar_e_clicar('imagens/ok.png')
    localizar_e_clicar('imagens/yes.png')
    # aguardar salvamento e fechar tudo
    aguardar_e_fechar()
    localizar_e_clicar('imagens/sair.png')


def Abrir_relatorios_acoes():
    # Abrir relatorios
    # localizar_e_clicar('imagens/relatorios.png')
    localizar_e_clicar_coordenada(botao_relatorios)
    # localizar_e_clicar('imagens/relatorios_acoes.png')
    localizar_e_clicar_coordenada(botao_relatorios_acoes)


def NS_abertas():
    localizar_e_clicar('imagens/filtro_acoes.png')
    localizar_e_clicar('imagens/dt_abertura_acao.png')
    coordenadas = [785, 390]
    ano = r'1999'
    preencher_datas(coordenadas, ano)
    coordenadas = [922, 392]
    ano = r'2050'
    preencher_datas(coordenadas, ano)
    localizar_e_clicar('imagens/status.png')
    localizar_e_clicar('imagens/aberto.png')
    localizar_e_clicar('imagens/imprimir_acoes.png')
    # Aguardar criação do relatorio e mandar imprimir
    time.sleep(10)
    imprimir()
    # Preparar impressão
    localizar_e_clicar('imagens/print_to_file.png')
    localizar_e_clicar('imagens/type.png')
    coordenadas = [1080, 614, 1080, 656]
    localizar_e_rolar(coordenadas)
    localizar_e_clicar('imagens/xlsx_date.png')
    localizar_e_clicar('imagens/....png')
    # preencher endereço e nome e salva no servidor
    endereco = acoes_abertas
    localizar_e_preencher('imagens/local_nome.png', endereco)
    localizar_e_clicar('imagens/salvar.png')
    localizar_e_clicar('imagens/ok.png')
    localizar_e_clicar('imagens/yes.png')
    # aguardar salvamento e fechar o relatorio de impressao
    aguardar_e_fechar()


def NS_fechadas():
    localizar_e_clicar('imagens/filtro_acoes_2.png')
    localizar_e_clicar('imagens/dt_fechamento_acao.png')
    # preenche o dia inicial da apuração
    coordenadas = [744, 395]
    preencher_datas(coordenadas, dia_inicial)
    # preenche o mes inicial da apuração
    coordenadas = [761, 390]
    preencher_datas(coordenadas, mes_inicial)
    # preenche o ano inicial da apuração
    coordenadas = [787, 393]
    preencher_datas(coordenadas, ano_inicial)
    localizar_e_clicar('imagens/aberto.png')
    localizar_e_clicar('imagens/todos.png')
    localizar_e_clicar('imagens/imprimir_acoes.png')
    # Aguardar criação do relatorio e mandar imprimir
    time.sleep(10)
    imprimir()
    # Preparar impressão
    localizar_e_clicar('imagens/print_to_file.png')
    localizar_e_clicar('imagens/type.png')
    coordenadas = [1080, 614, 1080, 656]
    localizar_e_rolar(coordenadas)
    localizar_e_clicar('imagens/xlsx_date.png')
    localizar_e_clicar('imagens/....png')
    # preencher endereço e nome e salva no servidor
    endereco = acoes_fechadas
    localizar_e_preencher('imagens/local_nome.png', endereco)
    localizar_e_clicar('imagens/salvar.png')
    localizar_e_clicar('imagens/ok.png')
    localizar_e_clicar('imagens/yes.png')
    # aguardar salvamento e fechar o relatorio de impressao
    aguardar_e_fechar()


def fechar_tudo():
    localizar_e_clicar('imagens/Fechar_x.png')


def gerar_relatorios_sgp():
    iniciar = p.confirm(text='Iniciar atualização automática dos relatorios do SGP?',
                        title='Aumotatização do Relatório do SGP', buttons=['OK', 'CANCELAR'])
    if iniciar == 'OK':
        if p.locateOnScreen('imagens/tela_inicial.png'):
            Relatorio_sgp_geral()
            Abrir_relatorios_acoes()
            NS_abertas()
            NS_fechadas()
            fechar_tudo()
            time.sleep(1)
            p.alert(text='Relatórios ATUALIZADOS!', title='Conclusão', button='OK')
        else:
            p.alert(text='SGP não localizado', title='CANCELADO', button='OK')
    else:
        p.alert(text='Atualização CANCELADA', title='CANCELADO', button='OK')


if __name__ == "__main__":
    gerar_relatorios_sgp()
