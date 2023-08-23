import time
import pyautogui as p
import keyboard as k
from datetime import datetime, timedelta
import acesso_bd as bd

# Constantes
p.PAUSE = 0.8
tempo_de_execucao = 0.3
data_atual = datetime.now()
mes_inicial = bd.mes_inicial
dia_nicial = bd.dia_inicial
ano_inicial = bd.ano_inicial
mes_final = bd.mes_final
ano_final = bd.ano_final
test = str(datetime.now().day)
dia_anterior_calculo = data_atual - timedelta(days=1)
dia_final = str(dia_anterior_calculo.day)
producao_analitica = bd.producao_analitica
protocolo = bd.protocolo
botao_relatorios = [90,31]
botao_relatorios_gerais = [127,54]

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
        elemento = p.locateCenterOnScreen("imagens/of2.png")
    impressora = p.locateCenterOnScreen("imagens/impressora.png")
    p.click(impressora[0], impressora[1])

def imprimir_analitico():
    elemento = None
    while not elemento:
        elemento = p.locateCenterOnScreen("imagens/page1.png")
    impressora = p.locateCenterOnScreen("imagens/impressora.png")
    p.click(impressora[0], impressora[1])

def localizar_e_preencher(imagem,texto):
    elemento = None
    while not elemento:
        elemento = p.locateCenterOnScreen(imagem)
    p.click(elemento[0], elemento[1], clicks = 2)
    k.write(texto)

def aguardar_e_fechar():
    elemento = p.locateCenterOnScreen("imagens/cancel.png")
    while elemento:
        elemento = p.locateCenterOnScreen("imagens/cancel.png")
    close = p.locateCenterOnScreen("imagens/close.png")
    p.click(close[0], close[1])

def preencher_datas(coordenadas, texto):
    p.click(coordenadas[0], coordenadas[1])
    time.sleep(0.5)
    p.write(texto)

def localizar_e_clicar_coordenada(coordenada):
    p.click(coordenada[0], coordenada[1])
def Relatorio_producao_analitica():
    # localizar_e_clicar('imagens/relatorios_sgp.png')
    # localizar_e_clicar('imagens/relatorios_gerais.png')
    localizar_e_clicar_coordenada(botao_relatorios)
    localizar_e_clicar_coordenada(botao_relatorios_gerais)
    time.sleep(100)
    localizar_e_clicar('imagens/producao_analitica.png')
    localizar_e_clicar('imagens/data_de_vencimento_geral.png')
    localizar_e_clicar('imagens/dt_conclusao_acao_analitco.png')
    #preencher dia inicial
    coordenadas = [683,358]
    data = "01"
    preencher_datas(coordenadas,data)
    #preencher mes inicial
    coordenadas = [704,357]
    data = "12"
    preencher_datas(coordenadas,data)
    #preencher ano inicial
    coordenadas = [729,354]
    data = "2022"
    preencher_datas(coordenadas,data)
    # Manda imprimir o relatorio
    localizar_e_clicar('imagens/imprimir_geral.png')
    # Solicita impressão do relatorio
    esperar_tela_aparecer('imagens/logo_cro.png')
    time.sleep(10)
    imprimir_analitico()
    # Preparar impressão
    localizar_e_clicar('imagens/print_to_file.png')
    localizar_e_clicar('imagens/type.png')
    coordenadas = [925, 528, 922, 566]
    localizar_e_rolar(coordenadas)
    localizar_e_clicar('imagens/xlsx_date.png')
    localizar_e_clicar('imagens/....png')
    #preencher endereço e nome e salva no servidor
    endereco = producao_analitica
    localizar_e_preencher('imagens/local_nome.png', endereco)
    localizar_e_clicar('imagens/salvar.png')
    localizar_e_clicar('imagens/ok.png')
    localizar_e_clicar('imagens/yes.png')
    #aguardar salvamento e fechar tudo
    aguardar_e_fechar()

def Relatorio_protocolo():
    localizar_e_clicar('imagens/protocolo.png')
    #preencher mes inicial
    coordenadas = [704,357]
    preencher_datas(coordenadas,mes_inicial)
    #preencher ano inicial
    coordenadas = [729,354]
    preencher_datas(coordenadas,ano_inicial)
    #preencher dia final
    coordenadas = [813,358]
    preencher_datas(coordenadas,dia_final)
    #preencher mes final
    coordenadas = [835,355]
    preencher_datas(coordenadas,mes_final)
    # preencher ano final
    coordenadas = [858, 362]
    preencher_datas(coordenadas, ano_final)
    # Manda imprimir o relatorio
    localizar_e_clicar('imagens/imprimir_geral.png')
    # Solicita impressão do relatorio
    esperar_tela_aparecer('imagens/logo_cro2.png')
    time.sleep(10)
    imprimir()
    # Preparar impressão
    localizar_e_clicar('imagens/print_to_file.png')
    localizar_e_clicar('imagens/type.png')
    coordenadas = [925, 528, 922, 566]
    localizar_e_rolar(coordenadas)
    localizar_e_clicar('imagens/xlsx_date.png')
    localizar_e_clicar('imagens/....png')
    #preencher endereço e nome e salva no servidor
    endereco = protocolo
    localizar_e_preencher('imagens/local_nome.png', endereco)
    localizar_e_clicar('imagens/salvar.png')
    localizar_e_clicar('imagens/ok.png')
    localizar_e_clicar('imagens/yes.png')
    #aguardar salvamento e fechar tudo
    aguardar_e_fechar()
    localizar_e_clicar('imagens/fechar.png')

def gerar_relatorios_producao():
    iniciar = p.confirm(text='Iniciar atualização automática dos relatorios do SGP?', title='Aumotatização do Relatório do SGP', buttons=['OK', 'CANCELAR'])
    if iniciar == 'OK':
        if p.locateOnScreen('imagens/tela_inicial.png'):
            Relatorio_producao_analitica()
            Relatorio_protocolo()
            p.alert(text='Relatórios ATUALIZADOS!', title='Conclusão', button='OK')
        else:
            p.alert(text='SGP não localizado', title='CANCELADO', button='OK')
    else:
        p.alert(text='Atualização CANCELADA', title='CANCELADO', button='OK')

if __name__ == "__main__":
    gerar_relatorios_producao()
