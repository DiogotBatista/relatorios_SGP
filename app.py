from tkinter import *
from Relatorios_SGP import gerar_relatorios_sgp
from Relatorios_gerais import gerar_relatorios_producao
from Atualizar_Periodo_db import janela_atualizar_periodo

def main():
    janela = Tk()
    altura_botão = 2
    largura_botao = 30
    altura = 10
    largura = 100
    fonte = ('Arial', 13)

    janela.title('Retirar relatórios do SGP')

    texto = Label(janela, text="Escolha qual relatório deseja retirar:", font=fonte)
    texto.grid(column=0, row=0, padx=largura, pady=altura)

    botao_relatorio_gerais = Button(janela, text='Relatórios Gerais',font=fonte, command=gerar_relatorios_sgp, width=largura_botao, height=altura_botão, activeforeground="green")
    botao_relatorio_gerais.grid(column=0, row=1, padx=largura, pady=altura)

    botao_relatorio_producao = Button(janela, text='Relatórios de Produção',font=fonte, command=gerar_relatorios_producao, width=largura_botao, height=altura_botão)
    botao_relatorio_producao.grid(column=0, row=2, padx=largura, pady=altura)

    botao_atualizar_bd = Button(janela, text='Atualizar Periodo',font=fonte, command=janela_atualizar_periodo, width=largura_botao, height=altura_botão)
    botao_atualizar_bd.grid(column=0, row=3, padx=largura, pady=altura)

    rodape = Label(janela, text="Desenvolvido por Diogo Batista ")
    rodape.grid(column=0, row=4, sticky='ne')

    janela.mainloop()

if __name__ == "__main__":
    main()
