import pyodbc
import tkinter as tk
from tkinter import simpledialog, messagebox  # Importar também o módulo 'messagebox'

# Função para atualizar a tabela no Access
def atualizar_tabela(valor):
    conn_str = r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=C:\GESTÃO\PYTHON\AUTOMACAO\RELATORIOS_SGP\BD\bd_relatorios.accdb;'
    connection = pyodbc.connect(conn_str)

    cursor = connection.cursor()
    cursor.execute("UPDATE Datas SET Periodo = ? WHERE Id = 1", (valor,))
    connection.commit()

    connection.close()

def janela_principal():
    # Criar uma janela principal
    root = tk.Tk()
    root.withdraw()  # Ocultar a janela principal para mostrar apenas a caixa de diálogo

    # Solicitar o valor diretamente
    valor = simpledialog.askstring("Atualização do Período", "Digite o valor do novo período de apuração dos relatórios:")
    if valor is not None and valor.strip() != "":
        valor += ".xlsx"  # Concatenar a extensão .xlsx ao valor
        atualizar_tabela(valor)
        messagebox.showinfo("Sucesso", "Valor atualizado com sucesso!")

    # Fechar a janela após o término da caixa de diálogo
    root.destroy()

def janela_atualizar_periodo():
    janela_principal()


if __name__ == "__main__":
    janela_atualizar_periodo()