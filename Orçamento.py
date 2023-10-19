import tkinter as tk
from fpdf import FPDF
from tkinter import *
from datetime import datetime
from random import randint
import win32com.client as win32


# Criar a janela da interface
janela = tk.Tk()
janela.title('Gerador de orçamento')
janela.geometry('390x270')
janela.config(bg='#f5f5f5')
janela.resizable(width=False, height=False)

# Criação dos widgets 
# Descrição do projeto
label_1 = (tk.Label(janela, text='Descrição do projeto', bg='#f5f5f5', font=('Roboto',10)))
label_1.grid(row=0, column=0, padx=5, pady=5, sticky='w')
entry_1 = (tk.Entry(janela, width = 30))
entry_1.grid(row=0, column=1, padx=5, pady=5, sticky='w')

# Horas
label_2 = (tk.Label(janela, text='Horas estimadas', bg='#f5f5f5', font=('Roboto',10)))
label_2.grid(row=1, column=0, padx=5, pady=5, sticky='w')
entry_2 = (tk.Entry(janela, width = 30))
entry_2.grid(row=1, column=1, padx=5, pady=5, sticky='w')

# Valor da hora
label_3 = (tk.Label(janela, text='Valor da hora trabalhada', bg='#f5f5f5', font=('Roboto',10)))
label_3.grid(row=2, column=0, padx=5, pady=5, sticky='w')
entry_3 = (tk.Entry(janela, width = 30))
entry_3.grid(row=2, column=1, padx=5, pady=5, sticky='w')

# Prazo de entrega
label_4 = (tk.Label(janela, text='Prazo de entrega', bg='#f5f5f5', font=('Roboto',10)))
label_4.grid(row=3, column=0, padx=5, pady=5, sticky='w')
entry_4 = (tk.Entry(janela, width = 30))
entry_4.grid(row=3, column=1, padx=5, pady=5, sticky='w')

# Cliente
label_5 = (tk.Label(janela, text='Cliente', bg='#f5f5f5', font=('Roboto',10)))
label_5.grid(row=4, column=0, padx=5, pady=5, sticky='w')
entry_5 = (tk.Entry(janela, width = 30))
entry_5.grid(row=4, column=1, padx=5, pady=5, sticky='w')

# Dados cliente 
label_6 = (tk.Label(janela, text='e-mail do cliente', bg='#f5f5f5', font=('Roboto',10)))
label_6.grid(row=5, column=0, padx=5, pady=5, sticky='w')
entry_6 = (tk.Entry(janela, width = 30))
entry_6.grid(row=5, column=1, padx=5, pady=5, sticky='w')

# Caminho do orçamento
label_7 = (tk.Label(janela, text='caminho orçamento', bg='#f5f5f5', font=('Roboto',10)))
label_7.grid(row=6, column=0, padx=5, pady=5, sticky='w')
entry_7 = (tk.Entry(janela, width = 30))
entry_7.grid(row=6, column=1, padx=5, pady=5, sticky='w')


# Criando um número para a proposta
data_atual = datetime.now()
data_formatada = data_atual.strftime('%d.%m')
numero = randint(0,10)


def gerador():
    # Criando as variáveis
    projeto = entry_1.get()
    horas_previstas = entry_2.get()
    valor_hora = entry_3.get()
    prazo = entry_4.get()
    cliente = entry_5.get()
    valor_total = int(valor_hora)*int(horas_previstas)
    nome_arquivo = f'orçamento.{cliente}.{numero}_{data_formatada}.pdf'
    
    # Imputando os dados no pdf
    pdf = FPDF() 
    pdf.add_page()
    pdf.set_font("Arial")
    pdf.image("template.png", x=0, y=0)

    pdf.text(115,145, projeto)
    pdf.text(115,160, horas_previstas)
    pdf.text(115,175, valor_hora)
    pdf.text(115,190, prazo)
    pdf.text(115,205, str(valor_total))

    pdf.output(nome_arquivo)
    print(f'Documento salvo como {nome_arquivo}')


def e_mail():
    # Criando a operação mandar e-mail
    outlook = win32.Dispatch('Outlook.Application')

    # criar um objeto de e-mail
    email = outlook.CreateItem(0)
    # Configurar o e-mail
    email.Subject = 'PROPOSTA COMERCIAL[DESENVOLVIMENTO PYTHON]'
    email.Body =  '''Prezado,

    segue em anexo, proposta referente ao projeto desenvolvimento Python

    att, 
    Pedro Matos'''
# Criando variáveis
    email.To = entry_6.get()
    anexo = entry_7.get()
    email.Attachments.Add(anexo)
    email.Send()
 
# botão de gerar pdf
button_1 =tk.Button(janela, text='Gerar PDF', command= gerador, bg ='#f5f5f5')
button_1.grid(row=7, column=0, padx=15, pady=10, sticky='w')

button_2 =tk.Button(janela, text='Mandar por e-mail', command= e_mail, bg ='#f5f5f5')
button_2.grid(row=7, column=1, padx=50, pady=10, sticky='w')

janela.mainloop()