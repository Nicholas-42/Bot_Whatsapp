# Importar bibliotecas necessárias
import tkinter as tk
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from ttkbootstrap.style import Style    
from urllib.parse import quote
from time import sleep
from datetime import datetime
import pandas as pd
import webbrowser
import threading


class InterfaceGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Disparador de Mensagem!")
        self.root.geometry("600x500")


        # TÍTULO
        app = self.root
        label = ttk.Label(app, text='Responda para começarmos!')
        label.pack(pady=35)
        label.config(font=('Arial', 20, 'bold'))
        Style(theme='superhero')

        # Campos de entrada
        self.pl_label = ttk.Label(root, text="Planilha:")
        self.pl_label.pack()
        self.pl_entry = ttk.Entry(root, width=30)
        self.pl_entry.pack()

        self.curso_label = ttk.Label(root, text="Curso:")
        self.curso_label.pack()
        self.curso_entry = ttk.Entry(root, width=30)
        self.curso_entry.pack()

        self.mensagem = ttk.Label(root, text="Mensagem a ser enviada:")
        self.mensagem.pack()
        self.mensagem_entry = ttk.Entry(root, width=30)
        self.mensagem_entry.pack()

        self.inicial_label = ttk.Label(root, text="Primeira linha a ser enviada:")
        self.inicial_label.pack()
        self.inicial_entry = ttk.Entry(root, width=30)
        self.inicial_entry.pack()

        self.max_label = ttk.Label(root, text="Quantidade de mensagens a serem disparadas:")
        self.max_label.pack()
        self.max_entry = ttk.Entry(root, width=30)
        self.max_entry.pack()


        # Criar um botão para enviar mensagens
        self.send_button = ttk.Button(root, text="Enviar Mensagens", command=self.start_sending)
        self.send_button.pack(pady=10)

        # Criar um botão para cancelar o código
        self.cancel_button = ttk.Button(root, text="Interromper código", bootstyle=DANGER, command=self.stop_sending)
        self.cancel_button.pack(pady=10)

        self.running = False

    def stop_sending(self):
        self.running = False
        print("Código interrompido.")


    def start_sending(self):
        if self.running:
            print("O envio de mensagens já está em andamento.")
            return
        self.running = True
        t = threading.Thread(target=self.send_messages)
        t.start()


    def send_messages(self):
        # Obter entrada do usuário
        input_pl = self.pl_entry.get().lower()
        input_curso = self.curso_entry.get()
        input_mensagem = self.mensagem_entry.get()
        input_inicial = int(self.inicial_entry.get())
        input_max = int(self.max_entry.get())
        

        # Ler o arquivo Excel

        alunos = pd.read_excel(f'alunos({input_pl}).xlsx')
        max_linhas = len(alunos)

        # Inicia o envio de mensagens

        self.running = True

        # Inicia variáveis de parada

        mensagem_enviada = 0
        linha_atual = 0
        achou = False

        if input_inicial >= max_linhas:
            return print(f"ERROR! A planilha tem apenas {max_linhas} linhas.")

        for x in range(input_inicial, max_linhas):
           
            if input_inicial <= 1:
                input_inicial = 1

            linha_atual = x + 1

            if not (self.running) or (mensagem_enviada >= input_max):

                self.running = False
                return print(f"Última linha a ser enviada foi: {linha_atual}")

            cursos = alunos.loc[x, "Dentre as opções qual curso gostaria de fazer?"]
        
            lista_cursos = str(cursos).split(sep=', ')
            
            #Olha cada item da lista criada, verifica se o curso de envio está dentro da lista, caso esteja, a planilha pega o nome e telefone do aluno, e faz o envio da mensagem.
            
            for curso in lista_cursos:
                if input_curso.upper() in curso.upper():
                    achou = True
                    # Ler planilha e guardar informações nome e telefone
                    nome = alunos.loc[x, 'Nome Completo']
                    telefone = alunos.loc[x, 'Whatsapp com DDD (somente números - sem espaço)']

                    mensagem = f"Olá {nome}. {input_mensagem}"
                    

                    # Criar links personalizados do whatsapp e enviar mensagens para cada cliente
                    link_mensagem_whatsapp = f'https://api.whatsapp.com/send?phone={telefone}&text={quote(mensagem)}'
                    webbrowser.open(link_mensagem_whatsapp)

                    mensagem_enviada += 1        
            
            if linha_atual == max_linhas:
                if achou is False:
                    return print(f"Não temos o curso '{input_curso}' disponível.")
                else:
                    print(f"Última linha a ser enviada foi: {linha_atual}")
                    return print(f"Chegamos ao fim da planilha!")

# Criar a GUI
root = tk.Tk()
gui = InterfaceGUI(root)
root.mainloop()
