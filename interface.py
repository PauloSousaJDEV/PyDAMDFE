from time import sleep
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter.scrolledtext import ScrolledText
from ttkbootstrap import Style
from ttkbootstrap.widgets import Button
from PIL import Image, ImageTk
import os
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
from selenium.webdriver.edge.service import Service as EdgeService
from selenium.webdriver.chrome.service import Service as ChromeService
from webdriver_manager.chrome import ChromeDriverManager
from webdriver_manager.microsoft import EdgeChromiumDriverManager
from webdriver import Navegador
from excel import Tabela_de_Dados
from damdfe import Damdfe

atualDamdfe = Damdfe()
atualExcel = Tabela_de_Dados()

class AutomacaoApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Emissão de DAMDFE - JTI")
        self.root.geometry("950x720")  # Aumentei um pouco o tamanho inicial
        self.root.resizable(False, False)

        self.style = Style("minty")
        self.excel_path = tk.StringVar()
        self.xml_paths = []
        self.logo = None

        self._create_widgets()
        self.driver = None

    def executar_automacao(self):
        try:
            if not self.excel_path.get():
                messagebox.showerror("Erro", "Selecione um arquivo Excel válido.")
                return

            if not self.xml_paths:
                messagebox.showerror("Erro", "Selecione os arquivos XML.")
                return

            excel_handler = Tabela_de_Dados(self.excel_path.get())
            excel_handler.coletar_dados(log_callback=self.log)

            # Aqui você pode seguir com a automação via Selenium, usando:
            # excel_handler.placas, excel_handler.Nome_do_Vendedor, excel_handler.localidade

        except Exception as e:
            self.log(f"Erro inesperado: {e}", tipo="erro")
            messagebox.showerror("Erro", str(e))


    def _create_widgets(self):
        # --- Seção do Logo ---
        logo_frame = tk.Frame(self.root, bg=self.style.colors.light)
        logo_frame.pack(pady=(15, 10))
        
        try:
            logo_img = Image.open("./Icone/Logo JTI.png")
            logo_img = logo_img.resize((180, 70), Image.LANCZOS)  # Reduzi um pouco o tamanho
            self.logo = ImageTk.PhotoImage(logo_img)
            logo_label = tk.Label(logo_frame, image=self.logo, bg=self.style.colors.light)
            logo_label.pack()
        except FileNotFoundError:
            messagebox.showerror("Erro", "Arquivo 'Logo JTI.png' não encontrado. Continuando sem logo.")

        # --- Seção de Seleção de Arquivos ---
        select_frame = tk.Frame(self.root, bg=self.style.colors.light, padx=20, pady=15)
        select_frame.pack(fill=tk.X)

        # Botão e Label para Excel
        excel_button = Button(select_frame, text="📊 Selecionar Excel", command=atualExcel.selecionar_excel, bootstyle="info")
        excel_button.pack(side=tk.LEFT, padx=(0, 10), fill=tk.X, expand=True)
        self.label_excel = tk.Label(select_frame, text="Nenhum Excel selecionado", bg=self.style.colors.light, anchor='w')
        self.label_excel.pack(side=tk.LEFT, fill=tk.X, expand=True)

        # Botão e Label para XMLs
        xml_button = Button(select_frame, text="📄 Selecionar XMLs", command=self.selecionar_xmls, bootstyle="info")
        xml_button.pack(side=tk.LEFT, padx=(10, 0), fill=tk.X, expand=True)
        self.label_xml = tk.Label(select_frame, text="Nenhum XML selecionado", bg=self.style.colors.light, anchor='w')
        self.label_xml.pack(side=tk.LEFT, fill=tk.X, expand=True)

        # --- Botão de Execução ---
        execute_button = Button(self.root, text="🚀 Executar Automação", command= atualDamdfe.cancelarDamdfe, bootstyle="success outline", padding=10)
        execute_button.pack(pady=20, padx=20, fill=tk.X)

        # --- Botão de CANCELAR ---
        execute_button1 = Button(self.root, text="🚀 CANCELAR", command= atualDamdfe.cancelarDamdfe, bootstyle="success outline", padding=10)
        execute_button1.pack(pady=20, padx=20, fill=tk.X)
        
        # --- Seção de Monitoramento ---
        log_frame = tk.LabelFrame(self.root, text="Monitoramento", labelanchor='n', font=('Arial', 12, 'bold'), padx=20, pady=10, bg=self.style.colors.light)
        log_frame.pack(pady=(10, 20), padx=20, fill=tk.BOTH, expand=True)
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(1, weight=1)
        log_frame.rowconfigure(3, weight=1)

        # Progresso
        tk.Label(log_frame, text="Progresso:", font=('Arial', 10, 'bold'), bg=self.style.colors.light).grid(row=0, column=0, sticky='w', pady=(0, 5))
        self.log_progresso = ScrolledText(log_frame, font=("Consolas", 9), bg=self.style.colors.secondary, fg=self.style.colors.light, wrap=tk.WORD)
        self.log_progresso.grid(row=1, column=0, padx=10, pady=5, sticky='nsew')

        # Erros
        tk.Label(log_frame, text="Erros:", font=('Arial', 10, 'bold'), bg=self.style.colors.light).grid(row=2, column=0, sticky='w', pady=(10, 5))
        self.log_erros = ScrolledText(log_frame, font=("Consolas", 9), foreground=self.style.colors.danger, bg=self.style.colors.secondary, fg=self.style.colors.danger, wrap=tk.WORD)
        self.log_erros.grid(row=3, column=0, padx=10, pady=5, sticky='nsew')

    def selecionar_excel(self):
        path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsm *.xlsx")])
        if path:
            self.excel_path.set(path)
            self.label_excel.config(text=f"Excel: {os.path.basename(path)}")

    def selecionar_xmls(self):
        paths = filedialog.askopenfilenames(filetypes=[("Arquivos XML", "*.xml")])
        if paths:
            if len(paths) > 260:
                messagebox.showwarning("Limite excedido", "Selecione no máximo 260 arquivos XML.")
                return
            self.xml_paths = list(paths)
            self.label_xml.config(text=f"{len(self.xml_paths)} XML(s) selecionado(s)")

    def log(self, mensagem, tipo="progresso"):
        if tipo == "progresso":
            self.log_progresso.insert(tk.END, mensagem + "\n")
            self.log_progresso.see(tk.END)
        elif tipo == "erro":
            self.log_erros.insert(tk.END, mensagem + "\n")
            self.log_erros.see(tk.END)
        self.root.update()

    
        