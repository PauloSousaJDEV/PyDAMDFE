# BKP unifk.py Funcionando!!!

from asyncio import sleep
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


class AutomacaoApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Emissão de DAMDFE - JTI")
        self.root.geometry("900x700")
        self.root.resizable(False, False)

        self.style = Style("minty")
        self.excel_path = tk.StringVar()
        self.xml_paths = []

        try:
            logo_img = Image.open("Logo JTI.png")
            logo_img = logo_img.resize((200, 80), Image.LANCZOS)
            self.logo = ImageTk.PhotoImage(logo_img)
            logo_label = tk.Label(self.root, image=self.logo, bg='white')
            logo_label.pack(pady=10)
        except FileNotFoundError:
            messagebox.showerror("Erro", "Arquivo 'Logo JTI.png' não encontrado. Continuando sem logo.")

        self.frame = tk.Frame(self.root, bg='white')
        self.frame.pack(pady=10)

        Button(self.frame, text="📊 Selecionar Excel", command=self.selecionar_excel, bootstyle="info").grid(row=0, column=0, padx=10, pady=5)
        self.label_excel = tk.Label(self.frame, text="Nenhum Excel selecionado", bg='white')
        self.label_excel.grid(row=1, column=0, columnspan=2)

        Button(self.frame, text="📄 Selecionar XMLs", command=self.selecionar_xmls, bootstyle="info").grid(row=0, column=1, padx=10, pady=5)
        self.label_xml = tk.Label(self.frame, text="Nenhum XML selecionado", bg='white')
        self.label_xml.grid(row=1, column=1, columnspan=2)

        Button(self.root, text="🚀 Executar Automação", command=self.executar_automacao, bootstyle="success outline").pack(pady=20)

        log_frame = tk.LabelFrame(self.root, text="Monitoramento", labelanchor='n', font=('Arial', 12, 'bold'))
        log_frame.pack(pady=10, padx=20, fill=tk.BOTH, expand=True)

        tk.Label(log_frame, text="Progresso:", font=('Arial', 10, 'bold')).grid(row=0, column=0, sticky='w')
        self.log_progresso = ScrolledText(log_frame, width=100, height=10, font=("Consolas", 9))
        self.log_progresso.grid(row=1, column=0, padx=10, pady=5)

        tk.Label(log_frame, text="Erros:", font=('Arial', 10, 'bold')).grid(row=2, column=0, sticky='w')
        self.log_erros = ScrolledText(log_frame, width=100, height=8, font=("Consolas", 9), foreground="red")
        self.log_erros.grid(row=3, column=0, padx=10, pady=5)

        self.driver = None

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

    def executar_automacao(self):
        caminho_arquivo = self.excel_path.get()
        if not caminho_arquivo:
            messagebox.showerror("Erro", "Por favor, selecione o arquivo Excel Valido.")
            return

        if not self.xml_paths:
            messagebox.showerror("Erro", "Por favor, selecione os arquivos XML.")
            return

        nome_aba = "Exportar Nota Fiscal"
        placas = []
        Nome_do_Vendedor = []
        localidade = None
       

        try:
            planilha = pd.ExcelFile(caminho_arquivo)
            for header in [None, 0, 1, 2, 3, 4, 5]:
                try:
                    dados = planilha.parse(sheet_name=nome_aba, header=header)
                    if 'PLACA' in dados.columns and 'Nome do Vendedor' in dados.columns:
                        placas = dados['PLACA'].dropna().astype(str).str.strip().tolist()
                        Nome_do_Vendedor = dados['Nome do Vendedor'].dropna().astype(str).str.strip().tolist()
                        break
                except Exception:
                    continue

            if not placas or not Nome_do_Vendedor:
                raise ValueError("Colunas 'PLACA' e/ou 'Nome do Vendedor' não encontradas ou vazias.")

            if len(placas) != len(Nome_do_Vendedor):
                raise ValueError("Número de placas e vendedores não correspondem.")

            # Lendo a célula M2 diretamente com tratamento de erros detalhado
            localidade = "ERRO AO LER LOCALIDADE"  # Valor padrão em caso de erro
            try:
                df_leitura = planilha.parse(sheet_name=nome_aba, header=None) # Lendo sem cabeçalho
                num_rows, num_cols = df_leitura.shape
                self.log(f"ℹ️ Dimensões da aba '{nome_aba}': {num_rows} linhas, {num_cols} colunas")

                linha_desejada = 1   # Segunda linha (índice 1)
                coluna_desejada = 12 # Décima terceira coluna (índice 12) - Coluna M

                if num_rows > linha_desejada and num_cols > coluna_desejada:
                    valor_localidade = df_leitura.iloc[linha_desejada, coluna_desejada]
                    self.log(f"ℹ️ Valor lido da célula M2 (cru): '{valor_localidade}'")
                    if isinstance(valor_localidade, str):
                        valor_localidade = valor_localidade.strip().upper()
                        if valor_localidade == "BRTE":
                            localidade = "PORTO ALEGRE"
                        elif valor_localidade == "BRTG":
                            localidade = "DUQUE DE CAXIAS"
                        else:
                            localidade = "OUTRA LOCALIDADE"
                    else:
                        localidade = "VALOR DE LOCALIDADE INVÁLIDO"
                else:
                    localidade = f"ERRO: A aba '{nome_aba}' não possui pelo menos {linha_desejada + 1} linhas e {coluna_desejada + 1} colunas."
            except Exception as e:
                localidade = f"ERRO INESPERADO AO LER LOCALIDADE: {e}"

            self.log(f"Localidade identificada: {localidade}")

            self.log("🚀 Iniciando automação com Selenium...")

            self.driver = webdriver.Chrome()
            driver = self.driver
            wait = WebDriverWait(driver, 20)
            sleep = time.sleep

            driver.get('https://mdfe-beta.hivecloud.com.br/')

            wait.until(EC.presence_of_element_located((By.XPATH, '//lib-form-control[1]//input'))).send_keys('Omar.Teixeira@jti.com')
            driver.find_element(By.XPATH, '//lib-form-control[2]//input').send_keys('17318208')
            driver.find_element(By.XPATH, '//div[2]/lib-button/button/span').click()
            

            if localidade == "PORTO ALEGRE":
                self.log(f"ℹ️ Executando comandos específicos para PORTO ALEGRE (Inicial).")
                sleep(2)
                try:
                    
                    self.log("🔍 Selecionando o Ambiente")
                    wait.until(EC.element_to_be_clickable((By.XPATH, '//lib-await-panel/div/div/div/div[2]/button/span'))).click()
                    wait.until(EC.element_to_be_clickable((By.XPATH, '//lib-company-selection/lib-await-panel/div/div/div/lib-company-selection-card[15]/div'))).click()
                    wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="menuLateral"]/div[2]/lib-sidenav-menu-item[2]/a'))).click()
                    if not search_input.is_displayed() or not search_input.is_enabled():
                        self.log("❌ Campo de pesquisa não está visível ou habilitado!")
                    else:
                        wait.until(EC.element_to_be_clickable((By.XPATH, '//lib-await-panel/div/div/div/div[2]/button/span'))).click()
                        wait.until(EC.element_to_be_clickable((By.XPATH, '//lib-company-selection/lib-await-panel/div/div/div/lib-company-selection-card[15]/div'))).click()
                        wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="menuLateral"]/div[2]/lib-sidenav-menu-item[2]/a'))).click()


                except Exception as e:
                    self.log(f"❌ Erro ao tentar preencher o campo de pesquisa: {e}")

            elif localidade == "DUQUE DE CAXIAS":
                self.log(f"ℹ️ Executando comandos específicos para DUQUE DE CAXIAS (Inicial).")
                # Navegar para a aba de Emissão
                sleep(2)
                                                                      
                try:                                                 
                    wait.until(EC.element_to_be_clickable((By.XPATH, '//lib-await-panel/div/div/div/div[2]/button/span'))).click()
                    wait.until(EC.element_to_be_clickable((By.XPATH, '//lib-company-selection-card[14]/div'))).click()
                    wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="menuLateral"]/div[2]/lib-sidenav-menu-item[2]/a'))).click()
                                                                      

                except Exception as e:
                    self.log(f"❌ Erro ao interagir com elementos em Duque de Caxias (Inicial): {e}", tipo="erro")

            else:
                self.log(f"ℹ️ Nenhuma ação específica definida para a localidade (Inicial): {localidade}")

            # Inicio da automação - Cancelar DAMDFE/

            for placa, vendedor in zip(placas, Nome_do_Vendedor):
                self.log(f"🔍 Processando (Cancelamento): Placa {placa}, Nome_do_Vendedor {vendedor}, Localidade M2: {localidade}")
                sleep(3)

                try:
                    search_input = wait.until(EC.presence_of_element_located((By.XPATH, "//input[@placeholder='Pesquisar MDFe']")))
                    search_input.clear()
                    search_input.send_keys(placa)
                    sleep(1)
                    search_input.send_keys(Keys.ENTER)

                    wait.until(EC.presence_of_element_located((By.XPATH, "//table//tr[1]")))

                    checkbox = wait.until(EC.element_to_be_clickable((By.XPATH, "//table/tbody/tr[1]//p-checkbox//div[@class='p-checkbox-box']")))
                    driver.execute_script("arguments[0].click();", checkbox)
                    self.log("✅ Checkbox marcado.")

                    sleep(2)
                    botao_encerrar = wait.until(EC.element_to_be_clickable((By.XPATH, "//button/span[contains(text(),'Encerrar')]")))
                    driver.execute_script("arguments[0].click();", botao_encerrar)
                    self.log("✅ Botão de Encerrar clicado.")

                    sleep(2)
                    confirmar_modal = wait.until(EC.element_to_be_clickable((By.XPATH, '//div[contains(@class,"cdk-overlay-pane")]//lib-button[2]/button')))
                    driver.execute_script("arguments[0].click();", confirmar_modal)
                    self.log("✅ Encerramento confirmado.")

                    sleep(2)
                    botao_fechar_modal = wait.until(EC.element_to_be_clickable((By.XPATH, '//app-mdfe-close-modal-response//button[span[text()="Fechar"]]')))
                    driver.execute_script("arguments[0].click();", botao_fechar_modal)
                    self.log("✅ Modal de encerramento fechado.")
                    self.log(f"🏁 Processamento de Placa {placa} e Vendedor {vendedor} (Cancelamento) finalizado.")
                    sleep(3)

                except Exception as e:
                    self.log(f"❌ Erro ao processar placa {placa} e vendedor {vendedor} (Cancelamento): {e}", tipo="erro")
                    continue

            # === NOVA ETAPA: EMISSÃO DA DAMDFE ===
            for placa, vendedor in zip(placas, Nome_do_Vendedor):
                self.log(f"📄 Iniciando emissão da DAMDFE para {vendedor} - {placa}, Localidade M2: {localidade}")

                # A navegação para a aba de emissão já ocorreu no bloco condicional inicial.

                if localidade == "PORTO ALEGRE":
                    self.log(f"ℹ️ Ação específica para Porto Alegre (Emissão).")

                    #buttun- MDFe
                    wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="menuLateral"]/div[2]/lib-sidenav-menu-item[2]/a'))).click()

                    


                    
                    pass
                elif localidade == "DUQUE DE CAXIAS":
                    self.log(f"ℹ️ Ação específica para Duque de Caxias (Emissão).")
                    # Adicione aqui o código específico para a EMISSÃO em Duque de Caxias
                    # Exemplo: Upload de XMLs, preenchimento de campos, etc.
                    # Você precisará identificar os XPaths dos elementos na página de emissão.
                    pass
                else:
                    self.log(f"ℹ️ Nenhuma ação específica para a localidade (Emissão): {localidade}")

                # Restante do seu código de emissão da DAMDFE (upload de XMLs, etc.)
                # ...

        except Exception as e:
            self.log(f"🚨 Erro geral durante execução: {e}", tipo="erro")

if __name__ == "__main__":
    root = tk.Tk()
    app = AutomacaoApp(root)
    root.mainloop()