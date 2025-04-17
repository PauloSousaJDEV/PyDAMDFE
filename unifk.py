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

    def _create_widgets(self):
        # --- Seção do Logo ---
        logo_frame = tk.Frame(self.root, bg=self.style.colors.light)
        logo_frame.pack(pady=(15, 10))

        try:
            logo_img = Image.open("Logo JTI.png")
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
        excel_button = Button(select_frame, text="📊 Selecionar Excel", command=self.selecionar_excel, bootstyle="info")
        excel_button.pack(side=tk.LEFT, padx=(0, 10), fill=tk.X, expand=True)
        self.label_excel = tk.Label(select_frame, text="Nenhum Excel selecionado", bg=self.style.colors.light, anchor='w')
        self.label_excel.pack(side=tk.LEFT, fill=tk.X, expand=True)

        # Botão e Label para XMLs
        xml_button = Button(select_frame, text="📄 Selecionar XMLs", command=self.selecionar_xmls, bootstyle="info")
        xml_button.pack(side=tk.LEFT, padx=(10, 0), fill=tk.X, expand=True)
        self.label_xml = tk.Label(select_frame, text="Nenhum XML selecionado", bg=self.style.colors.light, anchor='w')
        self.label_xml.pack(side=tk.LEFT, fill=tk.X, expand=True)

        # --- Botão de Execução ---
        execute_button = Button(self.root, text="🚀 Executar Automação", command=self.executar_automacao, bootstyle="success outline", padding=10)
        execute_button.pack(pady=20, padx=20, fill=tk.X)

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

            # mudar para Edg ou chome
            self.driver = webdriver.Edge()
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

                    
                    sleep(2)
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

                  
                    sleep(3.5)
                    botao_fechar_modal = wait.until(EC.element_to_be_clickable((By.XPATH, '//app-mdfe-close-modal-response//button[span[text()="Fechar"]]')))
                    driver.execute_script("arguments[0].click();", botao_fechar_modal)
                    self.log("✅ Modal de encerramento fechado.")
                    self.log(f"🏁 Processamento de Placa {placa} e Vendedor {vendedor} (Encerramento) finalizado.")
                    sleep(3.5)
                    
                except Exception as e:
                    self.log(f"❌ Erro ao processar placa {placa} e vendedor {vendedor} (Cancelamento): {e}", tipo="erro")
                    continue

            class automacaoEMTI: 
           # === NOVA ETAPA: EMISSÃO DA DAMDFE ===

                for placa, vendedor in zip(placas, Nome_do_Vendedor):
                    self.log(f"📄 Iniciando emissão da DAMDFE para {vendedor} - {placa}, Localidade M2: {localidade}")
                    try:
                        botao_encerrar = wait.until(EC.element_to_be_clickable((By.XPATH, "//button/span[contains(text(),'Encerrar')]")))
                        driver.execute_script("arguments[0].click();", botao_encerrar)
                        self.log("✅ Botão de Encerrar clicado.")


                      
                        
                        sleep(5) 
                    except Exception as e:
                        self.log(f"❌ Erro durante a emissão da DAMDFE para {vendedor} - {placa}: {e}", tipo="erro")

                   
                             
                self.log(f"Erro ao (Emitir) ")
        except Exception as e:
                self.new_method(e)
                self.log(f"🚨 Erro geral durante execução: {e}", tipo="erro")
        finally:   
                if self.driver:
                    self.driver.quit()
                    self.log("🛑 Driver do Selenium finalizado.")

    def new_method(self, e):
        messagebox.showerror("Erro Geral", f"Ocorreu um erro inesperado: {e}")

if __name__ == "__main__":
    root = tk.Tk()
    app = AutomacaoApp(root)
    root.mainloop() 