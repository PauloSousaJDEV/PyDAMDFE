from common_imports import *

class Tabela_de_Dados:
    
    def __init__(self, excel_path, nome_aba="Exportar Nota Fiscal"):
        self.excel_path       = excel_path
        self.nome_aba         = nome_aba
        self.placas           = []
        self.Nome_do_Vendedor = []
        self.localidade       = None

    def selecionar_excel(self):
        path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsm *.xlsx")])
        if path:
            AutomacaoApp.excel_path.set(path)
            AutomacaoApp.label_excel.config(text=f"Excel: {os.path.basename(path)}")  

       
    def coletar_dados(self, log_callback=None):


        
       
        try:
            planilha = pd.ExcelFile(self.excel_path)

            # 1) Ler colunas PLACA e Nome do Vendedor
            for header in [None, 0, 1, 2, 3, 4, 5]:
                try:
                    df = planilha.parse(sheet_name=self.nome_aba, header=header)
                    if 'PLACA' in df.columns and 'Nome do Vendedor' in df.columns:
                        self.placas = (
                            df['PLACA']
                            .dropna()
                            .astype(str)
                            .str.strip()
                            .tolist()
                        )
                        self.Nome_do_Vendedor = (
                            df['Nome do Vendedor']
                            .dropna()
                            .astype(str)
                            .str.strip()
                            .tolist()
                        )
                       
                        break
                except Exception:
                    continue

            if not self.placas or not self.Nome_do_Vendedor:
                raise ValueError("Colunas 'PLACA' e/ou 'Nome do Vendedor' não encontradas ou vazias.")
            if len(self.placas) != len(self.Nome_do_Vendedor):
                raise ValueError("Número de placas e vendedores não correspondem.")

            # 2) Ler valor de N2 (linha 1, coluna 13)
            df_leitura = planilha.parse(sheet_name=self.nome_aba, header=None)
            valor_localidade = df_leitura.iat[1, 13]
            if not isinstance(valor_localidade, str):
                raise ValueError(f"Valor de localidade inválido: {valor_localidade!r}")

            self.localidade = valor_localidade.strip()

            # 3) Logs de sucesso
            if log_callback:
                log_callback(f"Localidade identificada: {self.localidade}")
                log_callback("🚀 Coleta do Excel concluída com sucesso.")

        except Exception as e:
            if log_callback:
                log_callback(f"🚨 Erro ao processar Excel: {e}", tipo="erro")
            raise
