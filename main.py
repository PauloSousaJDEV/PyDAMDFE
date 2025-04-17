from common_imports import *

def main():
    root = tk.Tk()
    app = AutomacaoApp(root)
    excel_instancia = Tabela_de_Dados(root)  # <-- variável renomeada
    root.mainloop()

if __name__ == "__main__":
    main()
