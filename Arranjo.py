from tkinter import  ttk
import threading
import time
import tkinter as tk
from tkinter import filedialog, messagebox
import os
import sys

# ---- com terminal ----
#python -m nuitka ArranjoExcel.py --standalone --onefile --enable-plugin=tk-inter --include-data-file=app_icone.ico=app_icone.ico --include-data-file=rodar.py=rodar.py --include-data-file=TratamentoDePlanilha.py=TratamentoDePlanilha.py --include-module=pandas --include-module=openpyxl --windows-icon-from-ico=app_icone.ico --output-dir=build_nuitka --remove-output
# ---- sem terminal ----
# python -m nuitka ArranjoExcel.py --standalone --onefile --enable-plugin=tk-inter --include-data-file=app_icone.ico=app_icone.ico --include-data-file=rodar.py=rodar.py --include-data-file=TratamentoDePlanilha.py=TratamentoDePlanilha.py --include-module=pandas --include-module=openpyxl --windows-icon-from-ico=app_icone.ico --windows-disable-console --output-dir=build_nuitka --remove-output
#-- 2
# python -m nuitka ArranjoExcel.py --standalone --onefile --enable-plugin=tk-inter --include-data-file=app_icone.ico=app_icone.ico --include-data-file=rodar.py=rodar.py --include-data-file=TratamentoDePlanilha.py=TratamentoDePlanilha.py --include-module=pandas --include-module=openpyxl --windows-icon-from-ico=app_icone.ico --windows-console-mode=disable --output-dir=build_nuitka --remove-output
# Configuração para PyInstaller
def resource_path(relative_path):
    """Get absolute path to resource, works for dev and for PyInstaller"""
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)


# Teste inicial
print("Iniciando aplicação...")

try:
    # Verifica se os arquivos estão acessíveis
    files_to_check = ['TratamentoDePlanilha.py', 'app_icone.ico']
    for file in files_to_check:
        if os.path.exists(resource_path(file)):
            print(f"✓ {file} encontrado")
        else:
            print(f"✗ {file} NÃO encontrado")

except Exception as e:
    print(f"Erro ao verificar arquivos: {e}")

# Importe outros módulos usando resource_path se necessário
try:
    # Adiciona o path do PyInstaller ao sys.path
    sys.path.insert(0, resource_path('.'))

    # Importa os módulos
    from TratamentoDePlanilha import *
    from rodar import *

    print("✓ Todos os módulos importados com sucesso")
except Exception as e:
    print(f"✗ Erro ao importar módulos: {e}")
    import traceback

    traceback.print_exc()


class EditorPlanilhaApp:
    def __init__(self, root):
        self.root = root
        self.buscando_icon()
        self.root.title("Editor de Planilha (cc)")
        self.df = pd.DataFrame()
        self.root.geometry("400x200")
        self.root.resizable(False, False)
        self.root.configure(bg='#217346')
        self.processar = RunProcessos()

        # Frame dos botões
        frame_central = tk.Frame(self.root, bg='#217346')
        frame_central.pack(pady=20, padx=15, fill=tk.X)

        self.btn_carregar = tk.Button(
            frame_central, text="Selecionar Planilha", command=self.selecionar_e_processar,
            bg="#1E90FF", fg="white", activebackground="darkblue", activeforeground="yellow",
            height=1, width=16, font=("Arial", 14)
        )
        self.btn_carregar.pack(side=tk.LEFT, padx=5)

        self.btn_salvar = tk.Button(
            frame_central, text="Salvar Planilha", command=self.salvar_planilha,
            bg="#32CD32", fg="white", activebackground="darkblue", activeforeground="yellow",
            height=1, width=15, font=("Arial", 14)
        )
        self.btn_salvar.pack(side=tk.LEFT, padx=5)

        # Frame da barra de progresso
        frame_barra = tk.Frame(self.root, bg='#217346')
        frame_barra.pack(pady=20, padx=15, fill=tk.X)

        # Label "Processando..."
        self.label_status = tk.Label(frame_barra, text="", bg='#217346', fg="white", font=("Arial", 12, "bold"))
        self.label_status.pack()

        # Barra de progresso
        self.barra = ttk.Progressbar(frame_barra, orient="horizontal", length=400, mode="determinate")
        self.barra.pack(pady=10)
# pyinstaller --onefile --icon=app_icone.ico --name=ArranjoExcel --add-data="rodar.py;." --add-data="TratamentoDePlanilha.py;." --add-data="app_icone.ico;." --hidden-import=openpyxl --hidden-import=pandas Arranjo.py

    def buscando_icon(self):
        import sys
        # ... (seu resource_path)

        icon_ico = resource_path("app_icone.ico")
        icon_png = resource_path("app_icone.png")

        if sys.platform.startswith("win"):
            if os.path.exists(icon_ico):
                self.root.iconbitmap(icon_ico)
                # seu código de app ID...
        else:  # Linux / macOS
            if os.path.exists(icon_png):
                try:
                    icone = tk.PhotoImage(file=icon_png)
                    self.root.iconphoto(True, icone)
                    print("Ícone PNG carregado com sucesso!")
                except Exception as e:
                    print(f"Erro no PNG: {e}")

    def selecionar_e_processar(self):
        arquivo = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx *.xls")])
        if not arquivo:
            return
        self.arquivo_selecionado = arquivo
        self.df = pd.read_excel(arquivo)

        # Mostra o texto “Processando...”
        self.label_status.config(text="Processando...")

        # Rodar em thread para não travar a interface
        threading.Thread(target=self.processar_planilha).start()

    def processar_planilha(self):
        total = len(self.df)
        self.barra["maximum"] = total
        self.barra["value"] = 0

        for i, linha in self.df.iterrows():
            time.sleep(0.05)  # simula processamento
            self.barra["value"] += 1
            self.root.update_idletasks()

        # Chama seu processamento real
        self.processar.processar_tratamento(self.arquivo_selecionado)

        # Quando terminar
        self.label_status.config(text="Processamento concluído!")

    def salvar_planilha(self):
        self.processar.salvar_tratamento()
        messagebox.showinfo("Concluído", "Planilha salva com sucesso!")
