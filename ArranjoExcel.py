import tkinter as tk
from Arranjo import EditorPlanilhaApp

try:
    print("Iniciando interface gráfica...")
    root = tk.Tk()
    app = EditorPlanilhaApp(root)
    root.mainloop()
    print("Aplicação finalizada com sucesso")
except Exception as e:
    print(f"ERRO ao iniciar aplicação: {e}")
    import traceback
    traceback.print_exc()
    input("Pressione Enter para sair...")
