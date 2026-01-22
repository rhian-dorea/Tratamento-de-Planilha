# Tratamento-de-Planilha
Aplicativo em Python com Tkinter para automatizar o tratamento de planilhas da Claro. Lê, limpa e organiza dados em Excel usando pandas e openpyxl, gerando uma planilha final pronta. Empacotado em .exe com PyInstaller para uso direto no Windows. 
# pyinstaller --onefile --icon=app_icone.ico --name=ArranjoExcel --add-data="rodar.py;." --add-data="TratamentoDePlanilha.py;." --add-data="app_icone.ico;." --hidden-import=openpyxl --hidden-import=pandas Arranjo.py
