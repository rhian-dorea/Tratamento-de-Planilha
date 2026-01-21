from TratamentoDePlanilha import *

class RunProcessos():
    def processar_tratamento(self, pl):
        df = pl
        # === Lendo a planilha original do sistema ===
        df = pd.read_excel(
            df,
            header=2,
            dtype=str  # ← força tudo como texto → mais simples e seguro para planilhas da Claro
        )
        # Depois, limpe apenas o que precisa (opcional)
        #df["Nº SÉRIE"] = df["Nº SÉRIE"].str.strip()

        self.planilha = PlanilhaClaro(df)

        self.planilha.editar_colunas()
        self.planilha.aplicar_planos()
        self.planilha.criar_coluna_loja()
        self.planilha.remover_imei_ap() # remove imeis que foram lançados em lançamentos de seguros
        self.planilha.editar_debito()
        self.planilha.editar_nomes_vendedor()
        self.planilha.criar_colunas_faturas()
        self.planilha.ajustar_faturas()
        self.planilha.criar_coluna_ap() # conta imeis restantes
        self.planilha.ordenar_colunas()

    def salvar_tratamento(self):
       self.planilha.salvar("planilha_tratada.xlsx")


