from TratamentoDePlanilha import *

class RunProcessos():
    def processar_tratamento(self, pl):
        df = pl
        # === Lendo a planilha original do sistema ===
        df = pd.read_excel(
            df,
            header=2,
            converters={
                "Nº SÉRIE": lambda x: str(x).strip(),
                "Telefone": str,
                "Nº Provisório": str,
                "CPF": str
            }
        )

        self.planilha = PlanilhaClaro(df)

        self.planilha.editar_colunas()
        self.planilha.aplicar_planos()
        self.planilha.criar_coluna_loja()
        self.planilha.criar_coluna_ap()
        self.planilha.editar_debito()
        self.planilha.editar_nomes_vendedor()
        self.planilha.criar_colunas_faturas()
        self.planilha.ajustar_faturas()
        self.planilha.remover_imei_ap()
        self.planilha.ordenar_colunas()

    def salvar_tratamento(self):
       self.planilha.salvar("planilha_tratada.xlsx")


