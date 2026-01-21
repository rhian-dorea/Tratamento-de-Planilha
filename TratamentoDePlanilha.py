import openpyxl
import pandas as pd
import re
import os
from pathlib import Path
import locale

# Configuração global do locale
LOCALE_PTBR = locale.setlocale(locale.LC_TIME, 'pt_BR.UTF-8')


class PlanilhaClaro:
    """
    CLASS QUE FAZ EDIÇÃO DE PLANILHA, REMOVENDO COLUNAS, RENOMEANDO,
    CRIANDO COLUNAS DE ACORDO COM VENDAS,PLANOS E VENDEDORES
    PARA NO FINAL DEIXAR NO PADRÃO REQUISITADO
    """
    def __init__(self, df=None, caminho_arquivo=None, header=2):
        if df is not None:
            self.df = df.copy()  # recebe um DataFrame já carregado
        elif caminho_arquivo is not None:
            self.df = pd.read_excel(caminho_arquivo, header=header)
        else:
            raise ValueError("Você deve fornecer um DataFrame ou o caminho do arquivo")
        # Padroniza colunas
        self.df.columns = self.df.columns.str.strip().str.upper()

    def editar_colunas(self):
        """
        REMOVE TODAS COLUNAS DESNECESSÁRIAS E RENOMEIA COLUNAS RESTANTES DE ACORDO COM PADRÃO
        :return:
        """
        colunas_remover = [
            "HORA", "ATENDIMENTO", "NF", "PROTOCOLO", "CAIXA",
            "NASCIMENTO", "TEL CONTATO", "OS", "NRC", "CONTRATO",
            "TEL RESIDENCIAL", "LANÇADO POR", "CIDADE", "BAIRRO", "COMBO",
            "LOCAL", "CARTÃO", "PARC", "VALOR PLANO ANTIGO",
            "E-MAIL", "CANAL DE VENDAS", "SEGURO", "OPERADORA", "DESCRIÇÃO"
        ]
        self.df = self.df.drop(columns=colunas_remover, errors="ignore")
        self.df.columns = self.df.columns.str.strip().str.upper()
        renomear = {
            "Nº SÉRIE": "ICCID/IMEI",
            "PROTOCOLO GED": "GED",
            "Nº PROVISÓRIO": "PROVISORIO",
            "1º VENCIMENTO": "FATURA 1",
            "PAGAMENTO": "DEBITO AUTOMATICO",
            "VENDEDOR": "CONSULTOR",
            "TELEFONE": "NUMERO"
        }
        renomear = {k.upper(): v for k, v in renomear.items()}
        self.df = self.df.rename(columns=renomear)
        self.df.columns = self.df.columns.str.strip().str.upper()
        # Garante que ICCID/IMEI seja sempre string
        if 'ICCID/IMEI' in self.df.columns:
            self.df['ICCID/IMEI'] = self.df['ICCID/IMEI'].fillna('').astype(str)

    def definir_plano_final(self, row):
        """
        VERIFICA CADA VALOR DAS COLUNAS ATIVAÇÃO, PLANO E PROVISSÓRIO RESULTANDO EM UM NOME
        DO PLANO DE ACORDO COM CADA VALOR DIFERENTE NAS COLUNAS DE VERIFICAÇÃO
        :param row:
        :return:
        """
        ativ = str(row.get('ATIVAÇÃO', '')).upper()
        plano = str(row.get('PLANO', '')).upper()
        prov = str(row.get('PROVISORIO', '')).strip()

        tipo = ""
        gb = ""
        extra = ""

        if "CONTROLE" in plano:
            tipo = "CONTROLE"
        elif "INTERNET" in plano:
            tipo = "INTERNET"
        elif "PÓS" in plano or "POS" in plano:
            tipo = "PÓS"
        elif "DEPENDENTE+ DADOS E VOZ" in plano:
            tipo = "DEPENDENTE TOTAL"
        elif "DEPENDENTE+ DADOS" in plano:
            tipo = "DEPENDENTE INTERNET"
        elif "CLARO CARTAO" in plano:
            tipo = "PRE PAGO"
        elif "CLARO PÓS ON 25GB COMBO CONVERGENTE" in plano:
            tipo = "PÓS 25GB ON/CONVERGÊNCIA"
        elif re.search(r"\bfácil\b", plano, re.IGNORECASE):
            tipo = "CONTROLE FÁCIL"
        else:
            tipo = "OUTRO"

        match_gb = re.search(r'(\d+\s*GB)', plano)
        if match_gb:
            gb = match_gb.group(1).replace(" ", "")

        match_extra = re.search(r'\b(ON|NOITES)\b', plano)
        if match_extra:
            extra = match_extra.group(1)

        resultado = ""
        ## ----------ESCREVENDO PLANOS----------
        if "NOVA ATIVAÇÃO" in ativ and tipo == "CONTROLE":
            resultado = f"{tipo} {gb} {extra}".strip()
        elif "MIGRAÇÃO PRÉ" in ativ and tipo == "CONTROLE":
            resultado = f"MIGRAÇÃO {gb} {extra}".strip()
        elif tipo == "CONTROLE FÁCIL":
            resultado = f"{tipo} {gb} {extra}".strip()
        elif "NOVA ATIVAÇÃO" in ativ and tipo == "INTERNET":
            resultado = f"INTERNET {gb} {extra}".strip()
        elif tipo == "DEPENDENTE TOTAL":
            resultado = "DEPENDENTE TOTAL"
        elif tipo == "DEPENDENTE INTERNET":
            resultado = "DEPENDENTE INTERNET"
        elif "NOVA ATIVAÇÃO" in ativ and tipo == "PÓS":
            resultado = f"PÓS {gb} {extra}".strip()
        elif "MIGRAÇÃO PRÉ" in ativ and tipo == "PÓS":
            resultado = f"MIGRAÇÃO PÓS {gb} {extra}".strip()
        elif "NOVA ATIVAÇÃO" in ativ and tipo == "PÓS 25GB ON/CONVERGÊNCIA":
            resultado = f"PÓS {gb} {extra}/CONVERGÊNCIA".strip()
        elif "MIGRAÇÃO PRÉ" in ativ and tipo == "PÓS 25GB ON/CONVERGÊNCIA":
            resultado = f"MIGRAÇÃO PÓS {gb} {extra}/CONVERGÊNCIA".strip()
        elif "TROCA DE SIM CARD" in ativ:
            resultado = "RESGATE"
        elif "NOVA ATIVAÇÃO" in ativ and tipo == "PRE PAGO":
            resultado = "PRÉ PAGO"
        elif "SEGURO PROTEÇÃO MÓVEL" in ativ:
            resultado = "SEGURO PROTEÇÃO MÓVEL"
        elif "TROCA DE APARELHO PÓS OU CONTROLE" in ativ:
            resultado = "TROCA APARELHO"
        elif "TROCA PLANO EQUIVALENTE" in ativ or "UPGRADE DE PLANO" in ativ:
            resultado = "TROCA PLANO"

        if any(char.isdigit() for char in prov) and resultado:
            resultado += "/PORT"

        return resultado

    def aplicar_planos(self):
        """
        CRIA COLUNA PLANOS E ADICIONA TODOS OS RESULTADOS DA FUNÇÃO definir_plano_final
        :return:
        """
        self.df["PLANOS"] = self.df.apply(self.definir_plano_final, axis=1)

    def criar_coluna_ap(self):
        """
        COLUNA APOIO (AP) PARA TODOS CLIENTES QUE COMPRARAM UM APARELHO
        Coloca 'AP' SOMENTE quando ICCID/IMEI tem menos de 19  caracteres (indicando aparelho)
        """
        # Garante string e limpa
        self.df['ICCID/IMEI'] = self.df['ICCID/IMEI'].fillna('').astype(str).str.strip()
        self.df['ATIVAÇÃO'] = self.df['ATIVAÇÃO'].fillna('').astype(
            str).str.strip().str.upper()  # ainda pode ser útil para outras lógicas

        # Cria a coluna AP de forma vetorizada (mais eficiente)
        self.df['AP'] = ''
        mask_aparelho = (self.df['ICCID/IMEI'].str.len() < 19) & (self.df['ICCID/IMEI'] != '')
        self.df.loc[mask_aparelho, 'AP'] = 'AP'

    def criar_coluna_loja(self):
        """
        RETORNA UMA COLUNA COM NOME DE CADA LOJA DE ACORDO COM COLUNA DE VENDEDORES
        :return:
        """
        vendedores = {
            "YESSA": "LOJA 05", "THYELLE": "LOJA 03", "TALLITA": "LOJA 01",
            "TACIANE": "LOJA 05", "PATRICIA": "LOJA 02", "MARLI": "LOJA 02",
            "MARGARIDA": "LOJA 02", "LUIZA": "LOJA 05", "LEONARDO": "LOJA 06",
            "KARINE": "IPIAU", "JULIANA": "LOJA 06", "JESSICA": "LOJA 05",
            "INGRID": "LOJA 03", "DANIELA": "IPIAU", "ALESSANDRA": "LOJA 02",
            "AILTON": "LOJA 03", "ITABUNA": "PAP", "ELCI": "PAP", "ILHEUS": "PAP",
            "JEQUIE": "PAP", "ALCYMAR": "LOJA 02", "ALANA": "LOJA 02", "RHIAN": "LOJA 01",
            "LUMA": "LOJA 01", "ANA": "LOJA 05", "ANYA": "LOJA 02", "CLICIA": "LOJA 03",
            "CLEITON": "LOJA 03", "LUANA": "LOJA 05", "ENDRIL": "PAP", "CONQUISTA": "PAP",
            "GEOVANA": 'LOJA 05', 'ROBERTA': 'LOJA 05', 'GISELE': 'LOJA 03'
        }
        self.df["LOJA"] = self.df["CONSULTOR"].str.split().str[0].str.upper().map(vendedores).fillna("")

    def editar_debito(self):
        """
        DEIXA SOMENTE QUEM ADICIONOU DÉBITO AUTOMATICO (OK) E RETIRA TODOS OS OUTROS VALORES
        :return:
        """
        self.df['DEBITO AUTOMATICO'] = self.df['DEBITO AUTOMATICO'].replace('Cartão de Débito', 'ok')
        self.df.loc[self.df['DEBITO AUTOMATICO'] != 'ok', 'DEBITO AUTOMATICO'] = ''

    def editar_nomes_vendedor(self):
        """
        DEIXA SOMENTE O PRIMEIRO NOME DOS VENDEDORES E FORMATA COLUNAS
        CPF, NUMERO, PROVISORIO, E ICCID/IMEI
        :return:
        """
        self.df['CONSULTOR'] = self.df['CONSULTOR'].str.split().str[0]
        # cols = ["CPF", "NUMERO", "PROVISORIO", "ICCID/IMEI"]
        #
        # for col in cols:
        #     self.df[col] = self.df[col].astype(str).str.replace(r"[.\-/\s]+", "", regex=True) # opcional: remove espaços

        self.df["CPF"] = self.df["CPF"].astype(str).str.replace(r"[.\-\/]", "", regex=True)

    def criar_colunas_faturas(self):
        """
        ADICIONA +2 COLUNAS DE FATURAS E FORMATADA COLUNAS DE DATA DA PLANILHA E
        DEIXA A ATIVAÇÃO QUANDO NÃO TEM DATA DE VENCIMENTO
        :return:
        """
        # --- ACRESCENTANDO ATÉ TERCEIRA FATURA ---
        self.df['FATURA 1'] = pd.to_datetime(self.df['FATURA 1'], dayfirst=True, errors='coerce')
        self.df['FATURA 2'] = self.df['FATURA 1'] + pd.DateOffset(months=1)
        self.df['FATURA 3'] = self.df['FATURA 1'] + pd.DateOffset(months=2)
        ## ---- FORMATANDO COLUNAS DATAS ----
        self.df['DATA'] = pd.to_datetime(self.df['DATA'], dayfirst=True, errors='coerce')
        self.df['FATURA 1'] = self.df['FATURA 1'].dt.strftime('%d/%b')
        self.df['FATURA 2'] = self.df['FATURA 2'].dt.strftime('%d/%b')
        self.df['FATURA 3'] = self.df['FATURA 3'].dt.strftime('%d/%b')
        self.df['DATA'] = self.df['DATA'].dt.strftime('%d/%b')

    def ajustar_faturas(self):
        """
        ADICIONA NOME DO PLANO NAS COLUNAS FATURAS QUANDO SERVIÇO NÃO TEM FATURA PARA PAGAR
        Usa abreviações específicas para alguns tipos
        """
        # Dicionário de mapeamento para exibição nas faturas
        abreviacoes_fatura = {
            "TROCA APARELHO": "TROCA A",
            "TROCA PLANO": "TROCA P",
            "RESGATE": "RESGATE",  # mantido igual
            "PRE PAGO": "PRE PAGO",  # mantido igual
            "SEGURO PROTEÇÃO MÓVEL": "SEGURO P"
            # Se surgirem mais no futuro, é só adicionar aqui
        }

        # Lista de planos que devem ter o nome (ou abreviação) nas colunas de fatura
        planos_fixos = list(abreviacoes_fatura.keys())

        # Criar máscara: verifica se o início do PLANOS está na lista
        # (considerando que pode ter /PORT no final)
        mask = self.df['PLANOS'].str.upper().str.split('/').str[0].isin(planos_fixos)

        # Aplicar a abreviação correta
        def get_abreviacao(plano):
            plano_base = plano.upper().split('/')[0].strip()
            return abreviacoes_fatura.get(plano_base, plano_base)

        faturas = ['FATURA 1', 'FATURA 2', 'FATURA 3']
        for col in faturas:
            self.df.loc[mask, col] = self.df.loc[mask, 'PLANOS'].apply(get_abreviacao)

    # Dentro da classe
    def remover_imei_ap(self):
        """
        Remove linhas onde a coluna 'Ativação' indica upgrade, troca equivalente ou seguro proteção móvel.
        Isso elimina IMEIs/equipamentos associados a esses tipos de ativação indesejados.
        """
        coluna_ativacao = 'ATIVAÇÃO'  # Mude aqui se o nome exato for diferente (ex: 'Tipo Ativação', 'Ativação Plano', etc.)

        if coluna_ativacao not in self.df.columns:
            print(f"Coluna '{coluna_ativacao}' não encontrada. Pulando remoção de linhas indesejadas.")
            return

        # Limpa espaços e converte para string
        self.df[coluna_ativacao] = self.df[coluna_ativacao].fillna('').astype(str).str.strip().str.upper()

        # Palavras-chave para remover (case-insensitive, mas usamos upper para facilitar)
        termos_remover = [
            'UPGRADE DE PLANO',
            'TROCA PLANO EQUIVALENTE',
            'SEGURO PROTEÇÃO MÓVEL'
        ]

        # Cria máscara: True se conter algum dos termos
        mask = self.df[coluna_ativacao].str.contains('|'.join(termos_remover), na=False)

        # Mostra quantas linhas seriam removidas (debug útil)
        qtd_remover = mask.sum()
        if qtd_remover > 0:
            print(f"Removendo {qtd_remover} linhas com ativações indesejadas (upgrade/troca/seguro).")
        else:
            print("Nenhuma linha encontrada com os termos de remoção na coluna 'Ativação'.")

        # Remove as linhas
        self.df = self.df[~mask].reset_index(drop=True)

        # Opcional: remove a coluna temporária se não precisar mais
        # self.df.drop(columns=[coluna_ativacao], inplace=True)  # só se quiser limpar



    def ordenar_colunas(self):
        """
        ORDENANDA COLUNAS QUE NA SEQUENCIA CERTA E SOMENTE COLUNAS NECESSÁRIAS
        :return:
        """
        ordem_colunas = [
            "DATA", "GED", "LOJA",
            "PLANOS", "AP", "ICCID/IMEI",
            "CONSULTOR","CLIENTE", "CPF",
            "NUMERO","PROVISORIO",
            "DEBITO AUTOMATICO", "FATURA 1", "FATURA 2", "FATURA 3"
        ]
        self.df = self.df[ordem_colunas]

    def get_desktop_path(self):
        """
        Salva pasta na área de trabalho, verifica se tem ou não OneDrive.
        """
        # Caminho padrão local da área de trabalho
        desktop_local = os.path.join(os.path.expanduser("~"), "Desktop")

        # Possível caminho do OneDrive para Desktop (pt-br e inglês)
        onedrive_desktop_pt = os.path.join(os.path.expanduser("~"), "OneDrive", "Área de Trabalho")
        onedrive_desktop_en = os.path.join(os.path.expanduser("~"), "OneDrive", "Desktop")

        # Verifica se existe pasta do OneDrive
        if os.path.isdir(onedrive_desktop_pt):
            return onedrive_desktop_pt
        elif os.path.isdir(onedrive_desktop_en):
            return onedrive_desktop_en
        else:
            return desktop_local

    def salvar(self, nome_arquivo="PlanilhaTratada.xlsx"):
        """
        PARTE FINAL, SALVA PLANILHA EM ÁREA DE TRABALHO COM CAMINHO PEGO PELA FUNÇÃO get_desktop_path
        :param nome_arquivo:
        :return:
        """
        desktop = Path(self.get_desktop_path())  # Converte para Path
        caminho_saida = desktop / nome_arquivo

        self.df.to_excel(caminho_saida, index=False, engine='openpyxl')
        print(f"Planilha salva em {caminho_saida}")

        return self.df.to_excel

"""if __name__ == '__main__':
    # === Lendo a planilha original do sistema ===
    df = pd.read_excel("pl_excel.xlsx", header=2, dtype={
        "ICCID/IMEI": str, "NUMERO": str, "PROVISORIO": str, "CPF": str
    })
    planilha = PlanilhaClaro(df)

    planilha.editar_colunas()
    planilha.aplicar_planos()
    planilha.criar_coluna_ap()
    planilha.criar_coluna_loja()
    planilha.editar_debito()
    planilha.editar_nomes_vendedor()
    planilha.criar_colunas_faturas()
    planilha.ajustar_faturas()
    planilha.ordenar_colunas()
    planilha.salvar("planilha_tratada.xlsx")"""