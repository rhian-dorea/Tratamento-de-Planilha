from PIL import Image


def converter_para_icone(caminho_entrada, caminho_saida, tamanhos=None):
    """
    Converte uma imagem (como PNG) para o formato ICO, que suporta múltiplos tamanhos.

    :param caminho_entrada: Caminho para o arquivo de imagem de origem (ex: 'meu_icone.png').
    :param caminho_saida: Caminho para salvar o arquivo de ícone de saída (ex: 'meu_icone.ico').
    :param tamanhos: Uma lista de tuplas (largura, altura) para os tamanhos de ícone a incluir.
                     Se for None, a Pillow usará um conjunto padrão de tamanhos.
    """
    try:
        # 1. Abre a imagem de origem
        img = Image.open(caminho_entrada)

        # O Pillow salva no formato ICO, que pode conter vários tamanhos para
        # que o Windows possa escolher o melhor ícone para cada exibição.

        if tamanhos is None:
            # Tamanhos padrão recomendados para ícones de aplicativo do Windows
            tamanhos = [(16, 16), (24, 24), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)]

        # 2. Salva a imagem no formato ICO com os tamanhos especificados
        img.save(caminho_saida, format='ICO', sizes=tamanhos)

        print(
            f"Sucesso! Imagem '{caminho_entrada}' convertida para ícone '{caminho_saida}' com os tamanhos: {tamanhos}")

    except FileNotFoundError:
        print(f"Erro: Arquivo não encontrado no caminho: {caminho_entrada}")
    except Exception as e:
        print(f"Ocorreu um erro durante a conversão: {e}")


# --- Configurações ---
# Coloque o nome do seu arquivo de imagem aqui (assumindo que está na mesma pasta do script)
nome_do_arquivo_png = "Gemini_Generated_Image_2gosnt2gosnt2gos.png"

# O nome que você quer dar ao seu arquivo de ícone
nome_do_arquivo_ico = "app_icone.ico"

# Chama a função para converter
converter_para_icone(nome_do_arquivo_png, nome_do_arquivo_ico)