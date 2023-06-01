import pandas as pd
import os
import re
import numpy as np
from PIL import Image, ImageDraw, ImageFont, ExifTags
from resizeimage import resizeimage

# Defina o diretório correto onde estão os arquivos
directory = 'C:/Users/CS80037/Desktop/CART/shell/'
os.chdir(directory)

try:
    # leia o arquivo Excel
    df = pd.read_excel('C:/Users/CS80037/Desktop/CART/shell/IIGDCARTEIRINHA.xlsx')
except Exception as e:
    print("Ocorreu um erro ao ler o arquivo Excel:", str(e))

# eliminar as colunas indesejadas
df = df.drop(['Carimbo de data/hora', 'PARA CONTINUAR COM SUA ATUALIZAÇÃO DE CADASTRO, É IMPORTANTE QUE LEIA ATENTAMENTE O TERMO ACIMA E NOS INFORME SE ESTÁ CIENTE, RESPONDENDO AS OPÇÕES ABAIXO:', 'NOME MÃE (COMPLETO)', 'CELULAR P/ RECADO (SOMENTE NÚMERO)', 'NOME DA PESSOA P/ RECADO', 'TELEFONE FIXO (OPCIONAL)', 'ENDEREÇO DO  E-MAIL ( OPCIONAL )', 'RUA EM QUE MORA ( Ñ COLOCAR Nº )', 'NUMERO CASA, PREDIO OU CONDOMINIO (Nº)', 'CIDADE( EM QUE MORA)', 'CEP(SOMENTE NÚMERO)', 'EM CASO DE APARTAMENTO (Nº APTO e BLOCO)', 'BATIZADO (a) NAS AGUAS ?', 'DATA DO BATISMO', 'PARTICIPA DE ALGUM MINISTERIO DA IIGD ?\n', 'email'], axis=1)

# transformar todas as letras da coluna 'C' em maiúsculas
df['NOME COMPLETO  (SEM ABREVIAÇÕES)'] = df['NOME COMPLETO  (SEM ABREVIAÇÕES)'].str.upper()

df['Nº DO R.G.  (SOMENTE NÚMERO)'] = df['Nº DO R.G.  (SOMENTE NÚMERO)'].apply(lambda x: re.sub(r'\D', '', str(x)))

# remover caracteres não numéricos da coluna 'F' e garantir que tem exatamente 11 dígitos ou é vazio
df['WHATSAPP(PRINCIPAL)'] = df['WHATSAPP(PRINCIPAL)'].apply(lambda x: re.sub(r'\D', '', str(x)))
df['WHATSAPP(PRINCIPAL)'] = df['WHATSAPP(PRINCIPAL)'].apply(lambda x: x[:11] if len(x) > 11 else x)
df['WHATSAPP(PRINCIPAL)'] = df['WHATSAPP(PRINCIPAL)'].apply(lambda x: x if len(x) == 11 else "")

# Tratar a coluna 'DATA DE NASCIMENTO'
df['DATA DE NASCIMENTO'] = pd.to_datetime(df['DATA DE NASCIMENTO'], errors='coerce')
df['DATA DE NASCIMENTO'] = df['DATA DE NASCIMENTO'].dt.strftime('%d/%m/%Y')

# Renomear as colunas antes de salvar o DataFrame para um arquivo Excel
df = df.rename(columns={
    'NOME COMPLETO  (SEM ABREVIAÇÕES)': 'nome',
    'Nº DO R.G.  (SOMENTE NÚMERO)': 'identidade',
    'WHATSAPP(PRINCIPAL)': 'telefone',
    'DATA DE NASCIMENTO': 'data_nascimento',
    'ESTADO CIVIL': 'estado_civil',
    'FUNÇÃO OU MINISTERIO ( SERÁ COLOCADO NA CARTEIRINHA)': 'funcao',
})

# Salvar o DataFrame tratado em um novo arquivo Excel
df.to_excel('C:/Users/CS80037/Desktop/CART/shell/IIGDCARTEIRINHA_tratado.xlsx', index=False)

print("Dados tratados com sucesso.")

# Especifique o caminho para a fonte que você deseja usar
font_path = "C:/Users/CS80037/Desktop/CART/shell/fontes/Poppins-Medium.ttf"
font_size = 14
font = ImageFont.truetype(font_path, font_size)

# Dicionário com as coordenadas dos campos na imagem da carteirinha
coords = {
    "nome": (192, 48),
    "identidade": (227, 106),
    "estado_civil": (108, 166),
    "data_nascimento": (351, 167),
    "funcao": (174, 211),
    "telefone": (406, 208),
    "foto": (9, 22),
}

# Função para corrigir a orientação da imagem
def corrigir_orientacao_imagem(image):
    try:
        for orientation in ExifTags.TAGS.keys():
            if ExifTags.TAGS[orientation] == 'Orientation':
                break
        exif = image._getexif()
        if exif is not None:
            exif = dict(exif.items())
            if orientation in exif:
                if exif[orientation] == 3:
                    image = image.rotate(180, expand=True)
                elif exif[orientation] == 6:
                    image = image.rotate(270, expand=True)
                elif exif[orientation] == 8:
                    image = image.rotate(90, expand=True)
    except Exception as e:
        print(f"Erro ao corrigir a orientação da imagem: {e}")
    return image

# Função para redimensionar a imagem
def redimensionar_imagem(image, size):
    return image.resize(size)

# Percorra as linhas do DataFrame
for index, row in df.iterrows():
    # Verifique o valor da coluna 'funcao'
    funcao = row['funcao']

    # Verifique se o valor da função é um número
    if pd.notnull(funcao) and not isinstance(funcao, (int, float)):
        # Escolha o modelo da carteirinha com base na função
        if 'PASTOR' in funcao.upper() or 'OBREIRO' in funcao.upper() or 'CAPELANIA' in funcao.upper() or '':
            frente_carteirinha_path = "C:/Users/CS80037/Desktop/CART/shell/carterinhaoficialfrente.jpg"
            tras_carteirinha_path = "C:/Users/CS80037/Desktop/CART/shell/carterinhaoficialtras.jpg"
        elif 'MEMBRO' in funcao.upper():
            frente_carteirinha_path = "C:/Users/CS80037/Desktop/CART/shell/carterinhamembrofrente.jpg"
            tras_carteirinha_path = "C:/Users/CS80037/Desktop/CART/shell/carterinhamembrotras.jpg"
        else:
            # Se não for nenhuma das funções especificadas, pule para a próxima iteração
            continue

        # Crie o diretório para salvar as fotos
        foto_directory = "C:/Users/CS80037/Desktop/CART/shell/fotos/"
        os.makedirs(foto_directory, exist_ok=True)

        # Função para redimensionar a imagem sem corrigir a orientação
        def resize_image(input_image_path, output_image_path, size):
            with Image.open(input_image_path) as image:
                image = corrigir_orientacao_imagem(image)  # Corrigir orientação da imagem
                image = redimensionar_imagem(image, size)  # Redimensionar imagem
                image.save(output_image_path, image.format)

        # Abra a imagem da frente da carteirinha
        frente_carteirinha = Image.open(frente_carteirinha_path)
        draw_frente = ImageDraw.Draw(frente_carteirinha)

        # Abra a imagem de trás da carteirinha
        tras_carteirinha = Image.open(tras_carteirinha_path)
        draw_tras = ImageDraw.Draw(tras_carteirinha)

        # Preencha os dados na frente e no verso da carteirinha
        for campo, coord in coords.items():
            if campo == "foto":
                # Verifique se o índice está dentro dos limites do DataFrame
                if index < len(df) and pd.notnull(row[campo]):
                    # Defina o nome do arquivo para abrir a foto
                    foto_filename = f"{index + 2}.jpg"
                    foto_path = os.path.join(foto_directory, foto_filename)

                    # Verifique se o arquivo da foto existe antes de prosseguir
                    if os.path.exists(foto_path):
                        try:
                            # Redimensionar a foto
                            tamanho_redimensionado = (108, 124)
                            foto_redimensionada_path = os.path.join(foto_directory, f"{index + 2}_redimensionada.jpg")
                            resize_image(foto_path, foto_redimensionada_path, tamanho_redimensionado)

                            # Abre a foto redimensionada
                            foto_redimensionada = Image.open(foto_redimensionada_path)

                            # Cola a foto redimensionada na carteirinha
                            frente_carteirinha.paste(foto_redimensionada, coord)
                        except FileNotFoundError:
                            print(f"A foto {foto_filename} não foi encontrada.")
                        except Exception as e:
                            print(f"Erro ao redimensionar a foto {foto_filename}: {e}")
            else:
                # Verificar se o campo não está vazio ou é NaN
                if index < len(df) and pd.notnull(row[campo]):
                    text = str(row[campo])
                    draw_frente.text(coord, text, fill='black', font=font)

        # Redimensionar a imagem de trás para ter as mesmas dimensões da frente
        tras_carteirinha_resized = redimensionar_imagem(tras_carteirinha, frente_carteirinha.size)

        # Juntar a frente e o verso da carteirinha
        carteirinha_completa = Image.fromarray(
            np.hstack((np.array(frente_carteirinha), np.array(tras_carteirinha_resized))))

        # Redimensionar a carteirinha final
        tamanho_redimensionado_final = (744, 245)
        carteirinha_final = redimensionar_imagem(carteirinha_completa, tamanho_redimensionado_final)

        # Salve a imagem final da carteirinha na pasta "C:\Users\CS80037\Desktop\CART\prontas"
        output_directory = "C:/Users/CS80037/Desktop/CART/prontas/"
        os.makedirs(output_directory, exist_ok=True)
        output_path = os.path.join(output_directory, f"carteirinha_{row['nome']}.jpg")
        carteirinha_final.save(output_path)

print("Todas as carteirinhas foram criadas com sucesso.")
