import os
import xml.etree.ElementTree as ET

# Diretório contendo os arquivos XML 
diretorio = 'C:/xml'

# Lista todos os arquivos XML no diretório
arquivos_xml = [arquivo for arquivo in os.listdir(diretorio) if arquivo.endswith('.xml')]
cont = 0

# Le cada arquivo XML
for arquivo_xml in arquivos_xml:
    caminho_arquivo = os.path.join(diretorio, arquivo_xml)
    tree = ET.parse(caminho_arquivo)
    root = tree.getroot()

    # Localiza a tag <fone_residencial>
    for fone_residencial_tag in root.iter('fone_residencial'):
        # Altera o valor para um novo número de telefone 
        fone_residencial_tag.text = '0'
        cont += 1

    # Salva as alterações de volta no arquivo
    tree.write(caminho_arquivo)

    print(f"Tag <fone_residencial> alterada em {arquivo_xml}")

print(f"{cont} Alterações concluídas!")

