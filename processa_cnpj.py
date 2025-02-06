def process_file(input_file, output_file):
    try:
        with open(input_file, 'r', encoding="utf8") as infile, open(output_file, 'w') as outfile:
            for line in infile:
                # Remove espaços extras e quebras de linha
                line = line.strip()
                # Separar pelo delimitador '|'
                parts = line.split('|')
                
                # Verifica se a linha possui pelo menos 6 campos
                if len(parts) >= 7:
                    # Verifica o campo na posição 6 (índice 5)
                    field = parts[6]
                    
                    # Verifica se contém exatamente 14 números
                    if len(field) < 14 and field.isdigit():
                        # Completa com zeros à esquerda
                        parts[6] = field.zfill(14)
                    elif not field.isdigit():
                        # Substituir valores não numéricos por 14 zeros
                        parts[6] = '0' * 14
                
                # Reconstrói a linha com os campos corrigidos
                new_line = '|'.join(parts)
                outfile.write(new_line + '\n')

        print(f"Arquivo processado com sucesso! Saída salva em '{output_file}'.")

    except FileNotFoundError:
        print("Erro: O arquivo de entrada não foi encontrado.")
    except Exception as e:
        print(f"Ocorreu um erro: {e}")

# entrada e saida do arquivo corrigido
input_file = 'C:/sit/exporta_sit.txt'
output_file = 'C:/sit/exporta_sit_corrigido.txt'
process_file(input_file, output_file)