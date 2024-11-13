from Bio import Entrez, SeqIO
import openpyxl
import os
import time

# Carregar a planilha de entrada
caminho_planilha = "C:\\planilha.xlsx"
workbook_genes = openpyxl.load_workbook(caminho_planilha)
codigos_genes = workbook_genes.active

# Especifique o caminho onde você deseja salvar os arquivos
save_path = "Insira aqui"

# Forneça seu e-mail para o NCBI  
Entrez.email = "seuemail@gmail.com"

# Iterar sobre as linhas preenchidas na coluna A
for i in range(1, codigos_genes.max_row + 1):
    genome_id = codigos_genes[f'A{i}'].value  # Ler o valor da célula na coluna A

    if genome_id:  # Verificar se a célula não está vazia
        genome_id = genome_id.strip()  # Remover espaços em branco ao redor do valor, se houver

        # Verificar se o genome_id parece válido
        if len(genome_id) == 0:
            print(f"ID vazio na linha {i}. Pulando...")
            continue

        print(f"Buscando informações para {genome_id}")

        try:
            # Usando Entrez para buscar o assembly no formato GenBank
            handle = Entrez.efetch(db="assembly", id=genome_id, rettype="docsum", retmode="xml")
            resultado = handle.read()

            if len(resultado) == 0:
                print(f"Nenhum resultado encontrado para {genome_id}. Pulando...")
                continue

            # Decodificar bytes para string antes de salvar no arquivo
            resultado_str = resultado.decode('utf-8')

            # Criar o caminho do arquivo com o separador correto
            file_path = os.path.join(save_path, f"{genome_id}.gbk")

            # Salvar o arquivo em formato GenBank no caminho especificado
            with open(file_path, "w") as gbk_file:
                gbk_file.write(resultado_str)

            handle.close()

            print(f"Arquivo {genome_id}.gbk baixado com sucesso em {save_path}!")
        except Exception as e:
            print(f"Erro ao buscar ou salvar o assembly {genome_id}: {e}")

        # Pausa para evitar sobrecarga no NCBI
        time.sleep(0.5)
