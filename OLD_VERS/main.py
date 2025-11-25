import pandas as pd

# --- 1. Definição dos caminhos ---
caminho_do_arquivo_servidor = 'banrisul_dados_gp2.xls'
caminho_do_arquivo_conta = 'banrisul_retorno_contas2.xls'
arquivo_de_saida_txt = 'saida_formatada.txt' # Nome do arquivo TXT final

# --- 2. Constantes de Layout (Baseado na imagem da máscara) ---
# Valores fixos ou que não temos nos arquivos de entrada
DATA_PAGAMENTO = '20220220' # Da sua imagem: (20220220 - DATA DO PAGAMENTO)
TIPO_EMPREGO = 'J'        # Da sua imagem: (PREENCHER COM A LETRA 'J')
VALOR_13_SALARIO = '0' * 15 # Campo 'NC02VALORCALCP15' (13o salário), não temos, preenchemos com zeros
COD_OCORRENCIA = '0' * 2    # 'NC02OCORRP02 PIC 99'. (BRANCOS) é contraditório, PIC 99 sugere zeros.
DESC_OCORRENCIA = ' ' * 82  # 'NC02DESCROCORRC82 CHAR(82)', (BRANCOS)
DATA_AGENDAMENTO = ' ' * 8   # 'NC02DATAAGENC08 CHAR(08)', (BRANCOS)
CNPJ_PAGADOR = ' ' * 14     # 'NC02EMPREGADORC14 CHAR(14)', não temos, preenchemos com espaços

try:
    # --- 3. Carregar o primeiro arquivo (banrisul_retorno_contas) ---
    colunas_como_texto_contas = {'cpf': str, 'banco': str, 'agencia': str, 'conta': str}
    df_contas = pd.read_excel(caminho_do_arquivo_conta, dtype=colunas_como_texto_contas)
    print(f"Arquivo '{caminho_do_arquivo_conta}' lido com sucesso.")

    # --- 4. Carregar o segundo arquivo (banrisul_dados_gp) ---
    df_dados = pd.read_excel(caminho_do_arquivo_servidor, dtype={'cpf': str})
    print(f"Arquivo '{caminho_do_arquivo_servidor}' lido com sucesso.")

    # --- 5. Filtrar o 'df_dados' para manter a maior matrícula ---
    df_dados_ordenado = df_dados.sort_values(by='matricula', ascending=False)
    df_dados_limpo = df_dados_ordenado.drop_duplicates(subset='cpf', keep='first')
    print("DataFrame de servidores filtrado, mantendo a maior matrícula.")

    # --- 6. Cruzar (Merge) os dados ---
    df_final = pd.merge(df_contas, df_dados_limpo, on='cpf', how='left')
    print("\n--- Processando e Gerando Arquivo TXT ---")

    # --- 7. NOVO: Formatar e Salvar no Arquivo TXT ---
    # Abre o arquivo de saída para escrita
    with open(arquivo_de_saida_txt, 'w', encoding='utf-8') as f:
        
        for indice, linha in df_final.iterrows():
            
            if pd.isna(linha['nome']):
                # Caso "NÃO ENCONTRADO"
                print(f"Linha {indice}: CPF={linha['cpf']} (Conta: {linha['conta']}) - SERVIDOR NÃO ENCONTRADO")
                # Escreve no TXT como um comentário (opcional, mas útil)
                f.write(f"# CPF NÃO ENCONTRADO: {linha['cpf']} (Conta: {linha['conta']})\n")
            
            else:
                # Caso "ENCONTRADO", formata e salva
                
                # ljust() -> Alinha à esquerda e preenche com espaços (para CHAR de texto)
                # rjust() -> Alinha à direita e preenche com '0' (para CHAR de número ou PIC 9)

                # NC02NOMEC46 CHAR(46)
                nome = str(linha['nome']).ljust(46, ' ')
                
                # NC02CPFC11 CHAR(11)
                cpf = str(linha['cpf']).rjust(11, '0')
                
                # NC02BANCOC03 CHAR(03)
                banco = str(linha['banco']).rjust(3, '0')
                
                # NC02AGENCIAC04 CHAR(04)
                agencia = str(linha['agencia']).rjust(4, '0')
                
                # NC02CONTAC10 CHAR(10)
                conta = str(linha['conta']).rjust(10, '0')
                
                # NC02NUMFUNCC15 CHAR(15)
                matricula = str(linha['matricula']).rjust(15, '0')
                
                # NC02VALORCALCP15 PIC '(13)9V99' (Salário)
                # Converte '171620' para 171620.00, remove o ponto (17162000) e preenche
                salario_val = 0.0 if pd.isna(linha['salario']) else float(linha['salario'])
                valor_salario = str(int(salario_val * 100)).rjust(15, '0')

                # Junta todos os campos formatados em uma única string
                linha_formatada = (
                    f"{nome}"
                    f"{cpf}"
                    f"{banco}"
                    f"{agencia}"
                    f"{conta}"
                    f"{matricula}"
                    f"{valor_salario}"
                    f"{VALOR_13_SALARIO}"
                    f"{COD_OCORRENCIA}"
                    f"{DESC_OCORRENCIA}"
                    f"{DATA_AGENDAMENTO}"
                    f"{DATA_PAGAMENTO}"
                    f"{TIPO_EMPREGO}"
                    f"{CNPJ_PAGADOR}"
                )
                
                # Escreve a linha final no arquivo, com uma quebra de linha
                f.write(linha_formatada + '\n')
                
                # Imprime no console para feedback
                print(f"Linha {indice}: CPF={linha['cpf']} - Nome={linha['nome']} ... (Formatado e salvo)")

    print(f"\nProcesso concluído! Arquivo salvo em: '{arquivo_de_saida_txt}'")

except FileNotFoundError as e:
    print(f"Erro: O arquivo não foi encontrado. Verifique o nome/caminho: {e.filename}")
except KeyError as e:
    print(f"Erro: Não foi possível encontrar a coluna {e}. Verifique se os nomes estão corretos nos arquivos XLS.")
except Exception as e:
    print(f"Ocorreu um erro ao ler ou processar os arquivos: {e}")