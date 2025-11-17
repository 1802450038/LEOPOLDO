import pandas as pd

# --- 1. Definição dos caminhos ---
caminho_do_arquivo_servidor = 'banrisul_dados_gp2.xlsx'
caminho_do_arquivo_conta = 'banrisul_retorno_contas2.xlsx'
arquivo_de_saida_txt = 'saida_formatada2.txt' # Nome do arquivo TXT final

# --- 2. Constantes de Layout (Baseado na imagem da máscara) ---
DATA_PAGAMENTO = '20220220'
TIPO_EMPREGO = 'J'
VALOR_13_SALARIO = '0' * 15
COD_OCORRENCIA = '0' * 2
DESC_OCORRENCIA = ' ' * 82
DATA_AGENDAMENTO = ' ' * 8
CNPJ_PAGADOR = '88131164000107'

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

    # --- 7. NOVO: Formatar e Salvar no Arquivo TXT (Com a nova regra) ---
    with open(arquivo_de_saida_txt, 'w', encoding='utf-8') as f:
        
        for indice, linha in df_final.iterrows():
            
            # --- Início da Lógica de Formatação ---
            
            # 1. Pega os dados que sempre existem (do df_contas)
            cpf = str(linha['cpf'])
            banco_orig = str(linha['banco'])
            agencia_orig = str(linha['agencia'])
            conta_orig = str(linha['conta'])

            # 2. Verifica se o servidor foi encontrado no merge
            servidor_encontrado = not pd.isna(linha['nome'])
            
            # 3. Pega os dados do servidor (ou define padrão se não encontrado)
            if servidor_encontrado:
                nome = str(linha['nome'])
                matricula = str(linha['matricula'])
                salario_val = 0.0 if pd.isna(linha['salario']) else float(linha['salario'])
            else:
                # Se não encontrou, usa valores padrão "em branco"
                nome = ''      # Vai virar espaços no .ljust()
                matricula = '0'  # Vai virar zeros no .rjust()
                salario_val = 0.0
            
            # --- AQUI APLICA A NOVA REGRA ---
            conta_limpa = conta_orig.strip() # Remove espaços extras
            
            if (not servidor_encontrado) or conta_limpa.startswith('39') or conta_limpa.startswith('38'):
                # CONDIÇÃO ATINGIDA: Servidor não encontrado OU conta começa com 39/38
                banco = '041'
                agencia = '0000'
                conta = '0000000000'
                
                if not servidor_encontrado:
                    print(f"Linha {indice}: CPF={cpf} - SERVIDOR NÃO ENCONTRADO. Aplicando regra BCO/AG/CTA zerados.")
                else:
                    print(f"Linha {indice}: CPF={cpf} - Conta inicia com 38/39. Aplicando regra BCO/AG/CTA zerados.")
            else:
                # CONDIÇÃO NORMAL: Usa os dados originais do arquivo de contas
                banco = banco_orig
                agencia = agencia_orig
                conta = conta_orig
                print(f"Linha {indice}: CPF={cpf} - Nome={nome} ... (Formatado e salvo)")
            
            # --- Fim da Lógica de Formatação ---

            # 4. Aplica a máscara/padding nos valores finais
            # nome_fmt = nome.ljust(46, ' ')
            nome_fmt = nome[:46].ljust(46, ' ')
            cpf_fmt = cpf.rjust(11, '0')
            banco_fmt = banco.rjust(3, '0')
            agencia_fmt = agencia.rjust(4, '0')
            conta_fmt = conta.rjust(10, '0')
            matricula_fmt = str(matricula).rjust(15, '0')
            valor_salario_fmt = str(int(salario_val * 100)).rjust(15, '0')

            # 5. Junta todos os campos formatados em uma única string
            linha_formatada = (
                f"{nome_fmt}"
                f"{cpf_fmt}"
                f"{banco_fmt}"
                f"{agencia_fmt}"
                f"{conta_fmt}"
                f"{matricula_fmt}"
                f"{valor_salario_fmt}"
                f"{VALOR_13_SALARIO}"
                f"{COD_OCORRENCIA}"
                f"{DESC_OCORRENCIA}"
                f"{DATA_AGENDAMENTO}"
                f"{DATA_PAGAMENTO}"
                f"{TIPO_EMPREGO}"
                f"{CNPJ_PAGADOR}"
            )
            
            # 6. Escreve a linha final no arquivo, com uma quebra de linha
            f.write(linha_formatada + '\n')

    print(f"\nProcesso concluído! Arquivo salvo em: '{arquivo_de_saida_txt}'")

except FileNotFoundError as e:
    print(f"Erro: O arquivo não foi encontrado. Verifique o nome/caminho: {e.filename}")
except KeyError as e:
    print(f"Erro: Não foi possível encontrar a coluna {e}. Verifique se os nomes estão corretos nos arquivos XLS.")
except Exception as e:
    print(f"Ocorreu um erro ao ler ou processar os arquivos: {e}")