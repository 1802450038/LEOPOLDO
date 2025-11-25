import pandas as pd
import tkinter as tk
import datetime
from tkinter import filedialog, messagebox

# ==============================================================================
#  PASSO 1: A LÓGICA CORRIGIDA
# ==============================================================================

def processar_arquivos(caminho_servidor, caminho_conta, caminho_saida, data_pagamento, status_callback):
    """
    Função principal que executa toda a lógica de processamento de arquivos.
    """
    try:
        # --- 2. Constantes de Layout ---
        DATA_PAGAMENTO = data_pagamento
        TIPO_EMPREGO = 'J'
        COD_OCORRENCIA = ' ' * 2
        DESC_OCORRENCIA = ' ' * 82
        DATA_AGENDAMENTO = ' ' * 8
        CNPJ_PAGADOR = '88131164000107'
        REMOVE_DUPLICADOS = False

        # --- 3. Carregar o arquivo de CONTAS ---
        status_callback(f"Lendo '{caminho_conta}'...")
        colunas_como_texto_contas = {'cpf': str, 'banco': str, 'agencia': str, 'conta': str}
        df_contas = pd.read_excel(caminho_conta, dtype=colunas_como_texto_contas)

        # --- 4. Carregar o arquivo de SERVIDORES (GP) ---
        status_callback(f"Lendo '{caminho_servidor}'...")
        df_dados = pd.read_excel(caminho_servidor, dtype={'cpf': str})

        if(REMOVE_DUPLICADOS):
            # --- 5. Filtrar o 'df_dados' para manter a maior matrícula ---
            status_callback("Filtrando CPFs duplicados (maior matrícula)...")
            df_dados_ordenado = df_dados.sort_values(by='matricula', ascending=False)        
            df_dados_limpo = df_dados_ordenado.drop_duplicates(subset='cpf', keep='first')
        else:
            status_callback("Ordenando Matriculas...")
            df_dados_limpo = df_dados.sort_values(by='matricula', ascending=False)

        # --- 6. Cruzar (Merge) os dados CORRIGIDO ---
        status_callback("Cruzando dados (Mantendo todos os funcionários)...")
        
        # CORREÇÃO AQUI: Invertemos a ordem. df_dados_limpo fica na esquerda.
        # Isso garante que TODOS os funcionários fiquem no resultado.
        df_final = pd.merge(df_dados_limpo, df_contas, on='cpf', how='left')

        # --- 7. Tratar quem ficou sem conta (Preencher com '0') ---
        # Quem não tinha conta ficou com NaN (vazio). Vamos colocar '0'.
        cols_bancarias = ['banco', 'agencia', 'conta']
        for col in cols_bancarias:
            df_final[col] = df_final[col].fillna('0')

        # --- 8. Ordenar o resultado final por nome ---
        status_callback("Ordenando resultado por nome...")
        df_final_ordenado = df_final.sort_values(by='nome', ascending=True, na_position='last')

        # --- 9. Formatar e Salvar no Arquivo TXT ---
        status_callback("Gerando arquivo de saída formatado...")
        total_linhas = len(df_final_ordenado)

        with open(caminho_saida, 'w', encoding='utf-8') as f:
            
            for indice, linha in df_final_ordenado.iterrows():
                
                # Dados do Funcionário (Sempre existem agora)
                nome = str(linha['nome'])
                cpf = str(linha['cpf'])
                matricula = str(linha['matricula'])
                salario_val = 0.0 if pd.isna(linha['salario']) else float(linha['salario'])
                
                # Dados Bancários
                # Se não tinha conta, o fillna colocou '0'
                conta_orig = str(linha['conta']).strip()
                banco_orig = str(linha['banco']).strip()
                agencia_orig = str(linha['agencia']).strip()
                
                # Lógica para Zerar Dados Bancários
                # Se a conta for '0' (não encontrada) OU cair nas regras de exclusão (38/39)
                if conta_orig == '0' or conta_orig.startswith('39') or conta_orig.startswith('38'):
                    banco = '041'         # Mantém o banco Banrisul
                    agencia = '0000'      # Zera agência
                    conta = '0000000000'  # Zera conta
                else:
                    banco = banco_orig
                    agencia = agencia_orig
                    conta = conta_orig
                
                # Aplica a máscara/padding
                nome_fmt = nome[:46].ljust(46, ' ') 
                cpf_fmt = cpf.rjust(11, '0')
                banco_fmt = banco.rjust(3, '0')
                agencia_fmt = agencia.rjust(4, '0')
                conta_fmt = conta.rjust(10, '0')
                matricula_fmt = str(matricula).rjust(15, '0')
                valor_salario_fmt = str(int(salario_val)).rjust(15, '0')

                linha_formatada = (
                    f"{nome_fmt}{cpf_fmt}{banco_fmt}{agencia_fmt}{conta_fmt}"
                    f"{matricula_fmt}{valor_salario_fmt}{valor_salario_fmt}"
                    f"{COD_OCORRENCIA}{DESC_OCORRENCIA}{DATA_AGENDAMENTO}"
                    f"{DATA_PAGAMENTO}{TIPO_EMPREGO}{CNPJ_PAGADOR}"
                )
                
                f.write(linha_formatada + '\n')

        status_callback(f"Processo concluído! {total_linhas} linhas salvas.")
        messagebox.showinfo("Sucesso", f"Processo concluído!\n{total_linhas} linhas salvas em:\n{caminho_saida}")

    except FileNotFoundError as e:
        status_callback(f"Erro: Arquivo não encontrado - {e.filename}")
        messagebox.showerror("Erro de Arquivo", f"Erro: Arquivo não encontrado:\n{e.filename}")
    except KeyError as e:
        status_callback(f"Erro: Coluna não encontrada {e}. Verifique os arquivos XLS.")
        messagebox.showerror("Erro de Coluna", f"Erro: Coluna não encontrada: {e}\n\nVerifique se os arquivos XLS têm os cabeçalhos corretos (cpf, nome, matricula, etc).")
    except Exception as e:
        status_callback(f"Erro inesperado: {e}")
        messagebox.showerror("Erro", f"Ocorreu um erro inesperado:\n{e}")

# ==============================================================================
#  PASSO 2: A INTERFACE GRÁFICA (Tkinter)
# ==============================================================================

class App:
    def __init__(self, root):
        self.root = root
        self.root.title("Processador de Arquivos Banrisul")
        self.root.geometry("600x300")

        # --- Frame principal ---
        frame_main = tk.Frame(root, padx=10, pady=10)
        frame_main.pack(fill=tk.BOTH, expand=True)

        # --- 1. Arquivo de Servidores (dados_gp) ---
        frame_servidor = tk.Frame(frame_main)
        frame_servidor.pack(fill=tk.X)
        
        lbl_servidor = tk.Label(frame_servidor, text="Arquivo Servidores (dados_gp):", width=25, anchor="w")
        lbl_servidor.pack(side=tk.LEFT, padx=(0, 5))
        
        self.entry_servidor = tk.Entry(frame_servidor)
        self.entry_servidor.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        btn_servidor = tk.Button(frame_servidor, text="Procurar...", command=self.procurar_servidor)
        btn_servidor.pack(side=tk.LEFT, padx=(5, 0))

        # --- 2. Arquivo de Contas (retorno_contas) ---
        frame_contas = tk.Frame(frame_main)
        frame_contas.pack(fill=tk.X, pady=10)

        lbl_contas = tk.Label(frame_contas, text="Arquivo Contas (retorno_contas):", width=25, anchor="w")
        lbl_contas.pack(side=tk.LEFT, padx=(0, 5))
        
        self.entry_contas = tk.Entry(frame_contas)
        self.entry_contas.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        btn_contas = tk.Button(frame_contas, text="Procurar...", command=self.procurar_contas)
        btn_contas.pack(side=tk.LEFT, padx=(5, 0))

        # --- 3. Arquivo de Saída (TXT) ---
        frame_saida = tk.Frame(frame_main)
        frame_saida.pack(fill=tk.X)
        
        lbl_saida = tk.Label(frame_saida, text="Salvar arquivo de saída (.txt) como:", width=25, anchor="w")
        lbl_saida.pack(side=tk.LEFT, padx=(0, 5))
        
        self.entry_saida = tk.Entry(frame_saida)
        self.entry_saida.pack(side=tk.LEFT, fill=tk.X, expand=True)

        self.entry_saida.insert(0,"saida_banrisul")


        # --- 4. Data do pagamento ---
        frame_data = tk.Frame(frame_main)
        frame_data.pack(fill=tk.X)

        lbl_data = tk.Label(frame_data, text="Data do pagamento AAAAMMDD :", width=25, anchor="w")
        lbl_data.pack(side=tk.LEFT,padx=(0,5))
        
        self.entry_data = tk.Entry(frame_data)
        self.entry_data.pack(side=tk.LEFT,fill=tk.X, expand=True)
        self.entry_data.insert(0,datetime.date.today().strftime("%Y%m%d"))
        
        
        # --- 4. Botão de Processar ---
        frame_processar = tk.Frame(frame_main)
        frame_processar.pack(pady=(20, 10))
        
        self.btn_processar = tk.Button(frame_processar, text="Processar e Salvar Arquivo", 
                                       font=("Helvetica", 12, "bold"), 
                                       command=self.processar,
                                       bg="#4CAF50", fg="white")
        self.btn_processar.pack()

        # --- 5. Status Bar ---
        frame_status = tk.Frame(frame_main, relief=tk.SUNKEN, bd=1)
        frame_status.pack(fill=tk.X, side=tk.BOTTOM, pady=(10, 0))
        
        self.status_var = tk.StringVar()
        self.status_var.set("Pronto. Selecione os arquivos e clique em processar.")
        
        lbl_status = tk.Label(frame_status, textvariable=self.status_var, anchor="w")
        lbl_status.pack(fill=tk.X, padx=5)

    def procurar_servidor(self):
        path = filedialog.askopenfilename(
            title="Selecione o arquivo de servidores (dados_gp)",
            filetypes=(("Arquivos Excel", "*.xls *.xlsx"), ("Todos os arquivos", "*.*"))
        )
        if path:
            self.entry_servidor.delete(0, tk.END)
            self.entry_servidor.insert(0, path)

    def procurar_contas(self):
        path = filedialog.askopenfilename(
            title="Selecione o arquivo de contas (retorno_contas)",
            filetypes=(("Arquivos Excel", "*.xls *.xlsx"), ("Todos os arquivos", "*.*"))
        )
        if path:
            self.entry_contas.delete(0, tk.END)
            self.entry_contas.insert(0, path)

    def procurar_saida(self):
        path = filedialog.asksaveasfilename(
            title="Definir local e nome do arquivo de saída",
            defaultextension=".txt",
            filetypes=(("Arquivo de Texto", "*.txt"), ("Todos os arquivos", "*.*"))
        )
        if path:
            self.entry_saida.delete(0, tk.END)
            self.entry_saida.insert(0, path)

    def atualizar_status(self, mensagem):
        self.status_var.set(mensagem)
        self.root.update_idletasks() # Força a GUI a atualizar o texto

    def processar(self):
        # 1. Obter os caminhos dos campos de entrada
        caminho_servidor = self.entry_servidor.get()
        caminho_conta = self.entry_contas.get()
        caminho_saida = self.entry_saida.get() + ".txt"
        data_pagamento = self.entry_data.get()

        # 2. Validar se os campos não estão vazios
        if not caminho_servidor or not caminho_conta or not caminho_saida:
            messagebox.showwarning("Campos Vazios", "Por favor, selecione todos os três arquivos antes de processar.")
            return
            
        # 3. Desabilitar o botão e chamar a lógica
        self.btn_processar.config(text="Processando...", state=tk.DISABLED)
        self.atualizar_status("Iniciando processamento...")
        
        # Chama a função de lógica
        processar_arquivos(caminho_servidor, caminho_conta, caminho_saida, data_pagamento, self.atualizar_status)
        
        # 4. Reabilitar o botão
        self.btn_processar.config(text="Processar e Salvar Arquivo", state=tk.NORMAL)

# ==============================================================================
#  PASSO 3: INICIAR A APLICAÇÃO
# ==============================================================================

if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()