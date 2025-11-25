import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os # <-- NOVO: Precisamos disso para gerar o nome do arquivo CSV

# ==============================================================================
#  PASSO 1: O MODELO DE DADOS PARA AS REGRAS (Sem alterações)
# ==============================================================================
class FieldRule:
    def __init__(self, field_name, source_column, length, padding_char, justification, is_fixed=False, fixed_value=""):
        self.field_name = field_name
        self.source_column = source_column
        self.length = tk.IntVar(value=length)
        self.padding_char = tk.StringVar(value=padding_char)
        self.justification = tk.StringVar(value=justification)
        self.is_fixed = tk.BooleanVar(value=is_fixed)
        self.fixed_value = tk.StringVar(value=fixed_value)

# ==============================================================================
#  PASSO 2: O MOTOR DE PROCESSAMENTO (Atualizado para salvar CSV)
# ==============================================================================
def processar_arquivos_dinamico(caminho_servidor, caminho_conta, caminho_saida, rules, status_callback):
    try:
        # --- Lógica de negócio (ainda fixa) ---
        status_callback(f"Lendo '{caminho_conta}'...")
        df_contas = pd.read_excel(caminho_conta, dtype=str)

        status_callback(f"Lendo '{caminho_servidor}'...")
        df_dados = pd.read_excel(caminho_servidor, dtype={'cpf': str})

        status_callback("Filtrando CPFs duplicados...")
        df_dados_ordenado = df_dados.sort_values(by='matricula', ascending=False)
        df_dados_limpo = df_dados_ordenado.drop_duplicates(subset='cpf', keep='first')

        status_callback("Cruzando e ordenando dados...")
        df_final = pd.merge(df_contas, df_dados_limpo, on='cpf', how='left')
        df_final_ordenado = df_final.sort_values(by='nome', ascending=True, na_position='last')
        
        status_callback("Gerando arquivo de saída formatado...")
        total_linhas = len(df_final_ordenado)
        
        validation_data_list = [] # <-- NOVO: Lista para guardar os dados do CSV

        with open(caminho_saida, 'w', encoding='utf-8') as f:
            for indice, linha in df_final_ordenado.iterrows():
                linha_formatada_final = []
                csv_row_data = {} # <-- NOVO: Dicionário para a linha do CSV
                
                # --- Regra de negócio da conta 38/39 (ainda fixa) ---
                servidor_encontrado = not pd.isna(linha.get('nome'))
                conta_orig = str(linha.get('conta', '')).strip()
                regra_banco_especial = (not servidor_encontrado) or conta_orig.startswith('39') or conta_orig.startswith('38')

                # --- Motor de Formatação Dinâmico ---
                for rule in rules:
                    # 1. Pega o valor (fixo ou do dataframe)
                    if rule.is_fixed.get():
                        valor_bruto = rule.fixed_value.get()
                        # NOVO: Salva no dict do CSV
                        csv_row_data[rule.field_name] = valor_bruto
                    else:
                        # Aplica a regra especial para os campos de banco
                        if regra_banco_especial and rule.source_column in ['banco', 'agencia', 'conta']:
                            if rule.source_column == 'banco': valor_bruto = '041'
                            elif rule.source_column == 'agencia': valor_bruto = '0000'
                            elif rule.source_column == 'conta': valor_bruto = '0' # O padding cuida do resto
                        else:
                            valor_bruto = str(linha.get(rule.source_column, ''))

                        # Tratamento especial para Salário (PIC 9V99)
                        if rule.source_column == 'salario':
                            try:
                                salario_float = float(valor_bruto)
                                csv_row_data['salario_decimal'] = salario_float # Salva o valor decimal no CSV
                                valor_bruto = str(int(salario_float * 100)) # Valor para o TXT
                            except (ValueError, TypeError):
                                csv_row_data['salario_decimal'] = 0.0
                                valor_bruto = '0'
                        else:
                            # NOVO: Salva o valor bruto (final) no dict do CSV
                            csv_row_data[rule.source_column] = valor_bruto

                    # 2. Pega as configurações da GUI
                    tamanho = rule.length.get()
                    preenchimento = rule.padding_char.get()
                    justificativa = rule.justification.get()
                    
                    # 3. Aplica a formatação
                    valor_bruto_cortado = valor_bruto[:tamanho] # Corta se for maior
                    if justificativa == 'Direita':
                        valor_formatado = valor_bruto_cortado.rjust(tamanho, preenchimento)
                    else: # Esquerda
                        valor_formatado = valor_bruto_cortado.ljust(tamanho, preenchimento)
                    
                    linha_formatada_final.append(valor_formatado)
                
                # 4. Junta tudo e salva no TXT
                f.write("".join(linha_formatada_final) + '\n')
                validation_data_list.append(csv_row_data) # <-- NOVO: Adiciona a linha de dados à lista do CSV

        # --- FIM DO LOOP PRINCIPAL ---

        # --- NOVO: Salvar o arquivo de validação CSV ---
        if validation_data_list:
            status_callback("Salvando arquivo de validação CSV...")
            df_validation = pd.DataFrame(validation_data_list)
            
            # Define o nome do arquivo CSV (ex: saida.txt -> saida_validacao.csv)
            base_name, _ = os.path.splitext(caminho_saida)
            csv_path = f"{base_name}_validacao.csv"
            
            # Reordena as colunas do CSV para uma leitura mais lógica
            colunas_prioritarias = [
                'cpf', 'nome', 'matricula', 'salario_decimal', 'banco', 'agencia', 'conta'
            ]
            colunas_existentes = list(df_validation.columns)
            colunas_finais_csv = [col for col in colunas_prioritarias if col in colunas_existentes]
            colunas_finais_csv += [col for col in colunas_existentes if col not in colunas_finais_csv]
            
            df_validation = df_validation[colunas_finais_csv]
            
            # Salva em CSV usando ; como separador (melhor para Excel em português)
            df_validation.to_csv(csv_path, index=False, sep=';', encoding='utf-8-sig')
        
        status_callback(f"Processo concluído! {total_linhas} linhas salvas.")
        messagebox.showinfo("Sucesso", 
                            f"Processo concluído!\n\n"
                            f"Arquivo TXT salvo em:\n{caminho_saida}\n\n"
                            f"Arquivo CSV de validação salvo em:\n{csv_path}")

    except Exception as e:
        status_callback(f"Erro: {e}")
        messagebox.showerror("Erro", f"Ocorreu um erro:\n{e}")


# ==============================================================================
#  PASSO 3: A INTERFACE GRÁFICA (Sem alterações)
# ==============================================================================
class App:
    def __init__(self, root):
        self.root = root
        self.root.title("Processador de Arquivos Banrisul - Construtor de Layout")
        self.root.geometry("950x700")

        self.rules = self._create_initial_rules()

        # --- Layout Principal ---
        top_frame = tk.Frame(root, padx=10, pady=10)
        top_frame.pack(fill=tk.X)
        self._create_file_widgets(top_frame)

        # --- Construtor de Regras ---
        rules_container = tk.LabelFrame(root, text="Configuração do Layout de Saída", padx=10, pady=10)
        rules_container.pack(fill=tk.BOTH, expand=True, padx=10)
        
        canvas = tk.Canvas(rules_container)
        scrollbar = ttk.Scrollbar(rules_container, orient="vertical", command=canvas.yview)
        self.scrollable_frame = ttk.Frame(canvas)

        self.scrollable_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        self._create_rules_header()
        self._populate_rules_widgets()

        # --- Botão e Status ---
        bottom_frame = tk.Frame(root, padx=10, pady=10)
        bottom_frame.pack(fill=tk.X)
        self._create_action_widgets(bottom_frame)

    def _create_initial_rules(self):
        # (Sem alterações)
        return [
            FieldRule('NC02NOMEC46', 'nome', 46, ' ', 'Esquerda'),
            FieldRule('NC02CPFC11', 'cpf', 11, '0', 'Direita'),
            FieldRule('NC02BANCOC03', 'banco', 3, '0', 'Direita'),
            FieldRule('NC02AGENCIAC04', 'agencia', 4, '0', 'Direita'),
            FieldRule('NC02CONTAC10', 'conta', 10, '0', 'Direita'),
            FieldRule('NC02NUMFUNCC15', 'matricula', 15, '0', 'Direita'),
            FieldRule('NC02VALORCALCP15 (Salário)', 'salario', 15, '0', 'Direita'),
            FieldRule('NC02VALORCALCP15 (13º)', '', 15, '0', 'Direita', is_fixed=True, fixed_value='0'),
            FieldRule('NC02OCORRP02', '', 2, '0', 'Direita', is_fixed=True, fixed_value='0'),
            FieldRule('NC02DESCROCORRC82', '', 82, ' ', 'Esquerda', is_fixed=True, fixed_value=' '),
            FieldRule('NC02DATAAGENC08', '', 8, ' ', 'Esquerda', is_fixed=True, fixed_value=' '),
            FieldRule('NC02DATAPGTC08', '', 8, '0', 'Direita', is_fixed=True, fixed_value='20220220'),
            FieldRule('NC02TIPOEMPREGC01', '', 1, ' ', 'Esquerda', is_fixed=True, fixed_value='J'),
            FieldRule('NC02EMPREGADORC14', '', 14, ' ', 'Esquerda', is_fixed=True, fixed_value='88131164000107'),
        ]

    def _create_file_widgets(self, parent):
        # (Sem alterações)
        frame_servidor = tk.Frame(parent)
        frame_servidor.pack(fill=tk.X)
        tk.Label(frame_servidor, text="Arquivo Servidores:", width=18, anchor="w").pack(side=tk.LEFT)
        self.entry_servidor = tk.Entry(frame_servidor)
        self.entry_servidor.pack(side=tk.LEFT, fill=tk.X, expand=True)
        tk.Button(frame_servidor, text="Procurar...", command=lambda: self._procurar_arquivo(self.entry_servidor, "Selecione arquivo de servidores")).pack(side=tk.LEFT, padx=(5, 0))

        frame_contas = tk.Frame(parent)
        frame_contas.pack(fill=tk.X, pady=5)
        tk.Label(frame_contas, text="Arquivo Contas:", width=18, anchor="w").pack(side=tk.LEFT)
        self.entry_contas = tk.Entry(frame_contas)
        self.entry_contas.pack(side=tk.LEFT, fill=tk.X, expand=True)
        tk.Button(frame_contas, text="Procurar...", command=lambda: self._procurar_arquivo(self.entry_contas, "Selecione arquivo de contas")).pack(side=tk.LEFT, padx=(5, 0))

        frame_saida = tk.Frame(parent)
        frame_saida.pack(fill=tk.X)
        tk.Label(frame_saida, text="Salvar saída como:", width=18, anchor="w").pack(side=tk.LEFT)
        self.entry_saida = tk.Entry(frame_saida)
        self.entry_saida.pack(side=tk.LEFT, fill=tk.X, expand=True)
        tk.Button(frame_saida, text="Salvar Como...", command=lambda: self._procurar_saida(self.entry_saida)).pack(side=tk.LEFT, padx=(5, 0))


    def _create_rules_header(self):
        # (Sem alterações)
        header_frame = ttk.Frame(self.scrollable_frame)
        header_frame.grid(row=0, column=0, sticky='ew', pady=(0, 5))
        headers = ['Nome do Campo', 'Valor Fixo?', 'Valor', 'Origem (Coluna XLS)', 'Tamanho', 'Preencher com', 'Alinhar']
        weights = [3, 1, 2, 2, 1, 1, 1]
        for i, (text, weight) in enumerate(zip(headers, weights)):
            header_frame.columnconfigure(i, weight=weight)
            ttk.Label(header_frame, text=text, font=('Helvetica', 10, 'bold')).grid(row=0, column=i, sticky='w')

    def _populate_rules_widgets(self):
        # (Sem alterações)
        for i, rule in enumerate(self.rules):
            row_frame = ttk.Frame(self.scrollable_frame)
            row_frame.grid(row=i+1, column=0, sticky='ew', pady=2)
            
            weights = [3, 1, 2, 2, 1, 1, 1]
            for c, w in enumerate(weights): row_frame.columnconfigure(c, weight=w)

            ttk.Label(row_frame, text=rule.field_name, wraplength=180).grid(row=0, column=0, sticky='w')
            chk_fixed = ttk.Checkbutton(row_frame, variable=rule.is_fixed)
            chk_fixed.grid(row=0, column=1)
            entry_fixed = ttk.Entry(row_frame, textvariable=rule.fixed_value)
            entry_fixed.grid(row=0, column=2, sticky='ew', padx=2)
            entry_source = ttk.Entry(row_frame, textvariable=rule.source_column)
            entry_source.grid(row=0, column=3, sticky='ew', padx=2)
            entry_len = ttk.Entry(row_frame, textvariable=rule.length, width=5)
            entry_len.grid(row=0, column=4, padx=2)
            opt_pad = ttk.OptionMenu(row_frame, rule.padding_char, ' ', ' ', '0')
            opt_pad.config(width=5)
            opt_pad.grid(row=0, column=5, padx=2)
            opt_just = ttk.OptionMenu(row_frame, rule.justification, 'Esquerda', 'Esquerda', 'Direita')
            opt_just.config(width=8)
            opt_just.grid(row=0, column=6, padx=2)

    def _create_action_widgets(self, parent):
        # (Sem alterações)
        self.btn_processar = ttk.Button(parent, text="Processar e Salvar Arquivo", command=self.processar)
        self.btn_processar.pack(side=tk.RIGHT, pady=(10,0))
        
        self.status_var = tk.StringVar(value="Pronto.")
        ttk.Label(parent, textvariable=self.status_var, relief=tk.SUNKEN).pack(side=tk.LEFT, fill=tk.X, expand=True, pady=(10,0))

    # --- Funções de Ação ---
    def _procurar_arquivo(self, entry_widget, title):
        # (Sem alterações)
        path = filedialog.askopenfilename(title=title, filetypes=(("Arquivos Excel", "*.xls *.xlsx"), ("Todos", "*.*")))
        if path:
            entry_widget.delete(0, tk.END)
            entry_widget.insert(0, path)

    def _procurar_saida(self, entry_widget):
        # (Sem alterações)
        path = filedialog.asksaveasfilename(title="Salvar arquivo de saída", defaultextension=".txt", filetypes=(("Arquivo de Texto", "*.txt"), ("Todos", "*.*")))
        if path:
            entry_widget.delete(0, tk.END)
            entry_widget.insert(0, path)

    def atualizar_status(self, mensagem):
        # (Sem alterações)
        self.status_var.set(mensagem)
        self.root.update_idletasks()

    def processar(self):
        # (Sem alterações)
        caminho_servidor = self.entry_servidor.get()
        caminho_conta = self.entry_contas.get()
        caminho_saida = self.entry_saida.get()

        if not all([caminho_servidor, caminho_conta, caminho_saida]):
            messagebox.showwarning("Campos Vazios", "Por favor, preencha todos os caminhos de arquivo.")
            return
        
        self.btn_processar.config(state=tk.DISABLED)
        processar_arquivos_dinamico(caminho_servidor, caminho_conta, caminho_saida, self.rules, self.atualizar_status)
        self.btn_processar.config(state=tk.NORMAL)

if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()