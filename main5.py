import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

# ==============================================================================
#  PASSO 1: O MODELO DE DADOS PARA AS REGRAS
#  Esta classe representa as configurações para UMA linha no arquivo de saída.
# ==============================================================================
class FieldRule:
    def __init__(self, field_name, source_column, length, padding_char, justification, is_fixed=False, fixed_value=""):
        self.field_name = field_name          # Nome do campo (ex: 'NC02NOMEC46')
        self.source_column = source_column    # Coluna do Excel (ex: 'nome')
        self.length = tk.IntVar(value=length) # Tamanho final
        self.padding_char = tk.StringVar(value=padding_char) # ' ' ou '0'
        self.justification = tk.StringVar(value=justification) # 'Esquerda' ou 'Direita'
        self.is_fixed = tk.BooleanVar(value=is_fixed) # É um valor fixo?
        self.fixed_value = tk.StringVar(value=fixed_value) # O valor fixo em si

# ==============================================================================
#  PASSO 2: O MOTOR DE PROCESSAMENTO (Agora dinâmico)
#  Esta função agora recebe a lista de regras da GUI.
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
        
        with open(caminho_saida, 'w', encoding='utf-8') as f:
            for indice, linha in df_final_ordenado.iterrows():
                linha_formatada_final = []
                
                # --- Regra de negócio da conta 38/39 (ainda fixa) ---
                servidor_encontrado = not pd.isna(linha.get('nome'))
                conta_orig = str(linha.get('conta', '')).strip()
                regra_banco_especial = (not servidor_encontrado) or conta_orig.startswith('39') or conta_orig.startswith('38')

                # --- Motor de Formatação Dinâmico ---
                for rule in rules:
                    # 1. Pega o valor (fixo ou do dataframe)
                    if rule.is_fixed.get():
                        valor_bruto = rule.fixed_value.get()
                    else:
                        # Aplica a regra especial para os campos de banco
                        if regra_banco_especial and rule.source_column in ['banco', 'agencia', 'conta']:
                            if rule.source_column == 'banco': valor_bruto = '041'
                            elif rule.source_column == 'agencia': valor_bruto = '0000'
                            elif rule.source_column == 'conta': valor_bruto = '0'
                        else:
                            valor_bruto = str(linha.get(rule.source_column, ''))

                    # 2. Pega as configurações da GUI
                    tamanho = rule.length.get()
                    preenchimento = rule.padding_char.get()
                    justificativa = rule.justification.get()
                    
                    # Tratamento especial para Salário (PIC 9V99)
                    if rule.source_column == 'salario' and not rule.is_fixed.get():
                        try:
                            salario_val = float(valor_bruto) * 100
                            valor_bruto = str(int(salario_val))
                        except (ValueError, TypeError):
                            valor_bruto = '0'
                            
                    # 3. Aplica a formatação
                    valor_bruto = valor_bruto[:tamanho] # Corta se for maior
                    if justificativa == 'Direita':
                        valor_formatado = valor_bruto.rjust(tamanho, preenchimento)
                    else: # Esquerda
                        valor_formatado = valor_bruto.ljust(tamanho, preenchimento)
                    
                    linha_formatada_final.append(valor_formatado)
                
                # 4. Junta tudo e salva
                f.write("".join(linha_formatada_final) + '\n')

        status_callback(f"Processo concluído! {total_linhas} linhas salvas.")
        messagebox.showinfo("Sucesso", f"Processo concluído!\n{total_linhas} linhas salvas em:\n{caminho_saida}")

    except Exception as e:
        status_callback(f"Erro: {e}")
        messagebox.showerror("Erro", f"Ocorreu um erro:\n{e}")


# ==============================================================================
#  PASSO 3: A INTERFACE GRÁFICA (Agora com o construtor de regras)
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
        # Aqui definimos o nosso layout padrão. A GUI vai ler isso.
        return [
            FieldRule('NOME', 'nome', 46, ' ', 'Esquerda'),
            FieldRule('CPF', 'cpf', 11, '0', 'Direita'),
            FieldRule('BANCO', 'banco', 3, '0', 'Direita'),
            FieldRule('AGENCIA', 'agencia', 4, '0', 'Direita'),
            FieldRule('CONTA', 'conta', 10, '0', 'Direita'),
            FieldRule('MATRICULA', 'matricula', 15, '0', 'Direita'),
            FieldRule('SALARIO', 'salario', 15, '0', 'Direita'),
            FieldRule('VALOR CAL  (13º)', '', 15, '0', 'Direita', is_fixed=True, fixed_value='0'),
            FieldRule('COD OCORRENCIA (BRANCOS*2)', '', 2, '0', 'Direita', is_fixed=True, fixed_value='0'),
            FieldRule('DESC OCORR (BRANCOS*82)', '', 82, ' ', 'Esquerda', is_fixed=True, fixed_value=''),
            FieldRule('DATA AGENC (BRANCOS*8)', '', 8, ' ', 'Esquerda', is_fixed=True, fixed_value=''),
            FieldRule('DATA PAGAMENTO', '', 8, '0', 'Direita', is_fixed=True, fixed_value='20220220'),
            FieldRule('TIPO EMPREGADOR(J)', '', 1, ' ', 'Esquerda', is_fixed=True, fixed_value='J'),
            FieldRule('EMPREGADOR(CNPJ)', '', 14, ' ', 'Esquerda', is_fixed=True, fixed_value=''),
        ]

    def _create_file_widgets(self, parent):
        # Widgets para seleção de arquivos (similar ao anterior)
        # ... (código omitido por brevidade, mas é o mesmo da Fase 1)
        # --- 1. Arquivo de Servidores (dados_gp) ---
        frame_servidor = tk.Frame(parent)
        frame_servidor.pack(fill=tk.X)
        tk.Label(frame_servidor, text="Arquivo Servidores:", width=18, anchor="w").pack(side=tk.LEFT)
        self.entry_servidor = tk.Entry(frame_servidor)
        self.entry_servidor.pack(side=tk.LEFT, fill=tk.X, expand=True)
        tk.Button(frame_servidor, text="Procurar...", command=lambda: self._procurar_arquivo(self.entry_servidor, "Selecione arquivo de servidores")).pack(side=tk.LEFT, padx=(5, 0))

        # --- 2. Arquivo de Contas (retorno_contas) ---
        frame_contas = tk.Frame(parent)
        frame_contas.pack(fill=tk.X, pady=5)
        tk.Label(frame_contas, text="Arquivo Contas:", width=18, anchor="w").pack(side=tk.LEFT)
        self.entry_contas = tk.Entry(frame_contas)
        self.entry_contas.pack(side=tk.LEFT, fill=tk.X, expand=True)
        tk.Button(frame_contas, text="Procurar...", command=lambda: self._procurar_arquivo(self.entry_contas, "Selecione arquivo de contas")).pack(side=tk.LEFT, padx=(5, 0))

        # --- 3. Arquivo de Saída (TXT) ---
        frame_saida = tk.Frame(parent)
        frame_saida.pack(fill=tk.X)
        tk.Label(frame_saida, text="Salvar saída como:", width=18, anchor="w").pack(side=tk.LEFT)
        self.entry_saida = tk.Entry(frame_saida)
        self.entry_saida.pack(side=tk.LEFT, fill=tk.X, expand=True)
        tk.Button(frame_saida, text="Salvar Como...", command=lambda: self._procurar_saida(self.entry_saida)).pack(side=tk.LEFT, padx=(5, 0))


    def _create_rules_header(self):
        header_frame = ttk.Frame(self.scrollable_frame)
        header_frame.grid(row=0, column=0, sticky='ew', pady=(0, 5))
        headers = ['Nome do Campo', 'Valor Fixo?', 'Valor', 'Origem (Coluna XLS)', 'Tamanho', 'Preencher com', 'Alinhar']
        weights = [3, 1, 2, 2, 1, 1, 1]
        for i, (text, weight) in enumerate(zip(headers, weights)):
            header_frame.columnconfigure(i, weight=weight)
            ttk.Label(header_frame, text=text, font=('Helvetica', 10, 'bold')).grid(row=0, column=i, sticky='w')

    def _populate_rules_widgets(self):
        for i, rule in enumerate(self.rules):
            row_frame = ttk.Frame(self.scrollable_frame)
            row_frame.grid(row=i+1, column=0, sticky='ew', pady=2)
            
            weights = [3, 1, 2, 2, 1, 1, 1]
            for c, w in enumerate(weights): row_frame.columnconfigure(c, weight=w)

            # Coluna 0: Nome do Campo
            ttk.Label(row_frame, text=rule.field_name, wraplength=180).grid(row=0, column=0, sticky='w')

            # Coluna 1: Checkbox 'Valor Fixo'
            chk_fixed = ttk.Checkbutton(row_frame, variable=rule.is_fixed)
            chk_fixed.grid(row=0, column=1)

            # Coluna 2: Entry 'Valor Fixo'
            entry_fixed = ttk.Entry(row_frame, textvariable=rule.fixed_value)
            entry_fixed.grid(row=0, column=2, sticky='ew', padx=2)

            # Coluna 3: Entry 'Origem'
            entry_source = ttk.Entry(row_frame, textvariable=rule.source_column)
            entry_source.grid(row=0, column=3, sticky='ew', padx=2)
            
            # Coluna 4: Entry 'Tamanho'
            entry_len = ttk.Entry(row_frame, textvariable=rule.length, width=5)
            entry_len.grid(row=0, column=4, padx=2)
            
            # Coluna 5: OptionMenu 'Preenchimento'
            opt_pad = ttk.OptionMenu(row_frame, rule.padding_char, ' ', ' ', '0')
            opt_pad.config(width=5)
            opt_pad.grid(row=0, column=5, padx=2)

            # Coluna 6: OptionMenu 'Alinhamento'
            opt_just = ttk.OptionMenu(row_frame, rule.justification, 'Esquerda', 'Esquerda', 'Direita')
            opt_just.config(width=8)
            opt_just.grid(row=0, column=6, padx=2)

    def _create_action_widgets(self, parent):
        self.btn_processar = ttk.Button(parent, text="Processar e Salvar Arquivo", command=self.processar)
        self.btn_processar.pack(side=tk.RIGHT, pady=(10,0))
        
        self.status_var = tk.StringVar(value="Pronto.")
        ttk.Label(parent, textvariable=self.status_var, relief=tk.SUNKEN).pack(side=tk.LEFT, fill=tk.X, expand=True, pady=(10,0))

    # --- Funções de Ação ---
    def _procurar_arquivo(self, entry_widget, title):
        path = filedialog.askopenfilename(title=title, filetypes=(("Arquivos Excel", "*.xls *.xlsx"), ("Todos", "*.*")))
        if path:
            entry_widget.delete(0, tk.END)
            entry_widget.insert(0, path)

    def _procurar_saida(self, entry_widget):
        path = filedialog.asksaveasfilename(title="Salvar arquivo de saída", defaultextension=".txt", filetypes=(("Arquivo de Texto", "*.txt"), ("Todos", "*.*")))
        if path:
            entry_widget.delete(0, tk.END)
            entry_widget.insert(0, path)

    def atualizar_status(self, mensagem):
        self.status_var.set(mensagem)
        self.root.update_idletasks()

    def processar(self):
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