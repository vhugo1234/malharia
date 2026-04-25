import customtkinter as ctk
from tkinter import filedialog
import os
import json
import pandas as pd

from core.orchestrator import RenderThread, ExportThread
from core.report_generator import ReportGenerator

# --- IDENTIDADE VISUAL DA MALHARIA (AZUL MARINHO) ---
COLOR_BG = "#0b1325"
COLOR_PRIMARY = "#1A2B4C"
COLOR_ACCENT = "#29467d"
COLOR_SUCCESS = "#00b56b"
COLOR_ALERT = "#ff5555"

class AppMainWindow(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Sistema de Malharia Automática - Dashboard Industrial Completo")
        self.geometry("1400x900")
        self.after(0, lambda: self.state('zoomed'))
        ctk.set_appearance_mode("Dark")
        self.configure(fg_color=COLOR_BG)

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)

        # CABEÇALHO
        self.header_frame = ctk.CTkFrame(self, height=80, fg_color=COLOR_PRIMARY, corner_radius=0)
        self.header_frame.grid(row=0, column=0, sticky="ew")
        ctk.CTkLabel(self.header_frame, text="MALHARIA ESPORTIVA 5.0", font=ctk.CTkFont(size=28, weight="bold"), text_color="white").pack(side="left", padx=30, pady=20)
        
        self.tabview = ctk.CTkTabview(self, fg_color="transparent", segmented_button_selected_color=COLOR_ACCENT)
        self.tabview.grid(row=1, column=0, padx=20, pady=10, sticky="nsew")
        self.tab_monitor = self.tabview.add("📊 Dashboard Fabril")
        self.tab_lote = self.tabview.add("🚀 Renderizador de Lotes")
        self.tab_regras = self.tabview.add("📐 Parametrização Inteligente (Fim do JSON)")
        self.tab_plotter = self.tabview.add("🖨️ Estação Plotter (Export em Massa)")

        self.build_tab_monitor()
        self.build_tab_lote()
        self.build_tab_plotter()
        self.build_tab_regras()
        self.render_thread = None

    def build_tab_monitor(self):
        self.tab_monitor.grid_columnconfigure((0, 1, 2, 3), weight=1)
        self.tab_monitor.grid_rowconfigure(1, weight=1)
        
        c1 = ctk.CTkFrame(self.tab_monitor, fg_color=COLOR_PRIMARY, corner_radius=15, height=150); c1.grid(row=0, column=0, padx=10, pady=20, sticky="ew")
        ctk.CTkLabel(c1, text="Peças Processadas", text_color="#A0B1C4", font=("Arial", 16)).pack(pady=(20, 0))
        self.lbl_pecas = ctk.CTkLabel(c1, text="0", text_color="white", font=ctk.CTkFont(size=48, weight="bold")); self.lbl_pecas.pack(pady=(10, 20))

        c2 = ctk.CTkFrame(self.tab_monitor, fg_color=COLOR_PRIMARY, corner_radius=15, height=150); c2.grid(row=0, column=1, padx=10, pady=20, sticky="ew")
        ctk.CTkLabel(c2, text="Erros Críticos Salvos", text_color="#A0B1C4", font=("Arial", 16)).pack(pady=(20, 0))
        self.lbl_erros = ctk.CTkLabel(c2, text="0", text_color=COLOR_ALERT, font=ctk.CTkFont(size=48, weight="bold")); self.lbl_erros.pack(pady=(10, 20))

        c3 = ctk.CTkFrame(self.tab_monitor, fg_color=COLOR_PRIMARY, corner_radius=15, height=150); c3.grid(row=0, column=2, padx=10, pady=20, sticky="ew")
        ctk.CTkLabel(c3, text="Tempo/Peça (Lead)", text_color="#A0B1C4", font=("Arial", 16)).pack(pady=(20, 0))
        self.lbl_tempo = ctk.CTkLabel(c3, text="0.0s", text_color=COLOR_SUCCESS, font=ctk.CTkFont(size=48, weight="bold")); self.lbl_tempo.pack(pady=(10, 20))
        
        c4 = ctk.CTkFrame(self.tab_monitor, fg_color=COLOR_PRIMARY, corner_radius=15, height=150); c4.grid(row=0, column=3, padx=10, pady=20, sticky="ew")
        ctk.CTkLabel(c4, text="Custo de Tinta Estimado", text_color="#A0B1C4", font=("Arial", 16)).pack(pady=(20, 0))
        self.lbl_custo = ctk.CTkLabel(c4, text="R$ 0,00", text_color="#F39C12", font=ctk.CTkFont(size=48, weight="bold")); self.lbl_custo.pack(pady=(10, 20))

        log_fr = ctk.CTkFrame(self.tab_monitor, fg_color=COLOR_BG, border_width=2, border_color=COLOR_PRIMARY)
        log_fr.grid(row=1, column=0, columnspan=4, padx=10, pady=20, sticky="nsew")
        log_fr.grid_columnconfigure(0, weight=1); log_fr.grid_rowconfigure(1, weight=1)
        
        ctk.CTkLabel(log_fr, text="Histórico de Rastreio Fabril", text_color="white").grid(row=0, column=0, sticky="w", padx=10, pady=10)
        ctk.CTkButton(log_fr, text="💾 Exportar Log", fg_color="#556F96", command=self.export_log).grid(row=0, column=0, sticky="e", padx=10)
        
        self.log_tb = ctk.CTkTextbox(log_fr, fg_color="black", text_color="white", font=("Consolas", 14))
        self.log_tb.grid(row=1, column=0, sticky="nsew", padx=10, pady=(0, 10))
        self.log_tb.tag_config("ERROR", foreground=COLOR_ALERT); self.log_tb.tag_config("SUCCESS", foreground=COLOR_SUCCESS); self.log_tb.tag_config("INFO", foreground="#00a8ff")

    def build_tab_lote(self):
        # Formularios de Entrada
        self.tab_lote.grid_columnconfigure(0, weight=1)
        fr = ctk.CTkFrame(self.tab_lote, fg_color=COLOR_PRIMARY); fr.pack(fill="x", padx=40, pady=40)
        fr.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(fr, text="1. Planilha XLSX:").grid(row=0, column=0, padx=20, pady=20, sticky="w")
        self.inp_excel = ctk.CTkEntry(fr); self.inp_excel.grid(row=0, column=1, sticky="ew", padx=10)
        ctk.CTkButton(fr, text="Procurar...", command=lambda: self._askfile(self.inp_excel, "*.xlsx")).grid(row=0, column=2, padx=20)
        
        ctk.CTkLabel(fr, text="2. Corel Base .CDR:").grid(row=1, column=0, padx=20, pady=20, sticky="w")
        self.inp_corel = ctk.CTkEntry(fr); self.inp_corel.grid(row=1, column=1, sticky="ew", padx=10)
        ctk.CTkButton(fr, text="Procurar...", command=lambda: self._askfile(self.inp_corel, "*.cdr")).grid(row=1, column=2, padx=20)

        ctk.CTkLabel(fr, text="3. Pasta de Saída:").grid(row=2, column=0, padx=20, pady=20, sticky="w")
        self.inp_out = ctk.CTkEntry(fr); self.inp_out.grid(row=2, column=1, sticky="ew", padx=10)
        ctk.CTkButton(fr, text="Procurar...", command=lambda: self._askdir(self.inp_out)).grid(row=2, column=2, padx=20)
        
        self.ck_ia = ctk.CTkCheckBox(fr, text="✨ Limpar Fundos Branco de Escudos (Inteligência Artificial RemBg Neural Offline)", fg_color="#8E44AD", hover_color="#732d91")
        self.ck_ia.grid(row=3, column=0, columnspan=3, pady=(0, 20))

        self.btn_run = ctk.CTkButton(self.tab_lote, text="▶ 1º PASSO: GERAR ARQUIVO MESTRE (ENCAIXE)", font=ctk.CTkFont(size=24, weight="bold"), height=80, fg_color=COLOR_SUCCESS, hover_color="#008f53", command=self.start_production)
        self.btn_run.pack(fill="x", padx=40, pady=20)
        self.pg_bar = ctk.CTkProgressBar(self.tab_lote, progress_color=COLOR_SUCCESS); self.pg_bar.pack(fill="x", padx=40); self.pg_bar.set(0)

    def build_tab_plotter(self):
        # Aba focada apenas em converter CDR revisado para JPEGs Finais
        fr = ctk.CTkFrame(self.tab_plotter, fg_color=COLOR_PRIMARY); fr.pack(fill="x", padx=40, pady=40)
        fr.grid_columnconfigure(1, weight=1)
        
        ctk.CTkLabel(fr, text="Este módulo explode um arquivo Máster CDR (E já revisado nas mangas) e joga todos os JPEG para a Ploter.").grid(row=0, column=0, columnspan=3, pady=20)

        ctk.CTkLabel(fr, text="1. Arquivo LOTE_MASTER.cdr:").grid(row=1, column=0, padx=20, pady=20, sticky="w")
        self.inp_master = ctk.CTkEntry(fr); self.inp_master.grid(row=1, column=1, sticky="ew", padx=10)
        ctk.CTkButton(fr, text="Procurar...", command=lambda: self._askfile(self.inp_master, "*.cdr")).grid(row=1, column=2, padx=20)

        ctk.CTkLabel(fr, text="2. Pasta de Saída (JPGs):").grid(row=2, column=0, padx=20, pady=20, sticky="w")
        self.inp_out_plotter = ctk.CTkEntry(fr); self.inp_out_plotter.grid(row=2, column=1, sticky="ew", padx=10)
        ctk.CTkButton(fr, text="Procurar...", command=lambda: self._askdir(self.inp_out_plotter)).grid(row=2, column=2, padx=20)

        self.btn_explode = ctk.CTkButton(self.tab_plotter, text="▶ 2º PASSO: EXPLODIR LOTE PARA A IMPRESSORA PLOTTER", font=ctk.CTkFont(size=24, weight="bold"), height=80, fg_color="#F39C12", hover_color="#d68910", text_color="white", command=self.start_export_plotter)
        self.btn_explode.pack(fill="x", padx=40, pady=20)

    def build_tab_regras(self):
        header = ctk.CTkFrame(self.tab_regras, fg_color="transparent")
        header.pack(fill="x", padx=40, pady=(20,0))
        ctk.CTkLabel(header, text="Painel Universal Dinâmico (Sem Código JSON)", font=ctk.CTkFont(size=24, weight="bold")).pack(side="left")
        ctk.CTkButton(header, text="📥 Recarregar Configurações Físicas", fg_color=COLOR_ACCENT, command=self.build_dynamic_settings).pack(side="right")
        
        # Repartindo a tela no meio
        body_area = ctk.CTkFrame(self.tab_regras, fg_color="transparent")
        body_area.pack(fill="both", expand=True, padx=40, pady=20)
        
        # Lado Esquerdo: O Auto-Construtor (Para o Operador)
        self.scroll_forms = ctk.CTkScrollableFrame(body_area, fg_color="transparent")
        self.scroll_forms.pack(side="left", fill="both", expand=True, padx=(0, 20))
        
        # Lado Direito: O Motor a Combustão JSON (Para o Dono/Engenheiro)
        right_panel = ctk.CTkFrame(body_area, fg_color="transparent", width=400)
        right_panel.pack(side="right", fill="both")
        
        ctk.CTkLabel(right_panel, text="Modo Engenheiro (Raiz JSON)", font=ctk.CTkFont(size=18, weight="bold"), text_color="#A0B1C4").pack(pady=(0,10))
        self.txt_json = ctk.CTkTextbox(right_panel, fg_color="black", text_color="white", font=("Consolas", 14), width=450)
        self.txt_json.pack(fill="both", expand=True)
        ctk.CTkButton(right_panel, text="💾 FORÇAR SALVAMENTO DO TXT", height=50, fg_color="#C0392B", hover_color="#922b21", command=self.save_raw_json).pack(fill="x", pady=10)
        
        self.dynamic_widgets = {}
        self.base_data = {}
        self.build_dynamic_settings()
        
        # Botao Gigante de baixo (Salva o lado esquerdo)
        self.btn_save_dyn = ctk.CTkButton(self.tab_regras, text="💾 SALVAR FORMULÁRIO DINÂMICO PARA O JSON", font=ctk.CTkFont(size=20, weight="bold"), height=60, fg_color=COLOR_SUCCESS, hover_color="#008f53", command=self.save_dynamic_settings)
        self.btn_save_dyn.pack(fill="x", padx=40, pady=(0,20))

    def save_raw_json(self):
        """ Um override direto: o humano mexeu na tela preta da direita e vai forçar a barra """
        texto_sujo = self.txt_json.get("1.0", "end")
        caminho = os.path.join(os.getcwd(), 'config', 'default_settings.json')
        try:
            # Valida se o cara engoliu uma aspas antes de salvar
            test_json = json.loads(texto_sujo) 
            with open(caminho, 'w', encoding='utf-8') as f:
                json.dump(test_json, f, indent=2)
            self.log_rich("✅ Arq. Mestre Overridden via Texto Raiz. Atualizando Formulário Visual!", "SUCCESS")
            self.build_dynamic_settings()
        except Exception as e:
            self.log_rich(f"🚨 ERRO CRÍTICO DE SINTAXE NO JSON: {e}", "ERROR")

    def build_dynamic_settings(self):
        """ Destrói o código TXT/JSON e constrói dezenas de formulários amigáveis varrendo os dicionários """
        for widget in self.scroll_forms.winfo_children():
            widget.destroy()
        self.dynamic_widgets.clear()
            
        caminho = os.path.join(os.getcwd(), 'config', 'default_settings.json')
        try:
            with open(caminho, 'r', encoding='utf-8') as f:
                self.base_data = json.load(f)
        except:
            self.base_data = {}

        # Criar os formulários agrupados por Catgoria Visualmente
        self._create_category_form("💰 Parâmetros Fiscais e Fábrica", "FINANCEIRO", self.base_data.get("FINANCEIRO", {}))
        self._create_category_form("👕 Tamanhos dos LOGOS / ESCUDOS", "LOGO_ESCUDO", self.base_data.get("LOGO_ESCUDO", {}))
        self._create_category_form("🔢 Tamanhos das FONTES DOS NÚMEROS", "NUMERO", self.base_data.get("NUMERO", {}))
        self._create_category_form("⚙️ Meta Dados Estruturais", "METADATA", self.base_data.get("METADATA", {}))
        
        # Retro-Sincroniza o painel Raiz também!
        self.txt_json.delete("1.0", "end")
        self.txt_json.insert("1.0", json.dumps(self.base_data, indent=2))
        
    def _create_category_form(self, group_title, dict_key, dict_data):
        frame = ctk.CTkFrame(self.scroll_forms, fg_color=COLOR_PRIMARY, corner_radius=10)
        frame.pack(fill="x", pady=15)
        
        header = ctk.CTkFrame(frame, fg_color="#121e33")
        header.pack(fill="x")
        ctk.CTkLabel(header, text=group_title, font=ctk.CTkFont(size=18, weight="bold"), text_color="#F39C12").pack(side="left", padx=20, pady=10)
        
        self.dynamic_widgets[dict_key] = {}
        
        # Itera TUDO que for achado no dicionario e transforma em inputs visuais limpos
        for sub_key, sub_val in dict_data.items():
            row_frame = ctk.CTkFrame(frame, fg_color="transparent")
            row_frame.pack(fill="x", padx=20, pady=5)
            ctk.CTkLabel(row_frame, text=f"Parâmetro: {sub_key.upper()}", font=ctk.CTkFont(weight="bold"), width=200, anchor="w").pack(side="left")
            
            if isinstance(sub_val, dict):
                # Se for um Dicionario (Exemplo: Tamanho P tem MAX_HEIGHT = 80 e MAX_WIDTH = 50)
                self.dynamic_widgets[dict_key][sub_key] = {}
                for att_name, att_val in sub_val.items():
                    ctk.CTkLabel(row_frame, text=f"{att_name}:").pack(side="left", padx=(10, 5))
                    inp = ctk.CTkEntry(row_frame, fg_color=COLOR_BG, width=100)
                    inp.insert(0, str(att_val))
                    inp.pack(side="left", padx=5)
                    self.dynamic_widgets[dict_key][sub_key][att_name] = inp
            else:
                # String, float nativos brutos (Ex: Custo Papel = 0.5)
                inp = ctk.CTkEntry(row_frame, fg_color=COLOR_BG, width=300)
                inp.insert(0, str(sub_val))
                inp.pack(side="left", padx=10)
                self.dynamic_widgets[dict_key][sub_key] = inp

    def save_dynamic_settings(self):
        """ Recolhe os textos digitados pelos humanos e remonta o JSON 100% blindado contra aspas erradas """
        for top_key, dict_group in self.dynamic_widgets.items():
            for sub_key, sub_widget_or_dict in dict_group.items():
                if isinstance(sub_widget_or_dict, dict):
                    # Loopa atributos da Camisa (size, width)
                    for att_name, ctk_entry in sub_widget_or_dict.items():
                        val = ctk_entry.get()
                        # Cast pra int/float se for numerico
                        if str(val).replace('.','',1).isdigit():
                            val = float(val) if '.' in str(val) else int(val)
                        self.base_data[top_key][sub_key][att_name] = val
                else:
                    # Direto
                    val = sub_widget_or_dict.get()
                    if str(val).replace('.','',1).isdigit():
                        val = float(val) if '.' in str(val) else int(val)
                    self.base_data[top_key][sub_key] = val
                    
        caminho = os.path.join(os.getcwd(), 'config', 'default_settings.json')
        with open(caminho, 'w', encoding='utf-8') as f:
            json.dump(self.base_data, f, indent=2)
            
        self.log_rich("✅ Arquivo Cérebro atualizado visualmente!", "SUCCESS")
        
        # Sicroniza a caixinha do hacker
        self.txt_json.delete("1.0", "end")
        self.txt_json.insert("1.0", json.dumps(self.base_data, indent=2))

    def _askfile(self, entry, filetype):
        path = filedialog.askopenfilename(filetypes=[("Files", filetype)])
        if path: entry.delete(0, "end"); entry.insert(0, path)
    def _askdir(self, entry):
        path = filedialog.askdirectory()
        if path: entry.delete(0, "end"); entry.insert(0, path)
        
    def export_log(self):
        p = filedialog.asksaveasfilename(defaultextension=".txt", filetypes=[("Log", "*.txt")])
        if p:
            open(p, "w", encoding="utf-8").write(self.log_tb.get("1.0", "end"))
            self.log_rich(f"Logs Salvos em {p}", "SUCCESS")
    def log_rich(self, m, t="INFO"):
        self.log_tb.insert("end", m + "\n", t); self.log_tb.see("end")

    def start_production(self):
        self.log_rich("[INFO] Calculadora de Custos e Fila Acionadas...", "INFO")
        self.tabview.set("📊 Dashboard Fabril")
        self.btn_run.configure(state="disabled")

        # Reset contadores do Dashboard
        self.lbl_pecas.configure(text="0")
        self.lbl_erros.configure(text="0")

        # Lê planilha com validação de colunas
        try:
             df = pd.read_excel(self.inp_excel.get())
             colunas_obrigatorias = ["NOME", "TAMANHO", "NUMERO"]
             faltando = [c for c in colunas_obrigatorias if c not in df.columns]
             if faltando:
                 self.log_rich(f"🚨 PLANILHA SEM COLUNAS OBRIGATÓRIAS: {', '.join(faltando)}", "ERROR")
                 self.log_rich(f"   Colunas encontradas: {list(df.columns)}", "ERROR")
                 self.btn_run.configure(state="normal")
                 return
             dl = df.to_dict('records')
             self.log_rich(f"[INFO] Planilha carregada: {len(dl)} peças encontradas.", "INFO")
        except Exception as e_excel:
             self.log_rich(f"🚨 ERRO AO LER PLANILHA: {e_excel}", "ERROR")
             self.btn_run.configure(state="normal")
             return
        
        template_val = self.inp_corel.get().strip()
        out_val = self.inp_out.get().strip()
        financ_dict = self.base_data.get("FINANCEIRO", {})
        
        self.render_thread = RenderThread(
            ui_callback=self.update_progress, 
            data_list=dl, 
            template_path=template_val, 
            output_dir=out_val,
            finance_data=financ_dict,
            usar_ia=self.ck_ia.get()
        )
        self.render_thread.start()

    def start_export_plotter(self):
        self.log_rich("[INFO] Acionando Explosão Final para Maquina Plotter...", "INFO")
        self.tabview.set("📊 Dashboard Fabril")
        self.btn_explode.configure(state="disabled")
        
        m_path = self.inp_master.get().strip()
        o_path = self.inp_out_plotter.get().strip()
        
        self.export_thread = ExportThread(ui_callback=self.update_progress, master_cdr_path=m_path, output_dir=o_path)
        self.export_thread.start()

    def update_progress(self, msg, val):
        """ Wrapper thread-safe: redireciona para a thread principal do Tkinter via after() """
        self.after(0, self._update_progress_safe, msg, val)

    def _update_progress_safe(self, msg, val):
        """ Executa na thread principal — seguro para tocar em widgets Tkinter """
        # Detecta mensagens de controle para atualizar os contadores do dashboard
        if "Operando clonagem" in msg:
            atual = int(self.lbl_pecas.cget("text") or 0)
            self.lbl_pecas.configure(text=str(atual + 1))
        if "⚠️ Erro" in msg or "[FALHA]" in msg or "[ERRO]" in msg:
            atual_err = int(self.lbl_erros.cget("text") or 0)
            self.lbl_erros.configure(text=str(atual_err + 1))

        # Tag automática por conteúdo da mensagem
        tag = "INFO"
        if any(x in msg for x in ["⚠️", "[FALHA]", "[ERRO]", "🚨"]):
            tag = "ERROR"
        elif any(x in msg for x in ["100%", "✅", "Finalizado"]):
            tag = "SUCCESS"

        self.log_rich(msg, tag)
        self.pg_bar.set(val)

        if val >= 1.0:
            self.btn_run.configure(state="normal")
            try: self.btn_explode.configure(state="normal")
            except: pass
            self.log_rich("✅ FIM DA EXECUÇÃO DAS PISTAS!", "SUCCESS")
