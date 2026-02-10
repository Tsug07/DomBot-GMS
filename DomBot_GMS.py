import customtkinter as ctk
import pandas as pd
from pywinauto.application import Application
from pywinauto.keyboard import send_keys
from pywinauto import findwindows, timings
import win32gui
import win32con
import time
import logging
from datetime import datetime
import os
import traceback
import threading
from typing import Optional, Tuple
import tkinter.messagebox as messagebox
from PIL import Image, ImageDraw


# Handler de log separado da classe principal
class GUILogHandler(logging.Handler):
    def __init__(self, gui):
        super().__init__()
        self.gui = gui

    def emit(self, record):
        msg = self.format(record)
        self.gui.window.after(0, lambda: self.gui.adicionar_log(msg, record.levelno))


class AutomacaoGUI:
    # Cores do tema
    CORES = {
        'sucesso': '#2ECC71',
        'erro': '#E74C3C',
        'aviso': '#F39C12',
        'info': '#3498DB',
        'texto': '#ECF0F1',
        'fundo_card': '#2C3E50',
        'fundo_escuro': '#1A252F',
        'destaque': '#1ABC9C',
        'processando': '#9B59B6',
    }

    def __init__(self):
        # Configuração do tema
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("green")

        self.window = ctk.CTk()
        self.window.title("DomBot - Taxa GMS v2.0")
        self.window.geometry("800x550")
        self.window.minsize(750, 500)
        self.window.protocol("WM_DELETE_WINDOW", self.ao_fechar)

        # Flags para controle de execução
        self.executando = False
        self.pausa_solicitada = False
        self.thread_automacao = None

        # Estatísticas
        self.stats = {
            'processados': 0,
            'sucesso': 0,
            'erros': 0,
            'puladas': 0,
            'tempo_inicio': None
        }

        # DataFrame carregado
        self.df_carregado = None

        # Configurar ícone
        self.set_window_icon()

        # Criar diretório de logs se não existir
        self.logs_dir = os.path.join(os.path.dirname(__file__), "logs")
        if not os.path.exists(self.logs_dir):
            os.makedirs(self.logs_dir)

        # Configurar logging para arquivos
        self.setup_file_logging()

        # Variáveis da interface
        self.arquivo_excel = ctk.StringVar()
        self.linha_inicial = ctk.StringVar(value="2")
        self.status_var = ctk.StringVar(value="Aguardando início...")

        # Variáveis de controle (mantidas para compatibilidade com DominioAutomation)
        self.total_linhas = 0
        self.linhas_processadas = 0
        self.linhas_com_erro = 0
        self.linhas_puladas = 0

        # Logger
        self.logger = logging.getLogger('AutomacaoDominio')
        self.logger.setLevel(logging.INFO)
        self.logger.handlers = []

        # Adicionar GUIHandler
        self.gui_handler = GUILogHandler(self)
        formatter = logging.Formatter('%(message)s')
        self.gui_handler.setFormatter(formatter)
        self.logger.addHandler(self.gui_handler)

        self.criar_interface()

    def setup_file_logging(self):
        """Configura o logging para arquivos"""
        data_atual = datetime.now().strftime("%Y-%m-%d")

        # Logger de sucesso
        self.success_logger = logging.getLogger('SuccessLog')
        self.success_logger.setLevel(logging.INFO)
        if not self.success_logger.handlers:
            success_handler = logging.FileHandler(
                os.path.join(self.logs_dir, f'success_{data_atual}.log'),
                encoding='utf-8'
            )
            success_handler.setFormatter(
                logging.Formatter('%(asctime)s - %(message)s', '%Y-%m-%d %H:%M:%S')
            )
            self.success_logger.addHandler(success_handler)

        # Logger de erro
        self.error_logger = logging.getLogger('ErrorLog')
        self.error_logger.setLevel(logging.ERROR)
        if not self.error_logger.handlers:
            error_handler = logging.FileHandler(
                os.path.join(self.logs_dir, f'error_{data_atual}.log'),
                encoding='utf-8'
            )
            error_handler.setFormatter(
                logging.Formatter('%(asctime)s - %(message)s', '%Y-%m-%d %H:%M:%S')
            )
            self.error_logger.addHandler(error_handler)

    def set_window_icon(self):
        """Configura o ícone da janela"""
        try:
            icon_path = os.path.join(os.path.dirname(__file__), "assets", "favicon.ico")
            if os.name == 'nt' and os.path.exists(icon_path):
                self.window.iconbitmap(icon_path)
        except Exception as e:
            print(f"Erro ao carregar ícone: {e}")

    def criar_interface(self):
        # Frame principal com grid
        self.window.grid_columnconfigure(0, weight=1)
        self.window.grid_rowconfigure(0, weight=1)

        main_frame = ctk.CTkFrame(self.window, fg_color="transparent")
        main_frame.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
        main_frame.grid_columnconfigure(0, weight=1)
        main_frame.grid_rowconfigure(3, weight=1)

        # === HEADER ===
        self.criar_header(main_frame)

        # === PAINEL DE CONFIGURAÇÃO ===
        self.criar_painel_config(main_frame)

        # === PAINEL DE ESTATÍSTICAS ===
        self.criar_painel_estatisticas(main_frame)

        # === ÁREA DE CONTEÚDO (Abas) ===
        self.criar_area_conteudo(main_frame)

    def criar_header(self, parent):
        """Cria o cabeçalho com título e status"""
        header_frame = ctk.CTkFrame(parent, fg_color=self.CORES['fundo_card'], corner_radius=8)
        header_frame.grid(row=0, column=0, sticky="ew", pady=(0, 6))
        header_frame.grid_columnconfigure(1, weight=1)

        # Ícone/Logo com fundo branco circular
        logo_path = os.path.join(os.path.dirname(__file__), "assets", "DomBot_New.png")
        if os.path.exists(logo_path):
            size = 66
            circle_size = 44
            # Criar canvas transparente no tamanho total
            bg = Image.new("RGBA", (size, size), (0, 0, 0, 0))
            # Criar círculo branco de 44px centralizado
            circle_mask = Image.new("L", (circle_size, circle_size), 0)
            ImageDraw.Draw(circle_mask).ellipse((0, 0, circle_size - 1, circle_size - 1), fill=255)
            circle = Image.new("RGBA", (circle_size, circle_size), (255, 255, 255, 255))
            circle_offset = (size - circle_size) // 2
            bg.paste(circle, (circle_offset, circle_offset), circle_mask)
            # Colar a logo no tamanho total por cima
            original = Image.open(logo_path).convert("RGBA")
            original = original.resize((size, size), Image.LANCZOS)
            bg.paste(original, (0, 0), original)
            logo_image = ctk.CTkImage(light_image=bg, dark_image=bg, size=(size, size))
            ctk.CTkLabel(header_frame, image=logo_image, text="").grid(row=0, column=0, padx=10, pady=8)
        else:
            logo_frame = ctk.CTkFrame(header_frame, fg_color=self.CORES['destaque'],
                                       width=44, height=44, corner_radius=22)
            logo_frame.grid(row=0, column=0, padx=10, pady=8)
            logo_frame.grid_propagate(False)
            ctk.CTkLabel(logo_frame, text="🤖", font=("Segoe UI Emoji", 18)).place(relx=0.5, rely=0.5, anchor="center")

        # Título
        ctk.CTkLabel(
            header_frame,
            text="DomBot - GMS",
            font=ctk.CTkFont(size=16, weight="bold"),
            text_color=self.CORES['texto']
        ).grid(row=0, column=1, sticky="w", padx=5)

        # Status indicator
        self.status_frame = ctk.CTkFrame(header_frame, fg_color="transparent")
        self.status_frame.grid(row=0, column=2, padx=10)

        self.status_indicator = ctk.CTkFrame(
            self.status_frame,
            fg_color="#7F8C8D",
            width=10, height=10,
            corner_radius=5
        )
        self.status_indicator.pack(side="left", padx=(0, 6))

        self.status_label = ctk.CTkLabel(
            self.status_frame,
            textvariable=self.status_var,
            font=ctk.CTkFont(size=11),
            text_color="#95A5A6"
        )
        self.status_label.pack(side="left")

    def criar_painel_config(self, parent):
        """Cria o painel de configuração"""
        config_frame = ctk.CTkFrame(parent, fg_color=self.CORES['fundo_card'], corner_radius=8)
        config_frame.grid(row=1, column=0, sticky="ew", pady=(0, 6))
        config_frame.grid_columnconfigure(1, weight=1)

        # Linha única com tudo
        inner_frame = ctk.CTkFrame(config_frame, fg_color="transparent")
        inner_frame.grid(row=0, column=0, sticky="ew", padx=10, pady=8)
        inner_frame.grid_columnconfigure(1, weight=1)

        # Arquivo Excel
        ctk.CTkLabel(
            inner_frame, text="📁", font=ctk.CTkFont(size=14)
        ).grid(row=0, column=0, padx=(0, 5))

        self.entry_arquivo = ctk.CTkEntry(
            inner_frame,
            textvariable=self.arquivo_excel,
            placeholder_text="Selecione o arquivo Excel...",
            height=32,
            font=ctk.CTkFont(size=11)
        )
        self.entry_arquivo.grid(row=0, column=1, sticky="ew", padx=(0, 8))

        ctk.CTkButton(
            inner_frame, text="Procurar", command=self.selecionar_arquivo,
            width=80, height=32, font=ctk.CTkFont(size=11),
            fg_color=self.CORES['info'], hover_color="#2980B9"
        ).grid(row=0, column=2, padx=(0, 15))

        # Linha inicial
        ctk.CTkLabel(
            inner_frame, text="Linha:", font=ctk.CTkFont(size=11), text_color="#BDC3C7"
        ).grid(row=0, column=3, padx=(0, 3))

        self.entry_linha = ctk.CTkEntry(
            inner_frame, textvariable=self.linha_inicial,
            width=50, height=32, font=ctk.CTkFont(size=11), justify="center"
        )
        self.entry_linha.grid(row=0, column=4, padx=(0, 15))

        # Botões de controle
        self.btn_iniciar = ctk.CTkButton(
            inner_frame, text="▶ Iniciar", command=self.iniciar_automacao_thread,
            width=90, height=32, font=ctk.CTkFont(size=11, weight="bold"),
            fg_color=self.CORES['sucesso'], hover_color="#27AE60"
        )
        self.btn_iniciar.grid(row=0, column=5, padx=3)

        self.btn_pausar = ctk.CTkButton(
            inner_frame, text="⏸ Pausar", command=self.pausar_automacao,
            width=90, height=32, font=ctk.CTkFont(size=11, weight="bold"),
            fg_color=self.CORES['aviso'], hover_color="#E67E22", state="disabled"
        )
        self.btn_pausar.grid(row=0, column=6, padx=3)

        self.btn_parar = ctk.CTkButton(
            inner_frame, text="⏹ Parar", command=self.parar_automacao,
            width=90, height=32, font=ctk.CTkFont(size=11, weight="bold"),
            fg_color=self.CORES['erro'], hover_color="#C0392B", state="disabled"
        )
        self.btn_parar.grid(row=0, column=7, padx=(3, 0))

    def criar_painel_estatisticas(self, parent):
        """Cria o painel de estatísticas"""
        stats_frame = ctk.CTkFrame(parent, fg_color=self.CORES['fundo_card'], corner_radius=8)
        stats_frame.grid(row=2, column=0, sticky="ew", pady=(0, 6))

        # Grid para os cards de estatísticas
        for i in range(5):
            stats_frame.grid_columnconfigure(i, weight=1)

        # Cards de estatísticas
        self.criar_stat_card(stats_frame, 0, "📊", "Total", "total_label", "0")
        self.criar_stat_card(stats_frame, 1, "✅", "Sucesso", "sucesso_label", "0", self.CORES['sucesso'])
        self.criar_stat_card(stats_frame, 2, "❌", "Erros", "erros_label", "0", self.CORES['erro'])
        self.criar_stat_card(stats_frame, 3, "🏢", "Empresa", "empresa_label", "-", self.CORES['info'])
        self.criar_stat_card(stats_frame, 4, "⏱", "Tempo", "tempo_label", "00:00:00", self.CORES['aviso'])

        # Barra de progresso
        progress_frame = ctk.CTkFrame(stats_frame, fg_color="transparent")
        progress_frame.grid(row=1, column=0, columnspan=5, sticky="ew", padx=10, pady=(2, 8))
        progress_frame.grid_columnconfigure(0, weight=1)

        self.progress_bar = ctk.CTkProgressBar(
            progress_frame, height=6, corner_radius=3, progress_color=self.CORES['destaque']
        )
        self.progress_bar.grid(row=0, column=0, sticky="ew")
        self.progress_bar.set(0)

        self.progress_label = ctk.CTkLabel(
            progress_frame, text="0%", font=ctk.CTkFont(size=10), text_color="#95A5A6"
        )
        self.progress_label.grid(row=0, column=1, padx=(8, 0))

    def criar_stat_card(self, parent, col, icon, titulo, attr_name, valor_inicial, cor=None):
        """Cria um card de estatística"""
        card = ctk.CTkFrame(parent, fg_color="transparent")
        card.grid(row=0, column=col, padx=5, pady=8)

        ctk.CTkLabel(
            card, text=f"{icon} {titulo}", font=ctk.CTkFont(size=10), text_color="#7F8C8D"
        ).pack()

        label = ctk.CTkLabel(
            card, text=valor_inicial, font=ctk.CTkFont(size=14, weight="bold"),
            text_color=cor if cor else self.CORES['texto']
        )
        label.pack()

        setattr(self, attr_name, label)

    def criar_area_conteudo(self, parent):
        """Cria a área de conteúdo com abas"""
        self.tabview = ctk.CTkTabview(
            parent, fg_color=self.CORES['fundo_card'],
            segmented_button_fg_color=self.CORES['fundo_escuro'],
            segmented_button_selected_color=self.CORES['destaque'],
            corner_radius=8, height=25
        )
        self.tabview.grid(row=3, column=0, sticky="nsew")

        tab_logs = self.tabview.add("📋 Logs")
        tab_preview = self.tabview.add("📊 Preview")

        self.criar_aba_logs(tab_logs)
        self.criar_aba_preview(tab_preview)

    def criar_aba_logs(self, parent):
        """Cria a aba de logs"""
        parent.grid_columnconfigure(0, weight=1)
        parent.grid_rowconfigure(0, weight=1)

        log_container = ctk.CTkFrame(parent, fg_color="transparent")
        log_container.grid(row=0, column=0, sticky="nsew", padx=3, pady=3)
        log_container.grid_columnconfigure(0, weight=1)
        log_container.grid_rowconfigure(0, weight=1)

        self.log_text = ctk.CTkTextbox(
            log_container, font=ctk.CTkFont(family="Consolas", size=11),
            fg_color=self.CORES['fundo_escuro'], corner_radius=6
        )
        self.log_text.grid(row=0, column=0, sticky="nsew")

        # Configurar tags de cores
        self.log_text._textbox.tag_config("sucesso", foreground=self.CORES['sucesso'])
        self.log_text._textbox.tag_config("erro", foreground=self.CORES['erro'])
        self.log_text._textbox.tag_config("aviso", foreground=self.CORES['aviso'])
        self.log_text._textbox.tag_config("info", foreground=self.CORES['info'])
        self.log_text._textbox.tag_config("processando", foreground=self.CORES['processando'])

        # Botões de controle do log
        btn_frame = ctk.CTkFrame(log_container, fg_color="transparent")
        btn_frame.grid(row=1, column=0, sticky="ew", pady=(5, 0))

        ctk.CTkButton(
            btn_frame, text="🗑 Limpar", command=self.limpar_logs,
            width=90, height=26, font=ctk.CTkFont(size=10),
            fg_color="#34495E", hover_color="#2C3E50"
        ).pack(side="left")

        ctk.CTkButton(
            btn_frame, text="💾 Exportar", command=self.exportar_logs,
            width=90, height=26, font=ctk.CTkFont(size=10),
            fg_color="#34495E", hover_color="#2C3E50"
        ).pack(side="left", padx=8)

    def criar_aba_preview(self, parent):
        """Cria a aba de preview do Excel"""
        parent.grid_columnconfigure(0, weight=1)
        parent.grid_rowconfigure(1, weight=1)

        info_frame = ctk.CTkFrame(parent, fg_color="transparent")
        info_frame.grid(row=0, column=0, sticky="ew", padx=3, pady=3)

        self.preview_info_label = ctk.CTkLabel(
            info_frame, text="Nenhum arquivo carregado",
            font=ctk.CTkFont(size=11), text_color="#95A5A6"
        )
        self.preview_info_label.pack(side="left")

        ctk.CTkButton(
            info_frame, text="🔄 Recarregar", command=self.carregar_preview,
            width=85, height=24, font=ctk.CTkFont(size=10),
            fg_color="#34495E", hover_color="#2C3E50"
        ).pack(side="right")

        self.preview_text = ctk.CTkTextbox(
            parent, font=ctk.CTkFont(family="Consolas", size=10),
            fg_color=self.CORES['fundo_escuro'], corner_radius=6
        )
        self.preview_text.grid(row=1, column=0, sticky="nsew", padx=3, pady=(0, 3))

    def selecionar_arquivo(self):
        """Abre diálogo para selecionar arquivo Excel"""
        filename = ctk.filedialog.askopenfilename(
            filetypes=[("Excel files", "*.xlsx *.xls")],
            title="Selecione o arquivo Excel"
        )
        if filename:
            self.arquivo_excel.set(filename)
            self.adicionar_log(f"Arquivo selecionado: {os.path.basename(filename)}", logging.INFO, "info")
            self.carregar_preview()

    def carregar_preview(self):
        """Carrega preview do arquivo Excel"""
        if not self.arquivo_excel.get():
            return

        try:
            self.df_carregado = pd.read_excel(self.arquivo_excel.get())
            total_linhas = len(self.df_carregado)

            # Atualizar info
            self.preview_info_label.configure(
                text=f"📄 {os.path.basename(self.arquivo_excel.get())} | {total_linhas} linhas | Colunas: {', '.join(self.df_carregado.columns[:5])}..."
            )

            # Atualizar estatística de total
            self.total_label.configure(text=str(total_linhas))

            # Mostrar preview
            self.preview_text.delete("1.0", "end")

            # Cabeçalho
            header = " | ".join([f"{col:^15}" for col in self.df_carregado.columns[:6]])
            self.preview_text.insert("end", f"{'─' * len(header)}\n")
            self.preview_text.insert("end", f"{header}\n")
            self.preview_text.insert("end", f"{'─' * len(header)}\n")

            # Dados (primeiras 50 linhas)
            for idx, row in self.df_carregado.head(50).iterrows():
                row_text = " | ".join([f"{str(val)[:15]:^15}" for val in row.values[:6]])
                self.preview_text.insert("end", f"{row_text}\n")

            if total_linhas > 50:
                self.preview_text.insert("end", f"\n... e mais {total_linhas - 50} linhas\n")

            # Validar colunas necessárias
            colunas_necessarias = ['Nº', 'Periodo', 'Salvar Como']
            colunas_faltando = [col for col in colunas_necessarias if col not in self.df_carregado.columns]

            if colunas_faltando:
                self.adicionar_log(f"Colunas obrigatórias não encontradas: {', '.join(colunas_faltando)}", logging.WARNING, "aviso")
            else:
                self.adicionar_log(f"Preview carregado: {total_linhas} linhas. Todas as colunas obrigatórias encontradas", logging.INFO, "sucesso")

        except Exception as e:
            self.adicionar_log(f"Erro ao carregar preview: {str(e)}", logging.ERROR, "erro")

    def limpar_logs(self):
        """Limpa a área de logs"""
        self.log_text.delete("1.0", "end")
        self.adicionar_log("Log limpo", logging.INFO, "info")

    def exportar_logs(self):
        """Exporta logs para arquivo"""
        try:
            filename = ctk.filedialog.asksaveasfilename(
                defaultextension=".txt",
                filetypes=[("Text files", "*.txt"), ("All files", "*.*")],
                initialfilename=f"logs_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
            )
            if filename:
                with open(filename, 'w', encoding='utf-8') as f:
                    f.write(self.log_text.get("1.0", "end"))
                self.adicionar_log(f"Logs exportados para: {filename}", logging.INFO, "sucesso")
        except Exception as e:
            self.adicionar_log(f"Erro ao exportar logs: {str(e)}", logging.ERROR, "erro")

    def atualizar_progresso(self, atual, total):
        """Atualiza a barra de progresso"""
        porcentagem = atual / total if total > 0 else 0
        self.progress_bar.set(porcentagem)
        self.progress_label.configure(text=f"{porcentagem * 100:.1f}%")
        self.status_var.set(f"Processando: {atual}/{total}")
        self.window.update_idletasks()

    def atualizar_estatisticas(self):
        """Atualiza os cards de estatísticas"""
        self.sucesso_label.configure(text=str(self.linhas_processadas))
        self.erros_label.configure(text=str(self.linhas_com_erro))
        self.stats['processados'] = self.linhas_processadas + self.linhas_com_erro

    def atualizar_tempo(self):
        """Atualiza o tempo decorrido"""
        if self.stats['tempo_inicio'] and self.executando:
            elapsed = datetime.now() - self.stats['tempo_inicio']
            hours, remainder = divmod(int(elapsed.total_seconds()), 3600)
            minutes, seconds = divmod(remainder, 60)
            self.tempo_label.configure(text=f"{hours:02d}:{minutes:02d}:{seconds:02d}")
            self.window.after(1000, self.atualizar_tempo)

    def atualizar_status_indicator(self, status):
        """Atualiza o indicador de status visual"""
        cores = {
            'aguardando': '#7F8C8D',
            'executando': self.CORES['sucesso'],
            'pausado': self.CORES['aviso'],
            'erro': self.CORES['erro'],
            'concluido': self.CORES['info']
        }
        self.status_indicator.configure(fg_color=cores.get(status, '#7F8C8D'))

    def adicionar_log(self, mensagem, level=logging.INFO, tag=None):
        """Adiciona mensagem ao log visual com cores"""
        try:
            timestamp = datetime.now().strftime('%H:%M:%S')

            # Determinar tag baseado no nível se não especificado
            if tag is None:
                if level >= logging.ERROR:
                    tag = "erro"
                elif level >= logging.WARNING:
                    tag = "aviso"
                elif "sucesso" in mensagem.lower() or "processad" in mensagem.lower():
                    tag = "sucesso"
                else:
                    tag = "info"

            # Prefixo visual
            prefixos = {
                "sucesso": "✅",
                "erro": "❌",
                "aviso": "⚠️",
                "info": "ℹ️",
                "processando": "⏳"
            }
            prefixo = prefixos.get(tag, "•")

            # Inserir mensagem
            self.log_text.insert("end", f"[{timestamp}] {prefixo} ", tag)
            self.log_text.insert("end", f"{mensagem}\n", tag)
            self.log_text.see("end")
            self.window.update_idletasks()
        except Exception:
            pass

    def validar_entrada(self) -> Tuple[bool, str]:
        """Valida os dados de entrada"""
        if not self.arquivo_excel.get():
            return False, "Selecione um arquivo Excel"

        if not os.path.exists(self.arquivo_excel.get()):
            return False, "Arquivo Excel não encontrado"

        try:
            linha_inicial = int(self.linha_inicial.get())
            if linha_inicial < 1:
                return False, "Linha inicial deve ser maior que 0"
        except ValueError:
            return False, "Linha inicial deve ser um número válido"

        # Validar se o arquivo pode ser lido
        try:
            df = pd.read_excel(self.arquivo_excel.get())
            if len(df) == 0:
                return False, "Arquivo Excel está vazio"

            if linha_inicial > len(df) + 1:  # +1 porque linha 1 é cabeçalho
                return False, f"Linha inicial ({linha_inicial}) é maior que o total de linhas do arquivo ({len(df) + 1})"

            # Verificar colunas obrigatórias
            colunas_necessarias = ['Nº', 'Periodo', 'Salvar Como']
            colunas_faltando = [col for col in colunas_necessarias if col not in df.columns]

            if colunas_faltando:
                return False, f"Colunas obrigatórias não encontradas: {', '.join(colunas_faltando)}"

        except Exception as e:
            return False, f"Erro ao ler arquivo Excel: {str(e)}"

        return True, "Validação OK"

    def iniciar_automacao_thread(self):
        """Inicia a automação em uma thread separada"""
        if self.executando:
            self.adicionar_log("Automação já em execução", logging.WARNING, "aviso")
            return

        # Validar entrada
        valido, mensagem = self.validar_entrada()
        if not valido:
            self.adicionar_log(f"Erro de validação: {mensagem}", logging.ERROR, "erro")
            messagebox.showerror("Erro de Validação", mensagem)
            return

        # Resetar estatísticas
        self.linhas_processadas = 0
        self.linhas_com_erro = 0
        self.linhas_puladas = 0
        self.stats = {'processados': 0, 'sucesso': 0, 'erros': 0, 'puladas': 0, 'tempo_inicio': datetime.now()}
        self.sucesso_label.configure(text="0")
        self.erros_label.configure(text="0")

        self.thread_automacao = threading.Thread(target=self.iniciar_automacao)
        self.thread_automacao.daemon = True
        self.thread_automacao.start()

        # Atualizar interface
        self.btn_iniciar.configure(state="disabled")
        self.btn_pausar.configure(state="normal")
        self.btn_parar.configure(state="normal")
        self.atualizar_status_indicator('executando')

        # Iniciar timer
        self.atualizar_tempo()

    def pausar_automacao(self):
        """Pausa/retoma a automação"""
        if self.executando:
            self.pausa_solicitada = not self.pausa_solicitada
            if self.pausa_solicitada:
                self.btn_pausar.configure(text="▶  Retomar")
                self.status_var.set("Pausado")
                self.atualizar_status_indicator('pausado')
                self.adicionar_log("Automação pausada", logging.INFO, "aviso")
            else:
                self.btn_pausar.configure(text="⏸  Pausar")
                self.status_var.set("Em execução...")
                self.atualizar_status_indicator('executando')
                self.adicionar_log("Automação retomada", logging.INFO, "info")

    def parar_automacao(self):
        """Para a execução da automação"""
        if self.executando:
            self.executando = False
            self.pausa_solicitada = False
            self.adicionar_log("Solicitação de parada enviada. Aguardando conclusão...", logging.INFO, "aviso")
            self.status_var.set("Interrompendo...")
            self.atualizar_status_indicator('erro')

    def ao_fechar(self):
        """Tratamento do fechamento da janela"""
        if self.executando:
            if messagebox.askyesno("Confirmação",
                                   "Existe uma automação em execução. Deseja realmente sair?"):
                self.executando = False
                self.pausa_solicitada = False
                self.window.after(1000, self.window.destroy)
        else:
            self.window.destroy()

    def iniciar_automacao(self):
        """Método principal de automação"""
        linha_inicial = int(self.linha_inicial.get())

        try:
            self.adicionar_log("Iniciando automação...", logging.INFO, "processando")
            self.status_var.set("Em execução...")
            self.executando = True

            # Carregar Excel
            df = pd.read_excel(self.arquivo_excel.get())

            # Ajustar linha inicial para índice do DataFrame (linha 2 = índice 1)
            inicio_indice = linha_inicial - 2
            df_processar = df.iloc[inicio_indice:]

            self.total_linhas = len(df_processar)
            self.adicionar_log(f"Arquivo carregado: {self.total_linhas} linhas para processar", logging.INFO, "info")
            self.adicionar_log(f"Iniciando da linha {linha_inicial} (índice {inicio_indice})", logging.INFO, "info")
            self.total_label.configure(text=str(self.total_linhas))

            # Resetar barra de progresso
            self.progress_bar.set(0)

            # Iniciar automação
            automacao = DominioAutomation(self.logger, self)

            # Conectar ao Domínio
            if not automacao.connect_to_dominio():
                self.adicionar_log("Não foi possível conectar ao Domínio", logging.ERROR, "erro")
                return

            # Processar linhas
            for idx, (original_index, row) in enumerate(df_processar.iterrows()):
                # Verificar se deve parar
                if not self.executando:
                    self.adicionar_log("Automação interrompida pelo usuário", logging.INFO, "aviso")
                    break

                # Verificar pausa
                while self.pausa_solicitada and self.executando:
                    time.sleep(0.5)

                if not self.executando:
                    break

                # Atualizar progresso
                self.atualizar_progresso(idx + 1, self.total_linhas)

                linha_excel = original_index + 2  # +2 porque: +1 para base 1, +1 para cabeçalho

                # Atualizar empresa no card
                empresa_num = str(int(row['Nº']))
                self.empresa_label.configure(text=empresa_num[:20])

                try:
                    self.adicionar_log(f"Processando linha {linha_excel} - Empresa {row['Nº']} - {row.get('EMPRESAS', 'N/A')}", logging.INFO, "processando")

                    success = automacao.processar_linha(row, original_index, linha_excel)

                    if success:
                        self.linhas_processadas += 1
                        self.success_logger.info(f"Linha {linha_excel} - Empresa {row['Nº']} - processada com sucesso")
                        self.adicionar_log(f"Linha {linha_excel} processada com sucesso", logging.INFO, "sucesso")
                    else:
                        self.linhas_com_erro += 1
                        self.error_logger.error(f"Linha {linha_excel} - Empresa {row['Nº']} - erro no processamento")
                        self.adicionar_log(f"Erro na linha {linha_excel}", logging.ERROR, "erro")

                    self.atualizar_estatisticas()
                    time.sleep(1)

                except Exception as e:
                    self.linhas_com_erro += 1
                    erro_msg = f"Linha {linha_excel} - Erro: {str(e)}"
                    self.error_logger.error(erro_msg)
                    self.adicionar_log(erro_msg, logging.ERROR, "erro")
                    self.atualizar_estatisticas()

            # Finalização
            if self.executando:
                self.status_var.set("Processamento concluído")
                self.progress_bar.set(1.0)
                self.progress_label.configure(text="100%")
                self.atualizar_status_indicator('concluido')
                self.adicionar_log("Automação concluída!", logging.INFO, "sucesso")
                self.adicionar_log(f"Resumo: {self.linhas_processadas} processadas, {self.linhas_com_erro} com erro, {self.linhas_puladas} puladas", logging.INFO, "info")

        except Exception as e:
            erro_msg = f"Erro crítico: {str(e)}"
            self.error_logger.error(erro_msg)
            self.adicionar_log(erro_msg, logging.ERROR, "erro")
            self.status_var.set("Erro no processamento")
            self.atualizar_status_indicator('erro')
        finally:
            self.executando = False
            self.pausa_solicitada = False
            self.btn_iniciar.configure(state="normal")
            self.btn_pausar.configure(state="disabled", text="⏸ Pausar")
            self.btn_parar.configure(state="disabled")

    def executar(self):
        self.window.mainloop()

class DominioAutomation:
    def __init__(self, logger, gui):
        timings.Timings.window_find_timeout = 20
        self.app = None
        self.main_window = None
        self.logger = logger
        self.gui = gui

    def log(self, message):
        self.logger.info(message)

    def should_stop(self) -> bool:
        """Verifica se deve parar a execução"""
        return not self.gui.executando

    def check_pause(self):
        """Verifica e aguarda se pausado"""
        while self.gui.pausa_solicitada and self.gui.executando:
            time.sleep(0.5)

    def smart_sleep(self, seconds: float):
        """Sleep interruptível que verifica pausa/parada"""
        interval = 0.5
        elapsed = 0
        while elapsed < seconds:
            if self.should_stop():
                return False
            self.check_pause()
            if self.should_stop():
                return False
            time.sleep(min(interval, seconds - elapsed))
            elapsed += interval
        return True

    def find_dominio_window(self) -> Optional[int]:
        """Encontra a janela do Domínio Folha"""
        try:
            # Procurar por qualquer janela que contenha "Domínio Folha" no título
            self.log("🔍 Procurando janela do Domínio Folha...")

            # Listar todas as janelas abertas para debug
            try:
                all_windows = findwindows.find_windows()
                self.log(f"📋 Total de janelas abertas: {len(all_windows)}")

                # Tentar encontrar janelas com "Domínio" no título
                for hwnd in all_windows:
                    try:
                        title = win32gui.GetWindowText(hwnd)
                        if "Domínio" in title and title:
                            self.log(f"🪟 Janela encontrada: '{title}'")
                            if "Folha" in title:
                                self.log(f"✅ Janela do Domínio Folha localizada!")
                                return hwnd
                    except Exception:
                        continue
            except Exception as e:
                self.log(f"⚠️ Erro ao listar janelas: {str(e)}")

            # Fallback: tentar o método original com regex
            windows = findwindows.find_windows(title_re=".*Domínio Folha.*")
            if windows:
                self.log(f"✅ Janela do Domínio encontrada via regex (total: {len(windows)})")
                return windows[0]

            self.log("❌ Nenhuma janela do Domínio Folha encontrada")
            return None
        except Exception as e:
            self.log(f"❌ Erro ao procurar janela do Domínio: {str(e)}")
            import traceback
            self.log(f"Traceback: {traceback.format_exc()}")
            return None

    def connect_to_dominio(self) -> bool:
        """Conecta à aplicação Domínio"""
        try:
            handle = self.find_dominio_window()
            if not handle:
                return False

            # Restaura e foca a janela
            if win32gui.IsIconic(handle):
                win32gui.ShowWindow(handle, win32con.SW_RESTORE)
                time.sleep(1)

            win32gui.SetForegroundWindow(handle)
            time.sleep(0.5)

            self.app = Application(backend="uia").connect(handle=handle)
            self.main_window = self.app.window(handle=handle)

            self.log("✅ Conectado ao Domínio Folha com sucesso")
            return True

        except Exception as e:
            self.log(f"❌ Erro ao conectar ao Domínio: {str(e)}")
            return False

    def wait_for_window_close(self, window, window_title: str, timeout: int = 30) -> bool:
        """Espera até que uma janela seja fechada"""
        start_time = time.time()
        while time.time() - start_time < timeout:
            if self.should_stop():
                return False
            self.check_pause()

            try:
                if not window.exists() or not window.is_visible():
                    self.log(f"✅ Janela '{window_title}' fechada")
                    return True
            except Exception:
                return True

            # Verificar se há diálogos de erro bloqueando
            self.handle_error_dialogs()

            time.sleep(0.5)

        self.log(f"⚠️ Timeout aguardando fechamento da janela '{window_title}'")
        return False

    def handle_empresa_change(self, empresa_num: str) -> bool:
        """Gerencia a troca de empresa"""
        try:
            if self.should_stop():
                return False

            # Enviar F8 para troca de empresas
            self.log("📞 Solicitando troca de empresa (F8)")
            send_keys('{F8}')
            if not self.smart_sleep(2):
                return False

            # Aguardar janela de troca
            max_attempts = 10
            troca_window = None

            for attempt in range(max_attempts):
                if self.should_stop():
                    return False
                self.check_pause()

                try:
                    troca_window = self.main_window.child_window(
                        title="Troca de empresas",
                        class_name="FNWND3190"
                    )

                    if troca_window.exists():
                        break

                    # Verificar se há diálogos de erro bloqueando
                    if not self.handle_error_dialogs():
                        self.cleanup_windows()
                        return False

                    if not self.smart_sleep(0.5):
                        return False
                except Exception:
                    if attempt == max_attempts - 1:
                        self.log("❌ Janela 'Troca de empresas' não encontrada (timeout)")
                        return False
                    if not self.smart_sleep(1):
                        return False

            if not troca_window:
                self.log("❌ Janela 'Troca de empresas' não encontrada")
                return False

            self.log(f"🏢 Alterando para empresa: {empresa_num}")

            # Enviar código da empresa
            send_keys(empresa_num)
            if not self.smart_sleep(0.5):
                return False
            send_keys('{ENTER}')
            if not self.smart_sleep(3):
                return False

            if not self.handle_error_dialogs():
                self.cleanup_windows()
                return False

            # Aguardar fechamento da janela de troca
            self.wait_for_window_close(troca_window, "Troca de empresas")

            # Fechar avisos de vencimento se existirem
            self.close_avisos_vencimento()

            return True

        except Exception as e:
            self.log(f"❌ Erro na troca de empresa: {str(e)}")
            return False

    def close_avisos_vencimento(self):
        """Fecha janela de avisos de vencimento se estiver aberta"""
        try:
            aviso_window = self.main_window.child_window(
                title="Avisos de Vencimento",
                class_name="FNWND3190"
            )

            if aviso_window.exists() and aviso_window.is_visible():
                self.log("📋 Fechando 'Avisos de Vencimento'")
                aviso_window.set_focus()
                send_keys('{ESC}')
                time.sleep(0.5)
                send_keys('{ESC}')
                time.sleep(0.5)
        except Exception:
            pass  # Não é crítico se não conseguir fechar

    def processar_linha(self, row, index: int, linha_excel: int) -> bool:
        """Processa uma linha do Excel"""
        try:
            if self.should_stop():
                return False

            # Reconectar se necessário
            handle = self.find_dominio_window()
            if not handle:
                self.log("❌ Não foi possível localizar a janela do Domínio")
                return False

            # Reconectar o app e main_window
            try:
                self.app = Application(backend="uia").connect(handle=handle)
                self.main_window = self.app.window(handle=handle)
                self.log("✅ Reconectado ao Domínio com sucesso")
            except Exception as e:
                self.log(f"❌ Erro ao reconectar: {str(e)}")
                return False

            if win32gui.IsIconic(handle):
                win32gui.ShowWindow(handle, win32con.SW_RESTORE)
                if not self.smart_sleep(1):
                    return False

            win32gui.SetForegroundWindow(handle)
            if not self.smart_sleep(0.5):
                return False

            # Troca de empresa
            empresa_num = str(int(row['Nº']))
            if not self.handle_empresa_change(empresa_num):
                return False

            if self.should_stop():
                return False
            self.check_pause()

            # Acessar relatórios
            self.log("📊 Acessando relatórios")
            self.main_window.set_focus()
            send_keys('%r')  # ALT+R
            if not self.smart_sleep(0.5):
                return False
            send_keys('i')  # Relatórios Integrados
            if not self.smart_sleep(0.5):
                return False
            send_keys('i')  # Relatórios Integrados
            if not self.smart_sleep(0.5):
                return False
            send_keys('{ENTER}')
            if not self.smart_sleep(1):
                return False

            # Processar no Gerenciador de Relatórios
            return self.processar_relatorio_taxa_gms(row, linha_excel)

        except Exception as e:
            self.log(f"❌ Erro ao processar linha {linha_excel}: {str(e)}")
            return False

    def processar_relatorio_taxa_gms(self, row, linha_excel: int) -> bool:
        """Processa o relatório de Taxa GMS"""
        try:
            if self.should_stop():
                return False

            # Aguardar Gerenciador de Relatórios
            max_attempts = 10
            relatorio_window = None

            for attempt in range(max_attempts):
                if self.should_stop():
                    return False
                self.check_pause()

                try:
                    relatorio_window = self.main_window.child_window(
                        title="Gerenciador de Relatórios",
                        class_name="FNWND3190"
                    )

                    if relatorio_window.exists():
                        break

                    # Verificar se há diálogos de erro bloqueando
                    if not self.handle_error_dialogs():
                        self.cleanup_windows()
                        return False

                    if not self.smart_sleep(1):
                        return False
                except Exception:
                    if attempt == max_attempts - 1:
                        self.log("❌ Gerenciador de Relatórios não encontrado (timeout)")
                        return False

            if not relatorio_window:
                self.log("❌ Gerenciador de Relatórios não encontrado")
                return False

            self.log("📋 Gerenciador de Relatórios localizado")

            if self.should_stop():
                return False
            self.check_pause()

            # Navegar até Taxa GMS
            self.log("🎯 Navegando para Taxa GMS")

            # Sequência de navegação otimizada
            navigation_keys = ['d'] * 6  # 6 vezes 'd' para navegar
            for key in navigation_keys:
                if self.should_stop():
                    return False
                send_keys(key)
                time.sleep(0.2)

            send_keys('{ENTER}')
            if not self.smart_sleep(0.5):
                return False
            send_keys('c')  # Selecionar relatório
            if not self.smart_sleep(0.5):
                return False

            # Preencher campos
            self.log("📝 Preenchendo parâmetros do relatório")

            # Navegar pelos campos e preencher
            send_keys('{TAB}')  # Pular primeiro campo
            time.sleep(0.2)

            send_keys('{TAB}22')  # Campo de código (assumindo valor fixo 22)
            time.sleep(0.3)

            send_keys('{TAB}8')  # Próximo campo
            time.sleep(0.2)

            # Período
            periodo = str(row['Periodo'])
            send_keys('{TAB}' + periodo)
            if not self.smart_sleep(0.5):
                return False

            if self.should_stop():
                return False
            self.check_pause()

            # Executar relatório
            self.log("⚡ Executando relatório")
            try:
                button_executar = relatorio_window.child_window(auto_id="1007", class_name="Button")
                button_executar.click_input()
                if not self.smart_sleep(4):
                    return False
            except Exception as e:
                self.log(f"⚠️ Erro ao clicar em executar, tentando via teclado: {str(e)}")
                send_keys('{F5}')  # Alternativa via teclado
                if not self.smart_sleep(4):
                    return False

            # Gerar PDF
            return self.gerar_pdf(row, linha_excel)

        except Exception as e:
            self.log(f"❌ Erro no processamento do relatório: {str(e)}")
            return False

    def gerar_pdf(self, row, linha_excel: int) -> bool:
        """Gera e salva o PDF do relatório"""
        try:
            if self.should_stop():
                return False

            # Verificar e tratar janela de erro
            if not self.handle_error_dialogs():
                self.cleanup_windows()
                return False

            self.log("📄 Gerando PDF")

            # Salvar como PDF usando Ctrl+D
            self.log("📄 Enviando Ctrl+D para salvar como PDF")
            send_keys('^d')  # Ctrl+D
            if not self.smart_sleep(2):
                return False

            # Verificar e tratar janela de erro
            if not self.handle_error_dialogs():
                self.cleanup_windows()
                return False

            # Aguardar janela de salvamento
            self.log("💾 Configurando salvamento do PDF")

            try:
                # Aguardar janela de salvamento aparecer com verificação de parada
                max_wait = 15
                elapsed = 0
                save_window = None

                while elapsed < max_wait:
                    if self.should_stop():
                        return False
                    self.check_pause()

                    try:
                        save_window = self.main_window.child_window(
                            title="Salvar em PDF",
                            class_name="#32770"
                        )
                        if save_window.exists():
                            break
                    except Exception:
                        pass

                    # Verificar se há diálogos de erro bloqueando
                    if not self.handle_error_dialogs():
                        self.cleanup_windows()
                        return False

                    time.sleep(0.5)
                    elapsed += 0.5

                if not save_window or not save_window.exists():
                    self.log("❌ Janela de salvamento não encontrada (timeout)")
                    return False

                if self.should_stop():
                    return False
                self.check_pause()

                # Preencher campos
                self.log("📝 Indo até a pasta correta...")

                # Navegar pelos campos e preencher
                send_keys('{TAB}')  # Pular primeiro campo
                time.sleep(0.2)

                send_keys('{TAB}')
                time.sleep(0.3)

                send_keys('{TAB}')  # Próximo campo
                time.sleep(0.2)

                send_keys('{TAB}')  # Próximo campo
                time.sleep(0.2)

                # Preencher campos
                self.log("📝 Acessando a pasta GMS...")

                # Navegar pelos campos e preencher
                send_keys('G')  # Drive
                time.sleep(0.2)
                send_keys('P')  # Pessoal
                time.sleep(0.2)
                send_keys('G')  # GMS
                time.sleep(0.2)

                # Preencher campos
                self.log("📝 Nomeando PDF...")

                # Navegar pelos campos e preencher
                send_keys('{TAB}')  # Pular primeiro campo
                time.sleep(0.2)

                send_keys('{TAB}')
                time.sleep(0.3)

                send_keys('{TAB}')  # Próximo campo
                time.sleep(0.2)

                send_keys('{TAB}')  # Próximo campo
                time.sleep(0.2)
                send_keys('{TAB}')  # Próximo campo
                time.sleep(0.2)

                nome_pdf = str(row['Salvar Como'])
                self.log(f"📝 Nome do arquivo: {nome_pdf}")

                # Definir nome do arquivo
                if not self.smart_sleep(0.5):
                    return False
                name_field = save_window.child_window(auto_id="1148", class_name="Edit")
                name_field.set_text(nome_pdf)
                if not self.smart_sleep(0.5):
                    return False

                if self.should_stop():
                    return False
                self.check_pause()

                # Salvar
                self.log("💾 Salvando PDF")
                button_salvar = save_window.child_window(auto_id="1", class_name="Button")
                button_salvar.click_input()
                if not self.smart_sleep(10):  # Aguardar salvamento
                    return False

            except Exception as e:
                self.log(f"❌ Erro durante salvamento: {str(e)}")
                return False

            # Fechar janelas e limpar
            self.cleanup_windows()

            return True

        except Exception as e:
            self.log(f"❌ Erro na geração do PDF: {str(e)}")
            return False

    def handle_error_dialogs(self) -> bool:
        """Trata diálogos de erro que podem aparecer. Retorna True se deve continuar, False se deve abortar."""
        try:
            # Lista de títulos possíveis de erro/aviso
            error_titles = ["Erro", "Erro léxico", "Aviso", "Atenção", "Informação", "Alerta", "Warning", "Error"]

            # Procurar diálogos de erro/aviso
            for title in error_titles:
                try:
                    # Tentar encontrar na janela principal
                    error_window = None
                    try:
                        error_window = self.app.window(title=title, class_name="#32770")
                    except Exception:
                        pass

                    # Tentar também como child window
                    if not error_window or not error_window.exists():
                        try:
                            error_window = self.main_window.child_window(title=title, class_name="#32770")
                        except Exception:
                            pass

                    # Tentar busca parcial no título
                    if not error_window or not error_window.exists():
                        try:
                            error_window = self.app.window(title_re=f".*{title}.*", class_name="#32770")
                        except Exception:
                            pass

                    if error_window and error_window.exists() and error_window.is_visible():
                        # Tentar obter o texto da mensagem
                        message = ""
                        try:
                            message = error_window.window_text()
                            # Também tentar pegar texto de controles estáticos dentro do diálogo
                            try:
                                static_texts = error_window.children(class_name="Static")
                                for static in static_texts:
                                    text = static.window_text()
                                    if text:
                                        message += " " + text
                            except Exception:
                                pass
                        except Exception:
                            pass

                        self.log(f"⚠️ Diálogo detectado: '{title}' - {message[:100] if message else 'sem mensagem'}")

                        # Verificar mensagens específicas que permitem continuar
                        mensagens_continuar = [
                            "Sem dados para emitir",
                            "Nenhum registro encontrado",
                            "Não há dados",
                            "Registro não encontrado"
                        ]

                        for msg in mensagens_continuar:
                            if msg.lower() in message.lower():
                                self.log(f"⚠️ Aviso não crítico: {msg}")
                                error_window.set_focus()
                                send_keys('{ENTER}')
                                time.sleep(1)
                                # Limpar janelas e continuar para próxima linha
                                for _ in range(4):
                                    send_keys('{ESC}')
                                    time.sleep(1)
                                return False  # Pular esta linha mas não travar

                        # Erro léxico - tentar fechar e continuar
                        if title == "Erro léxico":
                            self.log(f"⚠️ Erro léxico detectado, fechando...")
                            error_window.set_focus()
                            for _ in range(3):
                                send_keys('{ESC}')
                                time.sleep(1)
                            return True

                        # Para outros erros, tentar fechar com ENTER ou ESC
                        self.log(f"⚠️ Fechando diálogo '{title}'...")
                        error_window.set_focus()
                        time.sleep(0.3)

                        # Tentar clicar no botão OK se existir
                        try:
                            ok_button = error_window.child_window(title="OK", class_name="Button")
                            if ok_button.exists():
                                ok_button.click_input()
                                time.sleep(1)
                                continue
                        except Exception:
                            pass

                        # Senão, enviar ENTER
                        send_keys('{ENTER}')
                        time.sleep(1)

                        # Verificar se ainda existe e tentar ESC
                        if error_window.exists():
                            send_keys('{ESC}')
                            time.sleep(0.5)

                        # Para erros críticos como "Erro" ou "Aviso", abortar linha
                        if title in ["Erro", "Aviso"]:
                            return False

                except Exception as e:
                    pass

            return True

        except Exception as e:
            self.log(f"⚠️ Exceção ao verificar diálogos: {str(e)}")
            return True


    def cleanup_windows(self):
        """Limpa e fecha janelas abertas"""
        try:
            self.log("🧹 Limpando janelas")

            # Focar janela principal
            self.main_window.set_focus()

            # Enviar ESCs para garantir que todas as janelas sejam fechadas
            for _ in range(4):
                send_keys('{ESC}')
                time.sleep(1.5)

            # Verificar se o Gerenciador de Relatórios ainda está aberto
            try:
                relatorio_window = self.main_window.child_window(
                    title="Gerenciador de Relatórios",
                    class_name="FNWND3190"
                )

                if relatorio_window.exists() and relatorio_window.is_visible():
                    self.log("🔄 Fechando Gerenciador de Relatórios restante")
                    send_keys('{ESC}')
                    time.sleep(1)
            except Exception:
                pass

        except Exception as e:
            self.log(f"⚠️ Erro durante limpeza: {str(e)}")

def main():
    """Função principal"""
    try:
        gui = AutomacaoGUI()
        gui.executar()
    except Exception as e:
        print(f"Erro crítico na aplicação: {str(e)}")
        traceback.print_exc()

if __name__ == "__main__":
    main()
