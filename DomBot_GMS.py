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
from pywinauto.timings import wait_until
from typing import Optional, Tuple
import tkinter.messagebox as messagebox

class AutomacaoGUI:
    def __init__(self):
        # Configuração do tema
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("green")

        self.window = ctk.CTk()
        self.window.title("DomBot - Taxa GMS v2.0")
        self.window.geometry("800x600")
        self.window.protocol("WM_DELETE_WINDOW", self.ao_fechar)

        # Flags para controle de execução
        self.executando = False
        self.thread_automacao = None
        self.pausa_solicitada = False

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
        self.linha_inicial = ctk.StringVar(value="2")  # Corrigido: começa da linha 2 (primeira linha de dados)
        self.status_var = ctk.StringVar(value="Aguardando início...")

        # Variáveis de controle
        self.total_linhas = 0
        self.linhas_processadas = 0
        self.linhas_com_erro = 0
        self.linhas_puladas = 0

        # Logger
        self.logger = logging.getLogger('AutomacaoDominio')
        self.logger.setLevel(logging.INFO)
        self.logger.handlers = []

        # Configurar GUI Handler
        self.setup_gui_logging()

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

    def setup_gui_logging(self):
        """Configura o logging para a interface gráfica"""
        class GUIHandler(logging.Handler):
            def __init__(self, gui):
                super().__init__()
                self.gui = gui

            def emit(self, record):
                msg = self.format(record)
                # Usar after para thread-safety
                self.gui.window.after(0, lambda: self.gui.adicionar_log(msg))

        self.gui_handler = GUIHandler(self)
        formatter = logging.Formatter('%(message)s')
        self.gui_handler.setFormatter(formatter)
        self.logger.addHandler(self.gui_handler)

    def set_window_icon(self):
        """Configura o ícone da janela"""
        try:
            icon_path = os.path.join(os.path.dirname(__file__), "assets", "bot-folha-de-pagamento.ico")
            if os.name == 'nt' and os.path.exists(icon_path):
                self.window.iconbitmap(icon_path)
        except Exception as e:
            print(f"Erro ao carregar ícone: {e}")

    def selecionar_arquivo(self):
        """Seleciona arquivo Excel e mostra preview dos dados"""
        filename = ctk.filedialog.askopenfilename(
            filetypes=[("Excel files", "*.xlsx *.xls")],
            title="Selecione o arquivo Excel"
        )
        if filename:
            self.arquivo_excel.set(filename)
            self.adicionar_log(f"Arquivo selecionado: {filename}")

            # Preview dos dados
            try:
                df = pd.read_excel(filename)
                self.adicionar_log(f"Arquivo contém {len(df)} linhas de dados")

                # Mostrar colunas disponíveis
                colunas = ", ".join(df.columns.tolist())
                self.adicionar_log(f"Colunas: {colunas}")

                # Validar colunas necessárias
                colunas_necessarias = ['Nº', 'Periodo', 'Salvar Como']
                colunas_faltando = [col for col in colunas_necessarias if col not in df.columns]

                if colunas_faltando:
                    self.adicionar_log(f"⚠️ ATENÇÃO: Colunas obrigatórias não encontradas: {', '.join(colunas_faltando)}")
                else:
                    self.adicionar_log("✅ Todas as colunas obrigatórias encontradas")

            except Exception as e:
                self.adicionar_log(f"Erro ao ler arquivo: {str(e)}")

    def criar_interface(self):
        # Frame principal com scroll
        main_frame = ctk.CTkScrollableFrame(self.window)
        main_frame.pack(fill="both", expand=True, padx=20, pady=20)

        # Título
        title_label = ctk.CTkLabel(
            main_frame,
            text="DomBot - Automação Taxa GMS",
            font=ctk.CTkFont(size=20, weight="bold")
        )
        title_label.pack(pady=10)

        # Frame de configurações
        config_frame = ctk.CTkFrame(main_frame)
        config_frame.pack(fill="x", padx=10, pady=10)

        # ctk.CTkLabel(config_frame, text="Configurações", font=ctk.CTkFont(size=16, weight="bold")).pack(pady=5)

        # Seleção do arquivo Excel
        ctk.CTkLabel(config_frame, text="Arquivo Excel:", anchor="w").pack(fill="x", padx=10, pady=(10,2))

        file_frame = ctk.CTkFrame(config_frame)
        file_frame.pack(fill="x", padx=10, pady=5)

        self.arquivo_entry = ctk.CTkEntry(file_frame, textvariable=self.arquivo_excel, width=500)
        self.arquivo_entry.pack(side="left", padx=5, fill="x", expand=True)

        ctk.CTkButton(
            file_frame,
            text="Procurar",
            command=self.selecionar_arquivo,
            width=100
        ).pack(side="right", padx=5)

        # Frame para linha inicial e estatísticas
        linha_stats_frame = ctk.CTkFrame(config_frame)
        linha_stats_frame.pack(fill="x", padx=10, pady=10)

        # Linha inicial
        linha_frame = ctk.CTkFrame(linha_stats_frame)
        linha_frame.pack(side="left", fill="x", expand=True, padx=5)

        ctk.CTkLabel(linha_frame, text="Iniciar da linha (dados):").pack(pady=2)
        linha_spinbox = ctk.CTkEntry(linha_frame, textvariable=self.linha_inicial, width=100, justify="center")
        linha_spinbox.pack(pady=2)

        # Informação sobre numeração
        info_label = ctk.CTkLabel(
            linha_frame,
            text="Linha a se Iniciar o Excel",
            font=ctk.CTkFont(size=10),
            text_color="gray"
        )
        info_label.pack(pady=2)

        # Frame de estatísticas
        stats_frame = ctk.CTkFrame(linha_stats_frame)
        stats_frame.pack(side="right", padx=5)

        ctk.CTkLabel(stats_frame, text="Estatísticas da Sessão", font=ctk.CTkFont(weight="bold")).pack(pady=2)

        self.stats_labels = {
            'processadas': ctk.CTkLabel(stats_frame, text="Processadas: 0"),
            'erros': ctk.CTkLabel(stats_frame, text="Erros: 0"),
            'puladas': ctk.CTkLabel(stats_frame, text="Puladas: 0")
        }

        for label in self.stats_labels.values():
            label.pack(pady=1)

        # Botões de controle
        buttons_frame = ctk.CTkFrame(main_frame)
        buttons_frame.pack(fill="x", padx=10, pady=10)

        # ctk.CTkLabel(buttons_frame, text="Controles", font=ctk.CTkFont(size=16, weight="bold")).pack(pady=5)

        control_buttons_frame = ctk.CTkFrame(buttons_frame)
        control_buttons_frame.pack(fill="x", pady=10)

        self.btn_iniciar = ctk.CTkButton(
            control_buttons_frame,
            text="▶ Iniciar",
            command=self.iniciar_automacao_thread,
            height=40,
            font=ctk.CTkFont(size=14, weight="bold")
        )
        self.btn_iniciar.pack(side="left", expand=True, fill="x", padx=5)

        self.btn_pausar = ctk.CTkButton(
            control_buttons_frame,
            text="⏸ Pausar",
            command=self.pausar_automacao,
            height=40,
            state="disabled"
        )
        self.btn_pausar.pack(side="left", expand=True, fill="x", padx=5)

        self.btn_parar = ctk.CTkButton(
            control_buttons_frame,
            text="⏹ Parar",
            command=self.parar_automacao,
            height=40,
            state="disabled"
        )
        self.btn_parar.pack(side="left", expand=True, fill="x", padx=5)

        # Barra de Progresso
        progress_frame = ctk.CTkFrame(main_frame)
        progress_frame.pack(fill="x", padx=10, pady=10)

        # ctk.CTkLabel(progress_frame, text="Progresso", font=ctk.CTkFont(size=16, weight="bold")).pack(pady=5)

        self.progress_bar = ctk.CTkProgressBar(progress_frame, height=20)
        self.progress_bar.pack(fill="x", padx=10, pady=5)
        self.progress_bar.set(0)

        # Status
        self.status_label = ctk.CTkLabel(
            progress_frame,
            textvariable=self.status_var,
            font=ctk.CTkFont(size=12)
        )
        self.status_label.pack(pady=5)

        # Área de log
        log_frame = ctk.CTkFrame(main_frame)
        log_frame.pack(fill="both", expand=True, padx=10, pady=10)

        log_header_frame = ctk.CTkFrame(log_frame)
        log_header_frame.pack(fill="x", pady=(5,0))

        ctk.CTkLabel(log_header_frame, text="Log de Execução", font=ctk.CTkFont(size=16, weight="bold")).pack(side="left", pady=5)

        ctk.CTkButton(
            log_header_frame,
            text="🗑 Limpar",
            command=self.limpar_logs,
            width=80,
            height=25
        ).pack(side="right", padx=5, pady=5)

        self.log_text = ctk.CTkTextbox(log_frame, height=250)
        self.log_text.pack(fill="both", expand=True, padx=5, pady=5)

    def atualizar_estatisticas(self):
        """Atualiza as estatísticas na interface"""
        self.stats_labels['processadas'].configure(text=f"Processadas: {self.linhas_processadas}")
        self.stats_labels['erros'].configure(text=f"Erros: {self.linhas_com_erro}")
        self.stats_labels['puladas'].configure(text=f"Puladas: {self.linhas_puladas}")

    def limpar_logs(self):
        """Limpa a área de logs"""
        self.log_text.delete("1.0", "end")
        self.adicionar_log("📋 Log limpo")

    def atualizar_progresso(self, atual, total):
        """Atualiza a barra de progresso"""
        if total > 0:
            porcentagem = (atual / total)
            self.progress_bar.set(porcentagem)
            self.status_var.set(f"Processando: {atual}/{total} ({porcentagem*100:.1f}%)")
        self.window.update_idletasks()

    def adicionar_log(self, mensagem):
        """Adiciona mensagem ao log visual de forma thread-safe"""
        try:
            timestamp = datetime.now().strftime('%H:%M:%S')
            self.log_text.insert("end", f"[{timestamp}] {mensagem}\n")
            self.log_text.see("end")
            self.window.update_idletasks()
        except Exception:
            pass  # Ignora erros de thread safety

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
            self.adicionar_log("❌ Automação já em execução")
            return

        # Validar entrada
        valido, mensagem = self.validar_entrada()
        if not valido:
            self.adicionar_log(f"❌ Erro de validação: {mensagem}")
            messagebox.showerror("Erro de Validação", mensagem)
            return

        # Resetar estatísticas
        self.linhas_processadas = 0
        self.linhas_com_erro = 0
        self.linhas_puladas = 0
        self.atualizar_estatisticas()

        self.thread_automacao = threading.Thread(target=self.iniciar_automacao)
        self.thread_automacao.daemon = True
        self.thread_automacao.start()

        # Atualizar interface
        self.btn_iniciar.configure(state="disabled")
        self.btn_pausar.configure(state="normal")
        self.btn_parar.configure(state="normal")

    def pausar_automacao(self):
        """Pausa/resume a automação"""
        if self.executando:
            self.pausa_solicitada = not self.pausa_solicitada
            if self.pausa_solicitada:
                self.btn_pausar.configure(text="▶ Retomar")
                self.adicionar_log("⏸ Pausa solicitada - será pausado após a linha atual")
                self.status_var.set("Pausando...")
            else:
                self.btn_pausar.configure(text="⏸ Pausar")
                self.adicionar_log("▶ Retomando execução")

    def parar_automacao(self):
        """Para a execução da automação"""
        if self.executando:
            self.executando = False
            self.pausa_solicitada = False
            self.adicionar_log("⏹ Solicitação de parada enviada - aguardando conclusão da linha atual...")
            self.status_var.set("Parando...")

    def ao_fechar(self):
        """Tratamento do fechamento da janela"""
        if self.executando:
            resposta = messagebox.askyesno(
                "Confirmação",
                "Existe uma automação em execução.\n\nDeseja realmente sair?\nO processo será interrompido."
            )
            if resposta:
                self.executando = False
                self.pausa_solicitada = False
                self.window.after(1000, self.window.destroy)
        else:
            self.window.destroy()

    def iniciar_automacao(self):
        """Método principal de automação"""
        linha_inicial = int(self.linha_inicial.get())

        try:
            self.adicionar_log("🚀 Iniciando automação...")
            self.status_var.set("Carregando arquivo...")
            self.executando = True

            # Carregar Excel
            df = pd.read_excel(self.arquivo_excel.get())

            # Ajustar linha inicial para índice do DataFrame (linha 2 = índice 1)
            inicio_indice = linha_inicial - 2
            df_processar = df.iloc[inicio_indice:]

            self.total_linhas = len(df_processar)
            self.adicionar_log(f"📊 Arquivo carregado: {self.total_linhas} linhas para processar")
            self.adicionar_log(f"📍 Iniciando da linha {linha_inicial} (índice {inicio_indice})")

            # Resetar barra de progresso
            self.progress_bar.set(0)

            # Iniciar automação
            automacao = DominioAutomation(self.logger, self)

            # Conectar ao Domínio
            if not automacao.connect_to_dominio():
                self.adicionar_log("❌ Erro: Não foi possível conectar ao Domínio")
                return

            # Processar linhas
            for idx, (original_index, row) in enumerate(df_processar.iterrows()):
                # Verificar se deve parar
                if not self.executando:
                    self.adicionar_log("⏹ Automação interrompida pelo usuário")
                    break

                # Verificar pausa
                while self.pausa_solicitada and self.executando:
                    self.status_var.set("⏸ Pausado - clique em 'Retomar' para continuar")
                    time.sleep(1)

                if not self.executando:
                    break

                # Atualizar progresso
                self.atualizar_progresso(idx + 1, self.total_linhas)

                linha_excel = original_index + 2  # +2 porque: +1 para base 1, +1 para cabeçalho

                try:
                    self.adicionar_log(f"📝 Processando linha {linha_excel} - Empresa {row['Nº']} - {row.get('EMPRESAS', 'N/A')}")

                    success = automacao.processar_linha(row, original_index, linha_excel)

                    if success:
                        self.linhas_processadas += 1
                        self.success_logger.info(f"Linha {linha_excel} - Empresa {row['Nº']} - processada com sucesso")
                        self.adicionar_log(f"✅ Linha {linha_excel} processada com sucesso")
                    else:
                        self.linhas_com_erro += 1
                        self.error_logger.error(f"Linha {linha_excel} - Empresa {row['Nº']} - erro no processamento")
                        self.adicionar_log(f"❌ Erro na linha {linha_excel}")

                        # Opção de continuar ou parar em caso de erro
                        # Por enquanto, continua

                    self.atualizar_estatisticas()
                    time.sleep(1)  # Reduzido tempo entre processamentos

                except Exception as e:
                    self.linhas_com_erro += 1
                    erro_msg = f"Linha {linha_excel} - Erro: {str(e)}"
                    self.error_logger.error(erro_msg)
                    self.adicionar_log(f"💥 {erro_msg}")
                    self.atualizar_estatisticas()

            # Finalização
            if self.executando:
                self.status_var.set("✅ Processamento concluído")
                self.progress_bar.set(1.0)
                self.adicionar_log(f"🎉 Automação concluída!")
                self.adicionar_log(f"📊 Resumo: {self.linhas_processadas} processadas, {self.linhas_com_erro} com erro, {self.linhas_puladas} puladas")

        except Exception as e:
            erro_msg = f"💥 Erro crítico: {str(e)}"
            self.error_logger.error(erro_msg)
            self.adicionar_log(erro_msg)
            self.status_var.set("❌ Erro no processamento")
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

    def find_dominio_window(self) -> Optional[int]:
        """Encontra a janela do Domínio Folha"""
        try:
            windows = findwindows.find_windows(title_re=".*Domínio Folha.*")
            if windows:
                return windows[0]
            self.log("❌ Nenhuma janela do Domínio Folha encontrada")
            return None
        except Exception as e:
            self.log(f"❌ Erro ao procurar janela do Domínio: {str(e)}")
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
            try:
                if not window.exists() or not window.is_visible():
                    self.log(f"✅ Janela '{window_title}' fechada")
                    return True
            except Exception:
                return True
            time.sleep(0.5)

        self.log(f"⚠️ Timeout aguardando fechamento da janela '{window_title}'")
        return False

    def handle_empresa_change(self, empresa_num: str) -> bool:
        """Gerencia a troca de empresa"""
        try:
            # Enviar F8 para troca de empresas
            self.log("📞 Solicitando troca de empresa (F8)")
            send_keys('{F8}')
            time.sleep(2)

            # Aguardar janela de troca
            max_attempts = 5
            troca_window = None

            for attempt in range(max_attempts):
                try:
                    troca_window = self.main_window.child_window(
                        title="Troca de empresas",
                        class_name="FNWND3190"
                    )

                    if troca_window.exists():
                        break

                    time.sleep(0.5)
                except Exception:
                    if attempt == max_attempts - 1:
                        self.log("❌ Janela 'Troca de empresas' não encontrada")
                        return False
                    time.sleep(1)

            if not troca_window:
                return False

            self.log(f"🏢 Alterando para empresa: {empresa_num}")

            # Enviar código da empresa
            send_keys(empresa_num)
            time.sleep(0.5)
            send_keys('{ENTER}')
            time.sleep(3)

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
            # Reconectar se necessário
            handle = self.find_dominio_window()
            if not handle:
                return False

            if win32gui.IsIconic(handle):
                win32gui.ShowWindow(handle, win32con.SW_RESTORE)
                time.sleep(1)

            win32gui.SetForegroundWindow(handle)
            time.sleep(0.5)

            # Troca de empresa
            empresa_num = str(int(row['Nº']))
            if not self.handle_empresa_change(empresa_num):
                return False

            # Acessar relatórios
            self.log("📊 Acessando relatórios")
            self.main_window.set_focus()
            send_keys('%r')  # ALT+R
            time.sleep(0.5)
            send_keys('i')  # Relatórios Integrados
            time.sleep(0.5)
            send_keys('i')  # Relatórios Integrados
            time.sleep(0.5)
            send_keys('{ENTER}')
            time.sleep(1)

            # Processar no Gerenciador de Relatórios
            return self.processar_relatorio_taxa_gms(row, linha_excel)

        except Exception as e:
            self.log(f"❌ Erro ao processar linha {linha_excel}: {str(e)}")
            return False

    def processar_relatorio_taxa_gms(self, row, linha_excel: int) -> bool:
        """Processa o relatório de Taxa GMS"""
        try:
            # Aguardar Gerenciador de Relatórios
            max_attempts = 5
            relatorio_window = None

            for attempt in range(max_attempts):
                try:
                    relatorio_window = self.main_window.child_window(
                        title="Gerenciador de Relatórios",
                        class_name="FNWND3190"
                    )

                    if relatorio_window.exists():
                        break

                    time.sleep(1)
                except Exception:
                    if attempt == max_attempts - 1:
                        self.log("❌ Gerenciador de Relatórios não encontrado")
                        return False

            if not relatorio_window:
                return False

            self.log("📋 Gerenciador de Relatórios localizado")

            # Navegar até Taxa GMS
            self.log("🎯 Navegando para Taxa GMS")

            # Sequência de navegação otimizada
            navigation_keys = ['d'] * 6  # 6 vezes 'd' para navegar
            for key in navigation_keys:
                send_keys(key)
                time.sleep(0.2)

            send_keys('{ENTER}')
            time.sleep(0.5)
            send_keys('c')  # Selecionar relatório
            time.sleep(0.5)

            # Preencher campos
            self.log("📝 Preenchendo parâmetros do relatório")

            # Navegar pelos campos e preencher
            send_keys('{TAB}')  # Pular primeiro campo
            time.sleep(0.2)

            send_keys('{TAB}22')  # Campo de código (assumindo valor fixo 22)
            time.sleep(0.3)

            send_keys('{TAB}')  # Próximo campo
            time.sleep(0.2)

            # Período
            periodo = str(row['Periodo'])
            send_keys('{TAB}' + periodo)
            time.sleep(0.5)

            # Executar relatório
            self.log("⚡ Executando relatório")
            try:
                button_executar = relatorio_window.child_window(auto_id="1007", class_name="Button")
                button_executar.click_input()
                time.sleep(4)
            except Exception as e:
                self.log(f"⚠️ Erro ao clicar em executar, tentando via teclado: {str(e)}")
                send_keys('{F5}')  # Alternativa via teclado
                time.sleep(4)

            # Gerar PDF
            return self.gerar_pdf(row, linha_excel)

        except Exception as e:
            self.log(f"❌ Erro no processamento do relatório: {str(e)}")
            return False

    def gerar_pdf(self, row, linha_excel: int) -> bool:
        """Gera e salva o PDF do relatório"""
        try:
            # Verificar e tratar janela de erro
            if not self.handle_error_dialogs():
                self.cleanup_windows()
                return False
            
            self.log("📄 Gerando PDF")

            # Clicar no botão PDF
            try:
                button_pdf = self.main_window.child_window(auto_id="1014", class_name="FNUDO3190")
                button_pdf.click_input()
                time.sleep(2)
            except Exception as e:
                self.log(f"⚠️ Erro ao clicar no botão PDF: {str(e)}")
                # Alternativa via teclado
                send_keys('^p')  # Ctrl+P
                time.sleep(2)

            # Verificar e tratar janela de erro
            if not self.handle_error_dialogs():
                self.cleanup_windows()
                return False

            # Aguardar janela de salvamento
            self.log("💾 Configurando salvamento do PDF")

            try:
                # Aguardar janela de salvamento aparecer
                wait_until(
                    timeout=15,
                    retry_interval=0.5,
                    func=lambda: self.main_window.child_window(
                        title="Salvar em PDF",
                        class_name="#32770"
                    ).exists()
                )

                save_window = self.main_window.child_window(
                    title="Salvar em PDF",
                    class_name="#32770"
                )

                if save_window.exists():
                    
                    # Preencher campos
                    self.log("📝 Indo até a pasta correta...")

                    # Navegar pelos campos e preencher
                    send_keys('{TAB}')  # Pular primeiro campo
                    time.sleep(0.2)

                    send_keys('{TAB}')  # 
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

                    send_keys('{TAB}')  # 
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
                    time.sleep(0.5)
                    name_field = save_window.child_window(auto_id="1148", class_name="Edit")
                    name_field.set_text(nome_pdf)
                    time.sleep(0.5)

                    # Salvar
                    self.log("💾 Salvando PDF")
                    button_salvar = save_window.child_window(auto_id="1", class_name="Button")
                    button_salvar.click_input()
                    time.sleep(10)  # Aguardar salvamento
                
                else:
                    self.log("❌ Janela de salvamento não encontrada")
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
            # Verificar janela de erro genérica
            error_window = self.main_window.child_window(
                title="Erro",
                class_name="#32770"
            )

            if error_window.exists():
                self.log("⚠️ Janela de erro detectada, fechando")
                send_keys('{ENTER}')
                time.sleep(1)
                return False  # erro crítico, abortar

            # Verificar outras janelas de aviso
            aviso_titles = ["Aviso", "Atenção", "Informação"]
            for title in aviso_titles:
                aviso_window = self.main_window.child_window(
                    title=title,
                    class_name="#32770"
                )

                if aviso_window.exists():
                    self.log(f"⚠️ Janela '{title}' detectada, fechando")
                    send_keys('{ENTER}')
                    time.sleep(0.5)
                    if title == "Aviso":
                        return False  # abortar para próxima empresa

            return True  # se nada crítico, continuar

        except Exception:
            return True  # se der erro no tratamento, continuar mesmo assim

    def cleanup_windows(self):
        """Limpa e fecha janelas abertas"""
        try:
            self.log("🧹 Limpando janelas")

            # Focar janela principal
            self.main_window.set_focus()

            # Enviar ESCs para garantir que todas as janelas sejam fechadas
            for _ in range(3):
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