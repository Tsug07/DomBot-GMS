import customtkinter as ctk
import pandas as pd
from pywinauto.application import Application
from pywinauto.keyboard import send_keys
from pywinauto import findwindows, timings
import win32gui
import win32con
import time
import logging
import json
from datetime import datetime
import os
import traceback
import threading
import requests
import ctypes
from typing import Optional, Tuple
import tkinter.messagebox as messagebox
from PIL import Image, ImageDraw
from dotenv import load_dotenv

load_dotenv()


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

        # Configurar ícone
        self.set_window_icon()

        # Criar diretório de logs se não existir
        self.logs_dir = os.path.join(os.path.dirname(__file__), "logs")
        if not os.path.exists(self.logs_dir):
            os.makedirs(self.logs_dir)

        # Configurar logging para arquivos
        self.setup_file_logging()

        # Variáveis da interface
        self.status_var = ctk.StringVar(value="Aguardando início...")
        self.pasta_saida = ctk.StringVar()
        self.competencia_var = ctk.StringVar(value=datetime.now().strftime("%m/%Y"))
        self.data_vencimento_var = ctk.StringVar()
        self.emitir_var = ctk.BooleanVar(value=True)
        self.publicar_var = ctk.BooleanVar(value=True)

        # Lista de empresas cadastradas (gerenciada pela interface e salva em JSON).
        # Cada item: {"codigo": str, "nome": str, "ativo": bool}
        self.empresas = []
        self.empresas_json = os.path.join(
            os.path.dirname(os.path.abspath(__file__)), "empresas_gms.json"
        )
        self.empresa_widgets = []  # linhas (frame + checkbox var) renderizadas na aba

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
        self.carregar_empresas_json()
        self.render_lista_empresas()

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
        config_frame.grid_columnconfigure(0, weight=1)

        # --- Linha 1: resumo de empresas + competência + botões de controle ---
        inner_frame = ctk.CTkFrame(config_frame, fg_color="transparent")
        inner_frame.grid(row=0, column=0, sticky="ew", padx=10, pady=(8, 4))
        inner_frame.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(inner_frame, text="🏢", font=ctk.CTkFont(size=14)).grid(row=0, column=0, padx=(0, 5))

        self.lbl_resumo_empresas = ctk.CTkLabel(
            inner_frame, text="Nenhuma empresa cadastrada",
            font=ctk.CTkFont(size=11), text_color="#BDC3C7", anchor="w"
        )
        self.lbl_resumo_empresas.grid(row=0, column=1, sticky="ew", padx=(0, 15))

        # Competência
        ctk.CTkLabel(
            inner_frame, text="Competência:", font=ctk.CTkFont(size=11), text_color="#BDC3C7"
        ).grid(row=0, column=3, padx=(0, 5))

        ctk.CTkEntry(
            inner_frame, textvariable=self.competencia_var,
            width=90, height=32, font=ctk.CTkFont(size=11), justify="center",
            placeholder_text="MM/AAAA"
        ).grid(row=0, column=4, padx=(0, 15))

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

        # --- Linha 2: pasta de saída + vencimento + etapas (emitir/publicar) ---
        row2 = ctk.CTkFrame(config_frame, fg_color="transparent")
        row2.grid(row=1, column=0, sticky="ew", padx=10, pady=(0, 8))
        row2.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(row2, text="📂", font=ctk.CTkFont(size=14)).grid(row=0, column=0, padx=(0, 5))
        ctk.CTkEntry(
            row2, textvariable=self.pasta_saida, height=32, font=ctk.CTkFont(size=11),
            placeholder_text="Pasta de saída dos PDFs..."
        ).grid(row=0, column=1, sticky="ew", padx=(0, 8))
        ctk.CTkButton(
            row2, text="Pasta", command=self.selecionar_pasta_saida,
            width=70, height=32, font=ctk.CTkFont(size=11),
            fg_color=self.CORES['info'], hover_color="#2980B9"
        ).grid(row=0, column=2, padx=(0, 15))

        ctk.CTkLabel(row2, text="⏳", font=ctk.CTkFont(size=14)).grid(row=0, column=3, padx=(0, 5))
        ctk.CTkLabel(row2, text="Vencimento:", font=ctk.CTkFont(size=11), text_color="#BDC3C7").grid(
            row=0, column=4, padx=(0, 5)
        )
        ctk.CTkEntry(
            row2, textvariable=self.data_vencimento_var,
            width=100, height=32, font=ctk.CTkFont(size=11), justify="center",
            placeholder_text="DD/MM/AAAA"
        ).grid(row=0, column=5, padx=(0, 15))

        ctk.CTkCheckBox(
            row2, text="Emitir", variable=self.emitir_var,
            font=ctk.CTkFont(size=11), checkbox_width=18, checkbox_height=18
        ).grid(row=0, column=6, padx=(0, 12))

        ctk.CTkCheckBox(
            row2, text="Publicar", variable=self.publicar_var,
            font=ctk.CTkFont(size=11), checkbox_width=18, checkbox_height=18
        ).grid(row=0, column=7, padx=(0, 3))

    def selecionar_pasta_saida(self):
        """Abre o seletor de pasta nativo do Windows (SHBrowseForFolder) em uma
        thread dedicada com COM em modo STA próprio.

        Motivo: este app importa/usa pywinauto (backend uia), que deixa o COM da
        thread principal em um estado incompatível com o diálogo de shell do
        Windows. Chamar ctk.filedialog.askdirectory na thread da UI trava o
        programa ('Não está respondendo') sem nunca abrir o diálogo. Rodar o
        picker em uma thread separada que faz seu próprio CoInitialize (STA)
        isola o COM e evita o travamento. A thread da UI apenas aguarda o
        resultado sem bloquear o mainloop (usa update() em polling curto)."""
        resultado = {'pasta': None, 'done': False}

        inicial = self.pasta_saida.get().strip()
        if not inicial or not os.path.isdir(inicial):
            inicial = os.path.expanduser("~")

        # HWND top-level da janela principal, para o diálogo abrir à frente (owner)
        try:
            owner_hwnd = win32gui.GetAncestor(self.window.winfo_id(), 2)  # GA_ROOT
        except Exception:
            owner_hwnd = 0

        # BIF_NEWDIALOGSTYLE (0x40) não é exposto por shellcon nesta versão do
        # pywin32, por isso usamos o literal. Ativa o diálogo moderno (com
        # redimensionamento e botão "Nova pasta").
        BIF_RETURNONLYFSDIRS = 0x00000001
        BIF_NEWDIALOGSTYLE = 0x00000040

        def _abrir_dialogo():
            try:
                import pythoncom
                import win32com.shell.shell as shell
                pythoncom.CoInitialize()
                try:
                    pidl, _display, _img = shell.SHBrowseForFolder(
                        owner_hwnd,
                        None,
                        "Selecione a pasta de saída dos PDFs",
                        BIF_RETURNONLYFSDIRS | BIF_NEWDIALOGSTYLE,
                    )
                    if pidl is not None:
                        resultado['pasta'] = shell.SHGetPathFromIDList(pidl)
                finally:
                    pythoncom.CoUninitialize()
            except Exception as e:
                resultado['erro'] = str(e)
            finally:
                resultado['done'] = True

        self.adicionar_log("Abrindo seletor de pasta...", logging.INFO, "info")
        t = threading.Thread(target=_abrir_dialogo, daemon=True)
        t.start()

        # Aguarda o diálogo sem congelar o mainloop (mantém a UI responsiva)
        while not resultado['done']:
            try:
                self.window.update()
            except Exception:
                pass
            time.sleep(0.05)

        if resultado.get('erro'):
            self.adicionar_log(f"Erro ao selecionar pasta: {resultado['erro']}",
                               logging.ERROR, "erro")
            return

        pasta = resultado['pasta']
        if pasta:
            if isinstance(pasta, bytes):
                pasta = pasta.decode('utf-8', 'ignore')
            pasta = pasta.replace("/", "\\")
            self.pasta_saida.set(pasta)
            self.adicionar_log(f"Pasta de saída: {pasta}", logging.INFO, "info")
        else:
            self.adicionar_log("Seleção de pasta cancelada (nenhuma pasta escolhida)",
                               logging.INFO, "aviso")

    # ── Gerenciamento de empresas (JSON + interface) ──────────────────────────

    def carregar_empresas_json(self):
        """Carrega a lista de empresas do arquivo JSON, se existir."""
        try:
            if os.path.exists(self.empresas_json):
                with open(self.empresas_json, 'r', encoding='utf-8') as f:
                    dados = json.load(f)
                self.empresas = []
                for item in dados:
                    self.empresas.append({
                        'codigo': str(item.get('codigo', '')).strip(),
                        'nome': str(item.get('nome', '')).strip(),
                        'ativo': bool(item.get('ativo', True)),
                    })
                self.adicionar_log(f"{len(self.empresas)} empresa(s) carregada(s) do cadastro.",
                                   logging.INFO, "sucesso")
        except Exception as e:
            self.adicionar_log(f"Erro ao carregar cadastro de empresas: {e}", logging.ERROR, "erro")

    def salvar_empresas_json(self):
        """Grava a lista atual de empresas no arquivo JSON."""
        try:
            with open(self.empresas_json, 'w', encoding='utf-8') as f:
                json.dump(self.empresas, f, ensure_ascii=False, indent=2)
        except Exception as e:
            self.adicionar_log(f"Erro ao salvar cadastro de empresas: {e}", logging.ERROR, "erro")

    @staticmethod
    def _chave_ordenacao_codigo(emp):
        codigo = str(emp.get('codigo', '')).strip()
        try:
            return (0, int(codigo))
        except ValueError:
            return (1, codigo)

    def ordenar_empresas(self):
        """Ordena o cadastro em memória por código (numérico quando possível)."""
        self.empresas.sort(key=self._chave_ordenacao_codigo)

    def render_lista_empresas(self):
        """Redesenha a lista rolável de empresas com base em self.empresas."""
        self.ordenar_empresas()
        for w in self.lista_empresas_frame.winfo_children():
            w.destroy()
        self.empresa_widgets = []

        if not self.empresas:
            ctk.CTkLabel(
                self.lista_empresas_frame,
                text="Nenhuma empresa cadastrada.\nUse ➕ Adicionar ou 📥 Importar do Excel.",
                font=ctk.CTkFont(size=12), text_color="#7F8C8D", justify="center",
            ).grid(row=0, column=0, pady=30)
            self._atualizar_resumo_empresas()
            return

        for idx, emp in enumerate(self.empresas):
            linha = ctk.CTkFrame(self.lista_empresas_frame, fg_color=self.CORES['fundo_card'],
                                 corner_radius=6)
            linha.grid(row=idx, column=0, sticky="ew", padx=2, pady=2)
            linha.grid_columnconfigure(2, weight=1)

            var = ctk.BooleanVar(value=emp.get('ativo', True))
            chk = ctk.CTkCheckBox(
                linha, text="", variable=var, width=24,
                command=lambda i=idx, v=var: self._toggle_ativo(i, v),
                checkbox_width=20, checkbox_height=20,
            )
            chk.grid(row=0, column=0, padx=(8, 2), pady=6)

            ctk.CTkLabel(linha, text=emp['codigo'], width=90, anchor="w",
                         font=ctk.CTkFont(size=11)).grid(row=0, column=1, padx=6, sticky="w")
            ctk.CTkLabel(linha, text=emp['nome'], anchor="w",
                         font=ctk.CTkFont(size=11)).grid(row=0, column=2, padx=6, sticky="w")

            btns = ctk.CTkFrame(linha, fg_color="transparent")
            btns.grid(row=0, column=3, padx=6)
            ctk.CTkButton(btns, text="✏", width=32, height=26, font=ctk.CTkFont(size=12),
                          fg_color=self.CORES['info'], hover_color="#2980B9",
                          command=lambda i=idx: self.editar_empresa_dialog(i)).pack(side="left", padx=2)
            ctk.CTkButton(btns, text="🗑", width=32, height=26, font=ctk.CTkFont(size=12),
                          fg_color=self.CORES['erro'], hover_color="#C0392B",
                          command=lambda i=idx: self.remover_empresa(i)).pack(side="left", padx=2)

            self.empresa_widgets.append(var)

        self._atualizar_resumo_empresas()

    def _atualizar_resumo_empresas(self):
        total = len(self.empresas)
        ativas = sum(1 for e in self.empresas if e.get('ativo', True))
        if total == 0:
            texto = "Nenhuma empresa cadastrada"
        else:
            texto = f"{ativas} de {total} empresa(s) selecionada(s) para processar"
        try:
            self.lbl_resumo_empresas.configure(text=texto)
        except Exception:
            pass

    def _toggle_ativo(self, idx, var):
        if 0 <= idx < len(self.empresas):
            self.empresas[idx]['ativo'] = bool(var.get())
            self.salvar_empresas_json()
            self._atualizar_resumo_empresas()

    def marcar_todas(self, valor: bool):
        for emp in self.empresas:
            emp['ativo'] = valor
        self.salvar_empresas_json()
        self.render_lista_empresas()

    def _dialog_empresa(self, titulo, codigo="", nome=""):
        """Abre um diálogo modal para cadastrar/editar. Retorna dict ou None se cancelado."""
        dlg = ctk.CTkToplevel(self.window)
        dlg.title(titulo)
        dlg.geometry("420x220")
        dlg.transient(self.window)
        dlg.grab_set()
        dlg.resizable(False, False)

        resultado = {'valor': None}
        var_cod = ctk.StringVar(value=codigo)
        var_nome = ctk.StringVar(value=nome)

        frame = ctk.CTkFrame(dlg, fg_color="transparent")
        frame.pack(fill="both", expand=True, padx=20, pady=20)
        frame.grid_columnconfigure(1, weight=1)

        campos = [("Código:", var_cod), ("Nome:", var_nome)]
        entradas = []
        for i, (rot, v) in enumerate(campos):
            ctk.CTkLabel(frame, text=rot, font=ctk.CTkFont(size=12), width=70,
                         anchor="w").grid(row=i, column=0, padx=(0, 8), pady=8, sticky="w")
            e = ctk.CTkEntry(frame, textvariable=v, height=34, font=ctk.CTkFont(size=12))
            e.grid(row=i, column=1, sticky="ew", pady=8)
            entradas.append(e)
        entradas[0].focus()

        erro_lbl = ctk.CTkLabel(frame, text="", font=ctk.CTkFont(size=11),
                                text_color=self.CORES['erro'])
        erro_lbl.grid(row=2, column=0, columnspan=2, sticky="w")

        def confirmar():
            c, n = var_cod.get().strip(), var_nome.get().strip()
            if not c or not n:
                erro_lbl.configure(text="Preencha código e nome.")
                return
            resultado['valor'] = {'codigo': c, 'nome': n, 'ativo': True}
            dlg.destroy()

        botoes = ctk.CTkFrame(frame, fg_color="transparent")
        botoes.grid(row=3, column=0, columnspan=2, pady=(12, 0), sticky="e")
        ctk.CTkButton(botoes, text="Cancelar", command=dlg.destroy, width=100, height=32,
                      fg_color="#34495E", hover_color="#2C3E50").pack(side="left", padx=6)
        ctk.CTkButton(botoes, text="Salvar", command=confirmar, width=100, height=32,
                      fg_color=self.CORES['sucesso'], hover_color="#27AE60").pack(side="left")

        dlg.bind('<Return>', lambda _e: confirmar())
        dlg.wait_window()
        return resultado['valor']

    def adicionar_empresa_dialog(self):
        nova = self._dialog_empresa("Adicionar Empresa")
        if nova:
            self.empresas.append(nova)
            self.salvar_empresas_json()
            self.render_lista_empresas()
            self.adicionar_log(f"Empresa adicionada: {nova['codigo']} - {nova['nome']}",
                               logging.INFO, "sucesso")

    def editar_empresa_dialog(self, idx):
        if not (0 <= idx < len(self.empresas)):
            return
        emp = self.empresas[idx]
        editada = self._dialog_empresa("Editar Empresa", emp['codigo'], emp['nome'])
        if editada:
            editada['ativo'] = emp.get('ativo', True)  # preserva a seleção
            self.empresas[idx] = editada
            self.salvar_empresas_json()
            self.render_lista_empresas()
            self.adicionar_log(f"Empresa atualizada: {editada['codigo']} - {editada['nome']}",
                               logging.INFO, "info")

    def remover_empresa(self, idx):
        if not (0 <= idx < len(self.empresas)):
            return
        emp = self.empresas[idx]
        if messagebox.askyesno("Confirmar remoção",
                               f"Remover a empresa:\n\n{emp['codigo']} - {emp['nome']}?"):
            self.empresas.pop(idx)
            self.salvar_empresas_json()
            self.render_lista_empresas()
            self.adicionar_log(f"Empresa removida: {emp['codigo']} - {emp['nome']}",
                               logging.INFO, "aviso")

    def importar_empresas_excel(self):
        """Importa empresas do Excel gerado pelo M.E.G_ONE (colunas: Nº, EMPRESAS, ...)."""
        filename = ctk.filedialog.askopenfilename(
            filetypes=[("Excel files", "*.xlsx *.xls")],
            title="Selecione o Excel para importar as empresas",
        )
        if not filename:
            return
        try:
            df = pd.read_excel(filename)
            df = df.iloc[:, :2]
            df.columns = ['Codigo', 'Nome']
            # Pular linha de cabeçalho se existir (Excel sem header já lido acima)
            if str(df.iloc[0]['Codigo']).lower() in ('codigo', 'código', 'nº', 'code'):
                df = df.iloc[1:].reset_index(drop=True)

            existentes = {e['codigo'] for e in self.empresas}
            importadas = 0
            duplicadas = 0
            for _, row in df.iterrows():
                codigo = self._limpar_codigo(row['Codigo'])
                if not codigo:
                    continue
                if codigo in existentes:
                    duplicadas += 1
                    continue
                self.empresas.append({
                    'codigo': codigo,
                    'nome': str(row.get('Nome', '')).strip(),
                    'ativo': True,
                })
                existentes.add(codigo)
                importadas += 1

            self.salvar_empresas_json()
            self.render_lista_empresas()
            msg = f"Importação concluída: {importadas} adicionada(s)"
            if duplicadas:
                msg += f", {duplicadas} já existente(s) ignorada(s)"
            self.adicionar_log(msg, logging.INFO, "sucesso")
            messagebox.showinfo("Importação", msg)
        except Exception as e:
            self.adicionar_log(f"Erro ao importar do Excel: {e}", logging.ERROR, "erro")
            messagebox.showerror("Erro na importação", str(e))

    @staticmethod
    def _limpar_codigo(codigo) -> str:
        """Converte código para string limpa, removendo '.0' e espaços."""
        if codigo is None or pd.isna(codigo):
            return ""
        if isinstance(codigo, float) and codigo.is_integer():
            return str(int(codigo))
        codigo_str = str(codigo).strip()
        if codigo_str.endswith('.0'):
            codigo_str = codigo_str[:-2]
        return codigo_str

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

        tab_empresas = self.tabview.add("🏢 Empresas")
        tab_logs = self.tabview.add("📋 Logs")

        self.criar_aba_empresas(tab_empresas)
        self.criar_aba_logs(tab_logs)

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

    def criar_aba_empresas(self, parent):
        """Cria a aba de gerenciamento de empresas cadastradas"""
        parent.grid_columnconfigure(0, weight=1)
        parent.grid_rowconfigure(2, weight=1)

        # --- Barra de ações ---
        acoes = ctk.CTkFrame(parent, fg_color="transparent")
        acoes.grid(row=0, column=0, sticky="ew", padx=3, pady=(3, 2))

        ctk.CTkButton(
            acoes, text="➕ Adicionar", command=self.adicionar_empresa_dialog,
            width=100, height=28, font=ctk.CTkFont(size=11, weight="bold"),
            fg_color=self.CORES['sucesso'], hover_color="#27AE60",
        ).pack(side="left", padx=(0, 6))

        ctk.CTkButton(
            acoes, text="📥 Importar do Excel", command=self.importar_empresas_excel,
            width=150, height=28, font=ctk.CTkFont(size=11),
            fg_color=self.CORES['info'], hover_color="#2980B9",
        ).pack(side="left", padx=(0, 6))

        ctk.CTkButton(
            acoes, text="☑ Todas", command=lambda: self.marcar_todas(True),
            width=70, height=28, font=ctk.CTkFont(size=11),
            fg_color="#34495E", hover_color="#2C3E50",
        ).pack(side="left", padx=(0, 4))

        ctk.CTkButton(
            acoes, text="☐ Nenhuma", command=lambda: self.marcar_todas(False),
            width=90, height=28, font=ctk.CTkFont(size=11),
            fg_color="#34495E", hover_color="#2C3E50",
        ).pack(side="left")

        # --- Cabeçalho da lista ---
        header = ctk.CTkFrame(parent, fg_color=self.CORES['fundo_escuro'], corner_radius=6)
        header.grid(row=1, column=0, sticky="ew", padx=3, pady=(2, 0))
        for col, (txt, w) in enumerate([("", 30), ("Código", 90), ("Nome", 300), ("Ações", 120)]):
            ctk.CTkLabel(header, text=txt, width=w, font=ctk.CTkFont(size=10, weight="bold"),
                         text_color="#7F8C8D", anchor="w").grid(row=0, column=col, padx=6, pady=4, sticky="w")

        # --- Lista rolável de empresas ---
        self.lista_empresas_frame = ctk.CTkScrollableFrame(
            parent, fg_color=self.CORES['fundo_escuro'], corner_radius=6,
        )
        self.lista_empresas_frame.grid(row=2, column=0, sticky="nsew", padx=3, pady=(0, 3))
        self.lista_empresas_frame.grid_columnconfigure(0, weight=1)

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
        if not self.emitir_var.get() and not self.publicar_var.get():
            return False, "Selecione ao menos uma etapa: Emitir e/ou Publicar"

        if not self.empresas:
            return False, "Nenhuma empresa cadastrada. Adicione ou importe do Excel na aba 🏢 Empresas."

        if not any(e.get('ativo', True) for e in self.empresas):
            return False, "Nenhuma empresa selecionada. Marque ao menos uma na aba 🏢 Empresas."

        competencia = self.competencia_var.get().strip()
        if not competencia:
            return False, "Informe a competência (MM/AAAA)"
        try:
            datetime.strptime(competencia, "%m/%Y")
        except ValueError:
            return False, "Competência inválida. Use o formato MM/AAAA"

        if not self.pasta_saida.get().strip():
            return False, "Selecione a pasta de saída dos PDFs"

        if not os.path.isdir(self.pasta_saida.get().strip()):
            return False, "Pasta de saída não encontrada. Verifique o caminho"

        if self.publicar_var.get():
            venc = self.data_vencimento_var.get().strip()
            if not venc:
                return False, "Informe a data de vencimento (DD/MM/AAAA) para a publicação"
            try:
                datetime.strptime(venc, "%d/%m/%Y")
            except ValueError:
                return False, "Data de vencimento inválida. Use o formato DD/MM/AAAA"

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
        self.erros_detalhados = []
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
        pasta_saida = self.pasta_saida.get().strip()
        competencia = self.competencia_var.get().strip()  # MM/AAAA
        emitir = self.emitir_var.get()
        publicar = self.publicar_var.get()

        mes, ano = competencia.split('/')
        competencia_fmt = f"{mes}{ano}"

        try:
            self.adicionar_log("Iniciando automação...", logging.INFO, "processando")
            self.status_var.set("Em execução...")
            self.executando = True

            # Snapshot das empresas selecionadas (ativas) no momento do início
            empresas_processar = [e for e in self.empresas if e.get('ativo', True)]
            self.total_linhas = len(empresas_processar)
            self.adicionar_log(f"{self.total_linhas} empresa(s) selecionada(s) para processar", logging.INFO, "info")
            self.total_label.configure(text=str(self.total_linhas))

            # Resetar barra de progresso
            self.progress_bar.set(0)

            # Iniciar automação
            automacao = DominioAutomation(self.logger, self)

            # Conectar ao Domínio
            if not automacao.connect_to_dominio():
                self.adicionar_log("Não foi possível conectar ao Domínio", logging.ERROR, "erro")
                return

            documentos_para_publicar = []  # lista de (codigo, caminho_pdf)

            if emitir:
                for idx, emp in enumerate(empresas_processar):
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

                    numero = idx + 1
                    codigo = str(emp['codigo']).strip()
                    nome = str(emp.get('nome', 'N/A')).strip()

                    # Atualizar empresa no card
                    self.empresa_label.configure(text=codigo[:20])

                    nome_pdf = f"{codigo}-{nome}-{competencia_fmt}"
                    caminho_pdf = os.path.join(pasta_saida, f"{nome_pdf}.pdf")

                    try:
                        self.adicionar_log(f"[{numero}/{self.total_linhas}] Processando - Empresa {codigo} - {nome}", logging.INFO, "processando")

                        success = automacao.processar_linha(codigo, nome, competencia, numero, caminho_pdf)

                        if success:
                            self.linhas_processadas += 1
                            documentos_para_publicar.append((codigo, caminho_pdf))
                            self.success_logger.info(f"Empresa {codigo} - {nome} - processada com sucesso")
                            self.adicionar_log(f"{codigo} - {nome} processada com sucesso", logging.INFO, "sucesso")
                        else:
                            self.linhas_com_erro += 1
                            self.error_logger.error(f"Empresa {codigo} - {nome} - erro no processamento")
                            self.adicionar_log(f"Erro ao processar {codigo} - {nome}", logging.ERROR, "erro")
                            self.erros_detalhados.append({
                                'empresa': nome,
                                'numero': codigo,
                                'motivo': 'Erro no processamento'
                            })

                        self.atualizar_estatisticas()

                    except Exception as e:
                        self.linhas_com_erro += 1
                        erro_msg = f"Empresa {codigo} - {nome} - Erro: {str(e)}"
                        self.error_logger.error(erro_msg)
                        self.adicionar_log(erro_msg, logging.ERROR, "erro")
                        self.erros_detalhados.append({
                            'empresa': nome,
                            'numero': codigo,
                            'motivo': str(e)[:80]
                        })
                        self.atualizar_estatisticas()
            else:
                self.adicionar_log("Etapa de emissão pulada (checkbox 'Emitir' desmarcado)", logging.INFO, "aviso")
                for emp in empresas_processar:
                    codigo = str(emp['codigo']).strip()
                    nome = str(emp.get('nome', 'N/A')).strip()
                    nome_pdf = f"{codigo}-{nome}-{competencia_fmt}"
                    caminho_pdf = os.path.join(pasta_saida, f"{nome_pdf}.pdf")
                    documentos_para_publicar.append((codigo, caminho_pdf))

            # Publicação em lote: só após salvar todos os PDFs e fechar as janelas
            publicados, falhas_publicacao = 0, 0
            if self.executando and publicar and documentos_para_publicar:
                automacao.cleanup_windows()
                self.status_var.set("Publicando no portal...")
                self.atualizar_status_indicator('executando')
                data_venc = self.data_vencimento_var.get().strip()
                publicados, falhas_publicacao = automacao.publicar_lote_gms(documentos_para_publicar, data_venc)
                self.adicionar_log(
                    f"Publicação concluída: {publicados} publicado(s), {falhas_publicacao} falha(s)",
                    logging.INFO, "sucesso" if falhas_publicacao == 0 else "aviso"
                )
            elif self.executando and publicar and not documentos_para_publicar:
                self.adicionar_log("Nenhum documento disponível para publicar", logging.WARNING, "aviso")

            # Finalização
            if self.executando:
                self.status_var.set("Processamento concluído")
                self.progress_bar.set(1.0)
                self.progress_label.configure(text="100%")
                self.atualizar_status_indicator('concluido')
                self.adicionar_log("Automação concluída!", logging.INFO, "sucesso")
                self.adicionar_log(f"Resumo: {self.linhas_processadas} processadas, {self.linhas_com_erro} com erro, {self.linhas_puladas} puladas", logging.INFO, "info")

                # Enviar notificação ao Discord via webhook
                try:
                    mensagem = (
                        f"📋 **Emissão de Taxa GMS Finalizada**\n\n"
                        f"📊 **Quantidade emitida:** {self.linhas_processadas}\n"
                        f"❌ **Com erro:** {self.linhas_com_erro}\n"
                        f"⏭️ **Puladas:** {self.linhas_puladas}\n"
                    )
                    if publicar:
                        mensagem += (
                            f"🌐 **Publicados no Onvio:** {publicados}\n"
                            f"⚠️ **Falhas na publicação:** {falhas_publicacao}\n"
                        )
                    mensagem += (
                        f"📂 **Diretório dos PDFs:** `{pasta_saida}`\n\n"
                        f"✅ Emissão finalizada com sucesso!\n\n"
                        f"<@&1299044385899548752>"
                    )
                    if self.erros_detalhados:
                        tabela = "```\n"
                        tabela += f"{'Nº':<6} {'Empresa':<35} {'Motivo'}\n"
                        tabela += "-" * 80 + "\n"
                        for erro in self.erros_detalhados:
                            num = str(erro['numero'])[:5]
                            empresa = str(erro['empresa'])[:34]
                            motivo = str(erro['motivo'])[:38]
                            tabela += f"{num:<6} {empresa:<35} {motivo}\n"
                        tabela += "```"
                        mensagem += f"\n\n❌ **Empresas com erro:**\n{tabela}"
                    webhook_url = os.getenv("DISCORD_WEBHOOK_URL")
                    requests.post(webhook_url, json={"content": mensagem}, timeout=10)
                    self.adicionar_log("Notificação enviada ao Discord", logging.INFO, "sucesso")
                except Exception as e:
                    self.adicionar_log(f"Erro ao enviar notificação ao Discord: {str(e)}", logging.WARNING, "aviso")

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
        interval = 0.15
        elapsed = 0.0
        while elapsed < seconds:
            if self.should_stop():
                return False
            self.check_pause()
            if self.should_stop():
                return False
            sleep_time = min(interval, seconds - elapsed)
            time.sleep(sleep_time)
            elapsed += sleep_time
        return True

    def wait_for_condition(self, condition_fn, timeout: float = 30, poll_interval: float = 0.15, description: str = "") -> bool:
        """Polls condition_fn() até retornar True, ou timeout.
        Retorna True se condição foi atendida, False se timeout ou stop."""
        start = time.time()
        while time.time() - start < timeout:
            if self.should_stop():
                return False
            self.check_pause()
            try:
                if condition_fn():
                    if description:
                        self.log(f"{description} - concluido em {time.time() - start:.1f}s")
                    return True
            except Exception:
                pass
            time.sleep(poll_interval)
        if description:
            self.log(f"{description} - timeout apos {timeout}s")
        return False

    def _window_exists(self, title: str, class_name: str) -> bool:
        """Verifica se janela com título/classe existe via win32gui (rápido)."""
        try:
            result = [False]
            def callback(hwnd, _):
                if win32gui.IsWindowVisible(hwnd):
                    try:
                        if (win32gui.GetWindowText(hwnd) == title and
                                win32gui.GetClassName(hwnd) == class_name):
                            result[0] = True
                            return False
                    except Exception:
                        pass
                return True
            win32gui.EnumWindows(callback, None)
            return result[0]
        except Exception:
            return False

    def _save_dialog_exists(self) -> bool:
        """Verifica se a janela de salvamento existe procurando pelo elemento 'Salvar em:' (AutomationId 1091)."""
        try:
            result = [False]
            def callback(hwnd, _):
                if win32gui.IsWindowVisible(hwnd):
                    try:
                        if win32gui.GetClassName(hwnd) == "#32770":
                            child = win32gui.FindWindowEx(hwnd, 0, "Static", "Salvar em:")
                            if child:
                                result[0] = True
                                return False
                    except Exception:
                        pass
                return True
            win32gui.EnumWindows(callback, None)
            return result[0]
        except Exception:
            return False

    def _window_exists_partial(self, title_part: str, class_name: str) -> bool:
        """Verifica se janela com título parcial/classe existe via win32gui (rápido)."""
        try:
            result = [False]
            def callback(hwnd, _):
                if win32gui.IsWindowVisible(hwnd):
                    try:
                        if (title_part in win32gui.GetWindowText(hwnd) and
                                win32gui.GetClassName(hwnd) == class_name):
                            result[0] = True
                            return False
                    except Exception:
                        pass
                return True
            win32gui.EnumWindows(callback, None)
            return result[0]
        except Exception:
            return False

    def _any_error_dialog_visible(self) -> bool:
        """Verifica se há diálogo de erro visível via win32gui (rápido)."""
        error_keywords = ("erro", "aviso", "atenção", "alerta", "warning", "error", "informação")
        try:
            result = [False]
            def callback(hwnd, _):
                if win32gui.IsWindowVisible(hwnd):
                    try:
                        cls = win32gui.GetClassName(hwnd)
                        if cls == "#32770":
                            title = win32gui.GetWindowText(hwnd).lower()
                            for kw in error_keywords:
                                if kw in title:
                                    result[0] = True
                                    return False
                    except Exception:
                        pass
                return True
            win32gui.EnumWindows(callback, None)
            return result[0]
        except Exception:
            return False

    def _is_connection_alive(self) -> bool:
        """Verifica se a conexão pywinauto ainda é válida."""
        if self.app is None or self.main_window is None:
            return False
        try:
            hwnd = self.main_window.handle
            if not win32gui.IsWindow(hwnd):
                return False
            win32gui.GetWindowText(hwnd)
            return True
        except Exception:
            return False

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

            time.sleep(0.15)

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
            if not self.smart_sleep(1.5):
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

    def processar_linha(self, codigo: str, empresa_nome: str, periodo: str, linha_excel: int, caminho_pdf: str) -> bool:
        """Processa uma empresa do cadastro"""
        try:
            if self.should_stop():
                return False

            # Só reconectar se a conexão existente estiver quebrada
            if not self._is_connection_alive():
                handle = self.find_dominio_window()
                if not handle:
                    self.log("❌ Não foi possível localizar a janela do Domínio")
                    return False
                try:
                    self.app = Application(backend="uia").connect(handle=handle)
                    self.main_window = self.app.window(handle=handle)
                    self.log("✅ Reconectado ao Domínio com sucesso")
                except Exception as e:
                    self.log(f"❌ Erro ao reconectar: {str(e)}")
                    return False

            handle = self.main_window.handle

            if win32gui.IsIconic(handle):
                win32gui.ShowWindow(handle, win32con.SW_RESTORE)
                if not self.smart_sleep(0.5):
                    return False

            win32gui.SetForegroundWindow(handle)
            time.sleep(0.2)

            # Troca de empresa
            if not self.handle_empresa_change(codigo):
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
            return self.processar_relatorio_taxa_gms(periodo, linha_excel, caminho_pdf)

        except Exception as e:
            self.log(f"❌ Erro ao processar {codigo} - {empresa_nome}: {str(e)}")
            return False

    def processar_relatorio_taxa_gms(self, periodo: str, linha_excel: int, caminho_pdf: str) -> bool:
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
            except Exception as e:
                self.log(f"⚠️ Erro ao clicar em executar, tentando via teclado: {str(e)}")
                send_keys('{F5}')  # Alternativa via teclado

            # Aguardar janela do relatório carregar (título contém "Taxa GMS")
            if not self.wait_for_condition(
                lambda: self._window_exists_partial("Taxa GMS", "FNWND3190") or self._any_error_dialog_visible(),
                timeout=30,
                poll_interval=0.15,
                description="Aguardando relatório carregar"
            ):
                self.log("⚠️ Timeout aguardando relatório carregar")
                return False

            # Verificar se não foi um diálogo de erro
            if self._any_error_dialog_visible():
                if not self.handle_error_dialogs():
                    self.cleanup_windows()
                    return False

            # Gerar PDF
            return self.gerar_pdf(linha_excel, caminho_pdf)

        except Exception as e:
            self.log(f"❌ Erro no processamento do relatório: {str(e)}")
            return False

    def _find_confirmacao_substituir_hwnd(self) -> int:
        """Localiza o hwnd do diálogo de confirmação de substituição de arquivo.
        Reconhece pelo título ("Confirmar Salvar como" / "Salvar como") ou pelo
        texto ("já existe" / "substituir") em janelas de classe #32770."""
        result = [0]

        def cb(hwnd, _):
            if not win32gui.IsWindowVisible(hwnd):
                return True
            try:
                if win32gui.GetClassName(hwnd) != "#32770":
                    return True
                titulo = win32gui.GetWindowText(hwnd).lower()
                # A janela de salvamento em si também é #32770; distinguir pela
                # presença do botão "&Sim" (IDYES) ou pelo título de confirmação.
                if ("confirmar salvar" in titulo or "salvar como" in titulo
                        or "confirmar" in titulo):
                    # Confirmar que existe um botão Sim (evita confundir com a
                    # própria janela "Salvar em PDF")
                    yes_btn = win32gui.FindWindowEx(hwnd, 0, "Button", "&Sim")
                    if not yes_btn:
                        yes_btn = win32gui.FindWindowEx(hwnd, 0, "Button", "Sim")
                    if yes_btn:
                        result[0] = hwnd
                        return False
            except Exception:
                pass
            return True

        try:
            win32gui.EnumWindows(cb, None)
        except Exception:
            pass
        return result[0]

    def _confirmar_substituir_arquivo(self):
        """Se o arquivo já existir, o Windows abre 'Confirmar Salvar como'
        perguntando se deseja substituir. Aguarda brevemente por esse diálogo e
        confirma 'Sim' (sobrescreve). Se ele não aparecer, segue normalmente."""
        inicio = time.time()
        while time.time() - inicio < 3:
            if self.should_stop():
                return
            hwnd = self._find_confirmacao_substituir_hwnd()
            if hwnd:
                self.log("♻️ Arquivo já existe — confirmando substituição")
                try:
                    dlg = self.app.window(handle=hwnd)
                    dlg.set_focus()
                    time.sleep(0.2)
                    # Tenta clicar no botão "Sim" (IDYES = 6)
                    clicado = False
                    for titulo in ("&Sim", "Sim"):
                        try:
                            btn = dlg.child_window(title=titulo, class_name="Button")
                            if btn.exists():
                                btn.click_input()
                                clicado = True
                                break
                        except Exception:
                            continue
                    if not clicado:
                        try:
                            btn = dlg.child_window(auto_id="6", class_name="Button")
                            if btn.exists():
                                btn.click_input()
                                clicado = True
                        except Exception:
                            pass
                    if not clicado:
                        # Último recurso: Enter aciona o botão default ("Sim")
                        send_keys('{ENTER}')
                except Exception:
                    try:
                        win32gui.SetForegroundWindow(hwnd)
                        send_keys('{ENTER}')
                    except Exception:
                        pass
                time.sleep(0.3)
                return
            time.sleep(0.15)

    def gerar_pdf(self, linha_excel: int, caminho_pdf: str) -> bool:
        """Gera e salva o PDF do relatório"""
        try:
            if self.should_stop():
                return False

            # Verificar e tratar janela de erro
            if not self.handle_error_dialogs():
                self.cleanup_windows()
                return False

            self.log("📄 Gerando PDF")

            # Enviar Ctrl+D em loop aguardando a janela de salvamento
            self.log("📄 Aguardando janela de salvamento (enviando Ctrl+D periodicamente)...")
            timeout_total = 60
            intervalo_ctrl_d = 5
            inicio = time.time()
            janela_encontrada = False

            while time.time() - inicio < timeout_total:
                if self.should_stop():
                    return False

                # Verificar diálogos de erro/aviso
                if self._any_error_dialog_visible():
                    if not self.handle_error_dialogs():
                        self.cleanup_windows()
                        return False
                    # Aviso não crítico tratado, continua aguardando
                    self.log("🔄 Aviso tratado, continuando aguardo da janela de salvamento...")
                    time.sleep(0.5)
                    continue

                if self._window_exists("Salvar em PDF", "#32770") or self._save_dialog_exists():
                    janela_encontrada = True
                    break

                # Garantir foco e enviar Ctrl+D
                try:
                    self.main_window.set_focus()
                    time.sleep(0.3)
                except Exception:
                    pass
                send_keys('^d')  # Ctrl+D

                # Aguardar entre tentativas verificando a cada 0.25s
                for _ in range(int(intervalo_ctrl_d / 0.25)):
                    if self.should_stop():
                        return False
                    if self._any_error_dialog_visible():
                        break
                    if self._window_exists("Salvar em PDF", "#32770") or self._save_dialog_exists():
                        janela_encontrada = True
                        break
                    time.sleep(0.25)
                if janela_encontrada:
                    break

            if not janela_encontrada:
                self.log("❌ Timeout aguardando janela de salvamento após Ctrl+D")
                return False

            # Localizar janela de salvamento
            self.log("💾 Configurando salvamento do PDF")

            try:
                save_window = self.main_window.child_window(
                    title="Salvar em PDF",
                    class_name="#32770"
                )

                if not save_window.exists():
                    # Fallback: procura janela de salvamento pelo elemento "Salvar em:" (AutomationId 1091)
                    self.log("🔍 Procurando janela de salvamento alternativa...")
                    try:
                        save_window = self.main_window.child_window(
                            class_name="#32770",
                            found_index=0
                        )
                        salvar_em_label = save_window.child_window(
                            auto_id="1091",
                            class_name="Static"
                        )
                        if not salvar_em_label.exists():
                            self.log("❌ Janela de salvamento não encontrada")
                            return False
                        self.log("✅ Janela de salvamento encontrada via elemento 'Salvar em:'")
                    except Exception:
                        self.log("❌ Janela de salvamento não encontrada")
                        return False

                if self.should_stop():
                    return False
                self.check_pause()

                # Preenche o caminho completo (pasta + nome) via clipboard, evitando
                # a navegação manual pela árvore de pastas do diálogo "Salvar em PDF"
                self.log(f"📝 Salvando em: {caminho_pdf}")
                save_hwnd = save_window.handle
                self._force_focus(save_hwnd)
                self._set_clipboard(caminho_pdf)

                name_field = save_window.child_window(auto_id="1148", class_name="Edit")
                name_field.set_focus()
                time.sleep(0.1)
                send_keys('^a')
                time.sleep(0.1)
                send_keys('^v')
                time.sleep(0.3)

                if self.should_stop():
                    return False
                self.check_pause()

                # Salvar
                self.log("💾 Salvando PDF")
                button_salvar = save_window.child_window(auto_id="1", class_name="Button")
                button_salvar.click_input()

                # Se o arquivo já existir, o Windows abre "Confirmar Salvar como"
                # perguntando se deseja substituir — confirmar "Sim" e sobrescrever.
                self._confirmar_substituir_arquivo()

                # Esperar janela de salvamento fechar. O diálogo de substituição
                # pode aparecer com atraso; por isso re-checamos dentro do laço.
                inicio_espera = time.time()
                salvou = False
                while time.time() - inicio_espera < 15:
                    if self.should_stop():
                        return False
                    self.check_pause()
                    try:
                        if not save_window.exists() or not save_window.is_visible():
                            salvou = True
                            break
                    except Exception:
                        salvou = True
                        break
                    # Trata diálogo de substituição que possa ter surgido com atraso
                    if self._find_confirmacao_substituir_hwnd():
                        self._confirmar_substituir_arquivo()
                    time.sleep(0.2)

                if not salvou:
                    self.log("⚠️ Timeout aguardando salvamento do PDF")
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
        """Trata diálogos de erro que podem aparecer.
        Retorna True se deve continuar, False se deve abortar.
        Otimizado: usa win32gui.EnumWindows (uma única passagem) em vez de múltiplas buscas UIA."""
        try:
            error_titles_lower = {"erro", "erro léxico", "aviso", "atenção",
                                  "informação", "alerta", "warning", "error"}

            # Passagem única: enumerar todas as janelas via Win32 API (rápido)
            found_hwnd = None
            found_title = None

            def enum_callback(hwnd, _):
                nonlocal found_hwnd, found_title
                if not win32gui.IsWindowVisible(hwnd):
                    return True
                try:
                    title = win32gui.GetWindowText(hwnd)
                    if not title:
                        return True
                    title_lower = title.strip().lower()
                    for err_title in error_titles_lower:
                        if title_lower == err_title or err_title in title_lower:
                            if win32gui.GetClassName(hwnd) == "#32770":
                                found_hwnd = hwnd
                                found_title = title
                                return False
                except Exception:
                    pass
                return True

            win32gui.EnumWindows(enum_callback, None)

            if found_hwnd is None:
                return True  # Nenhum diálogo de erro, continuar normalmente

            # Encontrou diálogo — agora usar pywinauto apenas para esta janela específica
            try:
                error_window = self.app.window(handle=found_hwnd)
            except Exception:
                win32gui.SetForegroundWindow(found_hwnd)
                send_keys('{ENTER}')
                time.sleep(0.3)
                return True

            # Ler texto da mensagem
            message = ""
            try:
                message = error_window.window_text()
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

            self.log(f"⚠️ Diálogo detectado: '{found_title}' - {message[:100] if message else 'sem mensagem'}")

            message_lower = message.lower()

            # Mensagens de aviso que não impedem o salvamento (fecha e continua o fluxo)
            mensagens_continuar_salvamento = [
                "erro na gravação do relatório",
                "nome do caminho inválido",
                "caminho inválido",
                "caminho invalido",
                "caracteres não permitidos"
            ]

            for msg in mensagens_continuar_salvamento:
                if msg in message_lower:
                    self.log(f"⚠️ Aviso de gravação detectado (não crítico): {msg}")
                    error_window.set_focus()
                    send_keys('{ENTER}')
                    time.sleep(0.5)
                    return True  # Continua o fluxo — janela de salvamento vai abrir após fechar este aviso

            # Verificar mensagens que abortam a linha (sem dados)
            mensagens_abortar = [
                "sem dados para emitir",
                "nenhum registro encontrado",
                "não há dados",
                "registro não encontrado"
            ]

            for msg in mensagens_abortar:
                if msg in message_lower:
                    self.log(f"⚠️ Aviso não crítico: {msg}")
                    error_window.set_focus()
                    send_keys('{ENTER}')
                    time.sleep(0.5)
                    for _ in range(4):
                        send_keys('{ESC}')
                        time.sleep(0.5)
                    return False

            # Erro léxico — fechar e continuar
            if "léxico" in found_title.lower():
                self.log("⚠️ Erro léxico detectado, fechando...")
                error_window.set_focus()
                for _ in range(3):
                    send_keys('{ESC}')
                    time.sleep(0.5)
                return True

            # Erro genérico: tentar OK, depois ENTER, depois ESC
            self.log(f"⚠️ Fechando diálogo '{found_title}'...")
            error_window.set_focus()
            time.sleep(0.2)

            try:
                ok_button = error_window.child_window(title="OK", class_name="Button")
                if ok_button.exists():
                    ok_button.click_input()
                    time.sleep(0.5)
                    if found_title.lower() in ("erro", "aviso"):
                        return False
                    return True
            except Exception:
                pass

            send_keys('{ENTER}')
            time.sleep(0.5)

            try:
                if error_window.exists():
                    send_keys('{ESC}')
                    time.sleep(0.3)
            except Exception:
                pass

            if found_title.lower() in ("erro", "aviso"):
                return False

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
                time.sleep(0.5)

            # Verificar se o Gerenciador de Relatórios ainda está aberto
            try:
                relatorio_window = self.main_window.child_window(
                    title="Gerenciador de Relatórios",
                    class_name="FNWND3190"
                )

                if relatorio_window.exists() and relatorio_window.is_visible():
                    self.log("🔄 Fechando Gerenciador de Relatórios restante")
                    send_keys('{ESC}')
                    time.sleep(0.5)
            except Exception:
                pass

        except Exception as e:
            self.log(f"⚠️ Erro durante limpeza: {str(e)}")

    # ── Publicação em lote (Publicação de Documentos Externos) ────────────────
    # Portado de DomBot_Taxas/taxa_panificacao.py, adaptado para receber a lista
    # de documentos (código, caminho) já montada em memória durante a emissão.

    PASTA_PUBLICACAO = "Pessoal/GMS"
    PUB_LOTE_TITULO = "Publicação de Documentos Externos"
    PUB_LOTE_CLASSE = "FNWND3190"

    def _set_clipboard(self, text: str):
        """Coloca texto no clipboard do Windows via win32clipboard."""
        import win32clipboard
        win32clipboard.OpenClipboard()
        try:
            win32clipboard.EmptyClipboard()
            win32clipboard.SetClipboardText(text, win32clipboard.CF_UNICODETEXT)
        finally:
            win32clipboard.CloseClipboard()

    def _force_focus(self, hwnd: int):
        """Força foco para uma janela contornando a restrição do Windows 10."""
        try:
            ctypes.windll.user32.keybd_event(0x12, 0, 0, 0)
            ctypes.windll.user32.keybd_event(0x12, 0, 2, 0)
            win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
            win32gui.SetForegroundWindow(hwnd)
            time.sleep(0.3)
        except Exception:
            pass

    def _get_pub_lote_window(self):
        """Retorna o objeto pywinauto da janela 'Publicação de Documentos Externos'."""
        return self.main_window.child_window(
            title=self.PUB_LOTE_TITULO, class_name=self.PUB_LOTE_CLASSE
        )

    def _pub_lote_window_ok(self) -> bool:
        try:
            pub = self._get_pub_lote_window()
            return pub.exists() and pub.is_visible()
        except Exception:
            return False

    def _abrir_janela_pub_lote(self) -> bool:
        """Garante a janela de publicação em lote aberta (abre pelo botão-nuvem)."""
        if self._pub_lote_window_ok():
            self.log("📋 Janela de Publicação em Lote já está aberta")
            return True

        for tentativa in range(3):
            if self.should_stop():
                return False
            try:
                self.main_window.set_focus()
                btn_nuvem = self.main_window.child_window(
                    auto_id="picturePublicacaoDocumentosExternos"
                )
                if btn_nuvem.exists(timeout=2):
                    btn_nuvem.click_input()
                    self.log("☁️ Botão de publicação em lote clicado")
            except Exception as e:
                self.log(f"⚠️ Tentativa {tentativa + 1} de clicar no botão-nuvem falhou: {e}")

            if self.wait_for_condition(
                self._pub_lote_window_ok,
                timeout=10, poll_interval=0.3,
                description="Aguardando janela de Publicação em Lote",
            ):
                return True
        return False

    def _garantir_checkbox(self, pub, auto_id: str, nome: str) -> bool:
        """Garante que um checkbox esteja marcado (ToggleState On)."""
        try:
            chk = pub.child_window(auto_id=auto_id, class_name="Button")
            if not chk.exists(timeout=2):
                self.log(f"⚠️ Checkbox '{nome}' não encontrado")
                return False
            for _ in range(2):
                try:
                    estado = chk.get_toggle_state()
                except Exception:
                    estado = None
                if estado == 1:
                    return True
                self.log(f"☑ Marcando '{nome}'")
                try:
                    chk.click_input()
                except Exception:
                    chk.click()
                self.smart_sleep(0.3)
            try:
                return chk.get_toggle_state() == 1
            except Exception:
                return True
        except Exception as e:
            self.log(f"⚠️ Não foi possível marcar '{nome}': {e}")
            return False

    def _preencher_data_mascarada(self, edit_data, digitos: str, data_fmt: str) -> bool:
        """Preenche um campo de data com máscara (00/00/0000)."""
        def _ler():
            try:
                v = (edit_data.get_value() or "").strip()
            except Exception:
                try:
                    v = (edit_data.window_text() or "").strip()
                except Exception:
                    v = ""
            return v

        def _bate(valor):
            return ''.join(c for c in valor if c.isdigit()) == digitos

        try:
            edit_data.set_focus()
            time.sleep(0.1)
            edit_data.type_keys("{HOME}{LEFT 12}", set_foreground=False)
            time.sleep(0.1)
            edit_data.type_keys(digitos, set_foreground=False)
            time.sleep(0.2)
            if _bate(_ler()):
                return True
        except Exception as e:
            self.log(f"⚠️ Data (tentativa 1) falhou: {e}")

        try:
            edit_data.set_focus()
            edit_data.type_keys("{HOME}{LEFT 12}{DELETE 12}", set_foreground=False)
            time.sleep(0.1)
            try:
                edit_data.set_text(data_fmt)
            except Exception:
                edit_data.type_keys(digitos, set_foreground=False)
            time.sleep(0.2)
            valor = _ler()
            if _bate(valor):
                return True
            self.log(f"⚠️ Data lida do campo: '{valor}' (esperado {data_fmt})")
        except Exception as e:
            self.log(f"⚠️ Data (tentativa 2) falhou: {e}")

        return False

    def _configurar_envio_lote(self, pub, data_vencimento: str) -> bool:
        """Configuração feita uma vez: pasta, data de vencimento e 'Concluir atividade'."""
        try:
            self.log(f"📁 Configurando pasta: {self.PASTA_PUBLICACAO}")
            try:
                combo = pub.child_window(auto_id="1001", class_name="ComboBox")
                combo.set_focus()
                selecionado = False
                try:
                    combo.select(self.PASTA_PUBLICACAO)
                    selecionado = True
                except Exception:
                    try:
                        itens = combo.item_texts()
                        alvo = self.PASTA_PUBLICACAO.strip().lower()
                        for i, txt in enumerate(itens):
                            t = (txt or "").strip().lower()
                            if t == alvo or alvo in t:
                                combo.select(i)
                                selecionado = True
                                self.log(f"📁 Pasta selecionada da lista: '{txt}'")
                                break
                        if not selecionado:
                            self.log(f"⚠️ '{self.PASTA_PUBLICACAO}' não está na lista. Itens: {itens}")
                    except Exception as e2:
                        self.log(f"⚠️ Não foi possível ler os itens da lista de pastas: {e2}")
                if not selecionado:
                    try:
                        combo.set_edit_text(self.PASTA_PUBLICACAO)
                    except Exception:
                        pass
            except Exception as e:
                self.log(f"⚠️ Não foi possível definir a pasta: {e}")
            self.smart_sleep(0.4)

            self._garantir_checkbox(pub, "1006", "Data de vencimento")
            self.smart_sleep(0.3)

            self.log(f"📅 Definindo data de vencimento: {data_vencimento}")
            digitos = ''.join(ch for ch in data_vencimento if ch.isdigit())
            try:
                edit_data = pub.child_window(auto_id="1005", class_name="PBEDIT190")
                if not self._preencher_data_mascarada(edit_data, digitos, data_vencimento):
                    self.log("⚠️ Data pode ter ficado incorreta no campo")
                self.smart_sleep(0.3)
            except Exception as e:
                self.log(f"⚠️ Não foi possível definir a data de vencimento: {e}")

            self._garantir_checkbox(pub, "1004", "Concluir atividade")

            return True
        except Exception as e:
            self.log(f"❌ Erro ao configurar envio em lote: {e}")
            return False

    def _publicar_um_documento(self, pub, caminho_pdf: str, codigo: str) -> bool:
        """Publica um único documento na janela já aberta e confirma o OK."""
        try:
            self.log(f"📄 Caminho: {os.path.basename(caminho_pdf)}")
            try:
                campo_caminho = pub.child_window(auto_id="1013", class_name="Edit")
                if not campo_caminho.exists(timeout=3):
                    self.log("❌ Campo 'Caminho' não encontrado")
                    return False
                campo_caminho.set_focus()
                campo_caminho.type_keys("^a{DELETE}", set_foreground=False)
                time.sleep(0.3)
                campo_caminho.set_text(caminho_pdf)
            except Exception as e:
                self.log(f"❌ Não foi possível preencher o caminho: {e}")
                return False
            self.smart_sleep(0.5)

            self.log(f"🏢 Código da empresa: {codigo}")
            try:
                campo_codigo = pub.child_window(auto_id="1001", class_name="PBEDIT190")
                if not campo_codigo.exists(timeout=3):
                    self.log("❌ Campo 'Código' não encontrado")
                    return False
                campo_codigo.set_focus()
                campo_codigo.type_keys("^a{DELETE}", set_foreground=False)
                time.sleep(0.3)
                campo_codigo.set_text(codigo)
            except Exception as e:
                self.log(f"❌ Não foi possível preencher o código da empresa: {e}")
                return False
            self.smart_sleep(0.5)

            self.log("⚡ Publicando documento")
            try:
                botao_publicar = pub.child_window(auto_id="1003", class_name="Button")
                if not botao_publicar.exists(timeout=3):
                    self.log("❌ Botão 'Publicar' não encontrado")
                    return False
                botao_publicar.click()
                time.sleep(2)
            except Exception as e:
                self.log(f"❌ Erro ao clicar em 'Publicar': {e}")
                return False

            dialog = self._aguardar_confirmacao(timeout=15)
            if dialog is False:
                return False
            if dialog:
                if self._clicar_botao_ok(dialog):
                    time.sleep(1)
                    return True
                self.log("⚠️ Falha ao clicar no OK de confirmação")
                return False
            else:
                self.log("⚠️ Janela de confirmação não encontrada")
                return False

        except Exception as e:
            self.log(f"❌ Erro ao publicar documento {os.path.basename(caminho_pdf)}: {e}")
            return False

    def _aguardar_confirmacao(self, timeout=15):
        """Aguarda o diálogo de confirmação após 'Publicar'."""
        self.log("🔍 Procurando janela de confirmação...")
        inicio = time.time()
        while (time.time() - inicio) < timeout:
            if self.should_stop():
                self.log("⏹️ Busca por confirmação interrompida")
                return False
            self.check_pause()
            try:
                all_windows = findwindows.find_windows()
                for hwnd in all_windows:
                    try:
                        window = self.app.window(handle=hwnd)
                        if window.is_dialog() and window.is_visible():
                            titulo = window.window_text()
                            if titulo and any(p in titulo.lower() for p in
                                              ['atenção', 'confirmação', 'aviso', 'informação', 'sucesso']):
                                self.log(f"✅ Confirmação encontrada: '{titulo}'")
                                return window
                    except Exception:
                        continue
            except Exception:
                pass
            time.sleep(0.5)
        self.log("⚠️ Timeout: nenhuma janela de confirmação encontrada")
        return None

    def _clicar_botao_ok(self, dialog) -> bool:
        """Clica no OK/Confirmar/Sim do diálogo."""
        for texto in ["OK", "Ok", "Confirmar", "Sim", "Yes"]:
            try:
                botao = dialog.child_window(title=texto, control_type="Button")
                if botao.exists(timeout=2):
                    botao.click()
                    self.log(f"✅ Botão '{texto}' clicado")
                    return True
            except Exception:
                continue
        for auto_id in ["1", "2", "6", "1001", "2001"]:
            try:
                botao = dialog.child_window(auto_id=auto_id, control_type="Button")
                if botao.exists(timeout=2):
                    botao.click()
                    self.log(f"✅ Botão auto_id '{auto_id}' clicado")
                    return True
            except Exception:
                continue
        try:
            botoes = dialog.children(control_type="Button")
            if botoes:
                botoes[0].click()
                self.log("✅ Primeiro botão do diálogo clicado")
                return True
        except Exception:
            pass
        return False

    def publicar_lote_gms(self, documentos: list, data_vencimento: str) -> tuple:
        """
        Publica em lote os documentos informados na janela 'Publicação de
        Documentos Externos'. `documentos` é uma lista de (codigo, caminho_pdf)
        já montada em memória durante a emissão (não depende do nome do arquivo).
        A configuração de envio (pasta/data/concluir) é feita uma única vez.
        Retorna (publicados, falhas).
        """
        publicados = 0
        falhas = 0
        try:
            if not documentos:
                self.log("⚠️ Nenhum documento informado para publicar")
                return (0, 0)

            self.log(f"🌐 Iniciando publicação em lote: {len(documentos)} documento(s)")

            if not self._abrir_janela_pub_lote():
                self.log("❌ Não foi possível abrir a janela de Publicação em Lote")
                return (0, len(documentos))

            pub = self._get_pub_lote_window()
            try:
                pub.set_focus()
            except Exception:
                pass

            if not self._configurar_envio_lote(pub, data_vencimento):
                self.log("❌ Falha ao configurar o envio em lote")
                self.cleanup_windows()
                return (0, len(documentos))

            for codigo, caminho_pdf in documentos:
                if self.should_stop():
                    self.log("Publicação em lote interrompida pelo usuário")
                    break
                self.check_pause()

                if not os.path.exists(caminho_pdf):
                    self.log(f"⚠️ PDF não encontrado, pulando: {caminho_pdf}")
                    falhas += 1
                    continue

                nome_base = os.path.basename(caminho_pdf)
                if self._publicar_um_documento(pub, caminho_pdf, codigo):
                    publicados += 1
                    self.log(f"✅ Publicado: {nome_base}")
                else:
                    falhas += 1
                    self.log(f"❌ Falha ao publicar: {nome_base}")

            self.log(f"🌐 Publicação em lote concluída: {publicados} publicado(s), {falhas} falha(s)")
            self.cleanup_windows()
            return (publicados, falhas)

        except Exception as e:
            self.log(f"❌ Erro na publicação em lote: {e}\n{traceback.format_exc()}")
            try:
                self.cleanup_windows()
            except Exception:
                pass
            return (publicados, falhas)


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
