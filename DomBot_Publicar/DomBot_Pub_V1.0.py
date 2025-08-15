import customtkinter as ctk
from tkinter import filedialog, scrolledtext, messagebox
import threading
from PIL import Image, ImageTk  # Para carregar PNG como logo
import os
import time
import pandas as pd
from pywinauto import Application, findwindows
from pywinauto.findwindows import ElementNotFoundError

class DomBot:
    def __init__(self, log_callback=None, progress_callback=None, ui_reference=None):
        self.log_callback = log_callback or print
        self.progress_callback = progress_callback
        self.ui_reference = ui_reference  # Referência para verificar is_running
        try:
            self.app = Application(backend="uia").connect(
                title="Domínio Folha - Versão: 10.5A-07 - 08",
                class_name="FNWND3190",
                timeout=10
            )
            self.main_window = self.app.window(
                title="Domínio Folha - Versão: 10.5A-07 - 08",
                class_name="FNWND3190"
            )
            self.main_window.set_focus()
            self.log("✅ Conectado à janela principal do Domínio Folha")
        except Exception as e:
            self.log(f"❌ Erro ao conectar à janela principal: {str(e)}")
            raise

    def log(self, mensagem):
        if callable(self.log_callback):
            self.log_callback(mensagem)
        # Opcional: salvar logs em arquivo para depuração
        with open("publicacao_log.txt", "a", encoding="utf-8") as f:
            f.write(f"{time.strftime('%Y-%m-%d %H:%M:%S')} - {mensagem}\n")

    def update_progress(self, current, total, status=""):
        if callable(self.progress_callback):
            progress = (current / total) * 100 if total > 0 else 0
            self.progress_callback(progress, f"{current}/{total} - {status}")

    def aguardar_e_encontrar_janela_confirmacao_interruptivel(self, timeout=15):
        """Versão interruptível da função de aguardar confirmação"""
        self.log("🔍 Procurando janela de confirmação...")
        titulos_possiveis = ["Atenção", "Confirmação", "Aviso", "Informação", "Sucesso"]
        classes_possiveis = ["#32770", "Dialog", "FNWND3190"]
        
        inicio = time.time()
        while (time.time() - inicio) < timeout:
            # ✅ VERIFICAÇÃO DE INTERRUPÇÃO durante a espera
            if self.ui_reference and not self.ui_reference.is_running:
                self.log("⏹️ Busca por janela de confirmação interrompida")
                return False  # False indica interrupção
                
            try:
                # Estratégia 3: Buscar janelas do sistema
                try:
                    all_windows = findwindows.find_windows()
                    for hwnd in all_windows:
                        try:
                            window = self.app.window(handle=hwnd)
                            if window.is_dialog() and window.is_visible():
                                titulo = window.window_text()
                                if titulo and any(palavra in titulo.lower() for palavra in ['atenção', 'confirmação', 'aviso']):
                                    self.log(f"✅ Janela do sistema encontrada: '{titulo}'")
                                    return window
                        except:
                            continue
                except:
                    pass
            except Exception as e:
                self.log(f"🔍 Erro durante busca: {str(e)}")
            time.sleep(0.5)
        
        self.log("⚠️ Timeout: Nenhuma janela de confirmação encontrada")
        return None

    def clicar_botao_ok(self, dialog):
        textos_botao = ["OK", "Ok", "Confirmar", "Sim", "Yes"]
        auto_ids = ["1", "2", "6", "1001", "2001"]
        
        for texto in textos_botao:
            try:
                botao = dialog.child_window(title=texto, control_type="Button")
                if botao.exists(timeout=2):
                    botao.click()
                    self.log(f"✅ Botão '{texto}' clicado com sucesso")
                    return True
            except:
                continue
        
        for auto_id in auto_ids:
            try:
                botao = dialog.child_window(auto_id=auto_id, control_type="Button")
                if botao.exists(timeout=2):
                    botao.click()
                    self.log(f"✅ Botão com auto_id '{auto_id}' clicado com sucesso")
                    return True
            except:
                continue
        
        try:
            botoes = dialog.children(control_type="Button")
            if botoes:
                botoes[0].click()
                self.log("✅ Primeiro botão encontrado foi clicado")
                return True
        except:
            pass
        
        self.log("🔍 Debugando controles da janela:")
        try:
            dialog.print_control_identifiers()
        except:
            self.log("❌ Não foi possível imprimir controles")
        return False

    def ler_excel(self, caminho_arquivo):
        try:
            df = pd.read_excel(caminho_arquivo)
            self.log(f"📊 Arquivo contém {len(df)} linhas de dados")
            colunas_necessarias = ['Nº', 'Periodo', 'Salvar Como', 'Caminho']
            colunas_faltando = [c for c in colunas_necessarias if c not in df.columns]
            if colunas_faltando:
                self.log(f"⚠️ Colunas obrigatórias não encontradas: {', '.join(colunas_faltando)}")
                return None
            self.log("✅ Todas as colunas obrigatórias encontradas")
            return df
        except FileNotFoundError:
            self.log(f"❌ Arquivo não encontrado: {caminho_arquivo}")
            return None
        except Exception as e:
            self.log(f"❌ Erro ao ler arquivo: {str(e)}")
            return None

    def publicar_documentos(self, caminho_excel):
        df = self.ler_excel(caminho_excel)
        if df is None:
            self.log("❌ Não foi possível prosseguir devido a erro na leitura do Excel")
            return False

        total_documentos = len(df)
        documentos_processados = 0
        documentos_sucesso = 0

        try:
            self.main_window.set_focus()
            self.log("✅ Foco definido na janela principal")
            pub_window = self.main_window.child_window(
                title="Publicação de Documentos Externos",
                class_name="FNWND3190"
            )

            if not pub_window.exists() or not pub_window.is_visible():
                self.log("❌ Janela de Publicação de Documentos Externos não encontrada ou não visível")
                return False

            self.log("✅ Janela de Publicação de Documentos Externos encontrada")
            pub_window.set_focus()

            for index, row in df.iterrows():
                # ✅ VERIFICAÇÃO 1: Início de cada documento
                if self.ui_reference and not self.ui_reference.is_running:
                    self.log("⏹️ Processo interrompido pelo usuário")
                    break
                    
                documentos_processados += 1
                
                caminho_pdf = str(row['Caminho'])
                numero = str(row['Nº']) if pd.notnull(row['Nº']) else ""
                salvar_como = str(row['Salvar Como']) if pd.notnull(row['Salvar Como']) else ""

                self.update_progress(documentos_processados, total_documentos, f"Processando: {salvar_como}")

                if not os.path.exists(caminho_pdf):
                    self.log(f"⚠️ Arquivo PDF não encontrado: {caminho_pdf}")
                    continue
                if not numero:
                    self.log(f"⚠️ Valor inválido na coluna 'Nº' para a linha {index + 2}")
                    continue
                if not salvar_como:
                    self.log(f"⚠️ Valor inválido na coluna 'Salvar Como' para a linha {index + 2}")
                    continue

                self.log(f"📂 Processando linha {index + 1}: {salvar_como}")

                try:
                    # ✅ VERIFICAÇÃO 2: Antes de preencher campos
                    if self.ui_reference and not self.ui_reference.is_running:
                        self.log("⏹️ Processo interrompido pelo usuário")
                        break
                        
                    campo_caminho = pub_window.child_window(auto_id="1013", class_name="Edit")
                    if campo_caminho.exists(timeout=3):
                        campo_caminho.set_focus()
                        campo_caminho.type_keys("^a{DELETE}")
                        time.sleep(0.3)
                        campo_caminho.set_text(caminho_pdf)
                        self.log(f"✅ Caminho preenchido: {caminho_pdf}")
                    else:
                        self.log("❌ Campo 'Caminho' não encontrado")
                        continue

                    time.sleep(0.5)

                    # ✅ VERIFICAÇÃO 3: Antes do segundo campo
                    if self.ui_reference and not self.ui_reference.is_running:
                        self.log("⏹️ Processo interrompido pelo usuário")
                        break

                    campo_numero = pub_window.child_window(auto_id="1001", class_name="PBEDIT190")
                    if campo_numero.exists(timeout=3):
                        campo_numero.set_focus()
                        campo_numero.type_keys("^a{DELETE}")
                        time.sleep(0.3)
                        campo_numero.set_text(numero)
                        self.log(f"✅ Número preenchido: {numero}")
                    else:
                        self.log("❌ Campo 'Número' não encontrado")
                        continue

                    time.sleep(0.5)

                    # ✅ VERIFICAÇÃO 4: Antes de clicar em Publicar
                    if self.ui_reference and not self.ui_reference.is_running:
                        self.log("⏹️ Processo interrompido pelo usuário")
                        break

                    botao_publicar = pub_window.child_window(auto_id="1003", class_name="Button")
                    if botao_publicar.exists(timeout=3):
                        self.log("📤 Clicando no botão 'Publicar'...")
                        botao_publicar.click()
                        time.sleep(2)
                    else:
                        self.log("❌ Botão 'Publicar' não encontrado")
                        continue

                    # ✅ VERIFICAÇÃO 5: Antes de aguardar confirmação
                    if self.ui_reference and not self.ui_reference.is_running:
                        self.log("⏹️ Processo interrompido pelo usuário")
                        break

                    dialog = self.aguardar_e_encontrar_janela_confirmacao_interruptivel(timeout=15)
                    if dialog:
                        if self.clicar_botao_ok(dialog):
                            self.log(f"✅ Documento '{salvar_como}' publicado com sucesso")
                            documentos_sucesso += 1
                            time.sleep(1)
                        else:
                            self.log(f"❌ Falha ao clicar no botão OK para '{salvar_como}'")
                            continue
                    elif dialog is False:  # Interrompido
                        self.log("⏹️ Processo interrompido durante espera de confirmação")
                        break
                    else:
                        self.log(f"⚠️ Janela de confirmação não encontrada para '{salvar_como}'")
                        continue

                except ElementNotFoundError as e:
                    self.log(f"⚠️ Elemento não encontrado para {salvar_como}: {str(e)}")
                    continue
                except Exception as e:
                    self.log(f"⚠️ Erro ao processar {salvar_como}: {str(e)}")
                    continue

            # Verifica se foi interrompido ou concluído
            if self.ui_reference and not self.ui_reference.is_running:
                self.update_progress(documentos_processados, total_documentos, "Interrompido pelo usuário")
                self.log(f"⏹️ Processo interrompido! {documentos_sucesso}/{documentos_processados} documentos publicados.")
            else:
                self.update_progress(total_documentos, total_documentos, "Concluído!")
                self.log(f"🎉 Processamento concluído! {documentos_sucesso}/{total_documentos} documentos publicados com sucesso.")
            return True

        except Exception as e:
            self.log(f"❌ Erro na automação: {str(e)}")
            return False

class AppUI(ctk.CTk):
  
    def __init__(self):
        super().__init__()
        self.title("DomBot_Pub - Publicador Domínio Folha")
        self.geometry("520x420")
        self.resizable(False, False)
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("green")

        # Define o ícone da janela (apenas no executável ou quando rodar local)
        try:
            self.iconbitmap("./assets/DomBot_Pub.ico")
        except Exception as e:
            print(f"Não foi possível carregar ícone: {e}")

        self.is_running = False
        self.setup_ui()

    def setup_ui(self):
        main_frame = ctk.CTkFrame(self)
        main_frame.pack(fill="both", expand=True, padx=10, pady=10)

        try:
            img_logo = Image.open("./assets/DomBot_Pub.png")
            img_logo = img_logo.resize((64, 64))
            self.logo_ctk = ctk.CTkImage(light_image=img_logo, dark_image=img_logo, size=(64, 64))

            # Frame para colocar logo e texto lado a lado
            top_frame = ctk.CTkFrame(main_frame)
            top_frame.pack(pady=(10, 5))

            # Logo
            logo_label = ctk.CTkLabel(top_frame, image=self.logo_ctk, text="")
            logo_label.pack(side="left", padx=(0, 10))  # Espaço entre logo e texto

            # Texto ao lado
            text_label = ctk.CTkLabel(top_frame, text="DomBot_Pub", font=("Arial", 18))
            text_label.pack(side="left")

        except Exception as e:
            print(f"Não foi possível carregar logo: {e}")

        # title_label = ctk.CTkLabel(main_frame, text="DomBot_Pub",
        #                            font=ctk.CTkFont(size=18, weight="bold"))
        # title_label.pack(pady=(0, 20))

        # Frame de seleção de arquivo
        file_frame = ctk.CTkFrame(main_frame)
        file_frame.pack(fill="x", padx=10, pady=5)

        self.excel_path = ctk.StringVar(value="")
        self.btn_select = ctk.CTkButton(file_frame, text="📁 Selecionar Excel", 
                                        command=self.select_file, width=140)
        self.btn_select.pack(side="left", padx=10, pady=10)

        self.lbl_path = ctk.CTkLabel(file_frame, textvariable=self.excel_path, 
                                     wraplength=340, anchor="w")
        self.lbl_path.pack(side="left", fill="x", expand=True, padx=(0, 10), pady=10)

        # Frame de controles
        control_frame = ctk.CTkFrame(main_frame)
        control_frame.pack(fill="x", padx=10, pady=5)

        self.btn_run = ctk.CTkButton(control_frame, text="🚀 Publicar", 
                                     command=self.run_bot, width=120)
        self.btn_run.pack(side="left", padx=10, pady=10)

        self.btn_stop = ctk.CTkButton(control_frame, text="⏹️ Parar", 
                                      command=self.stop_bot, width=80, 
                                      fg_color="red", hover_color="darkred", state="disabled")
        self.btn_stop.pack(side="left", padx=(0, 10), pady=10)

        # Validação de arquivo
        self.btn_validate = ctk.CTkButton(control_frame, text="✅ Validar Excel", 
                                          command=self.validate_file, width=120)
        self.btn_validate.pack(side="right", padx=10, pady=10)

        # Frame de progresso
        progress_frame = ctk.CTkFrame(main_frame)
        progress_frame.pack(fill="x", padx=10, pady=5)

        self.progress_bar = ctk.CTkProgressBar(progress_frame)
        self.progress_bar.pack(fill="x", padx=10, pady=(10, 5))
        self.progress_bar.set(0)

        self.progress_label = ctk.CTkLabel(progress_frame, text="Pronto para iniciar")
        self.progress_label.pack(pady=(0, 10))

        # Log
        log_frame = ctk.CTkFrame(main_frame)
        log_frame.pack(fill="both", expand=True, padx=10, pady=(5, 10))

        log_title = ctk.CTkLabel(log_frame, text="Log de Execução", 
                                 font=ctk.CTkFont(size=12, weight="bold"))
        log_title.pack(pady=(10, 5))

        self.txt_log = scrolledtext.ScrolledText(log_frame, height=8, wrap="word", 
                                                 state="disabled", bg="#2b2b2b", 
                                                 fg="white", insertbackground="white")
        self.txt_log.pack(fill="both", expand=True, padx=10, pady=(0, 10))

    def log_message(self, msg):
        self.txt_log.config(state="normal")
        self.txt_log.insert("end", f"{time.strftime('%H:%M:%S')} - {msg}\n")
        self.txt_log.see("end")
        self.txt_log.config(state="disabled")
        self.update_idletasks()

    def update_progress(self, progress, status):
        self.progress_bar.set(progress / 100)
        self.progress_label.configure(text=status)
        self.update_idletasks()

    def select_file(self):
        file_path = filedialog.askopenfilename(
            title="Selecione o arquivo Excel",
            filetypes=[("Excel Files", "*.xlsx"), ("Excel Files", "*.xls")]
        )
        if file_path:
            self.excel_path.set(os.path.basename(file_path))  # Mostra apenas o nome
            self.full_path = file_path  # Guarda o caminho completo
            self.log_message(f"📄 Arquivo selecionado: {os.path.basename(file_path)}")

    def validate_file(self):
        if not hasattr(self, 'full_path'):
            messagebox.showwarning("Aviso", "Selecione um arquivo Excel primeiro.")
            return

        try:
            df = pd.read_excel(self.full_path)
            colunas_necessarias = ['Nº', 'Periodo', 'Salvar Como', 'Caminho']
            colunas_faltando = [c for c in colunas_necessarias if c not in df.columns]
            
            if colunas_faltando:
                messagebox.showerror("Erro de Validação", 
                                     f"Colunas obrigatórias não encontradas:\n{', '.join(colunas_faltando)}")
                self.log_message(f"❌ Validação falhou: colunas ausentes - {', '.join(colunas_faltando)}")
            else:
                # Verifica se os arquivos existem
                arquivos_nao_encontrados = []
                for index, row in df.iterrows():
                    caminho_pdf = str(row['Caminho'])
                    if not os.path.exists(caminho_pdf):
                        arquivos_nao_encontrados.append(f"Linha {index + 2}: {caminho_pdf}")
                
                if arquivos_nao_encontrados:
                    msg = f"⚠️ {len(arquivos_nao_encontrados)} arquivo(s) não encontrado(s)"
                    self.log_message(msg)
                    for arquivo in arquivos_nao_encontrados[:5]:  # Mostra apenas os 5 primeiros
                        self.log_message(f"   {arquivo}")
                    if len(arquivos_nao_encontrados) > 5:
                        self.log_message(f"   ... e mais {len(arquivos_nao_encontrados) - 5}")
                
                messagebox.showinfo("Validação", 
                                    f"✅ Arquivo válido!\n📊 {len(df)} documento(s) encontrado(s)\n"
                                    f"⚠️ {len(arquivos_nao_encontrados)} arquivo(s) não encontrado(s)")
                self.log_message(f"✅ Validação concluída: {len(df)} documentos, {len(arquivos_nao_encontrados)} arquivos não encontrados")
                
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao validar arquivo:\n{str(e)}")
            self.log_message(f"❌ Erro na validação: {str(e)}")

    def run_bot(self):
        if not hasattr(self, 'full_path'):
            messagebox.showwarning("Aviso", "Selecione um arquivo Excel primeiro.")
            return

        # Verifica se o software está aberto
        try:
            app = Application(backend="uia")
            app.connect(title="Domínio Folha - Versão: 10.5A-07 - 08", timeout=5)
        except Exception:
            messagebox.showerror("Erro", "O software Domínio Folha não está aberto.\nAbra-o e tente novamente.")
            return

        self.is_running = True
        self.btn_run.configure(state="disabled")
        self.btn_stop.configure(state="normal")
        self.btn_select.configure(state="disabled")
        self.btn_validate.configure(state="disabled")
        
        threading.Thread(target=self.execute_bot, daemon=True).start()

    def stop_bot(self):
        self.is_running = False
        self.log_message("⏹️ Solicitação de parada enviada...")

    def execute_bot(self):
        try:
            bot = DomBot(log_callback=self.log_message, 
                        progress_callback=self.update_progress,
                        ui_reference=self)  # Passa referência da UI
            success = bot.publicar_documentos(self.full_path)
            
            if success:
                messagebox.showinfo("Sucesso", "Processamento concluído com sucesso!")
            else:
                messagebox.showerror("Erro", "Erro durante o processamento. Verifique o log.")
                
        except Exception as e:
            self.log_message(f"❌ Erro fatal: {str(e)}")
            messagebox.showerror("Erro Fatal", f"Erro inesperado:\n{str(e)}")
        finally:
            self.is_running = False
            self.btn_run.configure(state="normal")
            self.btn_stop.configure(state="disabled")
            self.btn_select.configure(state="normal")
            self.btn_validate.configure(state="normal")
            self.update_progress(0, "Pronto para iniciar")

if __name__ == "__main__":
    app = AppUI()
    app.mainloop()