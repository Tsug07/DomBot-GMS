import customtkinter as ctk
from tkinter import filedialog, scrolledtext
import threading
import os
import time
import pandas as pd
from pywinauto import Application, findwindows
from pywinauto.findwindows import ElementNotFoundError

class DomBot:
    def __init__(self, log_callback=None):
        self.log_callback = log_callback or print
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

    def aguardar_e_encontrar_janela_confirmacao(self, timeout=15):
        self.log("🔍 Procurando janela de confirmação...")
        titulos_possiveis = ["Atenção", "Confirmação", "Aviso", "Informação", "Sucesso"]
        classes_possiveis = ["#32770", "Dialog", "FNWND3190"]
        
        inicio = time.time()
        while (time.time() - inicio) < timeout:
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
            return

        try:
            self.main_window.set_focus()
            self.log("✅ Foco definido na janela principal")
            pub_window = self.main_window.child_window(
                title="Publicação de Documentos Externos",
                class_name="FNWND3190"
            )

            if not pub_window.exists() or not pub_window.is_visible():
                self.log("❌ Janela de Publicação de Documentos Externos não encontrada ou não visível")
                return

            self.log("✅ Janela de Publicação de Documentos Externos encontrada")
            pub_window.set_focus()

            for index, row in df.iterrows():
                caminho_pdf = str(row['Caminho'])
                numero = str(row['Nº']) if pd.notnull(row['Nº']) else ""
                salvar_como = str(row['Salvar Como']) if pd.notnull(row['Salvar Como']) else ""

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

                    botao_publicar = pub_window.child_window(auto_id="1003", class_name="Button")
                    if botao_publicar.exists(timeout=3):
                        self.log("📤 Clicando no botão 'Publicar'...")
                        botao_publicar.click()
                        time.sleep(2)
                    else:
                        self.log("❌ Botão 'Publicar' não encontrado")
                        continue

                    dialog = self.aguardar_e_encontrar_janela_confirmacao(timeout=15)
                    if dialog:
                        if self.clicar_botao_ok(dialog):
                            self.log(f"✅ Documento '{salvar_como}' publicado com sucesso")
                            time.sleep(1)
                        else:
                            self.log(f"❌ Falha ao clicar no botão OK para '{salvar_como}'")
                            continue
                    else:
                        self.log(f"⚠️ Janela de confirmação não encontrada para '{salvar_como}'")
                        continue

                except ElementNotFoundError as e:
                    self.log(f"⚠️ Elemento não encontrado para {salvar_como}: {str(e)}")
                    continue
                except Exception as e:
                    self.log(f"⚠️ Erro ao processar {salvar_como}: {str(e)}")
                    continue

            self.log("🎉 Processamento concluído!")

        except Exception as e:
            self.log(f"❌ Erro na automação: {str(e)}")

class AppUI(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Publicador Domínio Folha")
        self.geometry("500x350")
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("blue")

        self.excel_path = ctk.StringVar(value="")
        self.btn_select = ctk.CTkButton(self, text="Selecionar Excel", command=self.select_file)
        self.btn_select.pack(pady=10)
        self.lbl_path = ctk.CTkLabel(self, textvariable=self.excel_path, wraplength=480)
        self.lbl_path.pack()
        self.btn_run = ctk.CTkButton(self, text="Publicar", command=self.run_bot)
        self.btn_run.pack(pady=10)
        self.txt_log = scrolledtext.ScrolledText(self, height=10, wrap="word", state="disabled")
        self.txt_log.pack(fill="both", expand=True, padx=10, pady=10)

    def log_message(self, msg):
        self.txt_log.config(state="normal")
        self.txt_log.insert("end", msg + "\n")
        self.txt_log.see("end")
        self.txt_log.config(state="disabled")
        self.update_idletasks()

    def select_file(self):
        file_path = filedialog.askopenfilename(
            title="Selecione o Excel",
            filetypes=[("Excel Files", "*.xlsx;*.xls")]
        )
        if file_path:
            self.excel_path.set(file_path)
            self.log_message(f"📄 Arquivo selecionado: {file_path}")

    def run_bot(self):
        if not self.excel_path.get():
            self.log_message("⚠️ Selecione um arquivo Excel primeiro.")
            return
        try:
            # Verifica se o software está aberto
            app = Application(backend="uia")
            app.connect(title="Domínio Folha - Versão: 10.5A-07 - 08", timeout=5)
        except Exception:
            self.log_message("❌ Erro: O software Domínio Folha não está aberto. Abra-o e tente novamente.")
            return
        threading.Thread(target=self.execute_bot, daemon=True).start()

    def execute_bot(self):
        try:
            bot = DomBot(log_callback=self.log_message)
            bot.publicar_documentos(self.excel_path.get())
        except Exception as e:
            self.log_message(f"❌ Erro fatal: {str(e)}")

if __name__ == "__main__":
    app = AppUI()
    app.mainloop()