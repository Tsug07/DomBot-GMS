import pandas as pd
import time
import os
from pywinauto import Application, findwindows
from pywinauto.findwindows import ElementNotFoundError

class DomBot:
    def __init__(self):
        # Inicializa a aplicação do Domínio Folha
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
            self.log_file = "publicacao_log.txt"
            self.log("✅ Conectado à janela principal do Domínio Folha")
        except Exception as e:
            self.log(f"❌ Erro ao conectar à janela principal: {str(e)}")
            raise

    def log(self, mensagem):
        """Registra mensagens no console e em um arquivo de log."""
        print(mensagem)
        with open(self.log_file, "a", encoding="utf-8") as f:
            f.write(f"{time.strftime('%Y-%m-%d %H:%M:%S')} - {mensagem}\n")

    def aguardar_e_encontrar_janela_confirmacao(self, timeout=15):
        """
        Aguarda e encontra a janela de confirmação usando múltiplas estratégias.
        """
        self.log("🔍 Procurando janela de confirmação...")
        
        # Lista de possíveis títulos para a janela de confirmação
        titulos_possiveis = [
            "Atenção",
            "Confirmação", 
            "Aviso",
            "Informação",
            "Sucesso"
        ]
        
        # Lista de possíveis class_names para diálogos
        classes_possiveis = [
            "#32770",
            "Dialog",
            "FNWND3190"
        ]
        
        inicio = time.time()
        while (time.time() - inicio) < timeout:
            try:
                # Estratégia 1: Buscar por título específico
                for titulo in titulos_possiveis:
                    for classe in classes_possiveis:
                        try:
                            dialog = self.app.window(title=titulo, class_name=classe)
                            if dialog.exists(timeout=1) and dialog.is_visible():
                                self.log(f"✅ Janela encontrada: '{titulo}' com classe '{classe}'")
                                return dialog
                        except:
                            continue
                
                # Estratégia 2: Buscar todas as janelas filhas da aplicação
                try:
                    windows = self.app.windows()
                    for window in windows:
                        try:
                            if window.is_dialog() and window.is_visible():
                                titulo = window.window_text()
                                if any(palavra in titulo.lower() for palavra in ['atenção', 'confirmação', 'aviso', 'sucesso']):
                                    self.log(f"✅ Diálogo encontrado: '{titulo}'")
                                    return window
                        except:
                            continue
                except:
                    pass
                
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
            
            time.sleep(0.5)  # Aguarda meio segundo antes de tentar novamente
        
        self.log("⚠️ Timeout: Nenhuma janela de confirmação encontrada")
        return None

    def clicar_botao_ok(self, dialog):
        """
        Tenta clicar no botão OK usando diferentes estratégias.
        """
        # Lista de possíveis textos do botão
        textos_botao = ["OK", "Ok", "Confirmar", "Sim", "Yes"]
        
        # Lista de possíveis auto_ids
        auto_ids = ["1", "2", "6", "1001", "2001"]
        
        for texto in textos_botao:
            try:
                # Estratégia 1: Por texto
                botao = dialog.child_window(title=texto, control_type="Button")
                if botao.exists(timeout=2):
                    botao.click()
                    self.log(f"✅ Botão '{texto}' clicado com sucesso")
                    return True
            except:
                continue
        
        for auto_id in auto_ids:
            try:
                # Estratégia 2: Por auto_id
                botao = dialog.child_window(auto_id=auto_id, control_type="Button")
                if botao.exists(timeout=2):
                    botao.click()
                    self.log(f"✅ Botão com auto_id '{auto_id}' clicado com sucesso")
                    return True
            except:
                continue
        
        try:
            # Estratégia 3: Primeiro botão encontrado
            botoes = dialog.children(control_type="Button")
            if botoes:
                botoes[0].click()
                self.log("✅ Primeiro botão encontrado foi clicado")
                return True
        except:
            pass
        
        # Se chegou até aqui, vamos debugar
        self.log("🔍 Debugando controles da janela:")
        try:
            dialog.print_control_identifiers()
        except:
            self.log("❌ Não foi possível imprimir controles")
        
        return False

    def ler_excel_com_coluna_extra(self, caminho_arquivo):
        """
        Lê um arquivo Excel e valida se todas as colunas obrigatórias existem.
        """
        try:
            df = pd.read_excel(caminho_arquivo)
            self.log(f"📊 Arquivo contém {len(df)} linhas de dados")

            colunas_necessarias = ['Nº', 'Periodo', 'Salvar Como', 'Caminho']

            colunas_faltando = [col for col in colunas_necessarias if col not in df.columns]
            if colunas_faltando:
                self.log(f"⚠️ ATENÇÃO: Colunas obrigatórias não encontradas: {', '.join(colunas_faltando)}")
                return None
            else:
                self.log("✅ Todas as colunas obrigatórias encontradas")

            return df

        except FileNotFoundError:
            self.log(f"❌ Arquivo não encontrado: {caminho_arquivo}")
            return None
        except Exception as e:
            self.log(f"❌ Erro ao ler arquivo: {str(e)}")
            return None

    def publicar_documentos(self, caminho_excel):
        """Publica documentos no Domínio Folha a partir de um arquivo Excel."""
        df = self.ler_excel_com_coluna_extra(caminho_excel)
        if df is None:
            self.log("❌ Não foi possível prosseguir devido a erro na leitura do Excel")
            return

        try:
            self.main_window.set_focus()
            self.log("✅ Foco definido na janela principal")

            # Encontrar a janela de Publicação de Documentos Externos
            pub_window = self.main_window.child_window(
                title="Publicação de Documentos Externos",
                class_name="FNWND3190"
            )

            if not pub_window.exists() or not pub_window.is_visible():
                self.log("❌ Janela de Publicação de Documentos Externos não encontrada ou não visível")
                return

            self.log("✅ Janela de Publicação de Documentos Externos encontrada")
            pub_window.set_focus()

            # Iterar sobre as linhas do DataFrame
            for index, row in df.iterrows():
                caminho_pdf = str(row['Caminho'])
                numero = str(row['Nº']) if pd.notnull(row['Nº']) else ""
                salvar_como = str(row['Salvar Como']) if pd.notnull(row['Salvar Como']) else ""
                
                # Validações
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
                    # Limpar campos antes de preencher
                    self.log("🧹 Limpando campos...")
                    
                    # Campo Caminho
                    campo_caminho = pub_window.child_window(auto_id="1013", class_name="Edit")
                    if campo_caminho.exists(timeout=3):
                        campo_caminho.set_focus()
                        campo_caminho.type_keys("^a{DELETE}")
                        # campo_caminho.type_keys("{DELETE}")
                        time.sleep(0.3)
                        campo_caminho.set_text(caminho_pdf)
                        self.log(f"✅ Caminho preenchido: {caminho_pdf}")
                    else:
                        self.log("❌ Campo 'Caminho' não encontrado")
                        continue

                    time.sleep(0.5)

                    # Campo Número
                    campo_numero = pub_window.child_window(auto_id="1001", class_name="PBEDIT190")
                    if campo_numero.exists(timeout=3):
                        campo_numero.set_focus()
                        campo_numero.type_keys("^a{DELETE}")
                        # campo_numero.type_keys("{DELETE}")
                        time.sleep(0.3)
                        campo_numero.set_text(numero)
                        self.log(f"✅ Número preenchido: {numero}")
                    else:
                        self.log("❌ Campo 'Número' não encontrado")
                        continue

                    time.sleep(0.5)

                    # Clicar no botão Publicar
                    botao_publicar = pub_window.child_window(auto_id="1003", class_name="Button")
                    if botao_publicar.exists(timeout=3):
                        self.log("📤 Clicando no botão 'Publicar'...")
                        botao_publicar.click()
                        time.sleep(2)  # Aguarda processamento
                    else:
                        self.log("❌ Botão 'Publicar' não encontrado")
                        continue

                    # Aguardar e processar janela de confirmação
                    dialog = self.aguardar_e_encontrar_janela_confirmacao(timeout=15)
                    
                    if dialog:
                        if self.clicar_botao_ok(dialog):
                            self.log(f"✅ Documento '{salvar_como}' publicado com sucesso")
                            time.sleep(1)  # Aguarda a janela fechar
                        else:
                            self.log(f"❌ Falha ao clicar no botão OK para '{salvar_como}'")
                            continue
                    else:
                        self.log(f"⚠️ Janela de confirmação não encontrada para '{salvar_como}'")
                        # Continua mesmo assim, pode ter sido publicado
                        
                except ElementNotFoundError as e:
                    self.log(f"⚠️ Elemento não encontrado para {salvar_como}: {str(e)}")
                    continue
                except Exception as e:
                    self.log(f"⚠️ Erro ao processar {salvar_como}: {str(e)}")
                    # Se houver erro, tenta continuar com o próximo item
                    continue

            self.log("🎉 Processamento concluído!")

        except Exception as e:
            self.log(f"❌ Erro na automação: {str(e)}")

# Exemplo de uso
if __name__ == "__main__":
    try:
        bot = DomBot()
        arquivo_excel = r"C:\Users\VM001\Documents\HUGO\testes\Publica_GMS_teste.xlsx"
        bot.publicar_documentos(arquivo_excel)
    except Exception as e:
        print(f"❌ Erro fatal: {str(e)}")
        input("Pressione Enter para sair...")