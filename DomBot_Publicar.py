import pandas as pd
import time
import os
from pywinauto import Application, findwindows

class DomBot:
    def __init__(self):
        # Inicializa a aplicação do Domínio Folha
        try:
            self.app = Application(backend="uia").connect(
                title="Domínio Folha - Versão: 10.5A-07 - 06",
                class_name="FNWND3190",
                timeout=10
            )
            self.main_window = self.app.window(
                title="Domínio Folha - Versão: 10.5A-07 - 06",
                class_name="FNWND3190"
            )
            self.main_window.set_focus()  # Foca na janela principal
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

    def ler_excel_com_coluna_extra(self, caminho_arquivo):
        """
        Lê um arquivo Excel e valida se todas as colunas obrigatórias existem.
        
        :param caminho_arquivo: Caminho para o arquivo Excel (.xlsx ou .xls)
        :return: DataFrame do pandas ou None em caso de erro
        """
        try:
            # Lê o arquivo Excel
            df = pd.read_excel(caminho_arquivo)
            self.log(f"📊 Arquivo contém {len(df)} linhas de dados")

            # Lista de colunas obrigatórias
            colunas_necessarias = ['Nº', 'Periodo', 'Salvar Como', 'Caminho']

            # Valida colunas
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
        # Lê o arquivo Excel
        df = self.ler_excel_com_coluna_extra(caminho_excel)
        if df is None:
            self.log("❌ Não foi possível prosseguir devido a erro na leitura do Excel")
            return

        try:
            # Foca na janela principal
            self.main_window.set_focus()
            self.log("✅ Foco definido na janela principal")

            # Encontrar a janela de Publicação de Documentos Externos (filha)
            pub_window = self.main_window.child_window(
                title="Publicação de Documentos Externos",
                class_name="FNWND3190"
            )

            if not pub_window.exists() or not pub_window.is_visible():
                self.log("❌ Janela de Publicação de Documentos Externos não encontrada ou não visível")
                return

            self.log("✅ Janela de Publicação de Documentos Externos encontrada")
            pub_window.set_focus()  # Foca na janela filha

            # Iterar sobre as linhas do DataFrame
            for index, row in df.iterrows():
                caminho_pdf = str(row['Caminho'])
                numero = str(row['Nº']) if pd.notnull(row['Nº']) else ""
                
                # Validar se o arquivo PDF existe
                if not os.path.exists(caminho_pdf):
                    self.log(f"⚠️ Arquivo PDF não encontrado: {caminho_pdf}")
                    continue

                # Validar se o número é válido
                if not numero:
                    self.log(f"⚠️ Valor inválido na coluna 'Nº' para a linha {index + 2}")
                    continue

                self.log(f"📂 Inserindo caminho: {caminho_pdf} e número: {numero}")
                try:
                    # Preencher o campo 'Caminho' (Onvio Processos)
                    campo_caminho = pub_window.child_window(
                        auto_id="1013",
                        class_name="Edit"
                    )
                    campo_caminho.set_text(caminho_pdf)
                    time.sleep(0.5)

                    # Preencher o campo 'Nº'
                    campo_numero = pub_window.child_window(
                        auto_id="1001",
                        class_name="PBEDIT190"
                    )
                    campo_numero.set_text(numero)
                    time.sleep(0.5)

                    # Clicar no botão 'Publicar'
                    botao_publicar = pub_window.child_window(
                        auto_id="1003",
                        class_name="Button"
                    )
                    botao_publicar.click()
                    time.sleep(1)  # Aguarda a ação ser processada

                    self.log(f"✅ Documento {row['Salvar Como']} publicado com sucesso")
                except findwindows.ElementNotFoundError:
                    self.log(f"⚠️ Erro: Campo ou botão não encontrado para {caminho_pdf}")
                except Exception as e:
                    self.log(f"⚠️ Erro ao preencher caminho {caminho_pdf} ou número {numero}: {str(e)}")

        except Exception as e:
            self.log(f"❌ Erro na automação: {str(e)}")

# Exemplo de uso
if __name__ == "__main__":
    bot = DomBot()
    arquivo_excel = r"C:\caminho\para\arquivo.xlsx"
    bot.publicar_documentos(arquivo_excel)