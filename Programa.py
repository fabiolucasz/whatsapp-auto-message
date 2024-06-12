import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
import urllib.parse
import time
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import emoji
import os

class PlanilhaReaderApp:
    def __init__(self, root):

        self.root = root
        self.root.title("Hermes")
        #self.root.iconbitmap("C:/Users/Fabio/PycharmProjects/ProgramaFinal/icone.ico".replace("/", "\\"))

        # Defina as cores
        cor_fundo = "#000000"  # Preto
        cor_texto = "#FFFFFF"  # Branco
        cor_botao = "#FF0000"  # Vermelho

        # Ajusta a cor do plano de fundo
        self.root.configure(bg=cor_fundo)

        # Cria um estilo ttk para os botões
        estilo = ttk.Style()
        estilo.configure("Botao.TButton", font=('Helvetica', 10), foreground=cor_texto, background=cor_botao)
        estilo.map("Botao.TButton", foreground=[('pressed', 'black'), ('active', 'white')])

        # Variáveis para armazenar os dados da planilha
        self.dados_planilha = None
        self.caminho_planilha = None
        self.midia = None

        # Colunas da planilha
        self.Nome = 'Nome'  # Substitua pelo nome real da coluna
        self.cel1 = 'cel1'  # Substitua pelo nome real da coluna
        self.cel2 = 'cel2'  # Substitua pelo nome real da coluna
        self.cel3 = 'cel3'  # Substitua pelo nome real da coluna
        self.Matricula = 'Matrícula'  # Substitua pelo nome real da coluna
        self.mensagem = 'Mensagem'  # Substitua pelo nome real da coluna


        # Rótulo para "Editar planilha"
        #self.rotulo_editar_planilha = tk.Label(root, text="Editar planilha", bg=cor_fundo, fg=cor_texto)
        #self.rotulo_editar_planilha.grid(row=0, column=1, pady=5, sticky='w')

        # Botão para selecionar e ler a planilha
        self.botao_selecionar = tk.Button(root, text="Selecionar Planilha", command=self.selecionar_planilha, bg=cor_botao, fg=cor_texto)
        self.botao_selecionar.grid(row=0, column=1, pady=5, sticky='w')

        # Rótulo para exibir o nome do arquivo selecionado
        self.rotulo_nome_arquivo = tk.Label(root, text="", bg=cor_fundo, fg=cor_texto)
        self.rotulo_nome_arquivo.grid(row=1, column=1, pady=5, sticky='w')

        # Rótulo para a caixa de texto abaixo
        self.rotulo_texto = tk.Label(root, text="Digite a mensagem que deseja salvar:", bg=cor_fundo, fg=cor_texto)
        self.rotulo_texto.grid(row=2, column=1, pady=5, sticky='w')

        # Texto para inserir mensagens com quebras de linha
        self.texto_mensagem = tk.Text(root, height=5, width=30)
        self.texto_mensagem.grid(row=3, column=1, padx=5, pady=5, sticky='w')

        # Botão para salvar a mensagem
        self.botao_salvar_mensagem = tk.Button(root, text="Salvar Mensagem", command=self.salvar_mensagem, bg=cor_botao, fg=cor_texto)
        self.botao_salvar_mensagem.grid(row=4, column=1, padx=5, pady=5, sticky='we')

        # Botão para excluir mensagem selecionada
        self.botao_excluir_mensagem = tk.Button(root, text="Excluir Mensagem", command=self.excluir_mensagem, bg=cor_botao, fg=cor_texto)
        self.botao_excluir_mensagem.grid(row=8, column=1, padx=5, pady=5, sticky='we')

        # Dropdown para selecionar mensagens salvas
        self.label_selecione_texto = tk.Label(root, text="Selecione a mensagem a ser enviada:", bg=cor_fundo, fg=cor_texto)
        self.dropdown_mensagens = ttk.Combobox(root, values=[], state="readonly")
        self.label_selecione_texto.grid(row=6, column=1, padx=5, pady=5, sticky='we')
        self.dropdown_mensagens.grid(row=7, column=1, padx=10, pady=5, sticky='we')

        # Rótulo de disparos
        self.rotulo_disparos = tk.Label(root, text="Selecione o método de disparo:", bg=cor_fundo, fg=cor_texto)
        self.rotulo_disparos.grid(row=10, column=1, pady=5, sticky='we')

        # Botão para iniciar o WebDriver Chrome
        self.botao_iniciar_webdriver = tk.Button(root, text="Mensagens personalizadas", command=self.iniciar_webdriver, bg=cor_botao, fg=cor_texto)
        self.botao_iniciar_webdriver.grid(row=12, column=1, padx=10, pady=10, sticky='we')

        # Botão para selecionar a midia
        self.botao_selecionar_midia = tk.Button(root, text="Selecionar Mídia", command=self.selecionar_midia,
                                          bg=cor_botao, fg=cor_texto)
        self.botao_selecionar_midia.grid(row=13, column=1, pady=5, sticky='w')

        # Rótulo para exibir o nome da midia
        self.rotulo_nome_midia = tk.Label(root, text="", bg=cor_fundo, fg=cor_texto)
        self.rotulo_nome_midia.grid(row=14, column=1, pady=5, sticky='w')

        # Botão para enviar midia
        self.botao_enviar_midia = tk.Button(root, text="Enviar Midia", command=self.enviar_midia,
                                                 bg=cor_botao, fg=cor_texto)
        self.botao_enviar_midia.grid(row=15, column=1, padx=10, pady=10, sticky='we')

        # Carregar mensagens salvas de um arquivo (se existir)
        self.carregar_mensagens_salvas()

    def selecionar_planilha(self):
        # Abre uma caixa de diálogo para selecionar a planilha
        caminho_planilha = filedialog.askopenfilename(filetypes=[("Planilhas Excel", "*.xlsx;*.xls")])

        # Verifica se o usuário selecionou uma planilha
        if caminho_planilha:
            # Obtém o nome da pasta pai
            pasta_pai = os.path.basename(os.path.dirname(caminho_planilha))

            # Atualiza o rótulo com o caminho encurtado
            self.rotulo_nome_arquivo.config(
                text=f"Arquivo Selecionado: .../{pasta_pai}/{os.path.basename(caminho_planilha)}")

            # Lê a planilha usando o pandas
            try:
                self.dados_planilha = pd.read_excel(caminho_planilha, engine='openpyxl')
                self.caminho_planilha = caminho_planilha

                print("Colunas disponíveis:", self.dados_planilha.columns)

                self.atualizar_texto_dados()
            except pd.errors.EmptyDataError:
                self.mostrar_mensagem("A planilha está vazia.")
            except Exception as e:
                self.mostrar_mensagem(f"Erro ao ler a planilha: {e}")
        else:
            self.mostrar_mensagem("Nenhuma planilha selecionada.")

    #def
    def selecionar_midia(self):
        # Abre a caixa de diálogo para selecionar arquivos
        filepath = filedialog.askopenfilename(
            filetypes=[("Arquivos de Imagem e Vídeo", "*.jpg;*.jpeg;*.png;*.gif;*.mp4")])
        return filepath

    def carregar_midia(self):
        global midia
        midia = selecionar_midia(self)
        print("Midia selecionada:", midia)



    def salvar_mensagem(self):
        # Obtém a mensagem da entrada e adiciona à lista de mensagens salvas
        nova_mensagem = self.texto_mensagem.get("1.0", tk.END).strip()
        if nova_mensagem:
            # Converte emojis para o formato apropriado
            nova_mensagem = emoji.emojize(nova_mensagem)
            self.mensagens_salvas.append(nova_mensagem)
            # Salva as mensagens em um arquivo
            self.salvar_mensagens_em_arquivo()
            # Atualiza o dropdown com as mensagens salvas
            self.atualizar_dropdown_mensagens()
            # Limpa a entrada após salvar a mensagem
            self.texto_mensagem.delete("1.0", tk.END)
        else:
            self.mostrar_mensagem("Digite uma mensagem antes de salvar.")

    def salvar_mensagens_em_arquivo(self):
        # Salva as mensagens em um arquivo
        with open("mensagens_salvas.txt", "w", encoding="utf-8") as file:
            for mensagem in self.mensagens_salvas:
                # Substitui quebras de linha por um marcador especial
                mensagem = mensagem.replace('\n', '<quebra_de_linha>')
                file.write(f"{mensagem}\n")

    def carregar_mensagens_salvas(self):
        try:
            # Tenta ler as mensagens salvas de um arquivo
            with open("mensagens_salvas.txt", "r", encoding="utf-8", errors="replace") as file:
                # Lê as linhas do arquivo
                linhas = file.read().splitlines()
                # Substitui o marcador especial de volta para quebras de linha
                self.mensagens_salvas = [linha.replace('<quebra_de_linha>', '\n') for linha in linhas]
        except FileNotFoundError:
            # Se o arquivo não existir, inicializa a lista de mensagens como vazia
            self.mensagens_salvas = []
        except Exception as e:
            # Em caso de outros erros, exibe uma mensagem de aviso
            print(f"Erro ao carregar mensagens salvas: {e}")

        # Atualiza o dropdown com as mensagens salvas
        self.atualizar_dropdown_mensagens()

    def excluir_mensagem(self):
        # Obtém a mensagem selecionada no dropdown
        mensagem_selecionada = self.dropdown_mensagens.get()

        # Verifica se uma mensagem foi selecionada
        if mensagem_selecionada:
            # Carrega as mensagens salvas do arquivo
            self.carregar_mensagens_salvas()

            # Remove a mensagem selecionada da lista
            self.mensagens_salvas.remove(mensagem_selecionada)

            # Salva as mensagens atualizadas no arquivo
            self.salvar_mensagens_em_arquivo()

            # Atualiza o dropdown após excluir a mensagem
            self.atualizar_dropdown_mensagens()
        else:
            self.mostrar_mensagem("Selecione uma mensagem para excluir.")

    def atualizar_dropdown_mensagens(self):
        # Atualiza o dropdown com as mensagens salvas
        self.dropdown_mensagens['values'] = self.mensagens_salvas


    def iniciar_webdriver(self):
        # Verifica se uma planilha foi carregada
        if self.dados_planilha is not None:
            # Inicia o WebDriver Chrome
            chrome_options = webdriver.ChromeOptions()
            chrome_options.add_argument("--headless")  # Modifique conforme necessário
            driver = webdriver.Chrome()

            # carrega o webdriver
            driver.get('https://web.whatsapp.com/')

            # whatsapp
            while len(driver.find_elements(By.ID, 'side')) < 1:
                time.sleep(1)
            time.sleep(1)


            try:
                for i, row in self.dados_planilha.iterrows():
                    pessoa = str(row[self.Nome]).split()[0]
                    matricula = row[self.Matricula]
                    numero = row[self.cel1]

                    for mensagem_atual in self.mensagens_salvas:
                        texto = urllib.parse.quote(f'Olá, {pessoa} (Matrícula {matricula}). {mensagem_atual}')
                        link = f"https://web.whatsapp.com/send?phone={numero}&text={texto}"
                        driver.get(link)

                    while len(driver.find_elements(By.ID, 'side')) < 1:
                        time.sleep(1)
                    time.sleep(1)

                    button_span_xpath = '//*[@id="main"]/footer/div[1]/div/span[2]/div/div[2]/div[2]/button/span'
                    button_span = WebDriverWait(driver, 120).until(
                        EC.visibility_of_element_located((By.XPATH, button_span_xpath))
                    )
                    button_span.click()
                    time.sleep(1)
            except Exception as e:
                print(f"Erro durante a iteração: {e}")

                self.mostrar_mensagem("Mensagens enviadas com sucesso")
        else:
            self.mostrar_mensagem("Carregue uma planilha antes de iniciar o WebDriver Chrome.")

    def enviar_midia(self):
        # Verifica se uma planilha foi carregada
        if self.dados_planilha is not None:
            # Inicia o WebDriver Chrome
            chrome_options = webdriver.ChromeOptions()
            chrome_options.add_argument("--headless")  # Modifique conforme necessário
            driver = webdriver.Chrome()

            # carrega o webdriver
            driver.get('https://web.whatsapp.com/')

            # whatsapp
            while len(driver.find_elements(By.ID, 'side')) < 1:
                time.sleep(1)
            time.sleep(1)

            try:
                for i, row in self.dados_planilha.iterrows():
                    pessoa = str(row[self.Nome]).split()[0]
                    matricula = row[self.Matricula]
                    numero = row[self.cel1]

                    for mensagem_atual in self.mensagens_salvas:
                        texto = urllib.parse.quote(f'Olá, {pessoa} (Matrícula {matricula}). {mensagem_atual}')
                        link = f"https://web.whatsapp.com/send?phone={numero}&text={texto}"
                        driver.get(link)



                while len(driver.find_elements(By.ID, 'side')) < 1:
                    time.sleep(1)

                    driver.find_element(By.CSS_SELECTOR, "span[data-icon='attach-menu-plus']").click()
                    attach = driver.find_element(By.CSS_SELECTOR,
                                                 "input[accept='image/*,video/mp4,video/3gpp,video/quicktime']")
                    attach.send_keys(midia)
                    # time.sleep(10)
                    # send = driver.find_element(By.CSS_SELECTOR, "span[data-icon='send']")
                    # send.click()

                    button_span_xpath = '//*[@id="main"]/footer/div[1]/div/span[2]/div/div[2]/div[2]/button/span'
                    button_span = WebDriverWait(driver, 120).until(
                        EC.visibility_of_element_located((By.XPATH, button_span_xpath))
                    )
                    button_span.click()
                    time.sleep(1)
            except Exception as e:
                print(f"Erro durante a iteração: {e}")

                self.mostrar_mensagem("Mensagens enviadas com sucesso")
        else:
            self.mostrar_mensagem("Carregue uma planilha antes de iniciar o WebDriver Chrome.")


    def atualizar_texto_dados(self):
        # Limpa o texto existente
        if hasattr(self, 'texto_dados'):
            self.texto_dados.delete(1.0, tk.END)
            # Exibe os dados da planilha no Text
            self.texto_dados.insert(tk.END, str(self.dados_planilha.head()))

    def mostrar_mensagem(self, mensagem):
        # Exibe uma mensagem em uma caixa de diálogo
        messagebox.showinfo("mensagem", mensagem)

if __name__ == "__main__":
    root = tk.Tk()
    app = PlanilhaReaderApp(root)
    root.mainloop()