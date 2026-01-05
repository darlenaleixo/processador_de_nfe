# ========================================
# MODIFICAÇÕES IMPLEMENTADAS:
# ========================================
# 1. Alterada a lógica de filtragem de arquivos XML
# 2. Em vez de usar data de modificação dos arquivos, agora usa a chave de acesso da NFe
# 3. A chave de acesso contém o ano/mês da emissão nas posições 3-6 (AAMM)
# 4. Adicionadas funções para extrair e validar chaves de acesso
# 5. Mantida toda a funcionalidade original do programa

import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
import os
import shutil
from datetime import datetime, timedelta
import locale
import zipfile
import subprocess
import smtplib
import csv
import xmltodict 
import threading
import queue
import configparser 
import re
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Agendador e Trabalhador de NFEs")
        self.geometry("800x650")
        # --- NOVO: Inicia o programa minimizado ---
        self.iconify() # Deixa a janela minimizada na barra de tarefas
        self.after(100, self.deiconify) # Traz a janela para frente após um instante para garantir que seja renderizada
        self.after(200, self.iconify) # Minimiza novamente (truque para funcionar em todos os sistemas)
        
        self.log_queue = queue.Queue() 
        self.create_widgets()
        self.process_log_queue() 

        # --- NOVO: Seleciona a aba "Execução" ao iniciar ---
        # As abas são numeradas a partir de 0, então "Execução" é a aba 1.
        self.notebook.select(1) 

        print("DEBUG: Chamando a função _load_config...")
        self._load_config() 
        # --- NOVO: Inicia a contagem regressiva ---
        self._start_countdown()

    # ========================================
    # NOVA FUNÇÃO: EXTRAÇÃO DA CHAVE DE ACESSO
    # ========================================
    def extrair_chave_de_acesso_do_xml(self, caminho_arquivo_xml):
        """
        Extrai a chave de acesso do arquivo XML da NFe.
        A chave está no atributo Id do elemento infNFe.
        Retorna a chave de 44 dígitos ou None se não encontrar.
        """
        try:
            with open(caminho_arquivo_xml, 'r', encoding='utf-8') as arquivo:
                xml_content = arquivo.read()
                
            # Procura pelo padrão: <infNFe Id="NFe{44 dígitos}">
            pattern = r'<infNFe\s+Id="NFe(\d{44})"'
            match = re.search(pattern, xml_content)
            
            if match:
                return match.group(1)
            
            # Procura em outros formatos possíveis
            pattern2 = r'Id="NFe(\d{44})"'
            match2 = re.search(pattern2, xml_content)
            if match2:
                return match2.group(1)
                
            return None
            
        except Exception as e:
            self.log_message(f"ERRO ao extrair chave de acesso de {os.path.basename(caminho_arquivo_xml)}: {e}")
            return None

    # ========================================
    # NOVA FUNÇÃO: VALIDAÇÃO POR CHAVE DE ACESSO
    # ========================================
    def pertence_ao_mes_referencia_por_chave(self, chave_acesso, mes_referencia_datetime):
        """
        Verifica se a NFe pertence ao mês de referência baseado na chave de acesso.
        A chave de acesso possui 44 dígitos: nas posições 3-6 estão AAMM (Ano e Mês).
        
        Args:
            chave_acesso (str): Chave de 44 dígitos da NFe
            mes_referencia_datetime (datetime): Data do mês de referência
            
        Returns:
            bool: True se a NFe pertence ao mês de referência
        """
        if not chave_acesso or len(chave_acesso) != 44:
            return False
        
        try:
            # Extrai AAMM das posições 3-6 (índices 2-5)
            aamm = chave_acesso[2:6]
            
            # AA (ano de 2 dígitos) + MM (mês)
            ano_2_digitos = aamm[:2]
            mes_str = aamm[2:4]
            
            # Converte ano de 2 dígitos para 4 dígitos (assumindo 20XX)
            ano_nfe = int("20" + ano_2_digitos)
            mes_nfe = int(mes_str)
            
            # Valida se o mês está no range correto
            if mes_nfe < 1 or mes_nfe > 12:
                return False
            
            # Verifica se ano e mês da NFe coincidem com o mês de referência
            return (ano_nfe == mes_referencia_datetime.year and 
                    mes_nfe == mes_referencia_datetime.month)
                    
        except (ValueError, IndexError):
            return False

    def create_widgets(self):
        print("DEBUG: Entrou em create_widgets. Criando as abas...")
        # Notebook para abas
        self.notebook = ttk.Notebook(self)
        self.notebook.pack(expand=True, fill="both", padx=10, pady=10)

        # Aba de Configurações
        self.settings_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.settings_frame, text="Configurações")
        self.create_settings_tab(self.settings_frame)

        # Aba de Execução
        self.execution_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.execution_frame, text="Execução")
        self.create_execution_tab(self.execution_frame)

        # Aba de Agendamento
        self.schedule_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.schedule_frame, text="Agendamento")
        self.create_schedule_tab(self.schedule_frame)
        print("DEBUG: Abas criadas com sucesso.")

    def _select_folder(self, entry_widget):
        """Abre uma caixa de diálogo para selecionar uma pasta e atualiza o campo de texto."""
        folder_path = filedialog.askdirectory()
        if folder_path: # Se o usuário selecionou uma pasta e não cancelou
            entry_widget.delete(0, tk.END) # Limpa o campo atual
            entry_widget.insert(0, folder_path) # Insere o novo caminho

    def _select_file(self, entry_widget):
        """Abre uma caixa de diálogo para selecionar um arquivo e atualiza o campo de texto."""
        file_path = filedialog.askopenfilename()
        if file_path: # Se o usuário selecionou um arquivo e não cancelou
            entry_widget.delete(0, tk.END) # Limpa o campo atual
            entry_widget.insert(0, file_path) # Insere o novo caminho

    def _start_countdown(self):
        """Inicia a contagem regressiva no botão de execução."""
        if self.execute_button['state'] == 'disabled':
            return # Não faz nada se um backup já estiver em andamento

        self.countdown_job = None
        self.remaining_time = 10 # 10 segundos
        self._update_countdown()

    def _update_countdown(self):
        """Atualiza o texto do botão a cada segundo."""
        if self.remaining_time > 0:
            self.execute_button.config(text=f"Executando em {self.remaining_time}s... (Clique para cancelar)")
            self.remaining_time -= 1
            self.countdown_job = self.after(1000, self._update_countdown) # Agenda para rodar de novo em 1 seg
        else:
            self.execute_button.config(text="Executando...")
            self.start_backup_thread() # Inicia o backup


    def _cancel_countdown(self):
        """Cancela a contagem regressiva se o botão for clicado."""
        if self.countdown_job:
            self.after_cancel(self.countdown_job)
            self.countdown_job = None
            self.execute_button.config(text="Backup Automático Cancelado (Clique para iniciar manualmente)")
            # Altera o comando do botão para iniciar manualmente
            self.execute_button.config(command=self.start_backup_thread)

    def _save_config(self):
        """Salva as configurações atuais da GUI no arquivo config.ini."""
        config = configparser.ConfigParser()
        
        config['Paths'] = {
            'pasta_origem': self.pasta_origem_entry.get(),
            'pasta_destino_base': self.pasta_destino_base_entry.get()
        }
        config['Rclone'] = {
            'rclone_path': self.rclone_path_entry.get(),
            'rclone_remote_name': self.rclone_remote_name_entry.get(),
            'pasta_base_drive': self.pasta_base_drive_entry.get(),
            'nome_cliente_especifico': self.nome_cliente_especifico_entry.get()
        }
        config['Email'] = {
            'smtp_server': self.smtp_server_entry.get(),
            'smtp_port': self.smtp_port_entry.get(),
            'smtp_username': self.smtp_username_entry.get(),
            'smtp_password': self.smtp_password_entry.get(),
            'email_from': self.email_from_entry.get(),
            'email_to': self.email_to_entry.get()
        }
        config['Options'] = {
            'enable_email': str(self.enable_email_var.get()),
            'enable_rclone_upload': str(self.enable_rclone_upload_var.get()),
            'enable_prerequisites_check': str(self.enable_prerequisites_check_var.get())
        }
        
        try:
            with open('config.ini', 'w') as configfile:
                config.write(configfile)
            self.log_message("SUCESSO: Configurações salvas em config.ini")
        except Exception as e:
            self.log_message(f"ERRO ao salvar configurações: {e}")

    def _load_config(self):
        """Carrega as configurações do config.ini e preenche a GUI."""
        print("DEBUG: Entrou em _load_config.")
        config = configparser.ConfigParser()
        
        if not os.path.exists('config.ini'):
            print("DEBUG: ERRO - config.ini não encontrado!") # <-- ADICIONE AQUI
            self.log_message("AVISO: Arquivo config.ini não encontrado. Usando valores padrão e criando um novo.")
            # Se o arquivo não existe, os valores já na tela (padrão do código) serão usados
            # e podemos criar um config.ini inicial salvando-os.
            self._save_config()
            return

        print("DEBUG: Lendo o arquivo config.ini...") # <-- ADICIONE AQUI
        config.read('config.ini')

        # Vamos testar a leitura de um valor específico
        valor_lido = config.get('Paths', 'pasta_origem', fallback='VALOR NÃO ENCONTRADO')
        print(f"DEBUG: Valor lido para [Paths] pasta_origem = {valor_lido}") # <-- ADICIONE AQUI

        
        # Função auxiliar para evitar repetição de código
        def set_entry_value(entry, section, option):
            if config.has_option(section, option):
                entry.delete(0, tk.END)
                entry.insert(0, config.get(section, option))

        # Preenche os campos de texto
        set_entry_value(self.pasta_origem_entry, 'Paths', 'pasta_origem')
        set_entry_value(self.pasta_destino_base_entry, 'Paths', 'pasta_destino_base')
        set_entry_value(self.rclone_path_entry, 'Rclone', 'rclone_path')
        set_entry_value(self.rclone_remote_name_entry, 'Rclone', 'rclone_remote_name')
        set_entry_value(self.pasta_base_drive_entry, 'Rclone', 'pasta_base_drive')
        set_entry_value(self.nome_cliente_especifico_entry, 'Rclone', 'nome_cliente_especifico')
        set_entry_value(self.smtp_server_entry, 'Email', 'smtp_server')
        set_entry_value(self.smtp_port_entry, 'Email', 'smtp_port')
        set_entry_value(self.smtp_username_entry, 'Email', 'smtp_username')
        set_entry_value(self.smtp_password_entry, 'Email', 'smtp_password')
        set_entry_value(self.email_from_entry, 'Email', 'email_from')
        set_entry_value(self.email_to_entry, 'Email', 'email_to')

        # Preenche as caixas de seleção
        self.enable_email_var.set(config.getboolean('Options', 'enable_email', fallback=True))
        self.enable_rclone_upload_var.set(config.getboolean('Options', 'enable_rclone_upload', fallback=True))
        self.enable_prerequisites_check_var.set(config.getboolean('Options', 'enable_prerequisites_check', fallback=True))
        
        self.log_message("Configurações carregadas de config.ini")

# ================================================================================
# ================================= COLE ESTA NOVA VERSÃO NO LUGAR DA SUA FUNÇÃO extrair_dados_de_xml
# ================================================================================
    def extrair_dados_de_xml(self, caminho_arquivo_xml, canceled_keys_set):
        """
        Lê um arquivo XML de NFe e retorna uma LISTA de dicionários, 
        um para cada item/produto na nota.
        """
        # Dicionário para traduzir os códigos de pagamento
        pagamento_map = {
            '01': 'Dinheiro', '02': 'Cheque', '03': 'Cartão de Crédito', '04': 'Cartão de Débito',
            '05': 'Crédito Loja', '10': 'Vale Alimentação', '11': 'Vale Refeição',
            '15': 'Boleto Bancário', '16': 'Depósito Bancário', '17': 'PIX',
            '90': 'Sem Pagamento', '99': 'Outros'
        }

        try:
            with open(caminho_arquivo_xml, 'r', encoding='utf-8') as arquivo:
                nfe_dict = xmltodict.parse(arquivo.read())
            
            infNFe = nfe_dict.get('nfeProc', {}).get('NFe', {}).get('infNFe', {})
            if not infNFe:
                infNFe = nfe_dict.get('NFe', {}).get('infNFe', {})
            if not infNFe:
                self.log_message(f"AVISO: Estrutura XML não reconhecida em {os.path.basename(caminho_arquivo_xml)}")
                return []
            
# ============================= NOVAS EXTRAÇÕES ==============================

            chave_acesso = infNFe.get('@Id', 'NFe').replace('NFe', '')
            status = 'Cancelada' if chave_acesso in canceled_keys_set else 'Autorizada'
            
            pagamento_info = infNFe.get('pag', {}).get('detPag', {})
            # Se houver múltiplos pagamentos, pegamos o primeiro
            if isinstance(pagamento_info, list):
                pagamento_info = pagamento_info[0]
            
            cod_pagamento = pagamento_info.get('tPag', '99')
            forma_pagamento = pagamento_map.get(cod_pagamento, 'Outros') 

            # 1. Pega os dados gerais da nota (que se repetem para cada produto)
            dados_gerais = {
                'arquivo': os.path.basename(caminho_arquivo_xml),
                'data_emissao': infNFe.get('ide', {}).get('dhEmi', 'N/A'),
                'numero_nfe': infNFe.get('ide', {}).get('nNF', 'N/A'),
                'emitente_nome': infNFe.get('emit', {}).get('xNome', 'N/A'),
                'emitente_cnpj': infNFe.get('emit', {}).get('CNPJ', 'N/A'),
                'destinatario_nome': infNFe.get('dest', {}).get('xNome', 'N/A'),
                'valor_total_nota': infNFe.get('total', {}).get('ICMSTot', {}).get('vNF', '0.00'),
                'status': status, # <-- NOVO CAMPO
                'forma_pagamento': forma_pagamento # <-- NOVO CAMPO
            }

            lista_produtos = []
            itens_nfe = infNFe.get('det', [])

            # Garante que 'itens_nfe' seja sempre uma lista, mesmo que haja um só produto
            if not isinstance(itens_nfe, list):
                itens_nfe = [itens_nfe]

            # 2. Itera sobre cada produto da nota
            for item in itens_nfe:
                prod = item.get('prod', {})
                
                dados_produto_especifico = {
                    'codigo_produto': prod.get('cProd', 'N/A'),
                    'descricao_produto': prod.get('xProd', 'N/A'),
                    'ncm': prod.get('NCM', 'N/A'),
                    'quantidade': prod.get('qCom', '0'),
                    'valor_unitario': prod.get('vUnCom', '0.00'),
                    'valor_total_produto': prod.get('vProd', '0.00'),
                }
                
                # 3. Junta os dados gerais com os dados do produto específico
                linha_completa = {**dados_gerais, **dados_produto_especifico}
                lista_produtos.append(linha_completa)
            
            return lista_produtos

        except Exception as e:
            self.log_message(f"ERRO ao processar o arquivo XML {os.path.basename(caminho_arquivo_xml)}: {e}")
            return [] # Retorna uma lista vazia em caso de erro 

    def salvar_dados_em_csv(self, lista_de_dados, caminho_arquivo_csv, total_geral):
        """
        Salva uma lista de dados de produtos (uma linha por produto) em um arquivo CSV.
        """
        if not lista_de_dados:
            self.log_message("Nenhum dado de NFe para salvar no resumo CSV.")
            return
        try:
            # --- NOVO CABEÇALHO PARA O RELATÓRIO DETALHADO ---
            cabecalho = [
                'status', 'arquivo', 'data_emissao', 'numero_nfe', 'emitente_nome', 'emitente_cnpj', 
                'destinatario_nome', 'forma_pagamento', 'codigo_produto', 'descricao_produto', 'ncm', 
                'quantidade', 'valor_unitario', 'valor_total_produto', 'valor_total_nota'
            ]
            
            with open(caminho_arquivo_csv, 'w', newline='', encoding='utf-8-sig') as arquivo_csv:
                escritor = csv.DictWriter(arquivo_csv, fieldnames=cabecalho, delimiter=';')
                escritor.writeheader()
                escritor.writerows(lista_de_dados)
                
                # Escreve a linha de total com o valor que já foi calculado antes
                escritor.writerow({})
                linha_total = {
                    'valor_total_produto': 'TOTAL GERAL DAS NOTAS:',
                    'valor_total_nota': f'{total_geral:.2f}'.replace('.', ',')
                }
                escritor.writerow(linha_total)

            self.log_message(f"SUCESSO: Resumo detalhado salvo em '{caminho_arquivo_csv}'")
        except Exception as e:
            self.log_message(f"ERRO ao salvar o arquivo CSV detalhado: {e}") 
        
    def create_settings_tab(self, parent_frame):
        # Exemplo de campo de configuração
        ttk.Label(parent_frame, text="Pasta Origem:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.pasta_origem_entry = ttk.Entry(parent_frame, width=50)
        self.pasta_origem_entry.grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        ttk.Button(parent_frame, text="Procurar...", command=lambda: self._select_folder(self.pasta_origem_entry)).grid(row=0, column=2, padx=5, pady=5)
        self.pasta_origem_entry.insert(0, "C:\\Mobility_POS\\Xml_IO") # Valor padrão do script

        ttk.Label(parent_frame, text="Pasta Destino Base:").grid(row=1, column=0, padx=5, pady=5, sticky="w")
        self.pasta_destino_base_entry = ttk.Entry(parent_frame, width=50)
        self.pasta_destino_base_entry.grid(row=1, column=1, padx=5, pady=5, sticky="ew")
        ttk.Button(parent_frame, text="Procurar...", command=lambda: self._select_folder(self.pasta_destino_base_entry)).grid(row=1, column=2, padx=5, pady=5)
        self.pasta_destino_base_entry.insert(0, "C:\\CopiaNotasFiscais") # Valor padrão do script

        ttk.Label(parent_frame, text="Caminho Rclone:").grid(row=2, column=0, padx=5, pady=5, sticky="w")
        self.rclone_path_entry = ttk.Entry(parent_frame, width=50)
        self.rclone_path_entry.grid(row=2, column=1, padx=5, pady=5, sticky="ew")
        ttk.Button(parent_frame, text="Procurar...", command=lambda: self._select_file(self.rclone_path_entry)).grid(row=2, column=2, padx=5, pady=5)
        self.rclone_path_entry.insert(0, "C:\\Ferramentas\\rclone\\rclone.exe") # Valor padrão do script

        ttk.Label(parent_frame, text="Nome Remoto Rclone:").grid(row=3, column=0, padx=5, pady=5, sticky="w")
        self.rclone_remote_name_entry = ttk.Entry(parent_frame, width=50)
        self.rclone_remote_name_entry.grid(row=3, column=1, padx=5, pady=5, sticky="ew")
        self.rclone_remote_name_entry.insert(0, "MeuGoogleDrive") # Valor padrão do script

        ttk.Label(parent_frame, text="Pasta Base Drive:").grid(row=4, column=0, padx=5, pady=5, sticky="w")
        self.pasta_base_drive_entry = ttk.Entry(parent_frame, width=50)
        self.pasta_base_drive_entry.grid(row=4, column=1, padx=5, pady=5, sticky="ew")
        self.pasta_base_drive_entry.insert(0, "CLIENTES") # Valor padrão do script

        ttk.Label(parent_frame, text="Nome Cliente Específico:").grid(row=5, column=0, padx=5, pady=5, sticky="w")
        self.nome_cliente_especifico_entry = ttk.Entry(parent_frame, width=50)
        self.nome_cliente_especifico_entry.grid(row=5, column=1, padx=5, pady=5, sticky="ew")
        

        # Configurações de E-mail
        ttk.Label(parent_frame, text="--- Configurações de E-mail ---").grid(row=6, column=0, columnspan=2, padx=5, pady=10, sticky="w")

        ttk.Label(parent_frame, text="Servidor SMTP:").grid(row=7, column=0, padx=5, pady=5, sticky="w")
        self.smtp_server_entry = ttk.Entry(parent_frame, width=50)
        self.smtp_server_entry.grid(row=7, column=1, padx=5, pady=5, sticky="ew")
        self.smtp_server_entry.insert(0, "smtp.gmail.com")

        ttk.Label(parent_frame, text="Porta SMTP:").grid(row=8, column=0, padx=5, pady=5, sticky="w")
        self.smtp_port_entry = ttk.Entry(parent_frame, width=50)
        self.smtp_port_entry.grid(row=8, column=1, padx=5, pady=5, sticky="ew")
        self.smtp_port_entry.insert(0, "587")

        ttk.Label(parent_frame, text="Usuário SMTP:").grid(row=9, column=0, padx=5, pady=5, sticky="w")
        self.smtp_username_entry = ttk.Entry(parent_frame, width=50)
        self.smtp_username_entry.grid(row=9, column=1, padx=5, pady=5, sticky="ew")
        self.smtp_username_entry.insert(0, "solracinformatica@gmail.com")

        ttk.Label(parent_frame, text="Senha SMTP (App): ").grid(row=10, column=0, padx=5, pady=5, sticky="w")
        self.smtp_password_entry = ttk.Entry(parent_frame, width=50, show="*") # Senha oculta
        self.smtp_password_entry.grid(row=10, column=1, padx=5, pady=5, sticky="ew")
        self.smtp_password_entry.insert(0, "yuue eaqj kitd ohrq")

        ttk.Label(parent_frame, text="E-mail De:").grid(row=11, column=0, padx=5, pady=5, sticky="w")
        self.email_from_entry = ttk.Entry(parent_frame, width=50)
        self.email_from_entry.grid(row=11, column=1, padx=5, pady=5, sticky="ew")
        self.email_from_entry.insert(0, "solracinformatica@gmail.com")

        ttk.Label(parent_frame, text="E-mail Para:").grid(row=12, column=0, padx=5, pady=5, sticky="w")
        self.email_to_entry = ttk.Entry(parent_frame, width=50)
        self.email_to_entry.grid(row=12, column=1, padx=5, pady=5, sticky="ew")
        self.email_to_entry.insert(0, "solracinformatica@gmail.com")

        # Opções de ativação/desativação de funcionalidades
        ttk.Label(parent_frame, text="--- Opções de Funcionalidades ---").grid(row=13, column=0, columnspan=2, padx=5, pady=10, sticky="w")

        self.enable_email_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(parent_frame, text="Ativar Notificações por E-mail", variable=self.enable_email_var).grid(row=14, column=0, columnspan=2, padx=5, pady=2, sticky="w")

        self.enable_rclone_upload_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(parent_frame, text="Ativar Upload para Google Drive (rclone)", variable=self.enable_rclone_upload_var).grid(row=15, column=0, columnspan=2, padx=5, pady=2, sticky="w")

        self.enable_prerequisites_check_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(parent_frame, text="Ativar Verificação de Pré-requisitos", variable=self.enable_prerequisites_check_var).grid(row=16, column=0, columnspan=2, padx=5, pady=2, sticky="w")

        #----------- Botão Salvar ------------------------------------
        save_button = ttk.Button(parent_frame, text="Salvar Configurações", command=self._save_config)
        save_button.grid(row=17, column=0, columnspan=3, padx=5, pady=20)

        parent_frame.columnconfigure(1, weight=1) # Faz a coluna 1 expandir

    def create_execution_tab(self, parent_frame):
        # Botão de execução
        self.execute_button = ttk.Button(parent_frame, text="Executar Backup Agora", command=self._cancel_countdown)
        self.execute_button.pack(pady=20)

        # Barra de progresso
        self.progress_bar = ttk.Progressbar(parent_frame, mode='indeterminate')
        self.progress_bar.pack(pady=10, padx=20, fill="x")

        # Área de log
        ttk.Label(parent_frame, text="Log de Execução:").pack(anchor="w", padx=20)
        
        # Frame para o log com scrollbar
        log_frame = ttk.Frame(parent_frame)
        log_frame.pack(expand=True, fill="both", padx=20, pady=10)
        
        # Text widget com scrollbar
        self.log_text = tk.Text(log_frame, wrap="word", state="normal")
        scrollbar = ttk.Scrollbar(log_frame, orient="vertical", command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=scrollbar.set)
        
        # Pack o text e scrollbar
        self.log_text.pack(side="left", expand=True, fill="both")
        scrollbar.pack(side="right", fill="y")

    def create_schedule_tab(self, parent_frame):
        ttk.Label(parent_frame, text="Agendamento Automático").pack(pady=20)
        
        # Informações sobre o agendamento
        info_text = """
        O agendamento criará uma tarefa no Windows que executará
        automaticamente no dia 1º de cada mês às 08:00.
        
        A tarefa executará o backup das NFes do mês anterior
        e enviará os arquivos para o Google Drive.
        """
        
        ttk.Label(parent_frame, text=info_text, justify="center").pack(pady=10)
        
        # Botão para criar agendamento
        schedule_button = ttk.Button(parent_frame, text="Criar Agendamento", command=self.create_scheduled_task)
        schedule_button.pack(pady=20)

    def create_scheduled_task(self):
        """Cria uma tarefa agendada no Windows"""
        try:
            import sys
            script_path = sys.argv[0]
            
            # Comando para criar tarefa agendada
            cmd = f'''schtasks /create /tn "NFe Backup Mensal" /tr "python {script_path}" /sc monthly /d 1 /st 08:00 /f'''
            
            result = subprocess.run(cmd, shell=True, capture_output=True, text=True)
            
            if result.returncode == 0:
                self.log_message("SUCESSO: Tarefa agendada criada com sucesso!")
            else:
                self.log_message(f"ERRO ao criar tarefa agendada: {result.stderr}")
                
        except Exception as e:
            self.log_message(f"ERRO ao criar agendamento: {e}")

    def log_message(self, message):
        self.log_queue.put(message)
        
    def process_log_queue(self):
        """
        Verifica a fila de comunicação e atualiza a GUI (log e barra de progresso).
        """
        try:
            while not self.log_queue.empty():
                # Pega a mensagem da fila
                message = self.log_queue.get_nowait()
                
                # --- Lógica para interpretar a mensagem ---
                if isinstance(message, tuple):
                    # Se for uma tupla, é um comando para a barra de progresso
                    command, value = message
                    if command == "__PROGRESS_START_INDETERMINATE__":
                        self.progress_bar.config(mode='indeterminate')
                        self.progress_bar.start()
                    elif command == "__PROGRESS_SETUP_DETERMINATE__":
                        self.progress_bar.stop()
                        self.progress_bar.config(mode='determinate')
                        self.progress_bar['maximum'] = value
                        self.progress_bar['value'] = 0
                    elif command == "__PROGRESS_STEP__":
                        self.progress_bar.step(value)
                
                elif message == "__TASK_COMPLETE__":
                    # Mensagem especial para reativar o botão e parar a barra
                    self.execute_button.config(state="normal")
                    self.progress_bar.stop()
                    self.progress_bar['value'] = 0
                    self.log_message("--- TAREFA CONCLUÍDA ---")
                
                else:
                    # Se for uma string normal, apenas mostra no log
                    self.log_text.insert(tk.END, message + "\n")
                    self.log_text.see(tk.END)
        finally:
            self.after(100, self.process_log_queue)

    def start_backup_thread(self):
        """
        Inicia a tarefa de backup em uma nova thread para não congelar a GUI.
        """
        # Desativa o botão para evitar múltiplos cliques
        self.execute_button.config(state="disabled")
        
        # Cria e inicia a thread
        backup_thread = threading.Thread(target=self.execute_backup)
        backup_thread.daemon = True # Permite que a thread feche com o programa
        backup_thread.start()
        
    def get_settings(self):
        settings = {
            "pasta_origem": self.pasta_origem_entry.get(),
            "pasta_destino_base": self.pasta_destino_base_entry.get(),
            "rclone_path": self.rclone_path_entry.get(),
            "rclone_remote_name": self.rclone_remote_name_entry.get(),
            "pasta_base_drive": self.pasta_base_drive_entry.get(),
            "nome_cliente_especifico": self.nome_cliente_especifico_entry.get(),
            "smtp_server": self.smtp_server_entry.get(),
            "smtp_port": int(self.smtp_port_entry.get()),
            "smtp_username": self.smtp_username_entry.get(),
            "smtp_password": self.smtp_password_entry.get(),
            "email_from": self.email_from_entry.get(),
            "email_to": self.email_to_entry.get(),
            "enable_email": self.enable_email_var.get(),
            "enable_rclone_upload": self.enable_rclone_upload_var.get(),
            "enable_prerequisites_check": self.enable_prerequisites_check_var.get(),
        }
        return settings

    def send_script_email(self, subject, body, attachment_path=None):
        if not self.enable_email_var.get():
            self.log_message("Envio de e-mail desativado nas configurações.")
            return

        settings = self.get_settings()
        try:
            msg = MIMEMultipart()
            msg["From"] = settings["email_from"]
            msg["To"] = settings["email_to"]
            msg["Subject"] = subject

            msg.attach(MIMEText(body, "plain"))

            if attachment_path and os.path.exists(attachment_path):
                file_size_mb = os.path.getsize(attachment_path) / (1024 * 1024)
                if file_size_mb <= 25: # Limite do Gmail
                    with open(attachment_path, "rb") as attachment:
                        part = MIMEBase("application", "octet-stream")
                        part.set_payload(attachment.read())
                    encoders.encode_base64(part)
                    part.add_header("Content-Disposition", f"attachment; filename= {os.path.basename(attachment_path)}")
                    msg.attach(part)
                else:
                    self.log_message(f"AVISO: Arquivo de log muito grande ({file_size_mb:.2f}MB) para anexar ao e-mail.")
                    body += f"\n\nNOTA: Arquivo de log muito grande para anexar. Verifique o log localmente em: {attachment_path}"
                    msg = MIMEMultipart()
                    msg["From"] = settings["email_from"]
                    msg["To"] = settings["email_to"]
                    msg["Subject"] = subject
                    msg.attach(MIMEText(body, "plain"))

            with smtplib.SMTP(settings["smtp_server"], settings["smtp_port"]) as server:
                server.starttls()
                server.login(settings["smtp_username"], settings["smtp_password"])
                server.sendmail(settings["email_from"], settings["email_to"], msg.as_string())
            self.log_message(f"E-mail de notificação enviado com sucesso: {subject}")
        except Exception as e:
            self.log_message(f"ERRO ao enviar e-mail: {e}")

    # ========================================
    # FUNÇÃO PRINCIPAL MODIFICADA
    # ========================================
    def execute_backup(self):
        self.log_message("--- INICIANDO EXECUÇÃO DA LÓGICA DO TRABALHADOR ---")
        settings = self.get_settings()
        script_status = "SUCESSO"
        error_messages = []
        start_time = datetime.now()
        caminho_resumo_csv = "" # Inicializa a variável para EVITAR erros

        log_file = os.path.join(settings["pasta_destino_base"], "log_copia_nfe.log")

        if not os.path.exists(os.path.dirname(log_file)):
            os.makedirs(os.path.dirname(log_file))

        try:
            # 1. LÓGICA DE PRÉ-REQUISITOS
            if settings["enable_prerequisites_check"]:
                self.log_message("Verificando pré-requisitos...")
                if not os.path.exists(settings["rclone_path"]):
                    error_messages.append(f"rclone não encontrado em '{settings['rclone_path']}'")
                if not os.path.exists(settings["pasta_origem"]):
                    error_messages.append(f"Pasta origem não encontrada: '{settings['pasta_origem']}'")
                
                config_path = os.path.join(os.path.dirname(settings["rclone_path"]), "rclone.conf")
                if os.path.exists(settings["rclone_path"]) and not os.path.exists(config_path):
                    error_messages.append(f"Arquivo de configuração do rclone não encontrado: '{config_path}'")
                
                if os.path.exists(settings["rclone_path"]) and os.path.exists(config_path):
                    try:
                        result = subprocess.run([settings["rclone_path"], "listremotes", "--config", config_path], capture_output=True, text=True, check=True)
                        if f"{settings['rclone_remote_name']}:" not in result.stdout:
                            error_messages.append(f"Remote '{settings['rclone_remote_name']}' não encontrado na configuração do rclone")
                    except subprocess.CalledProcessError as e:
                        error_messages.append(f"Erro ao verificar configuração do rclone: {e.stderr}")

                if error_messages:
                    raise Exception("Pré-requisitos não atendidos. Verifique o log para detalhes.")
                self.log_message("Pré-requisitos verificados com sucesso.")

            # 2. LÓGICA DE DATAS E PASTAS
            try:
                locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')
            except locale.Error:
                locale.setlocale(locale.LC_ALL, 'Portuguese_Brazil.1252')
            hoje = datetime.now()
            primeiro_dia_mes_atual = hoje.replace(day=1, hour=0, minute=0, second=0, microsecond=0)
            mes_de_referencia = (primeiro_dia_mes_atual - timedelta(days=1)).replace(day=1)
            
            self.log_message(f"NOVA LÓGICA: Buscando NFes do mês {mes_de_referencia.strftime('%m/%Y')} baseado na CHAVE DE ACESSO")
            
            nome_mes_referencia = mes_de_referencia.strftime('%B').upper()
            nome_pasta_destino_local = f"{mes_de_referencia.strftime('%Y-%m')}_{nome_mes_referencia}"
            pasta_destino_completa = os.path.join(settings["pasta_destino_base"], nome_pasta_destino_local)
            if not os.path.exists(pasta_destino_completa):
                os.makedirs(pasta_destino_completa)

            # ========================================
            # 3. NOVA LÓGICA DE BUSCA DE ARQUIVOS
            # ========================================
            self.log_message(f"Procurando arquivos .xml em '{settings['pasta_origem']}'...")
            self.log_message("ATENÇÃO: Agora filtrando por CHAVE DE ACESSO (mês/ano da emissão) em vez de data do arquivo!")
            
            canceled_keys = self._find_canceled_invoices(settings['pasta_origem'])
            
            # Busca TODOS os arquivos XML
            todos_arquivos_xml = [
                os.path.join(root, file)
                for root, _, files in os.walk(settings["pasta_origem"])
                for file in files
                if file.endswith(".xml")
            ]
            
            self.log_message(f"Encontrados {len(todos_arquivos_xml)} arquivos XML no total. Verificando chaves de acesso...")
            
            # Filtra apenas os arquivos que pertencem ao mês de referência
            arquivos_para_copiar = []
            arquivos_verificados = 0
            
            for arquivo_xml in todos_arquivos_xml:
                arquivos_verificados += 1
                if arquivos_verificados % 100 == 0:  # Log a cada 100 arquivos
                    self.log_message(f"Verificados {arquivos_verificados}/{len(todos_arquivos_xml)} arquivos...")
                
                try:
                    chave_acesso = self.extrair_chave_de_acesso_do_xml(arquivo_xml)
                    if chave_acesso and self.pertence_ao_mes_referencia_por_chave(chave_acesso, mes_de_referencia):
                        arquivos_para_copiar.append(arquivo_xml)
                        # Log individual para debugar
                        aamm = chave_acesso[2:6]
                        self.log_message(f"✅ INCLUÍDO: {os.path.basename(arquivo_xml)} (AAMM: {aamm})")
                except Exception as e:
                    # Se der erro ao ler o XML, pula o arquivo
                    self.log_message(f"AVISO: Erro ao verificar {os.path.basename(arquivo_xml)}: {e}")
                    continue

            # 4. PROCESSO PRINCIPAL
            if arquivos_para_copiar:
                # Configura a barra de progresso
                self.log_message(f"🎯 RESULTADO: {len(arquivos_para_copiar)} arquivos do mês {mes_de_referencia.strftime('%m/%Y')} serão copiados!")
                total_steps = 1 + len(arquivos_para_copiar) + 1 + 1 
                if settings["enable_rclone_upload"]:
                    total_steps += 2
                self.log_message(("__PROGRESS_SETUP_DETERMINATE__", total_steps))
                self.log_message(("__PROGRESS_STEP__", 1))

                # Cópia dos Arquivos
                caminhos_arquivos_copiados = []
                for arquivo in arquivos_para_copiar:
                    try:
                        shutil.copy(arquivo, pasta_destino_completa)
                        caminhos_arquivos_copiados.append(os.path.join(pasta_destino_completa, os.path.basename(arquivo)))
                        self.log_message(("__PROGRESS_STEP__", 1))
                    except Exception as e:
                        error_messages.append(f"Erro ao copiar '{os.path.basename(arquivo)}': {e}")
                
                # Geração do CSV
                lista_dados_extraidos = []
                total_geral_notas = 0.0
                notas_ja_somadas = set()
                if caminhos_arquivos_copiados:
                    self.log_message("Iniciando extração de dados das NFes...")
                    for caminho_nfe in caminhos_arquivos_copiados:
                        produtos_da_nota = self.extrair_dados_de_xml(caminho_nfe, canceled_keys)
                        if produtos_da_nota:
                            lista_dados_extraidos.extend(produtos_da_nota)
                            info_nota = produtos_da_nota[0]
                            if info_nota['arquivo'] not in notas_ja_somadas and info_nota['status'] == 'Autorizada':
                                try:
                                    total_geral_notas += float(info_nota['valor_total_nota'].replace(',', '.'))
                                    notas_ja_somadas.add(info_nota['arquivo'])
                                except (ValueError, TypeError):
                                    pass
                    nome_resumo_csv = f"Resumo_Detalhado_NFEs_{nome_pasta_destino_local}.csv"
                    caminho_resumo_csv = os.path.join(settings["pasta_destino_base"], nome_resumo_csv)
                    self.salvar_dados_em_csv(lista_dados_extraidos, caminho_resumo_csv, total_geral_notas)
                    self.log_message(("__PROGRESS_STEP__", 1))

                # Compactação
                nome_arquivo_zip = f"NFEs_{mes_de_referencia.strftime('%b').upper()}_{mes_de_referencia.strftime('%Y')}.zip"
                caminho_arquivo_zip = os.path.join(settings["pasta_destino_base"], nome_arquivo_zip)
                if os.path.exists(caminho_arquivo_zip): os.remove(caminho_arquivo_zip)
                self.log_message(f"Iniciando a compactação para '{caminho_arquivo_zip}'...")
                with zipfile.ZipFile(caminho_arquivo_zip, 'w', zipfile.ZIP_DEFLATED) as zipf:
                    for f in caminhos_arquivos_copiados:
                        zipf.write(f, os.path.basename(f))
                self.log_message("Compactação concluída com sucesso.")
                self.log_message(("__PROGRESS_STEP__", 1))

                # Upload com Rclone
                if settings["enable_rclone_upload"]:
                    self.log_message("Iniciando upload para o Google Drive...")
                    caminho_destino_drive = f"{settings['pasta_base_drive']}/{settings['nome_cliente_especifico']}/{nome_pasta_destino_local}/"
                    config_path = os.path.join(os.path.dirname(settings["rclone_path"]), "rclone.conf")
                    try:
                        args_zip = ["copy", caminho_arquivo_zip, f"{settings['rclone_remote_name']}:{caminho_destino_drive}", "--config", config_path, "--progress"]
                        subprocess.run([settings["rclone_path"]] + args_zip, check=True, capture_output=True, text=True)
                        self.log_message(f"SUCESSO: Upload do arquivo '{os.path.basename(caminho_arquivo_zip)}' concluído.")
                        self.log_message(("__PROGRESS_STEP__", 1))
                    except subprocess.CalledProcessError as e:
                        error_messages.append(f"Upload do ZIP falhou: {e.stderr}")
                    
                    if os.path.exists(caminho_resumo_csv):
                        try:
                            args_csv = ["copy", caminho_resumo_csv, f"{settings['rclone_remote_name']}:{caminho_destino_drive}", "--config", config_path, "--progress"]
                            subprocess.run([settings["rclone_path"]] + args_csv, check=True, capture_output=True, text=True)
                            self.log_message(f"SUCESSO: Upload do arquivo '{os.path.basename(caminho_resumo_csv)}' concluído.")
                            self.log_message(("__PROGRESS_STEP__", 1))
                        except subprocess.CalledProcessError as e:
                            error_messages.append(f"Upload do CSV falhou: {e.stderr}")
            else:
                self.log_message(f"⚠️  Nenhum arquivo XML do mês {mes_de_referencia.strftime('%m/%Y')} foi encontrado com base na chave de acesso.")

        except Exception as e:
            script_status = "FALHA"
            error_messages.append(f"ERRO INESPERADO: {e}")
            self.log_message(f"ERRO INESPERADO: {e}")
        finally:
            end_time = datetime.now()
            duration = end_time - start_time
            
            body_details = f"""
            Computador: {os.environ.get('COMPUTERNAME', 'N/A')}
            Usuário: {os.environ.get('USERNAME', 'N/A')}
            Horário de início: {start_time.strftime('%d/%m/%Y %H:%M:%S')}
            Horário de término: {end_time.strftime('%d/%m/%Y %H:%M:%S')}
            Duração: {duration}
            Cliente: {settings['nome_cliente_especifico']}
            Método de Filtragem: CHAVE DE ACESSO (Nova implementação)
            """
            if script_status == "SUCESSO" and not error_messages:
                subject = f"SUCESSO: Copia de NFEs - {settings['nome_cliente_especifico']}"
                body_success = body_details + "\n\nA cópia e upload de NFEs foi concluída com sucesso usando a nova filtragem por chave de acesso!"
                self.send_script_email(subject, body_success)
            else:
                script_status = "FALHA" # Garante que o status seja FALHA se houver erros
                subject = f"FALHA: Copia de NFEs - {settings['nome_cliente_especifico']}"
                body_failure = body_details + "\n\nFALHAS DETECTADAS:\n\n" + ("\n".join(error_messages))
                self.send_script_email(subject, body_failure, log_file)
            
            self.log_message(f"Duração total da execução: {duration}")
            self.log_message("__TASK_COMPLETE__")

    def _find_canceled_invoices(self, folder_path):
        """
        Varre uma pasta para encontrar todos os XMLs de evento de cancelamento
        e retorna um conjunto com as Chaves de Acesso das notas canceladas.
        """
        canceled_keys = set()
        self.log_message("Iniciando verificação de notas canceladas...")
        for root, _, files in os.walk(folder_path):
            for file in files:
                if file.endswith(".xml"):
                    file_path = os.path.join(root, file)
                    try:
                        with open(file_path, 'r', encoding='utf-8') as f:
                            xml_content = f.read()
                        # O XML de evento é diferente do XML da nota
                        if '<evento' in xml_content.lower() and '<tpevento>110111</tpevento>' in xml_content:
                            # Extrai a chave de acesso do evento de cancelamento
                            chave_match = re.search(r'<chNFe>(\d{44})</chNFe>', xml_content)
                            if chave_match:
                                canceled_keys.add(chave_match.group(1))
                    except Exception as e:
                        continue  # Ignora arquivos com problemas
        
        self.log_message(f"Encontradas {len(canceled_keys)} notas canceladas.")
        return canceled_keys

if __name__ == "__main__":
    app = App()
    app.mainloop()