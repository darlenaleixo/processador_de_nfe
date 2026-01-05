#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Interface gráfica principal da aplicação NFe
"""

import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
import os
import shutil
import zipfile
import threading
import queue
import locale
from datetime import datetime, timedelta
from typing import Dict, Any, Optional


# Imports dos módulos do projeto
import sys
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from config.settings import ConfigManager
from nfe.parser import NFeParser
from services.email_service import EmailService
from services.rclone_service import RcloneService
from services.scheduler_service import SchedulerService

class NFeMainWindow(tk.Tk):
    """Janela principal da aplicação NFe"""
    
    def __init__(self):
        super().__init__()
        
        # Configuração da janela
        self.title("Agendador e Trabalhador de NFEs")
        self.geometry("800x650")
        
        # Inicia minimizado
        self.iconify()
        self.after(100, self.deiconify)
        self.after(200, self.iconify)
        
        # Inicializa componentes
        self.config_manager = ConfigManager()
        self.nfe_parser = NFeParser()
        self.email_service = EmailService()
        self.rclone_service = RcloneService()
        self.scheduler_service = SchedulerService()
        
        # Fila de comunicação para logs
        self.log_queue = queue.Queue()
        
        # Variáveis de controle
        self.countdown_job = None
        self.remaining_time = 10
        
        # Cria interface
        self.create_widgets()
        self.process_log_queue()
        
        # Seleciona aba de execução ao iniciar
        self.notebook.select(1)
        
        # Carrega configurações
        self.load_config()
        
        # Inicia contagem regressiva
        self.start_countdown()
    
    def create_widgets(self):
        """Cria os widgets da interface"""
        print("DEBUG: Criando interface...")
        
        # Notebook para abas
        self.notebook = ttk.Notebook(self)
        self.notebook.pack(expand=True, fill="both", padx=10, pady=10)

        # Aba de Configurações
        self.settings_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.settings_frame, text="Configurações")
        self.create_settings_tab()

        # Aba de Execução
        self.execution_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.execution_frame, text="Execução")
        self.create_execution_tab()

        # Aba de Agendamento
        self.schedule_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.schedule_frame, text="Agendamento")
        self.create_schedule_tab()
        
        print("DEBUG: Interface criada com sucesso.")
    
    def create_settings_tab(self):
        """Cria a aba de configurações"""
        parent = self.settings_frame
        
        # Campos de Pasta
        ttk.Label(parent, text="Pasta Origem:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.pasta_origem_entry = ttk.Entry(parent, width=50)
        self.pasta_origem_entry.grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        ttk.Button(parent, text="Procurar...", 
                  command=lambda: self.select_folder(self.pasta_origem_entry)).grid(row=0, column=2, padx=5, pady=5)

        ttk.Label(parent, text="Pasta Destino Base:").grid(row=1, column=0, padx=5, pady=5, sticky="w")
        self.pasta_destino_base_entry = ttk.Entry(parent, width=50)
        self.pasta_destino_base_entry.grid(row=1, column=1, padx=5, pady=5, sticky="ew")
        ttk.Button(parent, text="Procurar...", 
                  command=lambda: self.select_folder(self.pasta_destino_base_entry)).grid(row=1, column=2, padx=5, pady=5)

        # Campos Rclone
        ttk.Label(parent, text="Caminho Rclone:").grid(row=2, column=0, padx=5, pady=5, sticky="w")
        self.rclone_path_entry = ttk.Entry(parent, width=50)
        self.rclone_path_entry.grid(row=2, column=1, padx=5, pady=5, sticky="ew")
        ttk.Button(parent, text="Procurar...", 
                  command=lambda: self.select_file(self.rclone_path_entry)).grid(row=2, column=2, padx=5, pady=5)

        ttk.Label(parent, text="Nome Remoto Rclone:").grid(row=3, column=0, padx=5, pady=5, sticky="w")
        self.rclone_remote_name_entry = ttk.Entry(parent, width=50)
        self.rclone_remote_name_entry.grid(row=3, column=1, padx=5, pady=5, sticky="ew")

        ttk.Label(parent, text="Pasta Base Drive:").grid(row=4, column=0, padx=5, pady=5, sticky="w")
        self.pasta_base_drive_entry = ttk.Entry(parent, width=50)
        self.pasta_base_drive_entry.grid(row=4, column=1, padx=5, pady=5, sticky="ew")

        ttk.Label(parent, text="Nome Cliente Específico:").grid(row=5, column=0, padx=5, pady=5, sticky="w")
        self.nome_cliente_especifico_entry = ttk.Entry(parent, width=50)
        self.nome_cliente_especifico_entry.grid(row=5, column=1, padx=5, pady=5, sticky="ew")

        # Seção E-mail
        ttk.Label(parent, text="--- Configurações de E-mail ---").grid(row=6, column=0, columnspan=2, padx=5, pady=10, sticky="w")

        ttk.Label(parent, text="Servidor SMTP:").grid(row=7, column=0, padx=5, pady=5, sticky="w")
        self.smtp_server_entry = ttk.Entry(parent, width=50)
        self.smtp_server_entry.grid(row=7, column=1, padx=5, pady=5, sticky="ew")

        ttk.Label(parent, text="Porta SMTP:").grid(row=8, column=0, padx=5, pady=5, sticky="w")
        self.smtp_port_entry = ttk.Entry(parent, width=50)
        self.smtp_port_entry.grid(row=8, column=1, padx=5, pady=5, sticky="ew")

        ttk.Label(parent, text="Usuário SMTP:").grid(row=9, column=0, padx=5, pady=5, sticky="w")
        self.smtp_username_entry = ttk.Entry(parent, width=50)
        self.smtp_username_entry.grid(row=9, column=1, padx=5, pady=5, sticky="ew")

        ttk.Label(parent, text="Senha SMTP (App):").grid(row=10, column=0, padx=5, pady=5, sticky="w")
        self.smtp_password_entry = ttk.Entry(parent, width=50, show="*")
        self.smtp_password_entry.grid(row=10, column=1, padx=5, pady=5, sticky="ew")

        ttk.Label(parent, text="E-mail De:").grid(row=11, column=0, padx=5, pady=5, sticky="w")
        self.email_from_entry = ttk.Entry(parent, width=50)
        self.email_from_entry.grid(row=11, column=1, padx=5, pady=5, sticky="ew")

        ttk.Label(parent, text="E-mail Para:").grid(row=12, column=0, padx=5, pady=5, sticky="w")
        self.email_to_entry = ttk.Entry(parent, width=50)
        self.email_to_entry.grid(row=12, column=1, padx=5, pady=5, sticky="ew")

        # Opções
        ttk.Label(parent, text="--- Opções de Funcionalidades ---").grid(row=13, column=0, columnspan=2, padx=5, pady=10, sticky="w")

        self.enable_email_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(parent, text="Ativar Notificações por E-mail", 
                       variable=self.enable_email_var).grid(row=14, column=0, columnspan=2, padx=5, pady=2, sticky="w")

        self.enable_rclone_upload_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(parent, text="Ativar Upload para Google Drive (rclone)", 
                       variable=self.enable_rclone_upload_var).grid(row=15, column=0, columnspan=2, padx=5, pady=2, sticky="w")

        self.enable_prerequisites_check_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(parent, text="Ativar Verificação de Pré-requisitos", 
                       variable=self.enable_prerequisites_check_var).grid(row=16, column=0, columnspan=2, padx=5, pady=2, sticky="w")

        # Botão Salvar
        ttk.Button(parent, text="Salvar Configurações", 
                  command=self.save_config).grid(row=17, column=0, columnspan=3, padx=5, pady=20)

        parent.columnconfigure(1, weight=1)
    
    def create_execution_tab(self):
        """Cria a aba de execução"""
        parent = self.execution_frame
        
        # Botão de execução
        self.execute_button = ttk.Button(parent, text="Executar Backup Agora", command=self.cancel_countdown)
        self.execute_button.pack(pady=20)

        # Barra de progresso
        self.progress_bar = ttk.Progressbar(parent, mode='indeterminate')
        self.progress_bar.pack(pady=10, padx=20, fill="x")

        # Área de log
        ttk.Label(parent, text="Log de Execução:").pack(anchor="w", padx=20)
        
        log_frame = ttk.Frame(parent)
        log_frame.pack(expand=True, fill="both", padx=20, pady=10)
        
        self.log_text = tk.Text(log_frame, wrap="word", state="normal")
        scrollbar = ttk.Scrollbar(log_frame, orient="vertical", command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=scrollbar.set)
        
        self.log_text.pack(side="left", expand=True, fill="both")
        scrollbar.pack(side="right", fill="y")
    
    def create_schedule_tab(self):
        """Cria a aba de agendamento"""
        parent = self.schedule_frame
        
        ttk.Label(parent, text="Agendamento Automático").pack(pady=20)
        
        info_text = """
        O agendamento criará uma tarefa no Windows que executará
        automaticamente no dia 1º de cada mês às 08:00.
        
        A tarefa executará o backup das NFes do mês anterior
        e enviará os arquivos para o Google Drive.
        """
        
        ttk.Label(parent, text=info_text, justify="center").pack(pady=10)
        
        ttk.Button(parent, text="Criar Agendamento", 
                  command=self.create_scheduled_task).pack(pady=20)
    
    def select_folder(self, entry_widget):
        """Abre diálogo para selecionar pasta"""
        folder_path = filedialog.askdirectory()
        if folder_path:
            entry_widget.delete(0, tk.END)
            entry_widget.insert(0, folder_path)
    
    def select_file(self, entry_widget):
        """Abre diálogo para selecionar arquivo"""
        file_path = filedialog.askopenfilename()
        if file_path:
            entry_widget.delete(0, tk.END)
            entry_widget.insert(0, file_path)
    
    # Continuação do arquivo gui/main_window.py
# Esta é a segunda parte com os métodos restantes

    def start_countdown(self):
        """Inicia contagem regressiva"""
        if self.execute_button['state'] == 'disabled':
            return
        
        self.countdown_job = None
        self.remaining_time = 10
        self.update_countdown()
    
    def update_countdown(self):
        """Atualiza contador do botão"""
        if self.remaining_time > 0:
            self.execute_button.config(text=f"Executando em {self.remaining_time}s... (Clique para cancelar)")
            self.remaining_time -= 1
            self.countdown_job = self.after(1000, self.update_countdown)
        else:
            self.execute_button.config(text="Executando...")
            self.start_backup_thread()
    
    def cancel_countdown(self):
        """Cancela contagem regressiva"""
        if self.countdown_job:
            self.after_cancel(self.countdown_job)
            self.countdown_job = None
            self.execute_button.config(text="Backup Automático Cancelado (Clique para iniciar manualmente)")
            self.execute_button.config(command=self.start_backup_thread)
    
    def load_config(self):
        """Carrega configurações"""
        settings = self.config_manager.load_config()
        default_settings = self.config_manager.get_default_settings()
        
        # Preenche campos Paths
        self.pasta_origem_entry.insert(0, settings.get('Paths', {}).get('pasta_origem', default_settings['Paths']['pasta_origem']))
        self.pasta_destino_base_entry.insert(0, settings.get('Paths', {}).get('pasta_destino_base', default_settings['Paths']['pasta_destino_base']))
        
        # Preenche campos Rclone
        self.rclone_path_entry.insert(0, settings.get('Rclone', {}).get('rclone_path', default_settings['Rclone']['rclone_path']))
        self.rclone_remote_name_entry.insert(0, settings.get('Rclone', {}).get('rclone_remote_name', default_settings['Rclone']['rclone_remote_name']))
        self.pasta_base_drive_entry.insert(0, settings.get('Rclone', {}).get('pasta_base_drive', default_settings['Rclone']['pasta_base_drive']))
        self.nome_cliente_especifico_entry.insert(0, settings.get('Rclone', {}).get('nome_cliente_especifico', default_settings['Rclone']['nome_cliente_especifico']))
        
        # Preenche campos Email
        self.smtp_server_entry.insert(0, settings.get('Email', {}).get('smtp_server', default_settings['Email']['smtp_server']))
        self.smtp_port_entry.insert(0, settings.get('Email', {}).get('smtp_port', default_settings['Email']['smtp_port']))
        self.smtp_username_entry.insert(0, settings.get('Email', {}).get('smtp_username', default_settings['Email']['smtp_username']))
        self.smtp_password_entry.insert(0, settings.get('Email', {}).get('smtp_password', default_settings['Email']['smtp_password']))
        self.email_from_entry.insert(0, settings.get('Email', {}).get('email_from', default_settings['Email']['email_from']))
        self.email_to_entry.insert(0, settings.get('Email', {}).get('email_to', default_settings['Email']['email_to']))
        
        # Preenche opções
        self.enable_email_var.set(settings.get('Options', {}).get('enable_email', 'True') == 'True')
        self.enable_rclone_upload_var.set(settings.get('Options', {}).get('enable_rclone_upload', 'True') == 'True')
        self.enable_prerequisites_check_var.set(settings.get('Options', {}).get('enable_prerequisites_check', 'True') == 'True')
        
        self.log_message("Configurações carregadas com sucesso")
    
    def save_config(self):
        """Salva configurações"""
        settings = {
            'Paths': {
                'pasta_origem': self.pasta_origem_entry.get(),
                'pasta_destino_base': self.pasta_destino_base_entry.get()
            },
            'Rclone': {
                'rclone_path': self.rclone_path_entry.get(),
                'rclone_remote_name': self.rclone_remote_name_entry.get(),
                'pasta_base_drive': self.pasta_base_drive_entry.get(),
                'nome_cliente_especifico': self.nome_cliente_especifico_entry.get()
            },
            'Email': {
                'smtp_server': self.smtp_server_entry.get(),
                'smtp_port': self.smtp_port_entry.get(),
                'smtp_username': self.smtp_username_entry.get(),
                'smtp_password': self.smtp_password_entry.get(),
                'email_from': self.email_from_entry.get(),
                'email_to': self.email_to_entry.get()
            },
            'Options': {
                'enable_email': str(self.enable_email_var.get()),
                'enable_rclone_upload': str(self.enable_rclone_upload_var.get()),
                'enable_prerequisites_check': str(self.enable_prerequisites_check_var.get())
            }
        }
        
        if self.config_manager.save_config(settings):
            self.log_message("SUCESSO: Configurações salvas em config.ini")
        else:
            self.log_message("ERRO ao salvar configurações")
    
    def get_settings_dict(self) -> Dict[str, Any]:
        """Retorna configurações atuais como dicionário"""
        return {
            "pasta_origem": self.pasta_origem_entry.get(),
            "pasta_destino_base": self.pasta_destino_base_entry.get(),
            "rclone_path": self.rclone_path_entry.get(),
            "rclone_remote_name": self.rclone_remote_name_entry.get(),
            "pasta_base_drive": self.pasta_base_drive_entry.get(),
            "nome_cliente_especifico": self.nome_cliente_especifico_entry.get(),
            "smtp_server": self.smtp_server_entry.get(),
            "smtp_port": int(self.smtp_port_entry.get() or 587),
            "smtp_username": self.smtp_username_entry.get(),
            "smtp_password": self.smtp_password_entry.get(),
            "email_from": self.email_from_entry.get(),
            "email_to": self.email_to_entry.get(),
            "enable_email": self.enable_email_var.get(),
            "enable_rclone_upload": self.enable_rclone_upload_var.get(),
            "enable_prerequisites_check": self.enable_prerequisites_check_var.get(),
        }
    
    def create_scheduled_task(self):
        """Cria tarefa agendada"""
        success = self.scheduler_service.create_monthly_task()
        if success:
            self.log_message("SUCESSO: Tarefa agendada criada!")
        else:
            self.log_message("ERRO ao criar tarefa agendada")
    
    def log_message(self, message):
        """Adiciona mensagem à fila de log"""
        self.log_queue.put(message)
    
    def process_log_queue(self):
        """Processa fila de mensagens"""
        try:
            while not self.log_queue.empty():
                message = self.log_queue.get_nowait()
                
                if isinstance(message, tuple):
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
                    self.execute_button.config(state="normal")
                    self.progress_bar.stop()
                    self.progress_bar['value'] = 0
                    self.log_message("--- TAREFA CONCLUÍDA ---")
                
                else:
                    self.log_text.insert(tk.END, message + "\n")
                    self.log_text.see(tk.END)
        finally:
            self.after(100, self.process_log_queue)
    
    def start_backup_thread(self):
        """Inicia thread de backup"""
        self.execute_button.config(state="disabled")
        backup_thread = threading.Thread(target=self.execute_backup)
        backup_thread.daemon = True
        backup_thread.start()

    # Continuação do arquivo gui/main_window.py
# Esta é a terceira parte com a função execute_backup

    def execute_backup(self):
        """Executa o processo de backup das NFes"""
        self.log_message("--- INICIANDO EXECUÇÃO DA LÓGICA DO TRABALHADOR ---")
        settings = self.get_settings_dict()
        script_status = "SUCESSO"
        error_messages = []
        start_time = datetime.now()
        caminho_resumo_csv = ""

        log_file = os.path.join(settings["pasta_destino_base"], "log_copia_nfe.log")
        if not os.path.exists(os.path.dirname(log_file)):
            os.makedirs(os.path.dirname(log_file))

        try:
            # 1. Verificação de pré-requisitos
            if settings["enable_prerequisites_check"]:
                self.log_message("Verificando pré-requisitos...")
                
                if not os.path.exists(settings["pasta_origem"]):
                    error_messages.append(f"Pasta origem não encontrada: '{settings['pasta_origem']}'")
                
                # Verifica rclone se upload estiver habilitado
                if settings["enable_rclone_upload"]:
                    rclone_errors = self.rclone_service.verify_prerequisites(
                        settings["rclone_path"], settings["rclone_remote_name"]
                    )
                    error_messages.extend(rclone_errors)

                if error_messages:
                    raise Exception("Pré-requisitos não atendidos. Verifique o log para detalhes.")
                
                self.log_message("Pré-requisitos verificados com sucesso.")

            # 2. Configuração de datas e pastas
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

            # 3. Busca e filtragem de arquivos
            arquivos_para_copiar = self.nfe_parser.filtrar_arquivos_por_chave(
                settings["pasta_origem"], mes_de_referencia, self.log_message
            )

            # 4. Processamento principal
            if arquivos_para_copiar:
                self.log_message(f"🎯 RESULTADO: {len(arquivos_para_copiar)} arquivos do mês {mes_de_referencia.strftime('%m/%Y')} serão copiados!")
                
                # Configura barra de progresso
                total_steps = 1 + len(arquivos_para_copiar) + 1 + 1
                if settings["enable_rclone_upload"]:
                    total_steps += 2
                self.log_message(("__PROGRESS_SETUP_DETERMINATE__", total_steps))
                self.log_message(("__PROGRESS_STEP__", 1))

                # Cópia dos arquivos
                caminhos_arquivos_copiados = []
                for arquivo in arquivos_para_copiar:
                    try:
                        shutil.copy(arquivo, pasta_destino_completa)
                        caminhos_arquivos_copiados.append(
                            os.path.join(pasta_destino_completa, os.path.basename(arquivo))
                        )
                        self.log_message(("__PROGRESS_STEP__", 1))
                    except Exception as e:
                        error_messages.append(f"Erro ao copiar '{os.path.basename(arquivo)}': {e}")

                # Geração do CSV
                lista_dados_extraidos = []
                total_geral_notas = 0.0
                notas_ja_somadas = set()
                
                if caminhos_arquivos_copiados:
                    self.log_message("Iniciando extração de dados das NFes...")
                    
                    # Busca notas canceladas
                    canceled_keys = self.nfe_parser.encontrar_notas_canceladas(settings["pasta_origem"])
                    
                    for caminho_nfe in caminhos_arquivos_copiados:
                        produtos_da_nota = self.nfe_parser.extrair_dados_de_xml(caminho_nfe, canceled_keys)
                        if produtos_da_nota:
                            lista_dados_extraidos.extend(produtos_da_nota)
                            info_nota = produtos_da_nota[0]
                            
                            if info_nota['arquivo'] not in notas_ja_somadas and info_nota['status'] == 'Autorizada':
                                try:
                                    total_geral_notas += float(info_nota['valor_total_nota'].replace(',', '.'))
                                    notas_ja_somadas.add(info_nota['arquivo'])
                                except (ValueError, TypeError):
                                    pass
                    
                    # Salva CSV
                    nome_resumo_csv = f"Resumo_Detalhado_NFEs_{nome_pasta_destino_local}.csv"
                    caminho_resumo_csv = os.path.join(settings["pasta_destino_base"], nome_resumo_csv)
                    self.nfe_parser.salvar_dados_em_csv(lista_dados_extraidos, caminho_resumo_csv, total_geral_notas)
                    self.log_message(("__PROGRESS_STEP__", 1))

                # Compactação
                nome_arquivo_zip = f"NFEs_{mes_de_referencia.strftime('%b').upper()}_{mes_de_referencia.strftime('%Y')}.zip"
                caminho_arquivo_zip = os.path.join(settings["pasta_destino_base"], nome_arquivo_zip)
                
                if os.path.exists(caminho_arquivo_zip):
                    os.remove(caminho_arquivo_zip)
                
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
                    
                    # Upload do ZIP
                    success_zip = self.rclone_service.upload_file(
                        settings["rclone_path"], caminho_arquivo_zip,
                        settings["rclone_remote_name"], caminho_destino_drive
                    )
                    if not success_zip:
                        error_messages.append("Upload do ZIP falhou")
                    
                    self.log_message(("__PROGRESS_STEP__", 1))
                    
                    # Upload do CSV
                    if os.path.exists(caminho_resumo_csv):
                        success_csv = self.rclone_service.upload_file(
                            settings["rclone_path"], caminho_resumo_csv,
                            settings["rclone_remote_name"], caminho_destino_drive
                        )
                        if not success_csv:
                            error_messages.append("Upload do CSV falhou")
                    
                    self.log_message(("__PROGRESS_STEP__", 1))
            else:
                self.log_message(f"⚠️  Nenhum arquivo XML do mês {mes_de_referencia.strftime('%m/%Y')} foi encontrado com base na chave de acesso.")

        except Exception as e:
            script_status = "FALHA"
            error_messages.append(f"ERRO INESPERADO: {e}")
            self.log_message(f"ERRO INESPERADO: {e}")
        
        finally:
            # Envio de e-mail e finalização
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
            
            if settings["enable_email"]:
                email_config = {
                    'smtp_server': settings['smtp_server'],
                    'smtp_port': settings['smtp_port'],
                    'smtp_username': settings['smtp_username'],
                    'smtp_password': settings['smtp_password'],
                    'email_from': settings['email_from'],
                    'email_to': settings['email_to']
                }
                
                if script_status == "SUCESSO" and not error_messages:
                    subject = f"SUCESSO: Copia de NFEs - {settings['nome_cliente_especifico']}"
                    body = body_details + "\n\nA cópia e upload de NFEs foi concluída com sucesso usando a nova filtragem por chave de acesso!"
                    self.email_service.send_notification_email(email_config, subject, body)
                else:
                    script_status = "FALHA"
                    subject = f"FALHA: Copia de NFEs - {settings['nome_cliente_especifico']}"
                    body = body_details + "\n\nFALHAS DETECTADAS:\n\n" + ("\n".join(error_messages))
                    self.email_service.send_notification_email(email_config, subject, body, log_file)
            
            self.log_message(f"Duração total da execução: {duration}")
            self.log_message("__TASK_COMPLETE__")