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