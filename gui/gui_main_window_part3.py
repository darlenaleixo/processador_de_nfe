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