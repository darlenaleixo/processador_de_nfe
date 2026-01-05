#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Serviço de envio de e-mails
"""

import os
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from typing import Optional

class EmailService:
    """Serviço responsável pelo envio de e-mails"""
    
    def __init__(self):
        pass
    
    def send_email(self, smtp_server: str, smtp_port: int, smtp_username: str,
                   smtp_password: str, email_from: str, email_to: str,
                   subject: str, body: str, attachment_path: Optional[str] = None) -> bool:
        """
        Envia um e-mail com opção de anexo.
        
        Args:
            smtp_server: Servidor SMTP
            smtp_port: Porta SMTP
            smtp_username: Usuário SMTP
            smtp_password: Senha SMTP
            email_from: E-mail remetente
            email_to: E-mail destinatário
            subject: Assunto do e-mail
            body: Corpo do e-mail
            attachment_path: Caminho do anexo (opcional)
            
        Returns:
            True se enviou com sucesso, False caso contrário
        """
        try:
            msg = MIMEMultipart()
            msg["From"] = email_from
            msg["To"] = email_to
            msg["Subject"] = subject

            msg.attach(MIMEText(body, "plain"))

            # Adiciona anexo se fornecido
            if attachment_path and os.path.exists(attachment_path):
                file_size_mb = os.path.getsize(attachment_path) / (1024 * 1024)
                
                if file_size_mb <= 25:  # Limite do Gmail
                    with open(attachment_path, "rb") as attachment:
                        part = MIMEBase("application", "octet-stream")
                        part.set_payload(attachment.read())
                    
                    encoders.encode_base64(part)
                    part.add_header(
                        "Content-Disposition",
                        f"attachment; filename= {os.path.basename(attachment_path)}"
                    )
                    msg.attach(part)
                else:
                    print(f"AVISO: Arquivo de log muito grande ({file_size_mb:.2f}MB) para anexar ao e-mail.")
                    body += f"\n\nNOTA: Arquivo de log muito grande para anexar. Verifique o log localmente em: {attachment_path}"
                    
                    # Recria a mensagem sem anexo
                    msg = MIMEMultipart()
                    msg["From"] = email_from
                    msg["To"] = email_to
                    msg["Subject"] = subject
                    msg.attach(MIMEText(body, "plain"))

            # Envia o e-mail
            with smtplib.SMTP(smtp_server, smtp_port) as server:
                server.starttls()
                server.login(smtp_username, smtp_password)
                server.sendmail(email_from, email_to, msg.as_string())
            
            print(f"E-mail de notificação enviado com sucesso: {subject}")
            return True
            
        except Exception as e:
            print(f"ERRO ao enviar e-mail: {e}")
            return False
    
    def send_notification_email(self, config: dict, subject: str, body: str,
                               attachment_path: Optional[str] = None) -> bool:
        """
        Envia e-mail de notificação usando configurações fornecidas.
        
        Args:
            config: Dicionário com configurações de e-mail
            subject: Assunto do e-mail
            body: Corpo do e-mail
            attachment_path: Caminho do anexo (opcional)
            
        Returns:
            True se enviou com sucesso, False caso contrário
        """
        return self.send_email(
            smtp_server=config.get('smtp_server', ''),
            smtp_port=int(config.get('smtp_port', 587)),
            smtp_username=config.get('smtp_username', ''),
            smtp_password=config.get('smtp_password', ''),
            email_from=config.get('email_from', ''),
            email_to=config.get('email_to', ''),
            subject=subject,
            body=body,
            attachment_path=attachment_path
        )