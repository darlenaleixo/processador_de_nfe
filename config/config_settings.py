#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Gerenciamento de configurações da aplicação
"""

import os
import configparser
from typing import Dict, Any

class ConfigManager:
    """Gerencia as configurações da aplicação"""
    
    def __init__(self, config_file: str = 'config.ini'):
        self.config_file = config_file
        self.config = configparser.ConfigParser()
        
    def get_default_settings(self) -> Dict[str, Any]:
        """Retorna as configurações padrão"""
        return {
            'Paths': {
                'pasta_origem': 'C:\\Mobility_POS\\Xml_IO',
                'pasta_destino_base': 'C:\\CopiaNotasFiscais'
            },
            'Rclone': {
                'rclone_path': 'C:\\Ferramentas\\rclone\\rclone.exe',
                'rclone_remote_name': 'MeuGoogleDrive',
                'pasta_base_drive': 'CLIENTES',
                'nome_cliente_especifico': ''
            },
            'Email': {
                'smtp_server': 'smtp.gmail.com',
                'smtp_port': '587',
                'smtp_username': 'solracinformatica@gmail.com',
                'smtp_password': 'yuue eaqj kitd ohrq',
                'email_from': 'solracinformatica@gmail.com',
                'email_to': 'solracinformatica@gmail.com'
            },
            'Options': {
                'enable_email': 'True',
                'enable_rclone_upload': 'True',
                'enable_prerequisites_check': 'True'
            }
        }
    
    def load_config(self) -> Dict[str, Any]:
        """Carrega as configurações do arquivo"""
        if not os.path.exists(self.config_file):
            print(f"DEBUG: {self.config_file} não encontrado, criando com valores padrão...")
            self.save_config(self.get_default_settings())
            return self.get_default_settings()
        
        try:
            self.config.read(self.config_file, encoding='utf-8')
            return self._config_to_dict()
        except Exception as e:
            print(f"ERRO ao ler configurações: {e}")
            return self.get_default_settings()
    
    def save_config(self, settings: Dict[str, Any]) -> bool:
        """Salva as configurações no arquivo"""
        try:
            self.config.clear()
            for section_name, section_data in settings.items():
                self.config[section_name] = section_data
            
            with open(self.config_file, 'w', encoding='utf-8') as configfile:
                self.config.write(configfile)
            return True
        except Exception as e:
            print(f"ERRO ao salvar configurações: {e}")
            return False
    
    def _config_to_dict(self) -> Dict[str, Any]:
        """Converte ConfigParser para dicionário"""
        result = {}
        for section in self.config.sections():
            result[section] = dict(self.config[section])
        return result
    
    def get_setting(self, section: str, key: str, fallback: str = '') -> str:
        """Obtém uma configuração específica"""
        return self.config.get(section, key, fallback=fallback)
    
    def get_boolean_setting(self, section: str, key: str, fallback: bool = False) -> bool:
        """Obtém uma configuração booleana"""
        return self.config.getboolean(section, key, fallback=fallback)