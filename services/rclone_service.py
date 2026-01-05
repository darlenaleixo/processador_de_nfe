#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Serviço de upload usando rclone
"""

import os
import subprocess
from typing import List, Optional

class RcloneService:
    """Serviço responsável pelo upload de arquivos usando rclone"""
    
    def __init__(self):
        pass
    
    def verify_prerequisites(self, rclone_path: str, rclone_remote_name: str) -> List[str]:
        """
        Verifica se os pré-requisitos do rclone estão atendidos.
        
        Args:
            rclone_path: Caminho para o executável do rclone
            rclone_remote_name: Nome do remote configurado
            
        Returns:
            Lista de erros encontrados (vazia se tudo OK)
        """
        errors = []
        
        # Verifica se rclone existe
        if not os.path.exists(rclone_path):
            errors.append(f"rclone não encontrado em '{rclone_path}'")
            return errors
        
        # Verifica se arquivo de configuração existe
        config_path = os.path.join(os.path.dirname(rclone_path), "rclone.conf")
        if not os.path.exists(config_path):
            errors.append(f"Arquivo de configuração do rclone não encontrado: '{config_path}'")
            return errors
        
        # Verifica se o remote está configurado
        try:
            result = subprocess.run(
                [rclone_path, "listremotes", "--config", config_path],
                capture_output=True, text=True, check=True
            )
            if f"{rclone_remote_name}:" not in result.stdout:
                errors.append(f"Remote '{rclone_remote_name}' não encontrado na configuração do rclone")
        except subprocess.CalledProcessError as e:
            errors.append(f"Erro ao verificar configuração do rclone: {e.stderr}")
        
        return errors
    
    def upload_file(self, rclone_path: str, arquivo_local: str, rclone_remote_name: str,
                    caminho_destino_drive: str) -> bool:
        """
        Faz upload de um arquivo usando rclone.
        
        Args:
            rclone_path: Caminho para o executável do rclone
            arquivo_local: Caminho do arquivo local
            rclone_remote_name: Nome do remote configurado
            caminho_destino_drive: Caminho de destino no drive
            
        Returns:
            True se upload foi bem-sucedido, False caso contrário
        """
        if not os.path.exists(arquivo_local):
            print(f"ERRO: Arquivo local não existe: {arquivo_local}")
            return False
        
        config_path = os.path.join(os.path.dirname(rclone_path), "rclone.conf")
        
        try:
            args = [
                rclone_path, "copy", arquivo_local,
                f"{rclone_remote_name}:{caminho_destino_drive}",
                "--config", config_path, "--progress"
            ]
            
            result = subprocess.run(args, check=True, capture_output=True, text=True)
            print(f"SUCESSO: Upload do arquivo '{os.path.basename(arquivo_local)}' concluído.")
            return True
            
        except subprocess.CalledProcessError as e:
            print(f"ERRO no upload de {os.path.basename(arquivo_local)}: {e.stderr}")
            return False
    
    def upload_files(self, rclone_path: str, arquivos_locais: List[str],
                     rclone_remote_name: str, caminho_destino_drive: str) -> dict:
        """
        Faz upload de múltiplos arquivos.
        
        Args:
            rclone_path: Caminho para o executável do rclone
            arquivos_locais: Lista de caminhos dos arquivos locais
            rclone_remote_name: Nome do remote configurado
            caminho_destino_drive: Caminho de destino no drive
            
        Returns:
            Dicionário com resultado do upload de cada arquivo
        """
        resultados = {}
        
        for arquivo in arquivos_locais:
            if os.path.exists(arquivo):
                sucesso = self.upload_file(
                    rclone_path, arquivo, rclone_remote_name, caminho_destino_drive
                )
                resultados[arquivo] = sucesso
            else:
                print(f"AVISO: Arquivo não existe: {arquivo}")
                resultados[arquivo] = False
        
        return resultados