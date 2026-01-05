#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Serviço de agendamento de tarefas do Windows
"""

import subprocess
import sys
from typing import Optional

class SchedulerService:
    """Serviço responsável pelo agendamento de tarefas no Windows"""
    
    def __init__(self):
        pass
    
    def create_monthly_task(self, task_name: str = "NFe Backup Mensal",
                           day: int = 1, hour: str = "08:00") -> bool:
        """
        Cria uma tarefa agendada mensal no Windows.
        
        Args:
            task_name: Nome da tarefa
            day: Dia do mês para executar (1-31)
            hour: Hora para executar (formato HH:MM)
            
        Returns:
            True se criou com sucesso, False caso contrário
        """
        try:
            script_path = sys.argv[0]
            
            # Comando para criar tarefa agendada
            cmd = [
                'schtasks', '/create',
                '/tn', task_name,
                '/tr', f'python "{script_path}"',
                '/sc', 'monthly',
                '/d', str(day),
                '/st', hour,
                '/f'  # Force overwrite if exists
            ]
            
            result = subprocess.run(cmd, shell=True, capture_output=True, text=True)
            
            if result.returncode == 0:
                print("SUCESSO: Tarefa agendada criada com sucesso!")
                return True
            else:
                print(f"ERRO ao criar tarefa agendada: {result.stderr}")
                return False
                
        except Exception as e:
            print(f"ERRO ao criar agendamento: {e}")
            return False
    
    def delete_task(self, task_name: str = "NFe Backup Mensal") -> bool:
        """
        Remove uma tarefa agendada do Windows.
        
        Args:
            task_name: Nome da tarefa a ser removida
            
        Returns:
            True se removeu com sucesso, False caso contrário
        """
        try:
            cmd = ['schtasks', '/delete', '/tn', task_name, '/f']
            result = subprocess.run(cmd, shell=True, capture_output=True, text=True)
            
            if result.returncode == 0:
                print(f"SUCESSO: Tarefa '{task_name}' removida com sucesso!")
                return True
            else:
                print(f"ERRO ao remover tarefa: {result.stderr}")
                return False
                
        except Exception as e:
            print(f"ERRO ao remover agendamento: {e}")
            return False
    
    def check_task_exists(self, task_name: str = "NFe Backup Mensal") -> bool:
        """
        Verifica se uma tarefa existe no agendador.
        
        Args:
            task_name: Nome da tarefa
            
        Returns:
            True se a tarefa existe, False caso contrário
        """
        try:
            cmd = ['schtasks', '/query', '/tn', task_name]
            result = subprocess.run(cmd, shell=True, capture_output=True, text=True)
            return result.returncode == 0
        except Exception:
            return False