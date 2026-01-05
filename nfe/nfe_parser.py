#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Processamento de arquivos XML de NFe
Responsável por extrair dados, filtrar por chave de acesso e gerar relatórios
"""

import os
import re
import csv
import xmltodict
from datetime import datetime
from typing import List, Dict, Set, Optional, Tuple


class NFeParser:
    """Classe responsável pelo processamento de arquivos XML de NFe"""
    
    def __init__(self):
        self.pagamento_map = {
            '01': 'Dinheiro', '02': 'Cheque', '03': 'Cartão de Crédito', 
            '04': 'Cartão de Débito', '05': 'Crédito Loja', '10': 'Vale Alimentação', 
            '11': 'Vale Refeição', '15': 'Boleto Bancário', '16': 'Depósito Bancário', 
            '17': 'PIX', '90': 'Sem Pagamento', '99': 'Outros'
        }
    
    def extrair_chave_de_acesso(self, caminho_arquivo_xml: str) -> Optional[str]:
        """
        Extrai a chave de acesso do arquivo XML da NFe.
        
        Args:
            caminho_arquivo_xml: Caminho para o arquivo XML
            
        Returns:
            Chave de 44 dígitos ou None se não encontrar
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
            print(f"ERRO ao extrair chave de acesso de {os.path.basename(caminho_arquivo_xml)}: {e}")
            return None
    
    def pertence_ao_mes_referencia(self, chave_acesso: str, mes_referencia: datetime) -> bool:
        """
        Verifica se a NFe pertence ao mês de referência baseado na chave de acesso.
        
        Args:
            chave_acesso: Chave de 44 dígitos da NFe
            mes_referencia: Data do mês de referência
            
        Returns:
            True se a NFe pertence ao mês de referência
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
            return (ano_nfe == mes_referencia.year and 
                    mes_nfe == mes_referencia.month)
                    
        except (ValueError, IndexError):
            return False
    
    def filtrar_arquivos_por_chave(self, pasta_origem: str, mes_referencia: datetime, 
                                   log_callback=None) -> List[str]:
        """
        Filtra arquivos XML baseado na chave de acesso da NFe.
        
        Args:
            pasta_origem: Pasta onde estão os arquivos XML
            mes_referencia: Mês de referência para filtrar
            log_callback: Função para logging (opcional)
            
        Returns:
            Lista de caminhos dos arquivos que pertencem ao mês de referência
        """
        if log_callback:
            log_callback(f"Procurando arquivos .xml em '{pasta_origem}'...")
            log_callback("ATENÇÃO: Filtrando por CHAVE DE ACESSO (mês/ano da emissão)!")
        
        # Busca TODOS os arquivos XML
        todos_arquivos_xml = [
            os.path.join(root, file)
            for root, _, files in os.walk(pasta_origem)
            for file in files
            if file.endswith(".xml")
        ]
        
        if log_callback:
            log_callback(f"Encontrados {len(todos_arquivos_xml)} arquivos XML no total. Verificando chaves de acesso...")
        
        # Filtra apenas os arquivos que pertencem ao mês de referência
        arquivos_para_copiar = []
        arquivos_verificados = 0
        
        for arquivo_xml in todos_arquivos_xml:
            arquivos_verificados += 1
            if arquivos_verificados % 100 == 0 and log_callback:  # Log a cada 100 arquivos
                log_callback(f"Verificados {arquivos_verificados}/{len(todos_arquivos_xml)} arquivos...")
            
            try:
                chave_acesso = self.extrair_chave_de_acesso(arquivo_xml)
                if chave_acesso and self.pertence_ao_mes_referencia(chave_acesso, mes_referencia):
                    arquivos_para_copiar.append(arquivo_xml)
                    # Log individual para debugar
                    if log_callback:
                        aamm = chave_acesso[2:6]
                        log_callback(f"✅ INCLUÍDO: {os.path.basename(arquivo_xml)} (AAMM: {aamm})")
            except Exception as e:
                # Se der erro ao ler o XML, pula o arquivo
                if log_callback:
                    log_callback(f"AVISO: Erro ao verificar {os.path.basename(arquivo_xml)}: {e}")
                continue
        
        return arquivos_para_copiar
    
    def extrair_dados_de_xml(self, caminho_arquivo_xml: str, 
                            canceled_keys_set: Set[str]) -> List[Dict[str, str]]:
        """
        Lê um arquivo XML de NFe e retorna uma lista de dicionários,
        um para cada item/produto na nota.
        
        Args:
            caminho_arquivo_xml: Caminho para o arquivo XML
            canceled_keys_set: Conjunto de chaves de notas canceladas
            
        Returns:
            Lista de dicionários com dados dos produtos
        """
        try:
            with open(caminho_arquivo_xml, 'r', encoding='utf-8') as arquivo:
                nfe_dict = xmltodict.parse(arquivo.read())
            
            infNFe = nfe_dict.get('nfeProc', {}).get('NFe', {}).get('infNFe', {})
            if not infNFe:
                infNFe = nfe_dict.get('NFe', {}).get('infNFe', {})
            if not infNFe:
                print(f"AVISO: Estrutura XML não reconhecida em {os.path.basename(caminho_arquivo_xml)}")
                return []
            
            # Extrai informações da nota
            chave_acesso = infNFe.get('@Id', 'NFe').replace('NFe', '')
            status = 'Cancelada' if chave_acesso in canceled_keys_set else 'Autorizada'
            
            pagamento_info = infNFe.get('pag', {}).get('detPag', {})
            if isinstance(pagamento_info, list):
                pagamento_info = pagamento_info[0]
            
            cod_pagamento = pagamento_info.get('tPag', '99')
            forma_pagamento = self.pagamento_map.get(cod_pagamento, 'Outros')
            
            # Dados gerais da nota (que se repetem para cada produto)
            dados_gerais = {
                'arquivo': os.path.basename(caminho_arquivo_xml),
                'data_emissao': infNFe.get('ide', {}).get('dhEmi', 'N/A'),
                'numero_nfe': infNFe.get('ide', {}).get('nNF', 'N/A'),
                'emitente_nome': infNFe.get('emit', {}).get('xNome', 'N/A'),
                'emitente_cnpj': infNFe.get('emit', {}).get('CNPJ', 'N/A'),
                'destinatario_nome': infNFe.get('dest', {}).get('xNome', 'N/A'),
                'valor_total_nota': infNFe.get('total', {}).get('ICMSTot', {}).get('vNF', '0.00'),
                'status': status,
                'forma_pagamento': forma_pagamento
            }

            lista_produtos = []
            itens_nfe = infNFe.get('det', [])

            # Garante que 'itens_nfe' seja sempre uma lista
            if not isinstance(itens_nfe, list):
                itens_nfe = [itens_nfe]

            # Itera sobre cada produto da nota
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
                
                # Junta os dados gerais com os dados do produto específico
                linha_completa = {**dados_gerais, **dados_produto_especifico}
                lista_produtos.append(linha_completa)
            
            return lista_produtos

        except Exception as e:
            print(f"ERRO ao processar o arquivo XML {os.path.basename(caminho_arquivo_xml)}: {e}")
            return []
    
    def salvar_dados_em_csv(self, lista_de_dados: List[Dict[str, str]], 
                           caminho_arquivo_csv: str, total_geral: float) -> bool:
        """
        Salva uma lista de dados de produtos em um arquivo CSV.
        
        Args:
            lista_de_dados: Lista de dicionários com dados dos produtos
            caminho_arquivo_csv: Caminho onde salvar o CSV
            total_geral: Valor total geral das notas
            
        Returns:
            True se salvou com sucesso, False caso contrário
        """
        if not lista_de_dados:
            print("Nenhum dado de NFe para salvar no resumo CSV.")
            return False
            
        try:
            cabecalho = [
                'status', 'arquivo', 'data_emissao', 'numero_nfe', 'emitente_nome', 'emitente_cnpj',
                'destinatario_nome', 'forma_pagamento', 'codigo_produto', 'descricao_produto', 'ncm',
                'quantidade', 'valor_unitario', 'valor_total_produto', 'valor_total_nota'
            ]
            
            with open(caminho_arquivo_csv, 'w', newline='', encoding='utf-8-sig') as arquivo_csv:
                escritor = csv.DictWriter(arquivo_csv, fieldnames=cabecalho, delimiter=';')
                escritor.writeheader()
                escritor.writerows(lista_de_dados)
                
                # Escreve a linha de total
                escritor.writerow({})
                linha_total = {
                    'valor_total_produto': 'TOTAL GERAL DAS NOTAS:',
                    'valor_total_nota': f'{total_geral:.2f}'.replace('.', ',')
                }
                escritor.writerow(linha_total)

            print(f"SUCESSO: Resumo detalhado salvo em '{caminho_arquivo_csv}'")
            return True
            
        except Exception as e:
            print(f"ERRO ao salvar o arquivo CSV detalhado: {e}")
            return False
    
    def encontrar_notas_canceladas(self, pasta_origem: str) -> Set[str]:
        """
        Varre uma pasta para encontrar todos os XMLs de evento de cancelamento
        e retorna um conjunto com as chaves de acesso das notas canceladas.
        
        Args:
            pasta_origem: Pasta onde procurar os XMLs
            
        Returns:
            Conjunto com as chaves das notas canceladas
        """
        canceled_keys = set()
        print("Iniciando verificação de notas canceladas...")
        
        for root, _, files in os.walk(pasta_origem):
            for file in files:
                if file.endswith(".xml"):
                    file_path = os.path.join(root, file)
                    try:
                        with open(file_path, 'r', encoding='utf-8') as f:
                            xml_content = f.read()
                        
                        # Verifica se é XML de evento de cancelamento
                        if '<evento' in xml_content.lower() and '<tpEvento>110111</tpEvento>' in xml_content:
                            # Extrai a chave de acesso do evento de cancelamento
                            chave_match = re.search(r'<chNFe>(\d{44})</chNFe>', xml_content)
                            if chave_match:
                                canceled_keys.add(chave_match.group(1))
                    except Exception:
                        continue  # Ignora arquivos com problemas
        
        print(f"Encontradas {len(canceled_keys)} notas canceladas.")
        return canceled_keys