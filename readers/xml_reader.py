"""
Leitor de arquivos XML
Adaptado do seu código original
"""

import pandas as pd
import xml.etree.ElementTree as ET
from typing import Dict, Any, List
from .base_reader import BaseReader

class XMLReader(BaseReader):
    """Leitor especializado para arquivos XML"""
    
    def read(self, file_path: str, **kwargs) -> pd.DataFrame:
        """
        Lê arquivo XML e converte para DataFrame
        """
        tree = ET.parse(file_path)
        root = tree.getroot()
        
        # Converte XML para dicionário
        data = self._element_to_dict(root)
        
        # Encontra a maior lista para ser a tabela principal
        listas = self._find_all_lists(data)
        
        if not listas:
            # Se não encontrou lista, retorna como objeto único
            return pd.DataFrame([data] if isinstance(data, dict) else [{'valor': data}])
        
        # Pega a maior lista
        lista_principal = max(listas, key=lambda x: x['tamanho'])
        dados_processar = lista_principal['dados']
        
        # Achata cada item
        dados_achatados = []
        for item in dados_processar:
            if isinstance(item, dict):
                item_achatado = self._flatten_object(item)
            else:
                item_achatado = {'valor': str(item) if item is not None else ''}
            dados_achatados.append(item_achatado)
        
        return pd.DataFrame(dados_achatados)
    
    def _element_to_dict(self, element, prefixo=''):
        """Converte elemento XML em dicionário"""
        resultado = {}
        
        # Conta ocorrências de cada tag
        contagem_tags = {}
        for child in element:
            tag = child.tag
            contagem_tags[tag] = contagem_tags.get(tag, 0) + 1
        
        # Processa cada filho
        for child in element:
            tag = child.tag
            novo_prefixo = f"{prefixo}{tag}" if prefixo else tag
            
            if len(child) == 0:
                valor = child.text.strip() if child.text else ''
                if contagem_tags[tag] > 1:
                    if novo_prefixo not in resultado:
                        resultado[novo_prefixo] = []
                    resultado[novo_prefixo].append(valor)
                else:
                    resultado[novo_prefixo] = valor
            else:
                if contagem_tags[tag] > 1:
                    if novo_prefixo not in resultado:
                        resultado[novo_prefixo] = []
                    resultado[novo_prefixo].append(self._element_to_dict(child))
                else:
                    resultado.update(self._element_to_dict(child, f"{novo_prefixo}."))
        
        return resultado
    
    def _find_all_lists(self, obj, caminho=None):
        """Encontra todas as listas no dicionário"""
        if caminho is None:
            caminho = []
        
        listas = []
        
        if isinstance(obj, list) and len(obj) > 0:
            listas.append({
                'caminho': caminho.copy(),
                'tamanho': len(obj),
                'dados': obj
            })
        elif isinstance(obj, dict):
            for key, value in obj.items():
                caminho.append(key)
                listas.extend(self._find_all_lists(value, caminho))
                caminho.pop()
        
        return listas
    
    def _flatten_object(self, obj, prefixo=''):
        """Achata um objeto/dicionário"""
        resultado = {}
        
        if isinstance(obj, dict):
            for chave, valor in obj.items():
                novo_prefixo = f"{prefixo}{chave}" if prefixo else chave
                
                if isinstance(valor, dict):
                    resultado.update(self._flatten_object(valor, f"{novo_prefixo}."))
                elif isinstance(valor, list):
                    if len(valor) == 0:
                        resultado[novo_prefixo] = ''
                    elif len(valor) == 1 and isinstance(valor[0], (str, int, float, bool, type(None))):
                        resultado[novo_prefixo] = valor[0] if valor[0] is not None else ''
                    elif all(isinstance(item, (str, int, float, bool, type(None))) for item in valor):
                        valores = [str(v) for v in valor if v is not None]
                        resultado[novo_prefixo] = ', '.join(valores) if valores else ''
                    else:
                        for i, item in enumerate(valor, 1):
                            if isinstance(item, dict):
                                for sub_chave, sub_valor in self._flatten_object(item, f"{novo_prefixo}_{i}.").items():
                                    resultado[sub_chave] = sub_valor
                            elif item is not None:
                                resultado[f"{novo_prefixo}_{i}"] = str(item)
                else:
                    resultado[novo_prefixo] = valor if valor is not None else ''
        else:
            resultado[prefixo or 'valor'] = str(obj) if obj is not None else ''
        
        return resultado
    
    def get_supported_extensions(self) -> list:
        return ['.xml']