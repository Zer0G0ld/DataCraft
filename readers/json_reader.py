"""
Leitor de arquivos JSON
"""

import pandas as pd
import json
from .base_reader import BaseReader

class JSONReader(BaseReader):
    """Leitor especializado para arquivos JSON"""
    
    def read(self, file_path: str, **kwargs) -> pd.DataFrame:
        with open(file_path, 'r', encoding='utf-8') as f:
            data = json.load(f)
        
        # Encontra a maior lista
        listas = self._find_all_lists(data)
        
        if not listas:
            if isinstance(data, dict):
                return pd.DataFrame([data])
            elif isinstance(data, list):
                return pd.DataFrame(data)
            else:
                return pd.DataFrame([{'valor': data}])
        
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
    
    def _find_all_lists(self, obj, caminho=None):
        """Encontra todas as listas no objeto"""
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
        return ['.json']