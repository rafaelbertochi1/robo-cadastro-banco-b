"""Deixa os módulos do robô importáveis pelos testes (eles ficam na pasta
de cima) e monta planilhas de exemplo no formato do export da Plataforma.

Importar os scripts não abre navegador nem rede: eles só leem variáveis
de ambiente no topo. Todos os números de proposta, nomes e CPFs usados
nos testes são fictícios."""
import os
import sys

import pytest
from openpyxl import Workbook

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))


@pytest.fixture(autouse=True)
def ambiente_limpo(monkeypatch):
    """Nenhum teste herda o modo automático ou o piso de data da máquina."""
    for variavel in ("MODO_EXECUCAO", "DATA_MINIMA", "HEADLESS"):
        monkeypatch.delenv(variavel, raising=False)


@pytest.fixture
def planilha(tmp_path):
    """Grava um .xlsx com {nome_da_aba: [cabeçalho, linha, ...]} e devolve o caminho."""
    def criar(abas, nome="export.xlsx"):
        wb = Workbook()
        wb.remove(wb.active)
        for nome_aba, linhas in abas.items():
            ws = wb.create_sheet(nome_aba)
            for linha in linhas:
                ws.append(list(linha))
        caminho = tmp_path / nome
        wb.save(caminho)
        return str(caminho)
    return criar
