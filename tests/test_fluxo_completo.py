"""Fluxo completo (executar_fluxo_banco_b.py): ordem das etapas, o que
acontece quando uma falha e o código de saída que o Agendador de Tarefas
registra. As etapas de verdade são trocadas por funções falsas."""
import io
import os
import sys
import time

import pytest

import executar_fluxo_banco_b as fluxo


@pytest.fixture
def pasta(tmp_path, monkeypatch):
    monkeypatch.setattr(fluxo, "LOCK_FILE", str(tmp_path / "fluxo_em_execucao.lock"))
    monkeypatch.setattr(fluxo, "LOG_DIR", str(tmp_path / "logs"))
    return tmp_path


@pytest.fixture
def etapas(monkeypatch):
    """Troca o main() de cada etapa por uma função que só anota a chamada.
    `comportamento[n]` decide o que a etapa n faz (padrão: dá certo)."""
    chamadas = []
    comportamento = {}

    def falsa(numero, modulo):
        def main():
            chamadas.append(numero)
            acao = comportamento.get(numero)
            if isinstance(acao, Exception):
                raise acao
            if acao == "erro":
                print(f"[ERRO] proposta 1000000{numero} falhou")
            if numero == 1:
                return None if acao == "sem_excel" else "data/exports/export.xlsx"
        monkeypatch.setattr(modulo, "main", main)

    falsa(1, fluxo.exportar_status_banco_b)
    falsa(2, fluxo.cadastrar_central_gestao)
    falsa(3, fluxo.atribuir_engenheiro_central_gestao)
    falsa(4, fluxo.agendar_vistoria_central_gestao)
    falsa(5, fluxo.finalizar_vistorias_central_gestao)
    monkeypatch.setattr(fluxo, "preparar_sessao_central", lambda: comportamento.get("login", True))
    return chamadas, comportamento


def rodar(contador=None):
    contador = contador or fluxo.ContadorDeErros(io.StringIO())
    antigo = sys.stdout
    sys.stdout = contador
    problemas = {"quebraram": [], "com_erro": []}
    try:
        codigo = fluxo.executar_etapas("log.txt", contador, problemas)
    finally:
        sys.stdout = antigo
    return codigo, problemas


def test_tudo_certo(etapas):
    chamadas, _ = etapas
    assert rodar() == (fluxo.SAIDA_OK, {"quebraram": [], "com_erro": []})
    assert chamadas == [1, 2, 3, 4, 5]


def test_sem_excel_novo_nao_roda_o_resto(etapas):
    """Não cadastra nada em cima de um Excel antigo."""
    chamadas, comportamento = etapas
    comportamento[1] = "sem_excel"
    assert rodar()[0] == fluxo.SAIDA_FALHOU
    assert chamadas == [1]


def test_sem_login_na_central_para_antes_da_etapa_2(etapas):
    chamadas, comportamento = etapas
    comportamento["login"] = False
    assert rodar()[0] == fluxo.SAIDA_FALHOU
    assert chamadas == [1]


def test_etapa_que_quebra_nao_derruba_as_seguintes(etapas):
    """A Etapa 5 não depende do Excel: vale rodar mesmo com a 3 quebrada."""
    chamadas, comportamento = etapas
    comportamento[3] = RuntimeError("tela mudou")
    codigo, problemas = rodar()
    assert chamadas == [1, 2, 3, 4, 5]
    assert codigo == fluxo.SAIDA_FALHOU
    assert problemas["quebraram"] == [3]


def test_erro_no_log_vira_codigo_de_saida(etapas):
    """Antes o fluxo saía sempre com 0 e o Agendador registrava sucesso."""
    _, comportamento = etapas
    comportamento[4] = "erro"
    codigo, problemas = rodar()
    assert codigo == fluxo.SAIDA_COM_ERROS
    assert problemas["com_erro"] == [(4, 1)]


def test_contador_de_erros():
    saida = io.StringIO()
    contador = fluxo.ContadorDeErros(saida)
    contador.write("[ERRO] um\n[ERRO INESPERADO] dois\n")
    contador.write("[AVISO] não conta\n")
    assert contador.erros == 2
    assert saida.getvalue().count("\n") == 3


def test_resultado_na_ultima_linha(capsys):
    fluxo.imprimir_resultado(fluxo.SAIDA_COM_ERROS, {"quebraram": [], "com_erro": [(4, 2)]})
    assert capsys.readouterr().out.strip() == (
        "[FLUXO] Resultado: COM ERROS - rodou até o fim, mas tem [ERRO] pra conferir "
        "(código de saída 2) - [ERRO] no log: Etapa 4 (2x)")


def test_trava_contra_duas_execucoes(pasta):
    assert fluxo.adquirir_trava() is True
    assert fluxo.adquirir_trava() is False      # a primeira ainda está rodando
    fluxo.liberar_trava()
    assert fluxo.adquirir_trava() is True


def test_trava_velha_e_ignorada(pasta):
    """Execução que morreu sem limpar (queda de energia) não trava pra sempre."""
    trava = pasta / "fluxo_em_execucao.lock"
    trava.write_text("123", encoding="utf-8")
    seis_horas = 6 * 3600
    os.utime(trava, (time.time() - seis_horas, time.time() - seis_horas))
    assert fluxo.adquirir_trava() is True


def test_main_com_auto_limpa_argv_e_libera_a_trava(pasta, etapas, monkeypatch):
    """Visto no robô irmão do Banco A (21/09/2026): o "--auto" sobrava no
    sys.argv e as Etapas 2 a 4 o liam como caminho do Excel."""
    monkeypatch.setattr(sys, "argv", ["executar_fluxo_banco_b.py", "--auto"])
    stdout = sys.stdout
    assert fluxo.main() == fluxo.SAIDA_OK
    assert sys.stdout is stdout
    assert sys.argv == ["executar_fluxo_banco_b.py"]
    assert os.environ["MODO_EXECUCAO"] == "auto"
    assert not (pasta / "fluxo_em_execucao.lock").exists()
    [log] = os.listdir(pasta / "logs")
    assert "Resultado: OK" in (pasta / "logs" / log).read_text(encoding="utf-8")


def test_main_com_outra_execucao_rodando(pasta, etapas):
    chamadas, _ = etapas
    (pasta / "fluxo_em_execucao.lock").write_text("123", encoding="utf-8")
    assert fluxo.main() == fluxo.SAIDA_EM_ANDAMENTO
    assert chamadas == []
