"""Etapa 5 (finalizar_vistorias_central_gestao.py): só é finalizada a
proposta completa (engenheiro + Dt Ag. + Dt Vist.), e cada uma que fica
de fora sai no log com o motivo."""
import pytest

import finalizar_vistorias_central_gestao as finalizar


def linha(proposta, engenheiro="Ana Teste", dt_ag="01/09/2026 09:10", dt_vist="04/09/2026 15:00"):
    return {"id": proposta[-2:], "num_proposta": proposta, "engenheiro": engenheiro,
            "dt_ag": dt_ag, "dt_vist": dt_vist}


@pytest.mark.parametrize("texto, tem", [
    ("Ana Teste", True), ("", False), (" - ", False), ("--", False),
    ("Não definido", False), ("nao definido", False),
])
def test_marcadores_de_sem_engenheiro(texto, tem):
    assert finalizar._tem_engenheiro(texto) is tem


def test_classificar_linhas():
    completas, de_fora = finalizar.classificar_linhas([
        linha("10000001"),
        linha("10000002", engenheiro="-"),
        linha("10000003", dt_ag=" ", dt_vist=""),
    ])
    assert [l["num_proposta"] for l in completas] == ["10000001"]
    assert [(l["num_proposta"], l["motivo"]) for l in de_fora] == [
        ("10000002", "sem engenheiro"),
        ("10000003", "sem Dt Ag. / sem Dt Vist."),
    ]


def test_explicar_quem_ficou_de_fora(planilha):
    """Cruza o que a tela mostra com o Excel: diz se a Etapa 4 devia ter
    agendado ou se a Plataforma ainda não tem a data (caso normal).

    Diferente do Banco A, aqui não há cruzamento com o resultado da Etapa
    4: proposta com data no Excel que a Etapa 4 pulou com motivo (CEP
    vazio, por exemplo) também aparece como "ATENÇÃO"."""
    excel = planilha({"Inspeções": [
        ["Nº Proposta", "Identificador", "Data Agendamento", "Horário Agendamento", "Data Vistoria"],
        ["10000002", "INS2", "01/09/2026", "10:00", "05/09/2026"],
        ["10000003", "INS3", None, None, None],
    ]})
    _, de_fora = finalizar.classificar_linhas([
        linha("10000001", engenheiro="-"),
        linha("10000002", dt_vist=""),
        linha("10000003", dt_vist=""),
        linha("10000004", dt_vist=""),
    ])
    explicacoes = {l["num_proposta"]: l["explicacao"] for l in finalizar.explicar_de_fora(de_fora, excel)}
    assert explicacoes["10000001"].startswith("sem engenheiro na Gestão")
    assert explicacoes["10000002"].startswith("ATENÇÃO")
    assert explicacoes["10000003"].startswith("Plataforma ainda sem data de vistoria")
    assert explicacoes["10000004"].startswith("não está no export")


def test_sem_excel_pra_cruzar():
    _, de_fora = finalizar.classificar_linhas([linha("10000001", dt_vist="")])
    assert finalizar.explicar_de_fora(de_fora, None)[0]["explicacao"] == "(sem Excel pra cruzar)"


class DialogoFalso:
    message = "Tem certeza que deseja finalizar as propostas selecionadas?"
    aceito = False

    def accept(self):
        self.aceito = True


def test_confirm_do_finalizar_e_aceito():
    """Sem aceitar o confirm() da tela, o Playwright clica em Cancelar
    sozinho e nada é finalizado (184 confirmadas na prévia, 0 finalizadas)."""
    dialogo = DialogoFalso()
    finalizar._aceitar_dialogo(dialogo)
    assert dialogo.aceito is True
