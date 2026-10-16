"""Etapa 4 (agendar_vistoria_central_gestao.py): do Excel da Plataforma
pros 3 campos do modal Editar da Central de Gestão (data da vistoria,
horário e data/hora do agendamento) e o Tipo de Imóvel. Data ou hora
que não dá pra entender vira PULAR, nunca chute."""
from datetime import date, datetime, time

import pytest

import agendar_vistoria_central_gestao as agendar

CABECALHO = ["ID", "Nº Proposta", "Identificador", "Data Agendamento", "Horário Agendamento",
             "Data Vistoria", "Tipo Imovel"]


@pytest.mark.parametrize("valor, esperado", [
    ("04/09/2026", date(2026, 9, 4)),
    (" 04/09/2026 ", date(2026, 9, 4)),
    (datetime(2026, 9, 4, 15, 30), date(2026, 9, 4)),   # célula formatada como data
    (date(2026, 9, 4), date(2026, 9, 4)),
])
def test_parse_data_br(valor, esperado):
    assert agendar._parse_data_br(valor) == esperado


@pytest.mark.parametrize("valor", ["2026-09-04", "31/02/2026", ""])
def test_data_invalida_levanta_erro(valor):
    with pytest.raises(ValueError):
        agendar._parse_data_br(valor)


@pytest.mark.parametrize("valor, slot", [
    ("15:00", "15:00"),
    ("15:10", "15:00"),
    ("15:20", "15:30"),
    ("09:05:00", "09:00"),
    ("9", "09:00"),
    (time(14, 40), "14:30"),
    (datetime(2026, 9, 4, 14, 50), "15:00"),
    ("06:10", "07:00"),     # antes da 1ª opção do <select>
    ("21:00", "19:00"),     # depois da última
    # empate exato: o round() do Python arredonda pro par, então :15 e
    # :45 sempre vão pra hora cheia
    ("08:15", "08:00"),
    ("08:45", "09:00"),
])
def test_horario_vai_pro_slot_de_30_min(valor, slot):
    """O Horário da Vistoria é um <select> de 07:00 a 19:00, de 30 em 30 min."""
    assert agendar._arredondar_horario_30min(valor) == slot
    assert slot in agendar.HORARIOS_VALIDOS_MODAL


def test_opcoes_do_select():
    assert agendar.HORARIOS_VALIDOS_MODAL[0] == "07:00"
    assert agendar.HORARIOS_VALIDOS_MODAL[-1] == "19:00"
    assert len(agendar.HORARIOS_VALIDOS_MODAL) == 25


def item(**mudancas):
    base = {"num_proposta": "10000001", "data_vistoria_raw": "04/09/2026", "horario_raw": "15:10",
            "data_agendamento_raw": "01/09/2026", "tipo_imovel_raw": "Casa"}
    base.update(mudancas)
    return base


def test_valores_alvo_no_formato_do_formulario():
    assert agendar._valores_alvo(item()) == {
        "data_vistoria": "2026-09-04",
        "horario": "15:00",
        "data_agendamento": "2026-09-01T15:00",   # datetime-local leva o horário da vistoria
        "tipo_imovel": "Casa",
    }


@pytest.mark.parametrize("tipo_excel, texto", [
    ("Chácara", "Chácara"),
    ("  Galpão  ", "Galpão"),
    ("", None),
    (None, None),
])
def test_tipo_de_imovel_vai_como_texto_livre(tipo_excel, texto):
    """No Banco B o Tipo de Imóvel é um campo de texto livre (no Banco A é
    um <select> e precisa de tradução): vai o texto do Excel, sem mapa."""
    assert agendar._valores_alvo(item(tipo_imovel_raw=tipo_excel))["tipo_imovel"] == texto


def test_comparacao_ignora_o_carimbo_de_agendamento():
    """"Data e Hora do Agendamento" é o carimbo de quando o registro foi
    salvo: no Banco A, comparar com ele fazia 100% da reconferência falhar."""
    alvo = agendar._valores_alvo(item())
    atual = {"data_vistoria": "2026-09-04", "horario": "15:00", "data_agendamento": "2026-09-14T09:49"}
    assert agendar._valores_batem(atual, alvo) is True
    assert agendar._valores_batem(dict(atual, horario="15:30"), alvo) is False
    assert agendar._valores_batem(dict(atual, data_vistoria=""), alvo) is False


def test_vistoria_esperada_na_lista():
    assert agendar._formatar_vistoria_esperada(agendar._valores_alvo(item())) == "04/09/2026 15:00"


def test_le_so_linhas_com_os_tres_valores(planilha):
    caminho = planilha({"Inspeções": [
        CABECALHO,
        ["INS1", "10000001 ⭐", "INS1", "01/09/2026", "15:10", "04/09/2026", "Casa"],
        ["INS2", "10000002", "INS2", "01/09/2026", None, "05/09/2026", "Casa"],   # sem horário
        ["INS3", None, "INS3", "01/09/2026", "10:00", "05/09/2026", "Casa"],      # sem proposta
    ]})
    linhas = agendar.ler_agendamentos(caminho)
    assert len(linhas) == 1
    assert linhas[0]["num_proposta"] == "10000001"
    assert linhas[0]["identificador"] == "INS1"


def test_horario_agendamento_nao_confunde_com_data_agendamento(planilha):
    """"Data Agendamento" e "Horário Agendamento" são achadas pelo
    cabeçalho exato, mesmo fora de ordem."""
    caminho = planilha({"Inspeções": [
        ["Horário Agendamento", "Data Vistoria", "Nº Proposta", "Data Agendamento"],
        ["10:00", "05/09/2026", "10000001", "01/09/2026"],
    ]})
    [linha] = agendar.ler_agendamentos(caminho)
    assert (linha["horario_raw"], linha["data_agendamento_raw"]) == ("10:00", "01/09/2026")
    assert linha["tipo_imovel_raw"] is None and linha["identificador"] == ""


def test_montar_itens(planilha):
    """Duas inspeções da mesma proposta são duas vistorias (Identificador
    diferente) - sem essa chave a segunda seria descartada em silêncio.
    Linha repetida de verdade entra uma vez só."""
    caminho = planilha({"Inspeções": [
        CABECALHO,
        ["INS1", "10000006", "INS1002", "01/09/2026", "09:00", "02/09/2026", "Apartamento"],
        ["INS2", "10000006", "INS1003", "01/09/2026", "14:00", "08/09/2026", "Apartamento"],
        ["INS2", "10000006", "INS1003", "01/09/2026", "14:00", "08/09/2026", "Apartamento"],
        ["INS3", "10000007", "INS1004", "01/09/2026", "14h", "08/09/2026", "Casa"],
    ]})
    itens = agendar.montar_itens_agendamento(caminho)
    assert [(i["identificador"], i["acao"]) for i in itens] == [
        ("INS1002", "AGENDAR"), ("INS1003", "AGENDAR"), ("INS1004", "PULAR"),
    ]
    assert itens[1]["valores_alvo"]["data_vistoria"] == "2026-09-08"
    assert itens[2]["motivo"].startswith("não entendi a data/hora do Excel")
