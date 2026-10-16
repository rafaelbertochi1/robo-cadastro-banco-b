"""Etapa 3 (atribuir_engenheiro_central_gestao.py): qual engenheiro vai
pra cada proposta. A regra é nunca chutar: o nome só traz candidatos, e
quem decide é o CPF, que tem de bater com exatamente um deles."""
import pytest

import atribuir_engenheiro_central_gestao as atribuir


class RespostaFalsa:
    def __init__(self, dados, status=200):
        self._dados = dados
        self.status = status
        self.ok = status == 200

    def json(self):
        return self._dados


class PaginaFalsa:
    """Só o page.request.get da busca de engenheiros, com respostas por termo."""

    def __init__(self, por_termo, status=200):
        self.por_termo = por_termo
        self.status = status
        self.buscas = []
        self.request = self

    def get(self, url, params):
        assert url == atribuir.BUSCAR_ENGENHEIROS_URL
        self.buscas.append(params["q"])
        return RespostaFalsa(self.por_termo.get(params["q"], []), self.status)


def par(proposta, nome, cpf):
    return {"num_proposta": proposta, "nome_inspetor": nome, "cpf_inspetor": cpf}


ANA = {"id": 11, "nome": "Ana Teste", "cpf": "111.444.777-35"}
ANA_HOMONIMA = {"id": 12, "nome": "Ana Teste", "cpf": "529.982.247-25"}


def test_le_nome_e_cpf_das_colunas_certas(planilha):
    """"CPF Inspetor" também contém "Inspetor": a coluna do nome é achada
    pelo cabeçalho exato."""
    caminho = planilha({"Inspeções": [
        ["Nº Proposta", "CPF Inspetor", "Inspetor"],
        ["10000001", "111.444.777-35", "Ana Teste"],
        ["10000002", None, "Bruno Exemplo"],        # sem CPF: não dá pra confirmar
        [10000003, "52998224725", " Carla Modelo "],
    ]})
    assert atribuir.ler_propostas_e_engenheiros(caminho) == [
        par("10000001", "Ana Teste", "111.444.777-35"),
        par("10000003", "Carla Modelo", "52998224725"),
    ]


def test_aba_sem_as_colunas_e_pulada(planilha):
    assert atribuir.ler_propostas_e_engenheiros(planilha({"Informações": [["Usuário"], ["teste"]]})) == []


@pytest.mark.parametrize("entrada, esperado", [
    ("111.444.777-35", "11144477735"), ("111 444 777 35", "11144477735"), (None, ""), (11144477735, "11144477735"),
])
def test_normalizar_cpf(entrada, esperado):
    """O cadastro de engenheiros tem CPF com e sem pontuação."""
    assert atribuir._normalizar_cpf(entrada) == esperado


def test_normalizar_num_proposta():
    assert atribuir._normalizar_num_proposta(" 1000 0001 ⭐️") == "10000001"


def test_cpf_tem_de_bater_com_exatamente_um_candidato():
    assert atribuir.achar_correspondencia_exata_cpf([ANA, ANA_HOMONIMA], "11144477735") == ANA
    assert atribuir.achar_correspondencia_exata_cpf([ANA_HOMONIMA], "11144477735") is None
    assert atribuir.achar_correspondencia_exata_cpf([ANA, dict(ANA, id=99)], "11144477735") is None


def test_nome_curto_nem_busca():
    pagina = PaginaFalsa({})
    assert atribuir.buscar_candidatos_por_nome(pagina, " Al ") == []
    assert pagina.buscas == []


def test_resolver_engenheiros():
    pagina = PaginaFalsa({"Ana Teste": [ANA, ANA_HOMONIMA], "Bruno Exemplo": [ANA_HOMONIMA]})
    decisoes = atribuir.resolver_engenheiros(pagina, [
        par("10000001", "Ana Teste", "111.444.777-35"),
        par("10000002", "Bruno Exemplo", "222.333.444-05"),
        par("10000003", "Carla Modelo", "123.456.789-09"),
        par("10000004", "Dora Fictícia", "123.456"),
        par("10000005", "Ana Teste", "111.444.777-35"),
    ])
    acoes = [(d["num_proposta"], d["acao"], d["engenheiro"] and d["engenheiro"]["id"]) for d in decisoes]
    assert acoes == [
        ("10000001", "ATRIBUIR", 11),
        ("10000002", "PULAR", None),     # candidato existe, mas CPF diferente
        ("10000003", "PULAR", None),     # ninguém com esse nome
        ("10000004", "PULAR", None),     # CPF incompleto
        ("10000005", "ATRIBUIR", 11),
    ]
    assert "nenhum com CPF EXATAMENTE igual" in decisoes[1]["motivo"]
    assert decisoes[2]["motivo"] == "nenhum candidato encontrado buscando pelo nome"
    assert "11 dígitos" in decisoes[3]["motivo"]
    # o mesmo inspetor em duas propostas: uma busca só (cache por CPF);
    # CPF inválido nem chega a buscar
    assert pagina.buscas == ["Ana Teste", "Bruno Exemplo", "Carla Modelo"]


def test_erro_na_busca_pula_com_o_motivo():
    decisoes = atribuir.resolver_engenheiros(PaginaFalsa({}, status=500),
                                             [par("10000001", "Ana Teste", "111.444.777-35")])
    assert decisoes[0]["acao"] == "PULAR"
    assert "status 500" in decisoes[0]["motivo"]


def test_nome_com_espaco_duplo_confere():
    """"Ana  Teste" (espaço duplo no cadastro) dava "não confirmada" falso."""
    assert atribuir._sem_espacos_repetidos("Ana  Teste ") == atribuir._sem_espacos_repetidos("ana teste")
