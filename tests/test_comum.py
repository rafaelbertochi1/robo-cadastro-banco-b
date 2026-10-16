"""Funções compartilhadas pelos robôs (comum.py): configuração por variável
de ambiente, código do App Autenticador, escolha do Excel certo na pasta,
trava de confirmação e nova tentativa ao abrir página."""
import base64
import io
import os
import time

import pytest
from playwright.sync_api import Error as PlaywrightError
from playwright.sync_api import TimeoutError as PlaywrightTimeout

import comum


# --- configuração por variável de ambiente ----------------------------------

@pytest.mark.parametrize("valor, esperado", [
    (None, True),        # padrão: sem janela
    ("false", False),
    ("0", False),
    ("Não", False),
    ("true", True),
    ("qualquer", True),
])
def test_headless_configurado(monkeypatch, valor, esperado):
    if valor is not None:
        monkeypatch.setenv("HEADLESS", valor)
    assert comum.headless_configurado() is esperado


@pytest.mark.parametrize("valor, esperado", [(None, False), ("auto", True), (" AUTO ", True), ("confirmar", False)])
def test_modo_automatico(monkeypatch, valor, esperado):
    if valor is not None:
        monkeypatch.setenv("MODO_EXECUCAO", valor)
    assert comum.modo_automatico() is esperado


@pytest.mark.parametrize("valor, esperado", [(None, 10000), ("20000", 20000), ("0", 10000), ("abc", 10000), ("-5", 10000)])
def test_espera_carregamento_gestao(monkeypatch, valor, esperado):
    """Só número positivo muda a pausa; o resto mantém o padrão validado (10s)."""
    if valor is None:
        monkeypatch.delenv("ESPERA_CARREGAMENTO_GESTAO_MS", raising=False)
    else:
        monkeypatch.setenv("ESPERA_CARREGAMENTO_GESTAO_MS", valor)
    assert comum.espera_carregamento_gestao() == esperado


@pytest.mark.parametrize("valor, esperado", [(None, 6), ("3", 3), ("0", 6), ("seis", 6)])
def test_paralelismo_configurado(monkeypatch, valor, esperado):
    if valor is None:
        monkeypatch.delenv("PARALELISMO_TESTE", raising=False)
    else:
        monkeypatch.setenv("PARALELISMO_TESTE", valor)
    assert comum.paralelismo_configurado("PARALELISMO_TESTE") == esperado


def test_credenciais_precisam_de_usuario_e_senha(monkeypatch):
    monkeypatch.setenv("TESTE_USUARIO", " usuario@exemplo.com ")
    monkeypatch.delenv("TESTE_SENHA", raising=False)
    assert comum.credenciais_configuradas("TESTE") == (None, None)
    monkeypatch.setenv("TESTE_SENHA", "senha-de-teste")
    assert comum.credenciais_configuradas("TESTE") == ("usuario@exemplo.com", "senha-de-teste")


def test_url_sem_cache(monkeypatch):
    """No robô irmão do Banco A a Central de Gestão devolvia página em
    cache depois de salvar (é o mesmo sistema): toda URL ganha um
    parâmetro que muda a cada chamada."""
    monkeypatch.setattr(comum.time, "time", lambda: 1_790_000_000.123)
    assert comum.url_sem_cache("https://gestao.exemplo.test/lista") == \
        "https://gestao.exemplo.test/lista?_=1790000000123"
    assert comum.url_sem_cache("https://gestao.exemplo.test/lista?pagina=2") == \
        "https://gestao.exemplo.test/lista?pagina=2&_=1790000000123"


# --- App Autenticador (TOTP) -----------------------------------------------

CHAVE_RFC = base64.b32encode(b"12345678901234567890").decode()


@pytest.mark.parametrize("instante, codigo", [
    (59, "94287082"), (1111111109, "07081804"), (1234567890, "89005924"), (20000000000, "65353130"),
])
def test_totp_bate_com_os_vetores_da_rfc_6238(instante, codigo):
    """Mesmo cálculo do Google/Microsoft Authenticator (vetores SHA-1 da RFC)."""
    assert comum.gerar_codigo_totp(CHAVE_RFC, instante=instante, digitos=8) == codigo


def test_totp_padrao_tem_6_digitos_com_zero_a_esquerda():
    assert comum.gerar_codigo_totp(CHAVE_RFC, instante=1111111109) == "081804"


@pytest.mark.parametrize("entrada", [
    "jbsw y3dp ehpk 3pxp",
    "JBSW-Y3DP-EHPK-3PXP",
    "otpauth://totp/Plataforma:usuario.teste?secret=JBSWY3DPEHPK3PXP&issuer=Plataforma",
])
def test_normalizar_chave_totp(entrada):
    assert comum.normalizar_chave_totp(entrada) == "JBSWY3DPEHPK3PXP"


def test_chave_autenticador_invalida_e_ignorada(monkeypatch):
    monkeypatch.setenv("TESTE_CHAVE_AUTENTICADOR", "chave!invalida#")
    assert comum.chave_autenticador_configurada("TESTE") is None
    monkeypatch.setenv("TESTE_CHAVE_AUTENTICADOR", "jbsw y3dp ehpk 3pxp")
    assert comum.chave_autenticador_configurada("TESTE") == "JBSWY3DPEHPK3PXP"


# --- Excel da pasta de exports -----------------------------------------------

def test_excel_mais_recente_ignora_bloqueio_e_relatorios(tmp_path):
    """O "~$" que o Excel cria com a planilha aberta e os resultado_*.xlsx
    das etapas são mais novos que o export, mas não são a lista de
    propostas."""
    export = tmp_path / "Relatorio_Analitico.xlsx"
    export.write_bytes(b"x")
    os.utime(export, (time.time() - 100, time.time() - 100))
    (tmp_path / "~$Relatorio_Analitico.xlsx").write_bytes(b"x")
    (tmp_path / "resultado_agendamentos_20261001_120000.xlsx").write_bytes(b"x")
    (tmp_path / "anotacoes.txt").write_bytes(b"x")
    assert comum.achar_excel_mais_recente(str(tmp_path)) == str(export)


def test_excel_mais_recente_aceita_xls_antigo(tmp_path):
    antigo = tmp_path / "a.xlsx"
    antigo.write_bytes(b"x")
    os.utime(antigo, (time.time() - 100, time.time() - 100))
    novo = tmp_path / "b.xls"
    novo.write_bytes(b"x")
    assert comum.achar_excel_mais_recente(str(tmp_path)) == str(novo)


def test_pasta_sem_excel(tmp_path):
    assert comum.achar_excel_mais_recente(str(tmp_path)) is None


def test_contar_linhas_excel(planilha):
    """Conta as linhas de dados de todas as abas, sem cabeçalho e sem linha vazia."""
    caminho = planilha({
        "Inspeções": [["ID", "Nº Proposta"], ["INS1", "10000001"], [None, None], ["INS2", "10000002"]],
        "Outra": [["ID"], ["INS3"]],
    })
    assert comum.contar_linhas_excel(caminho) == 3


def test_contar_linhas_de_arquivo_que_nao_e_xlsx(tmp_path):
    assert comum.contar_linhas_excel(str(tmp_path / "export.xls")) is None


def test_coluna_por_cabecalho_exata_ou_por_trecho(planilha):
    """"CPF Inspetor" contém "Inspetor": sem a busca exata, o nome do
    inspetor seria lido da coluna do CPF."""
    from openpyxl import load_workbook
    ws = load_workbook(planilha({"Aba": [["Nº Proposta", "CPF Inspetor", "Inspetor"]]}))["Aba"]
    assert comum.achar_coluna_por_cabecalho(ws, "proposta") == 1
    assert comum.achar_coluna_por_cabecalho(ws, "Inspetor") == 2
    assert comum.achar_coluna_por_cabecalho(ws, "Inspetor", exato=True) == 3
    assert comum.achar_coluna_por_cabecalho(ws, "Engenheiro") is None


# --- trava de confirmação ----------------------------------------------------

@pytest.mark.parametrize("resposta, esperado", [("CONFIRMAR", True), ("confirmar", False), ("", False), ("sim", False)])
def test_pedir_confirmacao_exige_a_palavra_exata(monkeypatch, resposta, esperado):
    """Nada é enviado de verdade pra Central sem o OK explícito."""
    monkeypatch.setattr("builtins.input", lambda _pergunta: resposta)
    assert comum.pedir_confirmacao("Vou cadastrar 3 propostas.") is esperado


def test_modo_automatico_confirma_sem_perguntar(monkeypatch):
    monkeypatch.setenv("MODO_EXECUCAO", "auto")

    def nao_pode_perguntar(_pergunta):
        raise AssertionError("não devia pedir confirmação no modo automático")
    monkeypatch.setattr("builtins.input", nao_pode_perguntar)
    assert comum.pedir_confirmacao("Vou cadastrar 3 propostas.") is True


# --- log ---------------------------------------------------------------------

def test_tee_escreve_nos_dois_e_da_flush_a_cada_escrita():
    """Sem o flush a cada escrita, o log copiado com o robô rodando vinha
    incompleto."""
    class Saida(io.StringIO):
        flushes = 0

        def flush(self):
            Saida.flushes += 1
    terminal, arquivo = Saida(), Saida()
    tee = comum.Tee(terminal, arquivo)
    tee.write("linha 1\n")
    assert terminal.getvalue() == arquivo.getvalue() == "linha 1\n"
    assert Saida.flushes == 2


# --- nova tentativa ao abrir página ------------------------------------------

class PaginaFalsa:
    def __init__(self, *falhas):
        self.falhas = list(falhas)
        self.chamadas = 0

    def goto(self, url, **kwargs):
        self.chamadas += 1
        if self.falhas:
            raise self.falhas.pop(0)
        return "resposta"


@pytest.fixture
def sem_pausa(monkeypatch):
    monkeypatch.setattr(comum.time, "sleep", lambda s: None)
    monkeypatch.delenv("TENTATIVAS_ABRIR_PAGINA", raising=False)


def test_abrir_pagina_tenta_de_novo_em_timeout(sem_pausa):
    """Lentidão passageira da Plataforma perdia a execução agendada inteira."""
    pagina = PaginaFalsa(PlaywrightTimeout("Timeout 30000ms exceeded"),
                         PlaywrightError("net::ERR_CONNECTION_RESET"))
    assert comum.abrir_pagina(pagina, "https://plataforma.exemplo.test") == "resposta"
    assert pagina.chamadas == 3


def test_abrir_pagina_desiste_na_ultima_tentativa(sem_pausa):
    pagina = PaginaFalsa(*[PlaywrightTimeout("Timeout") for _ in range(3)])
    with pytest.raises(PlaywrightTimeout):
        comum.abrir_pagina(pagina, "https://plataforma.exemplo.test")
    assert pagina.chamadas == 3


def test_erro_que_nao_e_de_rede_sobe_na_hora(sem_pausa):
    pagina = PaginaFalsa(PlaywrightError("Target page has been closed"))
    with pytest.raises(PlaywrightError):
        comum.abrir_pagina(pagina, "https://plataforma.exemplo.test")
    assert pagina.chamadas == 1


def test_numero_de_tentativas_configuravel(sem_pausa, monkeypatch):
    monkeypatch.setenv("TENTATIVAS_ABRIR_PAGINA", "1")
    pagina = PaginaFalsa(PlaywrightTimeout("Timeout"))
    with pytest.raises(PlaywrightTimeout):
        comum.abrir_pagina(pagina, "https://plataforma.exemplo.test")
    assert pagina.chamadas == 1


# --- busca da tela (List.js) ---------------------------------------------------

@pytest.mark.parametrize("numero, termo", [
    ("10000001", "10000001"),
    ("3.190.620", "3 190 620"),       # com ponto a List.js não achava nada
    ("10000021-1", "10000021 1"),     # nem com hífen
    ("...", "..."),                   # nada sobra: usa o número como veio
])
def test_termo_busca_listjs(numero, termo):
    """A List.js escapa "." e "-" e depois procura o texto literal; espaço
    funciona como "todas estas partes". Quem escolhe a linha certa
    continua sendo a comparação exata do nº depois do filtro."""
    assert comum.termo_busca_listjs(numero) == termo


# --- tela de gestão do cliente ---------------------------------------------------

class Resposta:
    def __init__(self, status):
        self.status = status


class Links:
    def __init__(self, hrefs):
        self.hrefs = hrefs

    def count(self):
        return len(self.hrefs)

    def nth(self, i):
        return self

    def get_attribute(self, nome):
        return self.hrefs.pop(0)


class PaginaGestao:
    """goto devolve o status configurado por URL; o menu tem os links dados."""

    def __init__(self, status_por_url, links_do_menu=()):
        self.status_por_url = status_por_url
        self.links_do_menu = list(links_do_menu)
        self.url = "about:blank"
        self.seletor_menu = None

    def goto(self, url, **kwargs):
        status = self.status_por_url.get(url, 404)
        self.url = "https://gestao.exemplo.test/login" if status == "login" else url
        return Resposta(200 if status == "login" else status)

    def locator(self, seletor):
        self.seletor_menu = seletor
        return Links(self.links_do_menu)


URL_PADRAO = "https://gestao.exemplo.test/gestao-banco-b"


def test_gestao_abre_na_url_padrao():
    pagina = PaginaGestao({URL_PADRAO: 200})
    assert comum.abrir_gestao(pagina, URL_PADRAO, "banco_b", 1000) == URL_PADRAO


def test_gestao_procura_o_link_no_menu_quando_a_url_falha():
    """A URL do Banco B foi deduzida do padrão do Banco A: se ela não
    responder, o caminho certo vem do menu - esta etapa grava no sistema,
    então nunca insiste num palpite."""
    certa = "https://gestao.exemplo.test/gestao_banco_b/consultar"
    pagina = PaginaGestao({URL_PADRAO: 404, certa: 200}, links_do_menu=[certa, certa])
    assert comum.abrir_gestao(pagina, URL_PADRAO, "banco_b", 1000) == certa
    assert pagina.seletor_menu == "a[href*='banco_b' i]"


def test_gestao_que_volta_pro_login_nao_conta():
    pagina = PaginaGestao({URL_PADRAO: "login"})
    assert comum.abrir_gestao(pagina, URL_PADRAO, "banco_b", 1000) is None
