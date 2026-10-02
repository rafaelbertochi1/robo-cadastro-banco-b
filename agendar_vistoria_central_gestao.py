"""
Etapa 4: preenche a data e horário de agendamento e a data da vistoria de
cada proposta já cadastrada em "Em Aberto" na Central de Gestão, usando 3
colunas do Excel exportado da Plataforma (mesmo arquivo usado pelas outras
etapas): "Data Agendamento" (dd/mm/aaaa), "Horário Agendamento" (HH:MM) e
"Data Vistoria" (dd/mm/aaaa).

Adaptado do robô irmão que já faz essa mesma etapa pro Banco A
(robo-cadastro-banco-a) - mesma mecânica (aberta, testada e ajustada lá
depois de execuções reais em volume). Os seletores abaixo foram
conferidos um a um contra o HTML real do Banco B
(mapear_central_gestao.py, 17/09/2026) - nada aqui foi chutado:

    #table (tabela)              Banco A usa #tabelaBancoA - troca confirmada
    #data_input / #hora_input /
    #data_ag_input (campos)      IDÊNTICOS aos do Banco A (mesmo nome/tipo)
    #editGestaoBancoBModal<id> (modal Editar) - Banco A usa
                                  #editGestaoBancoAModal<id>: mesmo padrão
                                  (nome do cliente + "Modal" + id), só
                                  troca o nome do cliente
    campos required no formulário: Identificador, CEP, Logradouro,
                                  Cidade - mesmos 4 do Banco A, confirmados
                                  na captura real (a Banco A também citava
                                  "Eng", não visto aqui, mas a checagem é
                                  genérica via form.checkValidity() e não
                                  depende de saber os nomes de antemão)

1. Cada proposta, na tela Consultar (gestao-banco-b), tem um botão
   "Editar" (ícone de lápis, title="Editar") no mesmo grupo de botões
   escondidos já usado pra abrir "Atribuir Engenheiro" (Etapa 3) -
   primeiro clica na seta (toggleBotoes) pra revelar os botões, depois
   no lápis, que abre o modal #editGestaoBancoBModal<id_interno>.
2. Esse modal é um formulário de EDIÇÃO GERAL da proposta (endereço,
   proponente, engenheiro etc.) - só nos interessam 3 campos dele:
     - #data_input  (name="data_visita_data",  type="date")           -> Data da Vistoria
     - #hora_input  (name="data_visita_hora",  <select>)              -> Horário da Vistoria
     - #data_ag_input (name="data_agendamento", type="datetime-local") -> Data e Hora do Agendamento
   Só preenchemos esses 3 - o resto do formulário já vem preenchido com
   os dados atuais da proposta (value="..." no HTML) e é reenviado do
   jeito que está, sem risco de sobrescrever nada.
3. O campo Horário da Vistoria é um <select> com opções FIXAS de 30 em
   30 minutos (07:00, 07:30, ..., 19:00) no Banco A - assumido igual aqui
   (mesmo sistema, mesmo tipo de campo). Se o horário do Excel não bater
   exato com uma opção, o robô arredonda pro slot mais próximo em vez de
   pular a proposta.
4. Não existe uma coluna "Horário Vistoria" separada no Excel - só
   "Horário Agendamento", que é o horário real da visita. Por isso ele é
   usado tanto pro Horário da Vistoria quanto como a parte "hora" do
   campo combinado Data e Hora do Agendamento (junto com "Data
   Agendamento" pra parte "data").

4b. IMPORTANTE (descoberto no Banco A, em execução real): apesar do nome,
   "Data e Hora do Agendamento" NÃO é um campo de agendamento que a
   gente controle - é o carimbo de quando o registro foi salvo, mantido
   pelo próprio sistema (mesmo valor repetido em lotes salvos junto,
   minutos sempre quebrados, casos com agendamento DEPOIS da vistoria).
   Por isso ele continua sendo enviado (inofensivo), mas fica FORA da
   reconferência - no Banco A, enquanto estava dentro, ela falhava em 100%
   dos casos mesmo com o salvamento tendo funcionado.
5. Salvar: o formulário desse modal específico faz
   POST https://central-gestao.example.com/gestao-banco-b/<id> (action ".../update",
   confirmado na captura real) - só existe 1 botão "Salvar" dentro desse
   formulário, mas mesmo assim o clique é sempre escopado dentro do
   modal certo.
6. O formulário Editar tem campos `required` que não são nossos
   (Identificador, CEP, Logradouro, Cidade, confirmados na captura real
   do Banco B). Se qualquer um estiver vazio, o navegador bloqueia o
   envio em silêncio: nenhuma navegação acontece e o robô fica esperando
   até estourar o timeout. Isso é checado ANTES de clicar em Salvar
   (conferir_campos_obrigatorios) e a proposta é pulada na hora, dizendo
   qual campo falta.
7. Falta SÓ o engenheiro (campo `engenheiro_atribuicao`): no Banco A isso é
   decisão interna de quem atribui, fora do alcance do robô - vira
   `DESCARTADA` (não `PULADO`), já que ninguém vai agir manualmente
   nisso. Mesma classificação aqui; ainda não vimos esse caso numa
   captura real do Banco B (a amostra de 17/09/2026 só tinha os 4
   campos de endereço/identificador como required), mas o robô é
   inofensivo se o nome do campo não bater - só cai no PULADO genérico.
8. Tipo Imóvel (`#edit_tipo`, name="tipo_imovel"): AQUI É DIFERENTE do
   Banco A. No Banco A é um `<select>` com opções fixas (daí o
   MAPA_TIPO_IMOVEL, traduzindo o texto do Excel pra uma opção exata).
   No Banco B, confirmado na captura real (`02_atribuir_datas/frame_0.html`,
   17/09/2026), é um `<input type="text">` de texto livre - não há
   opções fixas pra traduzir. Se o campo estiver vazio no sistema, o
   robô preenche direto com o texto da coluna "Tipo Imovel" do Excel
   (sem dicionário, sem tradução) - e só pula a proposta se o Excel
   também estiver vazio (nunca inventa um valor).

SEGURANÇA (isso mexe em dado real da vistoria):
- Antes de mexer em qualquer proposta, abre o modal Editar e lê os
  campos como estão HOJE - se já baterem com o que seria preenchido,
  pula (marca OK, "já estava correto") sem clicar em Salvar.
- Modo "prévia" sempre roda primeiro: mostra proposta -> valores que
  seriam preenchidos (ou motivo de ter pulado) pra CADA linha do Excel,
  sem abrir o navegador de verdade ainda, e só segue pro envio real se
  você digitar CONFIRMAR.
- Depois de cada Salvar, reconfere lendo a coluna "Dt Vist." da própria
  linha da lista - nunca confia só no clique. Lê da lista (e não
  reabrindo o modal) porque é bem mais barato: reabrir o modal custava
  uma busca + abertura de modal a mais por proposta.
- Quando algo não bate, o log mostra o valor ESPERADO e o valor LIDO.

PARALELISMO: igual à Etapa 3 (atribuir engenheiro) - a etapa de
preencher + salvar roda em várias guias ao mesmo tempo, cada uma com seu
próprio navegador Playwright. Login e a prévia continuam sequenciais,
numa guia só; as guias paralelas só abrem depois do CONFIRMAR. Ver
comum.paralelismo_configurado pra ajustar quantas guias
(PARALELISMO_AGENDAMENTO) sem editar o código - padrão 6, validado em produção.
No Banco A, 12 guias saturou a tela real (muitas propostas voltaram sem a
data mesmo sem erro no envio) - o mesmo aviso está aqui por precaução.

⚠️ Ainda não rodou contra o sistema real do Banco B - os seletores
foram conferidos contra uma captura real da tela (não é chute), mas o
fluxo completo de agendar/salvar precisa do primeiro teste real (com
$env:HEADLESS="false" e um Excel pequeno) pra validar antes de confiar
nele com volume.

Uso: python agendar_vistoria_central_gestao.py [caminho\\para\\arquivo.xlsx]
Se não passar o caminho, usa o Excel mais recente em data/exports/.
"""

import os
import re
import sys
import threading
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import date, datetime
from datetime import date as _date

from openpyxl import Workbook, load_workbook
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeout

from comum import (
    abrir_pagina,
    termo_busca_listjs,
    Tee,
    modo_automatico,
    diagnosticar_linha_nao_achada,
    pausar_para_usuario,
    pedir_confirmacao,
    headless_configurado,
    achar_coluna_por_cabecalho,
    achar_excel_mais_recente,
    espera_carregamento_gestao,
    bloquear_recursos_visuais,
    paralelismo_configurado,
    url_sem_cache,
    abrir_gestao,
    credenciais_configuradas,
    login_automatico,
)

LOGIN_URL = "https://central-gestao.example.com/login"
GESTAO_BANCO_B_URL = "https://central-gestao.example.com/gestao-banco-b"
PALAVRA_CHAVE_CLIENTE = "banco_b"
# URL de gestão efetivamente em uso nesta execução - ver mesma lógica em
# atribuir_engenheiro_central_gestao.py (comum.abrir_gestao).
_url_gestao_em_uso = GESTAO_BANCO_B_URL
PASTA_SCRIPT = os.path.dirname(os.path.abspath(__file__))
EXPORTS_DIR = os.path.join(PASTA_SCRIPT, "data", "exports")
# mesmo arquivo de sessão dos outros scripts da Central de Gestão.
SESSION_FILE = os.path.join(PASTA_SCRIPT, "sessao_central_gestao.json")
LOG_DIR = os.path.join(PASTA_SCRIPT, "logs")

NOME_COLUNA_PROPOSTA = "proposta"
NOME_COLUNA_DATA_AGENDAMENTO = "Data Agendamento"
NOME_COLUNA_HORARIO_AGENDAMENTO = "Horário Agendamento"
NOME_COLUNA_DATA_VISTORIA = "Data Vistoria"
NOME_COLUNA_TIPO_IMOVEL = "Tipo Imovel"
# desempata quando a mesma proposta tem mais de uma inspeção (mesma
# lógica do Banco A, mas só do lado do Excel - ver JS_REVELAR_LINHA abaixo
# sobre por que o desempate por docid não foi portado).
NOME_COLUNA_IDENTIFICADOR = "Identificador"

# Caixa "Procurar..." da tela Consultar - mesma confirmada em
# atribuir_engenheiro_central_gestao.py: NÃO tem name="search" no Banco B
# (diferente do Banco A, que tem).
CAMPO_BUSCA = "input.search[type='search']"

# opções reais do <select> do Horário da Vistoria no Banco A (07:00 até
# 19:00, de 30 em 30 min) - assumido igual aqui (mesmo sistema, mesmo
# tipo de campo), ainda não conferido contra o Banco B.
_SLOT_MIN = 7 * 60
_SLOT_MAX = 19 * 60
HORARIOS_VALIDOS_MODAL = [f"{m // 60:02d}:{m % 60:02d}" for m in range(_SLOT_MIN, _SLOT_MAX + 1, 30)]

TIMEOUT_LONGO = 45000
# a tela de Em Aberto (Consultar) já demorou muito além do razoável pra
# ficar pronta em execuções reais no fluxo irmão do Banco A - mesmo timeout
# folgado documentado em atribuir_engenheiro_central_gestao.py.
TIMEOUT_PAGINA_EM_ABERTO = 300000
# timeout só pra navegação COMEÇAR depois do clique em Salvar. Curto de
# propósito: o que demora nesta tela é carregar, não iniciar a navegação.
# Se em 1 min nem começou, o envio foi bloqueado (ver
# conferir_campos_obrigatorios) e esperar os 5 min só queima relógio -
# foi o que aconteceu em várias propostas numa execução real do Banco A.
TIMEOUT_ENVIO = 60000
# quanto esperar a lista reaparecer DEPOIS de salvar. Se em 2 min não
# voltou, quase sempre é página de erro - e como ela fica na mesma URL
# (/gestao-banco-b/<id>), checar a URL não denuncia.
TIMEOUT_POS_ENVIO = 120000
ESPERA_CARREGAMENTO_GESTAO = espera_carregamento_gestao()

# quantas guias preenchem e salvam agendamento em paralelo (etapa
# [3/3]) - mesma estratégia da Etapa 3, ver docstring do módulo.
PARALELISMO = paralelismo_configurado("PARALELISMO_AGENDAMENTO")
# acima disso o robô avisa antes de começar: no Banco A, com 12 guias, a
# tela saturou e a maioria das propostas voltou sem a data, mesmo sem dar
# erro visível no envio. Não é um teto rígido - você decide - só um aviso
# baseado no que já aconteceu de verdade lá.
LIMITE_GUIAS_RECOMENDADO = 6

HEADLESS = headless_configurado()

# protege os prints de mais de uma guia escrevendo ao mesmo tempo.
_TRAVA_LOG = threading.Lock()


def _log(msg):
    with _TRAVA_LOG:
        print(msg)


def novo_contexto_pagina(browser, com_sessao):
    context = browser.new_context(
        accept_downloads=True,
        viewport={"width": 1920, "height": 1080} if HEADLESS else None,
        no_viewport=None if HEADLESS else True,
        storage_state=SESSION_FILE if com_sessao else None,
    )
    if HEADLESS:
        bloquear_recursos_visuais(context)
    return context, context.new_page()


def _aguardar_pagina_pronta(page):
    """Pausa fixa depois da caixa de busca já estar visível - a página
    "parece" pronta antes de estar (mesmo motivo documentado nos outros
    scripts da Central de Gestão). Usa _log porque também é chamada de dentro
    de guias paralelas."""
    _log(f"      Aguardando {ESPERA_CARREGAMENTO_GESTAO // 1000}s pra página terminar de carregar de verdade...")
    page.wait_for_timeout(ESPERA_CARREGAMENTO_GESTAO)


def _recarregar_pagina_com_seguranca(page):
    """Sai e volta pra Consultar do zero - usado depois de um erro no
    meio do processamento de uma proposta, pra garantir que a próxima
    começa numa tela limpa (nunca deixa uma exceção daqui escapar)."""
    try:
        page.goto(url_sem_cache(_url_gestao_em_uso), timeout=TIMEOUT_PAGINA_EM_ABERTO, wait_until="domcontentloaded")
        page.wait_for_selector(CAMPO_BUSCA, timeout=TIMEOUT_PAGINA_EM_ABERTO)
        _aguardar_pagina_pronta(page)
    except Exception:
        pass


def _normalizar_num_proposta(texto):
    """Tira espaços e o marcador ⭐ que a tela às vezes mostra ao lado
    do número, pra comparar com o valor do Excel sem falso negativo."""
    texto = texto.replace("⭐️", "").replace("⭐", "")
    return re.sub(r"\s+", "", texto).strip()


# Procura a linha entre TODAS as linhas (não só as visíveis) e, se ela
# estiver escondida por algum filtro da tela, revela só ela antes de
# clicar.
#
# ⚠️ Adaptado por precaução do robô irmão do Banco A, onde isso resolveu um
# problema real: lá a tela esconde com display:none as linhas do tipo
# que não está selecionado num filtro (Financiamento/Consórcio) - depois
# de cada Salvar a página volta pro filtro padrão, e toda proposta do
# outro tipo virava "não achei essa proposta pra clicar". Ainda não
# confirmado se o Banco B tem um filtro equivalente, mas a função é
# inofensiva se não houver nada escondido.
#
# NÃO porta o desempate por `docid` que o Banco A usa quando a mesma
# proposta tem mais de uma inspeção na lista (input[name="docid"] dentro
# da <tr>): conferido na captura real do Banco B (17/09/2026,
# 6 linhas) e esse campo simplesmente não existe no HTML de nenhuma
# linha - não é falta de amostra (ele apareceria em toda linha, não só
# nas duplicadas). Sem evidência de que o mesmo padrão existe aqui,
# fica só o fallback já seguro: 2+ linhas com o mesmo nº de proposta e
# sem como desempatar = None (localizar_linha), classificado como
# "duplicada" por diagnosticar_linha_nao_achada -> PULADO pra revisão
# manual, nunca ação na linha errada.
JS_REVELAR_LINHA = """
(numProposta) => {
    const tabela = document.querySelector('#table');
    if (!tabela) return {qtd: 0};
    const limpa = (el) => (el ? el.textContent : '')
        .replace(/⭐️|⭐/g, '').replace(/\s+/g, '').trim();

    const achados = [];
    for (const tr of tabela.querySelectorAll(':scope > tbody.list > tr')) {
        const celula = tr.querySelector('td.num_proposta');
        if (celula && limpa(celula) === numProposta) achados.push(tr);
    }
    if (achados.length !== 1) return {qtd: achados.length};

    const tr = achados[0];
    const chk = tr.querySelector('input.checkbox-proposta');
    if (!chk) return {qtd: 0};
    if (tr.offsetParent === null) tr.style.display = '';
    return {qtd: 1, id: chk.value};
}
"""


def localizar_linha(page, num_proposta):
    """Usa a busca (List.js) da tela Consultar pra filtrar até sobrar
    só a linha da proposta, e devolve o id interno dela - ou None se não
    sobrar EXATAMENTE 1 linha com o número EXATO. Mesma mecânica (e
    mesmo motivo pra usar fill() + keyup em vez de digitar) já validada
    em atribuir_engenheiro_central_gestao.py."""
    campo_busca = page.locator(CAMPO_BUSCA)
    campo_busca.click(timeout=TIMEOUT_PAGINA_EM_ABERTO)
    # nunca digitar o nº cru: com ponto/hífen a busca da tela não acha
    # nada (ver comum.termo_busca_listjs). A comparação logo abaixo
    # continua sendo pelo nº EXATO.
    campo_busca.fill(termo_busca_listjs(num_proposta), timeout=TIMEOUT_PAGINA_EM_ABERTO)
    campo_busca.dispatch_event("keyup")
    page.wait_for_timeout(800)

    linhas_visiveis = page.locator("tbody.list > tr:visible")
    candidatos = []
    for i in range(linhas_visiveis.count()):
        linha = linhas_visiveis.nth(i)
        texto_coluna = _normalizar_num_proposta(linha.locator("td.num_proposta").inner_text())
        if texto_coluna == num_proposta:
            candidatos.append(linha)

    if len(candidatos) == 1:
        return candidatos[0].locator("input.checkbox-proposta").get_attribute("value")

    # nenhuma linha VISÍVEL bateu: pode ser uma proposta escondida por
    # algum filtro da tela. Procura entre todas e revela.
    if len(candidatos) == 0:
        achado = page.evaluate(JS_REVELAR_LINHA, num_proposta)
        if achado.get("qtd") == 1:
            page.wait_for_timeout(200)
            return achado["id"]
    return None


def abrir_modal_editar(page, linha_id):
    """Clica na seta (mostra os botões escondidos) e depois no lápis
    'Editar' dessa linha específica - mesmo grupo de botões usado pra
    abrir 'Atribuir Engenheiro' na Etapa 3."""
    page.locator(f'button[onclick="toggleBotoes({linha_id})"]').click(timeout=TIMEOUT_LONGO)
    botao_editar = page.locator(f'button[data-bs-target="#editGestaoBancoBModal{linha_id}"]')
    botao_editar.wait_for(state="visible", timeout=TIMEOUT_LONGO)
    botao_editar.click(timeout=TIMEOUT_LONGO)
    modal = page.locator(f"#editGestaoBancoBModal{linha_id}")
    modal.wait_for(state="visible", timeout=TIMEOUT_LONGO)
    return modal


def _fechar_modal_editar(page, linha_id, tentativas=3):
    """Fecha o modal Editar sem salvar nada - usado quando a proposta já
    está com os valores certos (nada a fazer) ou não dá pra salvar - e
    CONFERE que ele sumiu mesmo.

    Caso real (28/09/2026 11:5x, Guia 1): depois de "Já está com esses
    dados - pulando" na 10000019, a busca da proposta seguinte (10000020)
    ficou 5 min tentando clicar e desistiu com "div.col-md-3 from
    div.table-responsive subtree intercepts pointer events". Os modais
    desta tela ficam DENTRO da tabela, e os campos deles são div.col-md-3:
    era o modal Editar anterior, que continuou aberto por cima da busca.

    Causa, reproduzida com o Bootstrap 5 real: um fechar que chega antes
    da animação de abertura terminar (evento shown.bs.modal) é ignorado
    pelo Bootstrap, e o modal fica aberto de vez. Normalmente o clique
    cai depois da animação (o Playwright espera o botão parar de se
    mexer), e por isso só 1 das 53 vezes em que o robô fechou sem salvar
    deu errado. Agora, se o modal não sumir, clica de novo (a animação já
    acabou), e em último caso recarrega a tela."""
    modal = page.locator(f"#editGestaoBancoBModal{linha_id}")
    for _ in range(tentativas):
        try:
            modal.locator('button.btn-close[data-bs-dismiss="modal"]').click(timeout=TIMEOUT_LONGO)
        except Exception:
            page.keyboard.press("Escape")
        try:
            modal.wait_for(state="hidden", timeout=5000)
            return True
        except Exception:
            continue
    _log("      [AVISO] O modal Editar não fechou - recarregando a tela pra não travar a próxima proposta.")
    _recarregar_pagina_com_seguranca(page)
    return False


def ler_valores_modal_editar(page, linha_id):
    """Com o modal Editar já aberto pra essa linha: lê os valores atuais
    dos campos que nos interessam, no mesmo formato usado pra calcular
    o alvo - pra comparar direto (pré-checagem e reconferência).

    tipo_imovel: não é required no HTML, mas o Banco A descobriu que o
    servidor rejeita o Salvar quando está vazio (ver docstring do
    módulo, item 8). Lido aqui pra saber SE está vazio antes de decidir
    se precisa preencher."""
    modal = page.locator(f"#editGestaoBancoBModal{linha_id}")
    return {
        "data_vistoria": modal.locator("#data_input").input_value(timeout=TIMEOUT_LONGO),
        "horario": modal.locator("#hora_input").input_value(timeout=TIMEOUT_LONGO),
        "data_agendamento": modal.locator("#data_ag_input").input_value(timeout=TIMEOUT_LONGO),
        "tipo_imovel": modal.locator("#edit_tipo").input_value(timeout=TIMEOUT_LONGO),
    }


def conferir_campos_obrigatorios(page, linha_id):
    """Checa, ANTES de clicar em Salvar, se o formulário do modal passa na
    validação do próprio navegador - devolve a lista de campos vazios que
    bloqueariam o envio (lista vazia = pode enviar).

    Isso existe porque o formulário Editar tem campos required que não
    são nossos (Identificador, CEP, Logradouro, Cidade, confirmados na
    captura real do Banco B): se qualquer um estiver vazio, o
    navegador bloqueia o submit silenciosamente, nenhuma navegação
    acontece e o expect_navigation ficaria esperando o timeout inteiro.
    Detectamos na hora e pulamos a proposta explicando qual campo falta."""
    return page.evaluate(
        """(linhaId) => {
            const modal = document.querySelector('#editGestaoBancoBModal' + linhaId);
            if (!modal) return ['(modal não encontrado)'];
            const form = modal.querySelector('form');
            if (!form) return ['(formulário não encontrado)'];
            if (form.checkValidity()) return [];
            return Array.from(form.querySelectorAll(':invalid')).map(el => {
                const rotulo = form.querySelector('label[for="' + el.id + '"]');
                return (rotulo ? rotulo.textContent.trim() : (el.name || el.id || 'campo sem nome'));
            });
        }""",
        linha_id,
    )


def preencher_e_salvar_agendamento(page, linha_id, valores_alvo, tipo_imovel=None):
    """Com o modal Editar já aberto: preenche os campos de
    agendamento/vistoria (o resto do formulário mantém o que já veio
    preenchido) e clica Salvar DESSE modal - escopado, porque cada linha
    tem seu próprio modal Editar com os mesmos ids internos.

    `tipo_imovel`, quando informado, é preenchido também (campo de texto
    livre, #edit_tipo - ver item 8 da docstring do módulo) - só chega
    aqui quando o campo estava VAZIO no sistema. Nunca sobrescreve valor
    existente.

    Espera só 'domcontentloaded' (não o padrão 'load') - mesmo motivo já
    confirmado em atribuir_e_salvar (Etapa 3): a página é pesada e o
    clique já funciona de verdade bem antes do 'load' dar timeout.

    O timeout da navegação aqui é curto de propósito (TIMEOUT_ENVIO): o
    que demora nesta tela é a página CARREGAR, não a navegação COMEÇAR.
    Se em 1 min nem começou, foi o navegador que bloqueou o envio - não
    adianta esperar os 5 min do TIMEOUT_PAGINA_EM_ABERTO."""
    modal = page.locator(f"#editGestaoBancoBModal{linha_id}")
    if tipo_imovel:
        modal.locator("#edit_tipo").fill(tipo_imovel)
    modal.locator("#data_input").fill(valores_alvo["data_vistoria"])
    modal.locator("#hora_input").select_option(valores_alvo["horario"])
    modal.locator("#data_ag_input").fill(valores_alvo["data_agendamento"])
    botao_salvar = modal.locator("form button[type='submit']")
    with page.expect_navigation(timeout=TIMEOUT_ENVIO, wait_until="domcontentloaded"):
        botao_salvar.click(timeout=TIMEOUT_LONGO)


def ler_aviso_da_tela(page):
    """Lê a mensagem que o sistema mostra no topo depois de um envio
    (ex.: 'Vistoria atualizada com sucesso!'). Devolve '' se não houver
    nenhuma.

    É a única fonte confiável de "o servidor aceitou?" - no Banco A, uma
    execução com muitas guias em paralelo teve propostas que passaram
    pelo Salvar sem erro nenhum do Playwright e mesmo assim voltaram com
    a Dt Vist. vazia. Essa mensagem entra no log e no Excel de
    resultado."""
    try:
        return page.evaluate(
            """() => {
                const vistos = document.querySelectorAll('[class*=alert], [role=alert]');
                for (const el of vistos) {
                    if (el.offsetParent === null) continue;
                    const texto = (el.textContent || '').replace(/\\s+/g, ' ').trim();
                    if (texto) return texto.slice(0, 200);
                }
                return '';
            }"""
        )
    except Exception:
        return ""


def diagnosticar_pagina_atual(page):
    """Descreve em que página o robô parou - usado quando a tela não
    volta pra lista depois de salvar, pra não reportar só 'timeout'."""
    try:
        return f"url={page.url} | título={page.title()!r} | aviso na tela={ler_aviso_da_tela(page)!r}"
    except Exception as e:
        return f"não consegui nem ler o estado da página ({e})"


def ler_datas_da_linha(page, num_proposta):
    """Lê as colunas 'Dt Ag.' e 'Dt Vist.' direto da linha da lista, sem
    abrir modal nenhum - é assim que a reconferência é feita depois de
    salvar (bem mais barato que reabrir o modal Editar).

    Devolve None se não achar a linha. Lê por textContent e sem depender
    do filtro da tela, pelo mesmo motivo da Etapa 5."""
    return page.evaluate(
        """(numProposta) => {
            const tabela = document.querySelector('#table');
            if (!tabela) return null;
            const limpa = (el) => (el ? el.textContent : '').replace(/\\s+/g, ' ').trim();
            const cabecalhos = Array.from(tabela.querySelectorAll(':scope > thead th')).map(limpa);
            const indiceDe = (nome) => cabecalhos.findIndex(
                h => h.toLowerCase().replace(/\\.$/, '') === nome.toLowerCase().replace(/\\.$/, '')
            );
            const iAg = indiceDe('Dt Ag.');
            const iVist = indiceDe('Dt Vist.');
            if (iAg < 0 || iVist < 0) return null;

            for (const tr of tabela.querySelectorAll(':scope > tbody.list > tr')) {
                const celula = tr.querySelector('td.num_proposta');
                if (!celula) continue;
                const valor = limpa(celula).replace(/⭐️|⭐/g, '').replace(/\\s+/g, '').trim();
                if (valor !== numProposta) continue;
                const tds = Array.from(tr.children).filter(el => el.tagName === 'TD');
                return {dt_ag: limpa(tds[iAg]), dt_vist: limpa(tds[iVist])};
            }
            return null;
        }""",
        num_proposta,
    )


def _parse_data_br(valor):
    """Converte 'dd/mm/aaaa' (ou já um date/datetime, se o Excel vier
    formatado como data) pra date. Levanta erro claro se não entender."""
    if isinstance(valor, datetime):
        return valor.date()
    if isinstance(valor, _date):
        return valor
    texto = str(valor).strip()
    return datetime.strptime(texto, "%d/%m/%Y").date()


def _arredondar_horario_30min(valor):
    """Arredonda um horário ('HH:MM', ou datetime/time do Excel) pro
    slot de 30 em 30 min mais próximo dentro de 07:00-19:00 - o <select>
    da Central de Gestão só aceita esses valores fixos, não digitação livre
    (mesma decisão já tomada no fluxo do Banco A: arredondar em vez de
    pular)."""
    if isinstance(valor, (datetime,)):
        hora, minuto = valor.hour, valor.minute
    elif hasattr(valor, "hour"):  # datetime.time
        hora, minuto = valor.hour, valor.minute
    else:
        texto = str(valor).strip()
        partes = texto.split(":")
        hora, minuto = int(partes[0]), int(partes[1]) if len(partes) > 1 else 0

    total_min = hora * 60 + minuto
    total_min = max(_SLOT_MIN, min(_SLOT_MAX, total_min))
    total_min_arred = round(total_min / 30) * 30
    hora_arred, min_arred = divmod(total_min_arred, 60)
    return f"{hora_arred:02d}:{min_arred:02d}"


def _valores_alvo(item_excel):
    """Calcula os valores finais (formato pronto pra preencher no modal)
    a partir de uma linha do Excel - pode levantar exceção se a data/hora
    não fizer sentido (linha vira PULAR, nunca é chutado).

    tipo_imovel: diferente do Banco A (campo <select> com opções fixas,
    precisa de MAPA_TIPO_IMOVEL pra traduzir), o campo do Banco B é
    texto livre (#edit_tipo) - então aqui é só o texto do Excel, sem
    tradução nenhuma."""
    data_vistoria = _parse_data_br(item_excel["data_vistoria_raw"])
    horario = _arredondar_horario_30min(item_excel["horario_raw"])
    data_agendamento = _parse_data_br(item_excel["data_agendamento_raw"])
    tipo_excel = str(item_excel.get("tipo_imovel_raw") or "").strip()
    return {
        "data_vistoria": data_vistoria.strftime("%Y-%m-%d"),
        "horario": horario,
        "data_agendamento": f"{data_agendamento.strftime('%Y-%m-%d')}T{horario}",
        "tipo_imovel": tipo_excel or None,
    }


def _valores_batem(atual, alvo):
    """Compara os valores lidos do modal com o alvo calculado, mas SÓ nos
    campos que nós de fato controlamos: Data da Vistoria e Horário da
    Vistoria.

    'Data e Hora do Agendamento' fica FORA da comparação de propósito -
    no Banco A, esse campo se mostrou ser um carimbo de sistema, não um
    dado de agendamento controlável (ver docstring do módulo)."""
    return (
        atual["data_vistoria"] == alvo["data_vistoria"]
        and atual["horario"] == alvo["horario"]
    )


def _formatar_vistoria_esperada(alvo):
    """Monta como a coluna 'Dt Vist.' da lista deve ficar depois de salvar
    (ex.: '04/09/2026 15:00') - a lista mostra data e horário juntos."""
    ano, mes, dia = alvo["data_vistoria"].split("-")
    return f"{dia}/{mes}/{ano} {alvo['horario']}"


def ler_agendamentos(caminho_arquivo):
    """Lê (Nº Proposta, Data Agendamento, Horário Agendamento, Data
    Vistoria) de todas as abas do Excel - pula linha sem proposta ou sem
    os 3 valores completos (nada a agendar ali, nunca inventa um dado
    faltando)."""
    wb = load_workbook(caminho_arquivo, read_only=True, data_only=True)
    linhas = []
    for nome_aba in wb.sheetnames:
        ws = wb[nome_aba]
        col_proposta = achar_coluna_por_cabecalho(ws, NOME_COLUNA_PROPOSTA)
        col_data_ag = achar_coluna_por_cabecalho(ws, NOME_COLUNA_DATA_AGENDAMENTO, exato=True)
        col_hora_ag = achar_coluna_por_cabecalho(ws, NOME_COLUNA_HORARIO_AGENDAMENTO, exato=True)
        col_data_vist = achar_coluna_por_cabecalho(ws, NOME_COLUNA_DATA_VISTORIA, exato=True)
        # opcional: só é usado quando o Tipo de Imóvel está vazio no sistema
        col_tipo = achar_coluna_por_cabecalho(ws, NOME_COLUNA_TIPO_IMOVEL, exato=True)
        col_ident = achar_coluna_por_cabecalho(ws, NOME_COLUNA_IDENTIFICADOR, exato=True)
        if not all([col_proposta, col_data_ag, col_hora_ag, col_data_vist]):
            continue

        for row in ws.iter_rows(min_row=2):
            proposta = row[col_proposta - 1].value
            data_ag = row[col_data_ag - 1].value
            hora_ag = row[col_hora_ag - 1].value
            data_vist = row[col_data_vist - 1].value
            if not proposta or not data_ag or not hora_ag or not data_vist:
                continue
            linhas.append({
                "num_proposta": _normalizar_num_proposta(str(proposta).strip()),
                "data_agendamento_raw": data_ag,
                "horario_raw": hora_ag,
                "data_vistoria_raw": data_vist,
                "tipo_imovel_raw": row[col_tipo - 1].value if col_tipo else None,
                "identificador": str(row[col_ident - 1].value or "").strip() if col_ident else "",
            })
    return linhas


def montar_itens_agendamento(caminho_arquivo):
    """Lê o Excel e devolve a lista de itens prontos pra prévia: cada um
    já com os valores calculados, ou um motivo de ter sido pulado (dado
    faltando ou data/hora que não deu pra entender). Ignora item repetido
    (mesma chave em mais de uma aba/linha).

    A chave é o Identificador, não o nº da proposta: a mesma proposta
    pode ter mais de uma inspeção (mesmo padrão já visto no Banco A - ver
    docstring do módulo). Sem essa distinção, a segunda inspeção da
    mesma proposta seria descartada em silêncio no dedup."""
    itens = []
    vistos = set()
    for linha in ler_agendamentos(caminho_arquivo):
        num_proposta = linha["num_proposta"]
        identificador = linha.get("identificador", "")
        chave = identificador or num_proposta
        if chave in vistos:
            continue
        vistos.add(chave)
        try:
            alvo = _valores_alvo(linha)
        except Exception as e:
            itens.append({"num_proposta": num_proposta, "identificador": identificador,
                          "acao": "PULAR",
                          "motivo": f"não entendi a data/hora do Excel ({e})"})
            continue
        itens.append({"num_proposta": num_proposta, "identificador": identificador,
                      "acao": "AGENDAR", "valores_alvo": alvo})
    return itens


def agendar_lote_paralelo(indice, minha_lista):
    """Agenda uma fatia dos itens (só os com acao == 'AGENDAR') numa
    guia própria, em paralelo com as outras - mesma estratégia já
    validada em atribuir_engenheiro_central_gestao.py (cada guia abre seu
    próprio navegador Playwright do zero). Devolve a lista de resultados
    dessa fatia; nunca lança exceção pra fora."""
    prefixo = f"[Guia {indice + 1}]"
    resultados_meus = []

    with sync_playwright() as p:
        browser = p.chromium.launch(
            headless=HEADLESS,
            args=[] if HEADLESS else ["--start-maximized"],
        )
        _contexto, page = novo_contexto_pagina(browser, com_sessao=True)
        try:
            try:
                page.goto(url_sem_cache(_url_gestao_em_uso), timeout=TIMEOUT_PAGINA_EM_ABERTO, wait_until="domcontentloaded")
                page.wait_for_selector(CAMPO_BUSCA, timeout=TIMEOUT_PAGINA_EM_ABERTO)
                _aguardar_pagina_pronta(page)
            except Exception as e:
                _log(f"{prefixo} [ERRO] Não consegui abrir Em Aberto nesta guia ({e}) - "
                     f"marcando as {len(minha_lista)} proposta(s) dela pra revisão manual.")
                return [
                    {**item, "resultado": "ERRO", "detalhe": f"guia não conseguiu abrir a tela: {e}"}
                    for item in minha_lista
                ]

            for item in minha_lista:
                num_proposta = item["num_proposta"]
                alvo = item["valores_alvo"]
                _log(f"{prefixo} [{num_proposta}] Agendando {alvo['data_vistoria']} {alvo['horario']}...")

                try:
                    linha_id = localizar_linha(page, num_proposta)
                except Exception as e:
                    _log(f"{prefixo}    [ERRO] Não consegui nem localizar a proposta na tela: {e}")
                    resultados_meus.append({**item, "resultado": "ERRO", "detalhe": f"erro localizando a proposta: {e}"})
                    _recarregar_pagina_com_seguranca(page)
                    continue

                if linha_id is None:
                    try:
                        diag = diagnosticar_linha_nao_achada(page, num_proposta)
                        codigo, motivo = diag.get("codigo", "?"), diag.get("texto", "?")
                    except Exception as e:
                        codigo, motivo = "erro_diagnostico", f"não consegui nem diagnosticar ({e})"

                    # "não está mais em Em Aberto" NÃO é erro: é o estado
                    # normal de quem já foi agendada e finalizada numa
                    # execução anterior (mesma distinção já usada na
                    # Etapa 3 e no fluxo do Banco A).
                    if codigo == "fora_da_lista":
                        _log(f"{prefixo}    [OK] Nada a fazer: {motivo}")
                        resultados_meus.append({**item, "resultado": "FORA", "detalhe": motivo})
                    elif codigo in ("duplicada", "visivel_divergente", "escondida"):
                        _log(f"{prefixo}    [PULADO] {motivo}")
                        resultados_meus.append({**item, "resultado": "PULADO", "detalhe": motivo})
                    else:
                        _log(f"{prefixo}    [ERRO] Não achei essa proposta pra clicar: {motivo}")
                        resultados_meus.append({**item, "resultado": "ERRO",
                                                 "detalhe": f"proposta não localizada na tela: {motivo}"})
                    continue

                try:
                    abrir_modal_editar(page, linha_id)
                    atual = ler_valores_modal_editar(page, linha_id)
                except Exception as e:
                    _log(f"{prefixo}    [ERRO] Não consegui abrir/ler o modal Editar: {e}")
                    resultados_meus.append({**item, "resultado": "ERRO", "detalhe": f"erro abrindo modal Editar: {e}"})
                    _recarregar_pagina_com_seguranca(page)
                    continue

                if _valores_batem(atual, alvo):
                    _log(f"{prefixo}    Já está com esses dados - pulando (nada a fazer).")
                    resultados_meus.append({**item, "resultado": "OK", "detalhe": "já estava correto antes desta execução"})
                    _fechar_modal_editar(page, linha_id)
                    continue

                # o formulário Editar tem campos required que não são
                # nossos - se algum estiver vazio, o navegador bloqueia o
                # envio e ficaríamos esperando uma navegação que nunca vem
                faltando = conferir_campos_obrigatorios(page, linha_id)

                # Falta SÓ o engenheiro: decisão interna de quem atribui,
                # fora do alcance deste fluxo (mesmo critério já
                # confirmado com o usuário no robô irmão do Banco A). Não é
                # pendência nossa nem erro - DESCARTADA, registrada no
                # Excel mas fora da conta de "precisa de ação". Se faltar
                # QUALQUER outro campo junto, continua como PULADO.
                if faltando == ["engenheiro_atribuicao"]:
                    _log(f"{prefixo}    [DESCARTADA] Sem engenheiro atribuído - o sistema não deixa salvar. "
                         "Processo interno, fora do controle do robô.")
                    resultados_meus.append({**item, "resultado": "DESCARTADA",
                                             "detalhe": "sem engenheiro atribuído no sistema (processo interno)"})
                    _fechar_modal_editar(page, linha_id)
                    continue

                if faltando:
                    campos = ", ".join(faltando)
                    _log(f"{prefixo}    [PULADO] O formulário desta proposta tem campo obrigatório vazio "
                         f"({campos}) - o sistema não deixaria salvar. Preencha na mão e rode de novo.")
                    resultados_meus.append({**item, "resultado": "PULADO",
                                             "detalhe": f"campo obrigatório vazio no formulário: {campos}"})
                    _fechar_modal_editar(page, linha_id)
                    continue

                # Tipo de Imóvel: não é required no HTML, mas pode ser
                # exigido pelo servidor (mesma causa raiz já confirmada no
                # Banco A - ver item 8 da docstring do módulo). Se estiver
                # vazio no sistema, preenche direto com o texto do Excel
                # (campo de texto livre no Banco B, sem tradução); sem
                # texto no Excel também, pula explicando - nunca chuta.
                tipo_para_preencher = None
                if not atual["tipo_imovel"].strip():
                    tipo_para_preencher = alvo.get("tipo_imovel")
                    if not tipo_para_preencher:
                        _log(f"{prefixo}    [PULADO] Tipo de Imóvel está vazio no sistema e também vazio no "
                             "Excel - não dá pra preencher sem chutar.")
                        resultados_meus.append({**item, "resultado": "PULADO",
                                                 "detalhe": "Tipo de Imóvel vazio no sistema e no Excel"})
                        _fechar_modal_editar(page, linha_id)
                        continue
                    _log(f"{prefixo}    Tipo de Imóvel estava vazio - preenchendo com {tipo_para_preencher!r} do Excel.")

                try:
                    preencher_e_salvar_agendamento(page, linha_id, alvo, tipo_para_preencher)
                except Exception as e:
                    _log(f"{prefixo}    [ERRO] Falha tentando salvar: {e}")
                    resultados_meus.append({**item, "resultado": "ERRO", "detalhe": f"erro ao salvar: {e}"})
                    _recarregar_pagina_com_seguranca(page)
                    continue

                # reconferência lendo as colunas da própria lista (Dt Ag. /
                # Dt Vist.), sem reabrir o modal - o salvamento já recarrega
                # a página, então só precisamos esperar ela ficar pronta
                esperado = _formatar_vistoria_esperada(alvo)
                aviso_do_sistema = ""
                try:
                    # a mensagem aparece na página que veio logo depois do
                    # envio - lida ANTES de qualquer recarregamento, senão
                    # ela some
                    aviso_do_sistema = ler_aviso_da_tela(page)
                    # SEMPRE volta pra lista por uma URL sem cache, em vez
                    # de reaproveitar a página que o envio deixou (ver
                    # comum.url_sem_cache - mesma causa raiz de um bug real
                    # no Banco A: propostas salvas de verdade eram lidas com
                    # a data vazia porque o navegador servia a cópia
                    # anterior ao salvamento).
                    page.goto(url_sem_cache(_url_gestao_em_uso),
                              timeout=TIMEOUT_PAGINA_EM_ABERTO, wait_until="domcontentloaded")
                    try:
                        page.wait_for_selector(CAMPO_BUSCA, timeout=TIMEOUT_POS_ENVIO)
                    except PlaywrightTimeout:
                        # não adianta esperar os 5 min completos: se em 2
                        # min a lista não apareceu, quase sempre é página
                        # de erro (a tela fica na mesma URL, então checar a
                        # URL não denuncia). Diz onde parou.
                        raise RuntimeError(f"a tela não voltou pra lista depois de salvar - {diagnosticar_pagina_atual(page)}")
                    _aguardar_pagina_pronta(page)
                    datas = ler_datas_da_linha(page, num_proposta)
                    if datas is None:
                        raise RuntimeError("proposta não encontrada de novo depois de salvar")
                except Exception as e:
                    _log(f"{prefixo}    [AVISO] Salvei, mas deu erro técnico reconferindo ({e}) - confira manualmente.")
                    resultados_meus.append({**item, "resultado": "VERIFICAR", "detalhe": f"salvo, mas erro técnico ao reconferir: {e}"})
                    _recarregar_pagina_com_seguranca(page)
                    continue

                if datas["dt_vist"] == esperado:
                    _log(f"{prefixo}    Confirmado na tela: Vist. {datas['dt_vist']} "
                         f"(Ag. ficou {datas['dt_ag'] or 'vazio'}, carimbo do sistema).")
                    resultados_meus.append({**item, "resultado": "OK",
                                             "detalhe": f"confirmado na tela (Dt Vist. {datas['dt_vist']}, "
                                                        f"Dt Ag. do sistema: {datas['dt_ag'] or 'vazio'})"})
                else:
                    # loga o esperado, o lido E o que o sistema respondeu
                    resposta = f" | sistema disse: {aviso_do_sistema!r}" if aviso_do_sistema else " | sistema não disse nada"
                    _log(f"{prefixo}    [AVISO] Cliquei em Salvar, mas a Dt Vist. não bateu. "
                         f"Esperava {esperado!r}, tela mostra {datas['dt_vist']!r}{resposta}")
                    resultados_meus.append({**item, "resultado": "VERIFICAR",
                                             "detalhe": f"enviado, mas Dt Vist. não bateu: esperava {esperado}, "
                                                        f"tela mostra {datas['dt_vist'] or 'vazio'}{resposta}"})
        finally:
            browser.close()

    return resultados_meus


def _tentar_login_automatico(page):
    """Tenta logar sozinho com as credenciais de ambiente (CENTRAL_USUARIO /
    CENTRAL_SENHA). Só devolve True se realmente SAIU da tela de login -
    nunca confia em "o clique não deu erro". Sem credenciais
    configuradas, nem tenta."""
    usuario, senha = credenciais_configuradas("CENTRAL")
    if not usuario:
        return False
    print("      Sessão expirada - tentando login automático...")
    if not login_automatico(page, usuario, senha):
        print("      [AVISO] Não achei o formulário de login pra preencher sozinho.")
        return False
    try:
        page.wait_for_url(lambda url: "/login" not in url, timeout=30000)
    except PlaywrightTimeout:
        pass
    if "/login" in page.url:
        print("      [AVISO] Preenchi o login mas continuo na tela de login "
              "(credencial errada, ou o sistema pediu verificação extra).")
        return False
    return True


def main():
    os.makedirs(LOG_DIR, exist_ok=True)

    log_path = os.path.join(LOG_DIR, f"agendar_vistoria_{datetime.now():%Y%m%d_%H%M%S}.txt")
    log_file = open(log_path, "w", encoding="utf-8")
    stdout_original = sys.stdout
    sys.stdout = Tee(stdout_original, log_file)

    try:
        print("=" * 60)
        print(" AGENDAR VISTORIA - CENTRAL DE GESTÃO (BANCO B)")
        print("=" * 60)
        print(f"Log desta execução: {log_path}")

        caminho_arquivo = sys.argv[1] if len(sys.argv) > 1 else achar_excel_mais_recente(EXPORTS_DIR)
        if not caminho_arquivo or not os.path.exists(caminho_arquivo):
            print("\n[ERRO] Não achei nenhum Excel pra usar.")
            print("Exporte primeiro com exportar_status_banco_b.py, ou passe o caminho:")
            print(f"  python {os.path.basename(__file__)} caminho\\para\\arquivo.xlsx")
            return
        # mostra a data do arquivo junto do nome: no Banco A o robô já pegou
        # um export antigo, de outro período, e só deu pra perceber depois
        # pela quantidade de propostas
        idade = datetime.fromtimestamp(os.path.getmtime(caminho_arquivo))
        print(f"Arquivo lido: {caminho_arquivo}")
        print(f"      (gerado em {idade:%d/%m/%Y %H:%M})")
        if idade.date() != datetime.now().date():
            print("      [ATENÇÃO] Esse Excel não é de hoje. Se você acabou de exportar um novo,")
            print("      confira se ele está mesmo em data/exports/ - senão passe o caminho na mão:")
            print(f"        python {os.path.basename(__file__)} caminho\\para\\arquivo.xlsx")

        itens = montar_itens_agendamento(caminho_arquivo)
        if not itens:
            print("\nNenhuma linha com Nº Proposta + Data Agendamento + Horário Agendamento + "
                  "Data Vistoria preenchidos - nada a fazer.")
            return

        decisoes = itens

        with sync_playwright() as p:
            browser = p.chromium.launch(
                headless=HEADLESS,
                args=[] if HEADLESS else ["--start-maximized"],
            )
            sessao_existente = os.path.exists(SESSION_FILE)
            context, page = novo_contexto_pagina(browser, com_sessao=sessao_existente)

            def salvar_sessao():
                context.storage_state(path=SESSION_FILE)

            try:
                print("\n[1/3] Verificando sessão salva...")
                abrir_pagina(page, LOGIN_URL)
                if sessao_existente and "/login" not in page.url:
                    print("      Sessão válida, painel carregado direto.")
                elif _tentar_login_automatico(page):
                    salvar_sessao()
                    print("      Login automático OK - sessão renovada e salva.")
                elif HEADLESS or modo_automatico():
                    print("\n[ERRO] A sessão salva expirou e não consegui logar sozinho.")
                    print("Pra rodar sem ninguém na frente, configure as credenciais uma vez:")
                    print('  [Environment]::SetEnvironmentVariable("CENTRAL_USUARIO", "voce@exemplo.com", "User")')
                    print('  [Environment]::SetEnvironmentVariable("CENTRAL_SENHA", "suasenha", "User")')
                    print("Ou rode com a janela visível e logue na mão:")
                    print(f'  $env:HEADLESS="false"; python {os.path.basename(__file__)}')
                    return
                else:
                    pausar_para_usuario(
                        "A sessão salva não está mais válida (ou é a primeira vez).",
                        "1. Faça o login manualmente na janela do Chrome (e-mail e senha).",
                        "2. Espere a página carregar depois do login (a tela do Dashboard",
                        "   aparecer de vez) ANTES de voltar aqui e apertar ENTER.",
                    )
                    if "/login" in page.url:
                        try:
                            page.wait_for_url(lambda url: "/login" not in url, timeout=15000)
                        except PlaywrightTimeout:
                            pass
                    if "/login" in page.url:
                        print("\n[ERRO] Ainda estou vendo a tela de login.")
                        print("Encerrando esta execução.")
                        return
                    salvar_sessao()
                    print("      Sessão salva para as próximas execuções.")

                print("[2/3] Abrindo Vistorias > Banco B > Em Aberto (Consultar)...")
                global _url_gestao_em_uso
                url_gestao = abrir_gestao(
                    page, GESTAO_BANCO_B_URL, PALAVRA_CHAVE_CLIENTE, timeout=TIMEOUT_PAGINA_EM_ABERTO
                )
                if url_gestao is None:
                    print("\n[ERRO] Não consegui chegar na tela de gestão do Banco B.")
                    print("Encerrando esta execução sem mexer em nada.")
                    return
                _url_gestao_em_uso = url_gestao
                if "/login" in page.url:
                    print("\n[ERRO] Cheguei na tela de gestão mas voltei pra tela de login.")
                    print("Encerrando esta execução.")
                    return
                page.wait_for_selector(CAMPO_BUSCA, timeout=TIMEOUT_PAGINA_EM_ABERTO)
                _aguardar_pagina_pronta(page)

                total_agendar = sum(1 for d in decisoes if d["acao"] == "AGENDAR")
                total_pular = len(decisoes) - total_agendar

                print("\n" + "-" * 60)
                print(" PRÉVIA - nada foi salvo ainda")
                print("-" * 60)
                for d in decisoes:
                    if d["acao"] == "AGENDAR":
                        v = d["valores_alvo"]
                        print(f"  [AGENDAR] Proposta {d['num_proposta']}: Vistoria {v['data_vistoria']} "
                              f"{v['horario']} | Agendamento {v['data_agendamento']}")
                    else:
                        print(f"  [PULAR]   Proposta {d['num_proposta']}: {d['motivo']}")
                print("-" * 60)
                print(f"  {total_agendar} proposta(s) receberiam agendamento/vistoria, {total_pular} seriam puladas.")

                if total_agendar == 0:
                    print("\nNada pra agendar de verdade - encerrando sem mexer em nada.")
                    return

                confirmado = pedir_confirmacao(
                    f"{total_agendar} proposta(s) vão ter agendamento/vistoria preenchidos DE VERDADE agora.",
                    f"{total_pular} serão puladas (veja os motivos acima).",
                )
                if not confirmado:
                    print("\nCancelado pelo usuário - nada foi salvo.")
                    return

                print("\n      Sessão confirmada - fechando esta guia antes de abrir as guias")
                print("      paralelas (cada uma abre seu próprio navegador).")

            except PlaywrightTimeout as e:
                print("\n[ERRO] O robô travou esperando um elemento aparecer na tela.")
                print("Copie a mensagem abaixo e me envie, junto com o que estava")
                print("aparecendo na janela do Chrome nesse momento:")
                print(str(e))
                decisoes = []
            except Exception as e:
                print("\n[ERRO INESPERADO]")
                print(str(e))
                decisoes = []
            finally:
                try:
                    salvar_sessao()
                except Exception:
                    pass
                if not HEADLESS and not modo_automatico():
                    print("\n" + "#" * 60)
                    print("#  A AÇÃO É SUA AGORA - O ROBÔ ESTÁ PAUSADO")
                    print("#" * 60)
                    input(">>> Pressione ENTER aqui para fechar o navegador... ")
                browser.close()

        # a partir daqui, cada guia abre seu próprio navegador Playwright
        # (etapa [3/3], em paralelo) - mesma estratégia da Etapa 3.
        agendar_lista = [d for d in decisoes if d["acao"] == "AGENDAR"]
        resultados = [
            {**d, "resultado": "PULADO", "detalhe": d["motivo"]}
            for d in decisoes if d["acao"] != "AGENDAR"
        ]

        if agendar_lista:
            num_guias = min(PARALELISMO, len(agendar_lista))
            partes = [agendar_lista[i::num_guias] for i in range(num_guias)]
            print(f"\n[3/3] Agendando e salvando com {num_guias} guia(s) em paralelo "
                  f"({len(agendar_lista)} proposta(s) no total)...")
            if num_guias > LIMITE_GUIAS_RECOMENDADO:
                print(f"      [ATENÇÃO] {num_guias} guias é bastante. Numa execução real do Banco A com 12")
                print("      guias, a tela saturou: cliques e preenchimentos estouraram o tempo, a")
                print("      página levou mais de 5 min pra carregar e muitas propostas voltaram com")
                print("      a data vazia mesmo sem dar erro no envio. Se esta execução vier com muito")
                print(f"      ERRO/VERIFICAR, baixe pra {LIMITE_GUIAS_RECOMENDADO} ou menos e rode de novo -")
                print("      o que já salvou é pulado rapidinho na segunda passada.")

            with ThreadPoolExecutor(max_workers=num_guias) as executor:
                futuro_por_parte = {
                    executor.submit(agendar_lote_paralelo, i, parte): parte
                    for i, parte in enumerate(partes)
                }
                for futuro in as_completed(futuro_por_parte):
                    try:
                        resultados.extend(futuro.result())
                    except Exception as e:
                        parte_perdida = futuro_por_parte[futuro]
                        print(f"\n[ERRO INESPERADO] Uma guia falhou de forma inesperada ({e}) - "
                              f"marcando as {len(parte_perdida)} proposta(s) dela pra revisão manual.")
                        resultados.extend(
                            {**d, "resultado": "ERRO", "detalhe": f"guia falhou de forma inesperada: {e}"}
                            for d in parte_perdida
                        )

            ordem = {d["num_proposta"]: i for i, d in enumerate(decisoes)}
            resultados.sort(key=lambda r: ordem.get(r["num_proposta"], 0))

        if resultados:
            caminho_resultado = os.path.join(
                EXPORTS_DIR, f"resultado_agendamentos_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
            )
            wb_resultado = Workbook()
            ws_resultado = wb_resultado.active
            ws_resultado.title = "Resultado"
            ws_resultado.append(["Nº Proposta", "Data Vistoria", "Horário", "Data Agendamento", "Resultado", "Detalhe"])
            for r in resultados:
                v = r.get("valores_alvo") or {}
                ws_resultado.append([
                    r["num_proposta"],
                    v.get("data_vistoria", ""),
                    v.get("horario", ""),
                    v.get("data_agendamento", ""),
                    r["resultado"],
                    r["detalhe"],
                ])
            wb_resultado.save(caminho_resultado)

            conta = lambda tipo: sum(1 for r in resultados if r["resultado"] == tipo)
            total_ok, total_fora = conta("OK"), conta("FORA")
            total_descartada = conta("DESCARTADA")
            total_verificar, total_erro, total_pulado = conta("VERIFICAR"), conta("ERRO"), conta("PULADO")
            print("\nConcluído!")
            print("-" * 60)
            print(f"  {total_ok:4d} agendada(s) e confirmada(s)")
            print(f"  {total_fora:4d} já não estão em Em Aberto (nada a fazer - já finalizadas)")
            print(f"  {total_descartada:4d} sem engenheiro atribuído (processo interno, fora do controle do robô)")
            if total_verificar:
                print(f"  {total_verificar:4d} salva(s) mas não confirmada(s) - confira na tela")
            if total_pulado:
                print(f"  {total_pulado:4d} pulada(s) - precisam de ação manual (veja os motivos no Excel)")
            if total_erro:
                print(f"  {total_erro:4d} com erro técnico de verdade")
            resolvidas = total_ok + total_fora + total_descartada
            print(f"\n  Nada pendente de ação: {resolvidas} de {len(resultados)}")
            print(f"  Resultado detalhado: {caminho_resultado}")
            print("-" * 60)

    finally:
        sys.stdout = stdout_original
        log_file.close()


if __name__ == "__main__":
    main()
