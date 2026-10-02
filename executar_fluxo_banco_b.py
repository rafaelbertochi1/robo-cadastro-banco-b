"""
Roda o fluxo inteiro em sequência, com um único comando: exporta da
Plataforma (Etapa 1), cadastra em "Em Aberto" na Central de Gestão (Etapa 2),
atribui o engenheiro de cada proposta (Etapa 3), preenche agendamento +
vistoria de cada uma (Etapa 4) e finaliza as que ficaram completas
(Etapa 5).

Não tem lógica nova aqui - só chama, em ordem, o main() de cada script
que já existe (exportar_status_banco_b.py, cadastrar_central_gestao.py,
atribuir_engenheiro_central_gestao.py, agendar_vistoria_central_gestao.py,
finalizar_vistorias_central_gestao.py). Cada um continua funcionando
normalmente sozinho também, exatamente como antes. Adaptado do
orquestrador irmão do Banco A (executar_fluxo_banco_a.py) - mesma mecânica,
trocando só as etapas do Banco B.

A Etapa 5 é a única que não é guiada pelo Excel: ela olha a tela de Em
Aberto e finaliza toda proposta que estiver completa (engenheiro + as
duas datas), inclusive de lotes anteriores. Por isso ela vem por
último - depois que as Etapas 3 e 4 terminaram de completar o que dava
pra completar nesta rodada.

Dois modos, controlados pela variável de ambiente MODO_EXECUCAO (ver
comum.modo_automatico - mesmo padrão de variável de ambiente do
HEADLESS, nunca editar o .py na mão):

- Padrão (sem definir a variável, ou "confirmar"): pede a data inicial
  da Etapa 1 e exige digitar CONFIRMAR antes de qualquer envio real nas
  Etapas 2 a 5. Use esse modo até confiar no fluxo inteiro de ponta a
  ponta.
- "auto" (ou `--auto` na linha de comando): calcula a data inicial
  sozinho (últimos N dias - ver DIAS_JANELA_AUTOMATICA em
  exportar_status_banco_b.py, sobrepõe com a variável JANELA_DIAS) e
  confirma tudo sozinho, sem pausar esperando ninguém. Pensado pro
  fluxo agendado (ex.: Agendador de Tarefas do Windows rodando isso de
  6 em 6h) - só ligue depois de ter validado o fluxo várias vezes no
  modo padrão.

Login da Central de Gestão resolvido UMA VEZ (preparar_sessao_central), antes
das Etapas 2 a 5: elas compartilham o mesmo arquivo de sessão
(sessao_central_gestao.json), então sem isso cada uma tentaria resolver o
login por conta própria - em modo visível, uma pausa pedindo login
manual por etapa na mesma execução; em headless, um erro atrás do outro
até "Concluído" sem ter feito nada.

Cada etapa roda isolada (rodar_etapa): um erro inesperado numa delas não
derruba as seguintes - a Etapa 5, em especial, não depende do Excel
desta rodada, então vale a pena rodar mesmo que uma etapa anterior tenha
falhado.

Se a Etapa 1 não gerar um Excel novo nesta execução (sessão expirada,
erro, ou nenhum resultado no período), as Etapas 2 a 5 são puladas -
não faz sentido cadastrar propostas ou atribuir engenheiro em cima de
um Excel antigo só porque a exportação de agora não trouxe nada novo.

Uso:
    python executar_fluxo_banco_b.py           # pede CONFIRMAR em cada etapa
    python executar_fluxo_banco_b.py --auto    # roda sozinho (= MODO_EXECUCAO=auto)

Ver no README como agendar isso no Windows pra rodar sozinho de 6 em
6h com o Agendador de Tarefas.
"""

import os
import sys
from datetime import datetime

import agendar_vistoria_central_gestao
import atribuir_engenheiro_central_gestao
import cadastrar_central_gestao
import exportar_status_banco_b
import finalizar_vistorias_central_gestao
from playwright.sync_api import sync_playwright

from comum import Tee, abrir_pagina, modo_automatico, pausar_para_usuario

PASTA_SCRIPT = os.path.dirname(os.path.abspath(__file__))
LOG_DIR = os.path.join(PASTA_SCRIPT, "logs")
# trava contra duas execuções ao mesmo tempo (ex.: o Agendador disparar de
# novo enquanto a anterior ainda roda). Fora do git (.gitignore).
LOCK_FILE = os.path.join(PASTA_SCRIPT, "fluxo_em_execucao.lock")
# se a trava for mais velha que isso, a execução anterior morreu sem
# limpar (queda de energia, kill) - ignora e segue.
LOCK_MAX_HORAS = 5

# Código de saída do processo - é o que o Agendador de Tarefas registra
# ("Resultado da última execução") e o que um aviso de falha vai olhar.
# Antes o fluxo saía SEMPRE com 0: a execução agendada de 25/09/2026
# 17:15 perdeu a Etapa 1 inteira (timeout abrindo a Plataforma), pulou as
# Etapas 2 a 5, e mesmo assim o Agendador registrou sucesso.
SAIDA_OK = 0          # rodou tudo, nenhum [ERRO] no log
SAIDA_FALHOU = 1      # execução perdida: Etapa 1 sem Excel, sem login na Central ou etapa quebrou
SAIDA_COM_ERROS = 2   # rodou até o fim, mas alguma etapa/proposta deu [ERRO] - abrir o log
SAIDA_EM_ANDAMENTO = 3  # outra execução ainda rodando (trava) - não rodou nada


class ContadorDeErros:
    """Repassa tudo pro stdout de baixo e conta as linhas com "[ERRO"
    (inclui "[ERRO INESPERADO]"). Toda falha de verdade das etapas já sai
    com esse marcador, inclusive as das guias paralelas - e só elas: nas
    execuções limpas de 21 a 28/09/2026 ele não aparece nenhuma vez. Fica
    por baixo do Tee que cada etapa põe, então enxerga tudo que elas
    imprimem."""

    def __init__(self, stream):
        self.stream = stream
        self.erros = 0

    def write(self, dado):
        self.erros += dado.count("[ERRO")
        self.stream.write(dado)

    def flush(self):
        self.stream.flush()


def adquirir_trava():
    """Devolve True se pode rodar; False se outra execução está em
    andamento (trava recente)."""
    if os.path.exists(LOCK_FILE):
        idade_h = (datetime.now().timestamp() - os.path.getmtime(LOCK_FILE)) / 3600
        if idade_h < LOCK_MAX_HORAS:
            print(f"[FLUXO] Já existe uma execução em andamento (trava criada há {idade_h:.1f}h).")
            print(f"[FLUXO] Se tem certeza de que não há outra rodando, apague: {LOCK_FILE}")
            return False
        print(f"[FLUXO] Trava antiga ({idade_h:.1f}h) - a execução anterior deve ter morrido. Seguindo.")
    with open(LOCK_FILE, "w", encoding="utf-8") as f:
        f.write(f"{os.getpid()} {datetime.now():%Y-%m-%d %H:%M:%S}\n")
    return True


def liberar_trava():
    try:
        os.remove(LOCK_FILE)
    except OSError:
        pass


def preparar_sessao_central():
    """Resolve o login da Central de Gestão UMA VEZ, antes das Etapas 2 a 5.

    Sem isso, cada uma das quatro etapas descobria a sessão vencida por
    conta própria: em modo visível, quatro pausas pedindo login manual na
    mesma execução; em headless, quatro erros seguidos - o fluxo ia até o
    fim e terminava com "Concluído" sem ter feito absolutamente nada.

    As quatro etapas compartilham o mesmo arquivo de sessão
    (sessao_central_gestao.json), então basta deixá-lo válido aqui que
    todas encontram a sessão pronta. Devolve True se a sessão está boa."""
    c = cadastrar_central_gestao
    with sync_playwright() as p:
        browser = p.chromium.launch(
            headless=c.HEADLESS,
            args=[] if c.HEADLESS else ["--start-maximized"],
        )
        sessao_existente = os.path.exists(c.SESSION_FILE)
        context, page = c.novo_contexto_pagina(browser, com_sessao=sessao_existente)
        try:
            abrir_pagina(page, c.LOGIN_URL)

            if sessao_existente and "/login" not in page.url:
                print("      Sessão salva ainda vale - nada a fazer.")
                context.storage_state(path=c.SESSION_FILE)
                return True

            if c._tentar_login_automatico(page):
                context.storage_state(path=c.SESSION_FILE)
                print("      Login automático OK - sessão renovada e salva.")
                return True

            # modo automático também nunca pausa: mesmo com a janela
            # visível, quem pediu --auto não está ali pra digitar nada.
            if c.HEADLESS or modo_automatico():
                print("\n[ERRO] A sessão da Central de Gestão venceu e não consegui logar sozinho.")
                print("Configure as credenciais uma vez (depois disso o fluxo roda sem ninguém):")
                print('  [Environment]::SetEnvironmentVariable("CENTRAL_USUARIO", "voce@exemplo.com", "User")')
                print('  [Environment]::SetEnvironmentVariable("CENTRAL_SENHA", "suasenha", "User")')
                print("Ou rode esta vez com a janela visível pra logar na mão:")
                print('  $env:HEADLESS="false"; python executar_fluxo_banco_b.py')
                return False

            pausar_para_usuario(
                "A sessão da Central de Gestão venceu (ou é a primeira vez).",
                "1. Faça o login manualmente na janela do Chrome que abriu.",
                "2. Espere o Dashboard carregar de vez ANTES de voltar aqui.",
                "Você só precisa fazer isso UMA vez - as Etapas 2 a 5 vão",
                "reaproveitar esta mesma sessão.",
            )
            if "/login" in page.url:
                try:
                    page.wait_for_url(lambda url: "/login" not in url, timeout=15000)
                except Exception:
                    pass
            if "/login" in page.url:
                print("\n[ERRO] Ainda estou vendo a tela de login - não deu pra seguir.")
                return False

            context.storage_state(path=c.SESSION_FILE)
            print("      Sessão salva - as Etapas 2 a 5 vão usar esta mesma.")
            return True
        finally:
            browser.close()


def rodar_etapa(numero, titulo, funcao, *args, contador=None, problemas=None):
    """Roda uma etapa isolando falhas: um erro inesperado numa etapa não
    pode derrubar as seguintes. A Etapa 5, em especial, não depende do
    Excel - ela finaliza o que já estiver completo na tela, então vale a
    pena rodar mesmo que uma etapa anterior tenha falhado.

    Anota em `problemas` as etapas que quebraram ou imprimiram [ERRO],
    pro resumo final e o código de saída."""
    print("\n" + "#" * 60)
    print(f"# ETAPA {numero} - {titulo}")
    print("#" * 60)
    erros_antes = contador.erros if contador else 0
    try:
        return funcao(*args)
    except KeyboardInterrupt:
        raise
    except Exception as e:
        print(f"\n[FLUXO] A Etapa {numero} falhou com um erro inesperado: {e}")
        print(f"[FLUXO] Seguindo pras próximas etapas - o que ela deixou de fazer")
        print(f"[FLUXO] aparece como pendência na próxima execução.")
        if problemas is not None:
            problemas["quebraram"].append(numero)
        return None
    finally:
        if contador and problemas is not None:
            novos = contador.erros - erros_antes
            if novos:
                problemas["com_erro"].append((numero, novos))


def main():
    # --auto: mesmo efeito de MODO_EXECUCAO=auto, sem precisar fixar a
    # variável no Windows (fixar faria as execuções MANUAIS pararem de
    # pedir CONFIRMAR). É o que a tarefa agendada usa.
    #
    # O remove() NÃO é detalhe de estilo - é o que faz o resto do fluxo
    # funcionar. As Etapas 2, 3 e 4 rodam no MESMO processo e leem
    # sys.argv[1] pra saber se o usuário passou um Excel específico na
    # mão. Com "--auto" ainda em sys.argv, as três liam "--auto" como se
    # fosse o caminho do arquivo, não achavam, e desistiam com
    # "[ERRO] Não achei nenhum Excel pra usar" - mesmo com a Etapa 1
    # tendo acabado de gerar o export segundos antes. Confirmado numa
    # execução real do robô irmão do Banco A (21/09/2026): Etapa 1 exportou
    # 579 laudos e as Etapas 2, 3 e 4 não fizeram absolutamente nada.
    if "--auto" in sys.argv[1:]:
        os.environ["MODO_EXECUCAO"] = "auto"
        sys.argv.remove("--auto")

    if not adquirir_trava():
        return SAIDA_EM_ANDAMENTO

    os.makedirs(LOG_DIR, exist_ok=True)
    log_path = os.path.join(LOG_DIR, f"fluxo_completo_{datetime.now():%Y%m%d_%H%M%S}.txt")
    log_file = open(log_path, "w", encoding="utf-8")
    stdout_original = sys.stdout
    contador = ContadorDeErros(stdout_original)
    sys.stdout = Tee(contador, log_file)
    problemas = {"quebraram": [], "com_erro": []}

    try:
        try:
            codigo = executar_etapas(log_path, contador, problemas)
        except KeyboardInterrupt:
            raise
        except Exception as e:
            print(f"\n[ERRO INESPERADO] O fluxo parou fora de qualquer etapa: {e}")
            codigo = SAIDA_FALHOU
        imprimir_resultado(codigo, problemas)
        return codigo
    finally:
        sys.stdout = stdout_original
        log_file.close()
        liberar_trava()


def executar_etapas(log_path, contador, problemas):
    modo = "AUTOMÁTICO (sem pedir confirmação)" if modo_automatico() else "CONFIRMAR (pede OK antes de agir)"
    print("=" * 60)
    print(" FLUXO COMPLETO BANCO B - PLATAFORMA -> EM ABERTO -> ENGENHEIRO")
    print("=" * 60)
    print(f"Log deste resumo: {log_path}")
    print(f"Modo: {modo}")
    print(f"Início: {datetime.now():%d/%m/%Y %H:%M:%S}")

    caminho_excel = rodar_etapa(1, "Exportar da Plataforma", exportar_status_banco_b.main,
                                contador=contador, problemas=problemas)

    if not caminho_excel:
        print("\n[FLUXO] Etapa 1 não gerou um Excel novo nesta execução -")
        print("[FLUXO] pulando as Etapas 2 a 5 (nada novo pra processar).")
        return SAIDA_FALHOU

    print(f"\n[FLUXO] Etapa 1 concluída: {caminho_excel}")

    # login da Central de Gestão resolvido AQUI, uma vez só. As Etapas 2 a
    # 5 compartilham o mesmo arquivo de sessão, então sem isso cada
    # uma delas esbarrava no login por conta própria - quatro pausas
    # na mesma execução (modo visível), ou quatro erros seguidos
    # terminando em "Concluído" sem ter feito nada (headless).
    print("\n" + "#" * 60)
    print("# LOGIN - Central de Gestão (uma vez só, para as Etapas 2 a 5)")
    print("#" * 60)
    if not preparar_sessao_central():
        print("\n[FLUXO] Sem sessão válida na Central de Gestão - parando aqui.")
        print("[FLUXO] Não adianta rodar as Etapas 2 a 5: todas usam essa mesma")
        print("[FLUXO] sessão e iam parar no mesmo lugar, uma depois da outra.")
        print(f"[FLUXO] O export da Etapa 1 está salvo e continua valendo:")
        print(f"[FLUXO]   {caminho_excel}")
        return SAIDA_FALHOU

    for numero, titulo, funcao in (
        (2, "Cadastrar em Em Aberto (Central de Gestão)", cadastrar_central_gestao.main),
        (3, "Atribuir engenheiro (Central de Gestão)", atribuir_engenheiro_central_gestao.main),
        (4, "Agendar vistoria (Central de Gestão)", agendar_vistoria_central_gestao.main),
        (5, "Finalizar vistorias completas (Central de Gestão)", finalizar_vistorias_central_gestao.main),
    ):
        rodar_etapa(numero, titulo, funcao, contador=contador, problemas=problemas)

    print(f"\n[FLUXO] Concluído às {datetime.now():%d/%m/%Y %H:%M:%S}.")
    if problemas["quebraram"]:
        return SAIDA_FALHOU
    if problemas["com_erro"]:
        return SAIDA_COM_ERROS
    return SAIDA_OK


def imprimir_resultado(codigo, problemas):
    """Última linha do log: diz em uma frase se precisa olhar, e com o
    mesmo número que o Agendador de Tarefas vai mostrar."""
    detalhes = []
    if problemas["quebraram"]:
        detalhes.append("etapa(s) que quebraram: " + ", ".join(str(n) for n in problemas["quebraram"]))
    if problemas["com_erro"]:
        detalhes.append("[ERRO] no log: " + ", ".join(
            f"Etapa {n} ({q}x)" for n, q in problemas["com_erro"]))
    textos = {
        SAIDA_OK: "OK - rodou tudo, nenhum [ERRO]",
        SAIDA_FALHOU: "FALHOU - execução perdida, veja o motivo acima",
        SAIDA_COM_ERROS: "COM ERROS - rodou até o fim, mas tem [ERRO] pra conferir",
    }
    print(f"\n[FLUXO] Resultado: {textos[codigo]} (código de saída {codigo})"
          + (f" - {'; '.join(detalhes)}" if detalhes else ""))


if __name__ == "__main__":
    sys.exit(main())
