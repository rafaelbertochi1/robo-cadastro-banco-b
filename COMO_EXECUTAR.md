# Como executar (guia rápido)

Guia prático do dia a dia: que comando rodar, como mudar os parâmetros e o
que fazer quando algo sai do normal. A explicação de **por que** cada coisa
funciona do jeito que funciona (bugs encontrados, evidências, decisões)
está no `README.md`.

Todos os comandos abaixo são para o **PowerShell do VS Code**, aberto na
pasta do projeto:

```powershell
cd C:\caminho\para\robo-cadastro-banco-b
.venv\Scripts\Activate.ps1
```

(o `(.venv)` aparece no começo da linha quando o ambiente está ativo)

---

## 1. O comando do dia a dia

```powershell
git pull --autostash origin main
python executar_fluxo_banco_b.py --auto
```

Só isso. Roda as 5 etapas de uma vez, sem janela e sem pedir confirmação:

1. Exporta da Plataforma os laudos dos últimos 30 dias
2. Cadastra as propostas novas em Em Aberto (Central de Gestão)
3. Atribui o engenheiro
4. Agenda a vistoria
5. Finaliza as que ficaram completas

Os padrões já estão ajustados: **6 guias em paralelo** nas Etapas 3 e 4,
**10s de espera** a cada carregamento de página e **sem janela**. Não
precisa passar nenhuma variável.

Tempo real medido em 24/09/2026: **6 minutos** com 21 propostas novas e
**4 minutos** repetindo logo em seguida. Com muitas propostas novas, demora
mais. Pode rodar quantas vezes quiser: cada etapa
pula o que já foi feito. Rodar duas vezes seguidas não duplica nada.

O `git pull` só é necessário quando houver atualização do código. Rodar
sempre não faz mal.

---

## 2. Outras formas de rodar

| Quero... | Comando |
|---|---|
| Rodar tudo **pedindo confirmação** antes de cada etapa (pede a data inicial e `CONFIRMAR`) | `python executar_fluxo_banco_b.py` |
| **Ver o navegador** trabalhando | `$env:HEADLESS="false"; python executar_fluxo_banco_b.py --auto` |
| Rodar **só uma etapa**, pedindo confirmação | `python <arquivo da etapa>.py` (tabela abaixo) |
| Rodar **só uma etapa**, sem confirmação | `$env:MODO_EXECUCAO="auto"; python <arquivo>.py; $env:MODO_EXECUCAO=$null` |
| Usar um Excel específico nas Etapas 2, 3 ou 4 (em vez do mais recente de `data\exports`) | `python cadastrar_central_gestao.py "C:\caminho\arquivo.xlsx"` |
| **Continuar depois de uma interrupção** (fechou o VS Code, caiu a energia...) | É só rodar o comando do dia a dia de novo |

Arquivo de cada etapa:

| Etapa | Arquivo |
|---|---|
| 1. Exportar da Plataforma | `exportar_status_banco_b.py` |
| 2. Cadastrar em Em Aberto | `cadastrar_central_gestao.py` |
| 3. Atribuir engenheiro | `atribuir_engenheiro_central_gestao.py` |
| 4. Agendar vistoria | `agendar_vistoria_central_gestao.py` |
| 5. Finalizar completas | `finalizar_vistorias_central_gestao.py` |

---

## 3. Parâmetros (como mudar sem editar o código)

Tudo é ajustado por **variável de ambiente**. Nunca edite os `.py` pra
mudar um número: isso cria conflito no próximo `git pull`.

| Variável | Padrão | O que faz |
|---|---|---|
| `PARALELISMO_ATRIBUICAO` | `6` | Quantas guias atribuem engenheiro ao mesmo tempo (Etapa 3) |
| `PARALELISMO_AGENDAMENTO` | `6` | Quantas guias agendam vistoria ao mesmo tempo (Etapa 4) |
| `ESPERA_CARREGAMENTO_GESTAO_MS` | `10000` (10s) | Pausa depois de cada carregamento de página na Central de Gestão, em milissegundos |
| `HEADLESS` | `true` | `false` mostra o navegador |
| `JANELA_DIAS` | `30` | Quantos dias pra trás a Etapa 1 exporta (no `--auto`; no modo manual, é a data sugerida) |
| `DATA_MINIMA` | `01/09/2026` | Nunca exporta antes dessa data, mesmo que a janela peça. `""` desliga |
| `TENTATIVAS_ABRIR_PAGINA` | `3` | Quantas vezes tenta abrir a página de entrada de cada etapa quando o site não responde a tempo |
| `PAUSA_NOVA_TENTATIVA_S` | `30` | Segundos de pausa entre essas tentativas |
| `MODO_EXECUCAO` | (vazio = pede confirmação) | `auto` = não pergunta nada. No fluxo completo, use `--auto` |
| `IGNORAR_LOG_LOCAL` | (vazio) | Só pro teste de reenvio (ver README). Não use no dia a dia |

### Três jeitos de mudar

**Só nesta execução.** Coloque antes do comando, na mesma linha:
```powershell
$env:PARALELISMO_AGENDAMENTO="4"; python executar_fluxo_banco_b.py --auto
```
Atenção: vale até você fechar o terminal. Pra voltar ao padrão no mesmo
terminal:
```powershell
$env:PARALELISMO_AGENDAMENTO=$null
```

**Pra sempre, neste computador.** Depois de rodar, feche e reabra o VS Code
inteiro:
```powershell
[Environment]::SetEnvironmentVariable("PARALELISMO_AGENDAMENTO", "4", "User")
```
Pra desfazer, voltando ao padrão do código:
```powershell
[Environment]::SetEnvironmentVariable("PARALELISMO_AGENDAMENTO", $null, "User")
```

**Nunca** fixe `MODO_EXECUCAO` pra sempre: as execuções manuais deixariam
de pedir confirmação. Pra rodar sem perguntas, use `--auto`.

### Quando mexer em cada um

- **Apareceram vários erros seguidos** de "não consegui localizar a
  proposta" ou timeout no Salvar: o sistema está lento. Suba a espera
  (`ESPERA_CARREGAMENTO_GESTAO_MS="15000"`) ou baixe o paralelismo
  (`"4"`).
- **Paralelismo acima de 6 não é recomendado.** No robô do Banco A, 12 guias
  saturaram a tela e as propostas voltaram sem salvar.
- **Ficou muito tempo sem rodar** (mais de 30 dias): aumente a janela só
  naquela execução (`JANELA_DIAS="45"`). Laudo que sai da janela não é
  exportado nunca mais.

---

## 4. Credenciais (uma vez por computador)

O robô loga sozinho com estas variáveis do Windows. Configure uma vez,
feche e reabra o VS Code:

```powershell
[Environment]::SetEnvironmentVariable("PLATAFORMA_USUARIO", "seu-email", "User")
[Environment]::SetEnvironmentVariable("PLATAFORMA_SENHA", "sua-senha", "User")
[Environment]::SetEnvironmentVariable("CENTRAL_USUARIO", "seu-usuario", "User")
[Environment]::SetEnvironmentVariable("CENTRAL_SENHA", "sua-senha", "User")
```

O código da verificação em duas etapas da Plataforma vem do **App
Autenticador**. A chave é configurada com `configurar_autenticador.py`
(passo a passo no README, seção **Login**). Nunca coloque a chave em
código, em chat ou no git.

Pra conferir se está tudo definido:
```powershell
python testar_autenticador.py
"PLATAFORMA: $([bool]$env:PLATAFORMA_USUARIO)/$([bool]$env:PLATAFORMA_SENHA)  CENTRAL: $([bool]$env:CENTRAL_USUARIO)/$([bool]$env:CENTRAL_SENHA)"
```
O código mostrado tem que bater com o do celular, e a segunda linha tem
que mostrar `True` nos quatro.

---

## 5. Onde ver o resultado

- **`logs\fluxo_completo_<data>.txt`**: resumo das 5 etapas de uma
  execução. É o primeiro arquivo a olhar, ou a mandar quando pedir ajuda.
  A **última linha** diz em uma frase como foi:
  `[FLUXO] Resultado: OK` (código 0), `FALHOU` (1: execução perdida),
  `COM ERROS` (2: rodou, mas tem `[ERRO]` pra conferir). O código 3 quer
  dizer que outra execução já estava rodando. O Agendador de Tarefas
  mostra esse mesmo número em "Resultado da última execução".
- `logs\execucao_*.txt`, `cadastro_central_*.txt`,
  `atribuir_engenheiro_*.txt`, `agendar_vistoria_*.txt`: log detalhado de
  cada etapa.
- `data\exports\resultado_engenheiros_*.xlsx`,
  `resultado_agendamentos_*.xlsx`, `resultado_finalizacoes_*.xlsx`: uma
  linha por proposta, com o que aconteceu e o motivo.

No fim de cada etapa aparece `Nada pendente de ação: X de Y`. O que
**sobra** (Y − X) é o que precisa de gente:

| Resultado | Significa | Precisa de ação? |
|---|---|---|
| `OK` / atribuída / agendada / finalizada | Feito e conferido na tela | Não |
| já não está em Em Aberto (`FORA`) | Já foi finalizada antes | Não |
| `DESCARTADA` | Sem engenheiro no sistema (processo interno) | Não, pelo robô |
| `PULADO` | O robô não tinha como decidir sozinho (ex.: engenheiro não achado pelo nome/CPF, proposta duplicada) | **Sim**, o motivo está no Excel de resultado |
| `VERIFICAR` / salva mas não confirmada | Salvou, mas não conseguiu confirmar na tela | **Confira** na tela da Central |
| `ERRO` | Falha de verdade | **Sim**, mande o log |

---

## 6. Quando algo sai do normal

| Mensagem | O que fazer |
|---|---|
| `[MFA] ... Não consegui concluir a verificação` | Confira `python testar_autenticador.py` contra o celular. Se não bater, sincronize o relógio do Windows (Configurações > Hora e idioma > Data e hora > Sincronizar agora) ou reconfigure a chave |
| `[MFA] ... configure o App Autenticador` | A variável `PLATAFORMA_CHAVE_AUTENTICADOR` não está definida neste terminal. Reabra o VS Code; se continuar, rode `configurar_autenticador.py` |
| `Já existe uma execução em andamento` | Outra execução ainda está rodando. Se tiver **certeza** que não está, apague o arquivo `fluxo_em_execucao.lock` e rode de novo. Depois de 5 horas a trava vence sozinha |
| `A sessão salva expirou e não consegui logar sozinho` | Credenciais ausentes ou erradas neste terminal (seção 4). Pra destravar na hora, rode com janela e logue na mão: `$env:HEADLESS="false"; python executar_fluxo_banco_b.py` |
| `git pull` reclama de arquivos locais | Use sempre `git pull --autostash ...` (já está no comando acima) |
| Trocou de computador | Configure as credenciais (seção 4) e a chave do autenticador de novo. Se o computador antigo foi perdido, desligue e ligue o "Aplicativo de Autenticação" na Plataforma pra invalidar a chave antiga |

---

## 7. Rodar sozinho de 6 em 6 horas

O passo a passo do Agendador de Tarefas do Windows está no README, seção
**Agendando de 6 em 6 horas no Windows**. Os pré-requisitos já foram
provados em 24/09/2026: login automático na Plataforma (com o App
Autenticador) e na Central de Gestão, sem ninguém olhando.
