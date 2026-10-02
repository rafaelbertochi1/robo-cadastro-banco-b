"""
Ferramenta de mapeamento do sistema de gestão (Central de Gestão).

Abre o Chrome visível. Você navega manualmente até a tela que quer
mapear (fazendo login se pedir) e volta aqui no terminal pra dar um nome
à etapa. O script então salva, dentro de exploracao/<numero>_<nome>/:

  - screenshot.png            (foto da tela inteira)
  - frame_0.html, frame_1.html, ...  (HTML de cada frame/iframe da página)
  - resumo.json                (lista de botões, links e campos de cada
                                 frame, já filtrados - mais fácil de ler
                                 que o HTML puro)

Repita pra cada tela do fluxo (ex.: login, tela Cadastrar, campo
Plataforma). No fim, compacte a pasta `exploracao` inteira num .zip e
mande aqui no chat - uso isso pra pegar os seletores reais em vez de
chutar.

Uso: python mapear_central_gestao.py
"""

import json
import os
from datetime import datetime

from playwright.sync_api import sync_playwright

PASTA_SCRIPT = os.path.dirname(os.path.abspath(__file__))
SESSION_FILE = os.path.join(PASTA_SCRIPT, "sessao_central_gestao.json")
EXPLORACAO_DIR = os.path.join(PASTA_SCRIPT, "exploracao")
LOGIN_URL = "https://central-gestao.example.com/login"

JS_RESUMO_FRAME = """
() => {
    const pegar = (sel) => Array.from(document.querySelectorAll(sel)).map(el => ({
        texto: (el.innerText || el.value || "").trim().slice(0, 120),
        tag: el.tagName.toLowerCase(),
        id: el.id || null,
        name: el.name || null,
        classe: el.className || null,
        tipo: el.type || null,
        placeholder: el.placeholder || null,
        href: el.href || null,
    }));
    return {
        botoes: pegar("button, [role=button], input[type=submit], input[type=button]"),
        links: pegar("a"),
        campos: pegar("input, select, textarea"),
    };
}
"""


def capturar_etapa(page, numero, nome):
    pasta = os.path.join(EXPLORACAO_DIR, f"{numero:02d}_{nome}")
    os.makedirs(pasta, exist_ok=True)

    page.screenshot(path=os.path.join(pasta, "screenshot.png"), full_page=True)

    resumo_frames = []
    for i, frame in enumerate(page.frames):
        try:
            html = frame.content()
        except Exception:
            html = ""
        with open(os.path.join(pasta, f"frame_{i}.html"), "w", encoding="utf-8") as f:
            f.write(html)

        try:
            dados = frame.evaluate(JS_RESUMO_FRAME)
        except Exception:
            dados = {"botoes": [], "links": [], "campos": []}
        resumo_frames.append({"indice": i, "url": frame.url, **dados})

    resumo = {
        "etapa": nome,
        "url_pagina": page.url,
        "titulo_pagina": page.title(),
        "capturado_em": datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
        "frames": resumo_frames,
    }
    with open(os.path.join(pasta, "resumo.json"), "w", encoding="utf-8") as f:
        json.dump(resumo, f, ensure_ascii=False, indent=2)

    print(f"      Salvo em: {pasta}")


def main():
    os.makedirs(EXPLORACAO_DIR, exist_ok=True)

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False, args=["--start-maximized"])
        context = browser.new_context(
            no_viewport=True,
            storage_state=SESSION_FILE if os.path.exists(SESSION_FILE) else None,
        )
        page = context.new_page()
        page.goto(LOGIN_URL)

        print("=" * 60)
        print(" MAPEADOR - CENTRAL DE GESTÃO")
        print("=" * 60)
        print("Navegue manualmente na janela do Chrome até a tela que quer")
        print("mapear (faça login se pedir). Quando a tela estiver pronta,")
        print("volte aqui no terminal e aperte ENTER pra capturar essa tela.")

        numero = 1
        while True:
            entrada = input(
                "\nENTER = capturar esta tela (ou digite um nome curto pra "
                "ela antes do ENTER, tipo 'tela_cadastrar'). Pra encerrar, "
                "digite 'sair': "
            ).strip()
            if entrada.lower() in ("sair", "fim", "exit", "encerrar"):
                break
            nome_arquivo = entrada.lower().replace(" ", "_") if entrada else f"etapa_{numero}"
            capturar_etapa(page, numero, nome_arquivo)
            numero += 1

        context.storage_state(path=SESSION_FILE)
        browser.close()

    print("\nPronto! Agora compacte a pasta 'exploracao' em um .zip e manda")
    print("aqui no chat.")


if __name__ == "__main__":
    main()
