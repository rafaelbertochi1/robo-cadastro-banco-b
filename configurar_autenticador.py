"""Configura a chave do App Autenticador da Plataforma neste computador,
sem a chave aparecer na tela, no código ou no git.

A chave é o texto fixo que o QR code da Plataforma carrega (não o código
de 6 dígitos - esse muda a cada 30s e é calculado a partir da chave).
Com ela salva na variável de ambiente PLATAFORMA_CHAVE_AUTENTICADOR, o
robô calcula sozinho o código do momento quando a Plataforma pede a
verificação em duas etapas.

Duas formas de usar:

  1) A partir de um print do QR code (o jeito mais fácil):
       python configurar_autenticador.py caminho\\do\\print.png
     Precisa do leitor de QR uma vez: python -m pip install opencv-python-headless

  2) Colando a chave em texto (se a Plataforma mostrar):
       python configurar_autenticador.py
     A chave colada NÃO aparece na tela - é normal, cole e aperte ENTER.

Antes de salvar, mostra o código de 6 dígitos que o robô calcularia
agora pra você comparar com o app do celular - só salva se bater. A
chave nunca é impressa. Depois de salvar, apague o print do QR.
"""
import getpass
import os
import sys
import time

from comum import gerar_codigo_totp, normalizar_chave_totp

NOME_VARIAVEL = "PLATAFORMA_CHAVE_AUTENTICADOR"


def _decodificar_qr(cv2, imagem):
    """Tenta alguns leitores/escalas: o leitor padrão do OpenCV falhou em
    teste local com um QR de autenticador comum (texto longo = QR denso),
    o Aruco leu; e print de tela vem em tamanhos variados."""
    cinza = cv2.cvtColor(imagem, cv2.COLOR_BGR2GRAY)
    leitores = [cv2.QRCodeDetector()]
    if hasattr(cv2, "QRCodeDetectorAruco"):
        leitores.insert(0, cv2.QRCodeDetectorAruco())
    for escala in (1, 2, 0.5, 3):
        if escala == 1:
            img = cinza
        else:
            interp = cv2.INTER_CUBIC if escala > 1 else cv2.INTER_AREA
            img = cv2.resize(cinza, None, fx=escala, fy=escala, interpolation=interp)
        for leitor in leitores:
            try:
                texto = leitor.detectAndDecode(img)[0]
            except Exception:
                continue
            if texto:
                return texto
    return None


def ler_qr_da_imagem(caminho):
    try:
        import cv2
    except ImportError as erro:
        # caso real (24/09/2026): "pip install" instalou no Python global,
        # mas o script rodava pelo Python do .venv - por isso o comando
        # abaixo usa o MESMO Python que está rodando este script.
        print(f"Não consegui carregar o leitor de QR ({erro}).")
        print(f"Python em uso: {sys.executable}")
        print("Instale o leitor NESTE Python com:")
        print(f'  & "{sys.executable}" -m pip install opencv-python-headless')
        return None
    imagem = cv2.imread(caminho)
    if imagem is None:
        print(f"Não consegui abrir a imagem: {caminho}")
        return None
    texto = _decodificar_qr(cv2, imagem)
    if not texto:
        print("Não achei um QR code legível nesse print. Tire o print só da")
        print("área do QR, sem cortar as bordas, e tente de novo.")
        return None
    if not texto.lower().startswith("otpauth://"):
        print("O QR desse print não é de App Autenticador (não começa com otpauth://).")
        return None
    return texto


def salvar_variavel_do_usuario(valor):
    """Grava em HKEY_CURRENT_USER\\Environment - o mesmo que
    [Environment]::SetEnvironmentVariable(..., "User") faz no PowerShell -
    e avisa o Windows, pra programas abertos depois já enxergarem."""
    if os.name != "nt":
        print("Isso só grava sozinho no Windows. Aqui, defina a variável")
        print(f"{NOME_VARIAVEL} manualmente.")
        return False
    import ctypes
    import winreg
    with winreg.OpenKey(winreg.HKEY_CURRENT_USER, "Environment", 0, winreg.KEY_SET_VALUE) as chave_reg:
        winreg.SetValueEx(chave_reg, NOME_VARIAVEL, 0, winreg.REG_SZ, valor)
    resultado = ctypes.c_ulong()
    ctypes.windll.user32.SendMessageTimeoutW(
        0xFFFF, 0x001A, 0, "Environment", 0x0002, 5000, ctypes.byref(resultado)
    )
    return True


def main():
    if len(sys.argv) > 1:
        bruto = ler_qr_da_imagem(sys.argv[1])
        if not bruto:
            return
    else:
        bruto = getpass.getpass("Cole a chave do App Autenticador (não vai aparecer) e ENTER: ")

    chave = normalizar_chave_totp(bruto)
    try:
        gerar_codigo_totp(chave)
    except Exception:
        print("Isso não parece uma chave de App Autenticador (esperado: letras A-Z")
        print("e números 2-7). Nada foi salvo.")
        return

    restante = 30 - int(time.time() % 30)
    if restante < 10:
        # pouco tempo pra comparar com o celular - espera o próximo código
        time.sleep(restante + 0.5)
        restante = 30 - int(time.time() % 30)
    print(f"\nCódigo que o robô calcularia AGORA: {gerar_codigo_totp(chave)}  (vale mais {restante}s)")
    resposta = input("É o mesmo que o app do celular mostra pra Plataforma? (s/n): ").strip().lower()
    if resposta not in ("s", "sim"):
        print("Nada foi salvo. Se o celular mostrou o código seguinte/anterior, confira")
        print("o relógio do Windows (Configurações > Hora > Sincronizar agora) e rode de novo.")
        return

    if salvar_variavel_do_usuario(chave):
        print(f"\nPronto: {NOME_VARIAVEL} salva só no seu usuário do Windows.")
        print("1. Feche e reabra o VS Code INTEIRO (senão o terminal não enxerga).")
        print("2. Confira com: python testar_autenticador.py")
        if len(sys.argv) > 1:
            print(f"3. Apague o print do QR ({sys.argv[1]}) - ele equivale à chave.")


if __name__ == "__main__":
    main()
