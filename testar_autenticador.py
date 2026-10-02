"""Confere a chave do App Autenticador da Plataforma sem rodar o robô.

Mostra o código de 6 dígitos que o robô usaria AGORA (a partir de
PLATAFORMA_CHAVE_AUTENTICADOR) e quantos segundos ele ainda vale - pra
comparar com o app do celular, ou pra digitar na tela de configuração
do App Autenticador da Plataforma quando ela pedir um código de
confirmação. Nunca imprime a chave.

Uso:
    python testar_autenticador.py
"""
import time

from comum import chave_autenticador_configurada, gerar_codigo_totp


def main():
    chave = chave_autenticador_configurada("PLATAFORMA")
    if not chave:
        print("PLATAFORMA_CHAVE_AUTENTICADOR não está definida neste terminal (ou é inválida).")
        print("Configure com (troque pela chave mostrada na Plataforma):")
        print('  [Environment]::SetEnvironmentVariable("PLATAFORMA_CHAVE_AUTENTICADOR", "SUACHAVE", "User")')
        print("e feche/reabra o VS Code inteiro pra variável valer.")
        return
    restante = 30 - int(time.time() % 30)
    print(f"Código atual: {gerar_codigo_totp(chave)}  (vale por mais {restante}s)")
    print("Tem que ser igual ao que o app do celular mostra neste mesmo momento.")


if __name__ == "__main__":
    main()
