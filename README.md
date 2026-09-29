# Extratores SAP (S_ALR_87013019)

Automacao de extracao do relatorio S_ALR_87013019 via SAP GUI Scripting.

## Uso
1. Copie `.env.example` para `.env` e preencha `SAP_USER` / `SAP_PASSWORD`
   (opcional: `SAP_CLIENT`, `WATCHDOG_SEGURANCA`, `MEMORIZAR_SEGURANCA`).
2. Ative a VPN, abra o SAP Logon e rode:
   `python Automat_andre.py`
3. Saida em `C:\Automat\*.txt`.

## Recursos
- Conexao em cascata (reaproveita sessao aberta, OpenConnection via VBS,
  abertura via sapshcut.exe) - contorna o erro 605.
- Watchdog do popup "Seguranca SAPGUI": marca "Memorizar aqui" e clica
  em "Permitir" em thread paralela.

## Executavel
Baixe o `Automat_andre.exe` na aba **Releases** (nao precisa de Python).
