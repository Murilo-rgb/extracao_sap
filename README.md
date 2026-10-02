# Extratores SAP (S_ALR_87013019)

Automacao de extracao do relatorio S_ALR_87013019 via SAP GUI Scripting.

## Uso
1. Copie `.env.example` para `.env` e preencha `SAP_USER` / `SAP_PASSWORD`
   (opcional: `SAP_CLIENT`, `WATCHDOG_SEGURANCA`, `MEMORIZAR_SEGURANCA`,
   `SAP_OUTPUT_DIR`, `DUMP_TREE`, `SEGMENT_LIMIT`).
2. Ative a VPN, abra o SAP Logon e rode:
   `python Automat_andre.py`
3. Saida em `C:\Automat\*.txt` (a pasta e criada automaticamente; mude com
   `SAP_OUTPUT_DIR`).

## Ano extraido
- O script seleciona a pasta do ano pelo TEXTO na arvore "Variacao: Exercicio"
  (ex.: `2026 2026`), nao por ID fixo (000002/000003/000004/...). Assim funciona
  mesmo quando a posicao/ID do ano muda de projeto para projeto.
- Troque o ano em `TARGET_YEAR_TEXT` no `Automat_andre.py`.

## Diagnostico
- `DUMP_TREE=1`: grava `C:\Automat\_tree_<segmento>.txt` com todos os nos lidos
  da arvore (key + texto) - util para conferir o formato real dos nos.
- `SEGMENT_LIMIT=1`: roda so o 1o segmento (piloto) antes de rodar todos os 25.

## Recursos
- Conexao em cascata (reaproveita sessao aberta, OpenConnection via VBS,
  abertura via sapshcut.exe) - contorna o erro 605.
- Watchdog do popup "Seguranca SAPGUI": marca "Memorizar aqui" e clica
  em "Permitir" em thread paralela.

## Executavel
Baixe o `Automat_andre.exe` na aba **Releases** (nao precisa de Python).
