# Relatório v11 — Correção para deploy Streamlit Cloud

## Problemas corrigidos

1. **Só o ArtSoul coletava**
   - ArtSoul usa `requests` e não precisa de navegador.
   - Blombo, Gagosian e Saatchi usam Playwright/Chromium; em deploy o Chromium pode não estar instalado.
   - Foi adicionada a função `garantir_playwright_chromium()` no início do app e antes de rodar scrapers Playwright.
   - Foi adicionado `packages.txt` com bibliotecas Linux necessárias ao Chromium no Streamlit Cloud.

2. **Dólar não era puxado corretamente**
   - A cotação USD-BRL agora tenta três fontes:
     - AwesomeAPI
     - Frankfurter
     - open.er-api.com
   - A função valida se o valor está numa faixa plausível antes de retornar.

3. **pyautogui no deploy**
   - `pyautogui` foi removido do `requirements.txt`.
   - O runner dos scrapers antigos remove `import pyautogui` e troca alertas por `print`, evitando erro em ambiente sem tela gráfica.

## Validação

- `app.py` compilado com sucesso.
- Todos os scrapers compilados com sucesso.
- `requirements.txt` atualizado.
- `packages.txt` incluído.
