# Busca de legislação nas Assembleias Legislativas

## Instalar e rodar
```bash
pip install -r requirements.txt
playwright install chromium      # necessário para sites com JavaScript e para os estados da planilha 2
streamlit run app.py
```
Rode a partir de uma máquina com acesso normal à internet (os portais .gov.br precisam ser alcançáveis).

## Como funciona
**Planilha 1 (`data/estaduais.xlsx`)** — o link já é a busca por "escola". O app troca esse valor
pelo termo digitado (qualquer parâmetro `=escola`), baixa os resultados, segue a paginação
(link "próxima", `?page=N`, postback ASP.NET) e extrai título, ementa e link direto da norma.
Se o HTML vier sem resultados (site em JavaScript), repete com navegador headless.

**Planilha 2 (`data/estaduais2.xlsx`)** — só há a página geral. Estratégia em 3 camadas:
1. **SAPL**: se a casa usa o SAPL/Interlegis (Paraíba conhecida; demais descobertas por sondagem em
   `sapl.al.<uf>.leg.br`), monta a URL de busca por ementa e lê por HTTP.
2. **Navegador headless**: abre a página, acha o campo de busca, digita o termo, envia e lê o resultado
   (HTML renderizado + JSON das APIs internas do site).
3. **Busca manual**: se nada funcionar, o estado vai para o fim do Word com o link da página de busca.

## Se um estado falhar
Aba **Diagnóstico por estado** mostra método, páginas lidas e o erro, e permite baixar o HTML capturado.
Com esse HTML dá para escrever um extrator específico (função em `busca_leg/engine.py`).

## Estrutura
- `app.py` – interface Streamlit
- `busca_leg/parsers.py` – extração genérica (tabela e âncora+contexto), paginação
- `busca_leg/engine.py` – estratégia por estado (grupos 1 e 2)
- `busca_leg/browser.py` – Playwright (modo direto e modo formulário)
- `busca_leg/docgen.py` – Word (Estado | Ementa | Link da norma)
- `tests/test_fixtures.py` – testes com servidor local simulando os padrões de portal
