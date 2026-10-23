"""Busca de legislação nas Assembleias Legislativas estaduais.

Rodar:  streamlit run app.py
"""
from __future__ import annotations

import io
from concurrent.futures import ThreadPoolExecutor, as_completed

import pandas as pd
import streamlit as st

from busca_leg import net
from busca_leg.browser import Browser, PLAYWRIGHT_OK
from busca_leg.docgen import build_docx
from busca_leg.engine import Options, StateResult, load_sources, search_state
from busca_leg.ui import hero, inject

st.set_page_config(page_title="estaduais — busca de legislação", page_icon=(lambda p: str(p) if p.exists() else "⚖️")(__import__("pathlib").Path(__file__).parent / "assets" / "favicon.png"), layout="wide")
inject()


@st.cache_data
def sources():
    return load_sources()


SRC = sources()
ICON = {"ok": "✅", "vazio": "⚪", "erro": "❌"}

hero(len(SRC))

# ------------------------------------------------------------------ barra lateral
with st.sidebar:
    st.header("Opções")
    max_results = st.slider("Máx. de normas por estado", 5, 200, 30, 5)
    max_pages = st.slider("Máx. de páginas por estado", 1, 30, 8)
    timeout = st.slider("Tempo limite por requisição (s)", 10, 90, 30)
    use_browser = st.checkbox(
        "Usar navegador headless nos sites dinâmicos",
        value=PLAYWRIGHT_OK,
        disabled=not PLAYWRIGHT_OK,
        help="Necessário para sites feitos em JavaScript e para os estados do 2º arquivo. "
        "Requer: pip install playwright && playwright install chromium",
    )
    if not PLAYWRIGHT_OK:
        st.warning("Playwright não instalado: estados que dependem de JavaScript podem falhar.")
    verify_ssl = st.checkbox("Exigir certificado SSL válido", value=False, help="Sites .gov costumam ter cadeia incompleta.")
    workers = st.slider("Consultas simultâneas", 1, 12, 6)

    st.divider()
    st.subheader("Estados")
    g1 = [s["estado"] for s in SRC if s["grupo"] == 1]
    g2 = [s["estado"] for s in SRC if s["grupo"] == 2]
    sel1 = st.multiselect("Com busca direta (planilha 1)", g1, default=g1)
    sel2 = st.multiselect("Com estratégia adaptada (planilha 2)", g2, default=g2)

# ------------------------------------------------------------------ formulário
with st.form("busca"):
    c1, c2 = st.columns([4, 1])
    termo = c1.text_input("Tema / termo de busca", placeholder="ex.: escola, merenda, transporte escolar, meio ambiente")
    go = c2.form_submit_button("🔎 Buscar", width="stretch", type="primary")


def run(termo: str, chosen: list[dict], opt: Options) -> list[StateResult]:
    results: dict[str, StateResult] = {}
    bar = st.progress(0.0, text="Iniciando…")
    status_box = st.empty()
    total = len(chosen)
    done = 0

    def render(msg: str):
        rows = [
            {"": ICON[r.status], "Estado": r.estado, "Normas": len(r.items), "Método": r.metodo or "—", "Obs.": r.mensagem[:90]}
            for r in sorted(results.values(), key=lambda x: x.estado)
        ]
        status_box.dataframe(pd.DataFrame(rows), hide_index=True, width="stretch")
        bar.progress(min(done / (total * 1.0), 1.0), text=msg)

    # Fase 1 — HTTP em paralelo
    with ThreadPoolExecutor(max_workers=workers) as ex:
        futs = {ex.submit(search_state, s, termo, opt, net.make_session(), None, None, "http"): s for s in chosen}
        for f in as_completed(futs):
            s = futs[f]
            try:
                results[s["estado"]] = f.result()
            except Exception as e:  # noqa: BLE001
                results[s["estado"]] = StateResult(s["estado"], s["grupo"], mensagem=f"{type(e).__name__}: {e}", link_geral=s["link"], busca_url=s["link"])
            done += 1
            render(f"Fase 1/2 — HTTP: {done}/{total}")

    # Fase 2 — navegador, só para quem ainda não tem resultado
    pend = [s for s in chosen if not results[s["estado"]].items]
    if opt.use_browser and pend and PLAYWRIGHT_OK:
        try:
            with Browser(verify_ssl=opt.verify_ssl) as br:
                for i, s in enumerate(pend, 1):
                    bar.progress(i / len(pend), text=f"Fase 2/2 — navegador: {s['estado']} ({i}/{len(pend)})")
                    try:
                        results[s["estado"]] = search_state(
                            s, termo, opt, None, br, None, "browser", results[s["estado"]]
                        )
                    except Exception as e:  # noqa: BLE001
                        results[s["estado"]].mensagem += f" | Navegador: {type(e).__name__}: {e}"
                    render(f"Fase 2/2 — navegador: {i}/{len(pend)}")
        except Exception as e:  # noqa: BLE001
            st.warning(f"Não foi possível iniciar o navegador headless: {e}")
    bar.progress(1.0, text="Concluído")
    return [results[s["estado"]] for s in chosen]


if go:
    if not termo.strip():
        st.error("Digite um termo de busca.")
    else:
        chosen = [s for s in SRC if (s["grupo"] == 1 and s["estado"] in sel1) or (s["grupo"] == 2 and s["estado"] in sel2)]
        if not chosen:
            st.error("Selecione ao menos um estado.")
        else:
            opt = Options(max_results, max_pages, timeout, verify_ssl, use_browser)
            res = run(termo.strip(), chosen, opt)
            st.session_state["resultado"] = {"termo": termo.strip(), "res": res}

# ------------------------------------------------------------------ resultados
data = st.session_state.get("resultado")
if data:
    termo_r, res = data["termo"], data["res"]
    ok = [r for r in res if r.items]
    sem = [r for r in res if not r.items]
    total = sum(len(r.items) for r in ok)
    st.divider()
    m1, m2, m3 = st.columns(3)
    m1.metric("Normas encontradas", total)
    m2.metric("Estados com resultado", f"{len(ok)} de {len(res)}")
    m3.metric("Estados sem resultado", len(sem))

    docx_bytes = build_docx(termo_r, res)
    d1, d2, _ = st.columns([1, 1, 3])
    d1.download_button(
        "📄 Baixar Word (.docx)", docx_bytes, file_name=f"legislacao_{termo_r.replace(' ', '_')}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", type="primary",
    )
    df = pd.DataFrame(
        [{"Estado": r.estado, "Norma": it.titulo, "Ementa": it.ementa, "Link": it.link} for r in ok for it in r.items]
    )
    buf = io.BytesIO()
    df.to_excel(buf, index=False)
    d2.download_button("📊 Baixar Excel", buf.getvalue(), file_name=f"legislacao_{termo_r.replace(' ', '_')}.xlsx")

    tab1, tab2 = st.tabs(["Normas encontradas", "Diagnóstico por estado"])
    with tab1:
        if df.empty:
            st.info("Nenhuma norma extraída. Veja a aba de diagnóstico.")
        else:
            est = st.multiselect("Filtrar estados", sorted(df["Estado"].unique()))
            view = df[df["Estado"].isin(est)] if est else df
            st.dataframe(
                view, hide_index=True, width="stretch",
                column_config={"Link": st.column_config.LinkColumn("Link da norma", display_text="abrir"), "Ementa": st.column_config.TextColumn(width="large")},
            )
    with tab2:
        st.caption("Se um estado falhar, baixe o HTML capturado e envie para ajustar o extrator daquele portal.")
        for r in sorted(res, key=lambda x: x.estado):
            with st.expander(f"{ICON[r.status]} {r.estado} — {len(r.items)} norma(s) · {r.metodo or 'sem método bem-sucedido'}"):
                st.write(f"**Grupo:** {'busca direta' if r.grupo == 1 else 'estratégia adaptada'} · **páginas lidas:** {r.paginas} · **tempo:** {r.segundos:.1f}s")
                if r.mensagem:
                    st.write(f"**Observações:** {r.mensagem}")
                st.markdown(f"[Abrir a busca no site]({r.busca_url or r.link_geral})")
                if r.debug_html:
                    st.download_button("Baixar HTML capturado", r.debug_html.encode("utf-8"), file_name=f"debug_{r.estado}.html", key=f"dbg_{r.estado}")
else:
    st.info("Informe um tema e clique em **Buscar**.")
