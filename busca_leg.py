"""busca_leg — motor de busca, extração, Word e visual do app estaduais (módulo único).

Coloque este arquivo ao lado do app.py. Os imports do app.py continuam valendo:
    from busca_leg import net
    from busca_leg.browser import Browser, PLAYWRIGHT_OK
    from busca_leg.docgen import build_docx
    from busca_leg.engine import Options, StateResult, load_sources, search_state
    from busca_leg.ui import hero, inject
"""
from __future__ import annotations


# ======================================================================
# REDE
# ======================================================================
"""Camada HTTP: sessão com cabeçalhos de navegador, retry e tolerância a certificados ruins."""

import re
from urllib.parse import urlparse, parse_qsl, urlencode, urlunparse, quote_plus

import requests
import urllib3
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

UA = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/124.0 Safari/537.36"
)


def make_session() -> requests.Session:
    s = requests.Session()
    s.headers.update(
        {
            "User-Agent": UA,
            "Accept": "text/html,application/xhtml+xml,application/json;q=0.9,*/*;q=0.8",
            "Accept-Language": "pt-BR,pt;q=0.9,en;q=0.5",
        }
    )
    retry = Retry(total=2, backoff_factor=0.8, status_forcelist=(429, 500, 502, 503, 504), allowed_methods=None)
    s.mount("https://", HTTPAdapter(max_retries=retry))
    s.mount("http://", HTTPAdapter(max_retries=retry))
    return s


def get(session: requests.Session, url: str, timeout: int = 30, verify_ssl: bool = True, **kw) -> requests.Response:
    """GET com fallback automático para verify=False quando o certificado do site é inválido
    (comum em sites governamentais com cadeia incompleta)."""
    method = kw.pop("method", "GET")
    try:
        r = session.request(method, url, timeout=timeout, verify=verify_ssl, **kw)
    except requests.exceptions.SSLError:
        urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
        r = session.request(method, url, timeout=timeout, verify=False, **kw)
    r.raise_for_status()
    ct = r.headers.get("content-type", "").lower()
    if "charset" not in ct and "json" not in ct:
        r.encoding = r.apparent_encoding or "utf-8"
    return r


def substitute_term(url: str, term: str, old: str = "escola") -> str:
    """Troca o valor 'escola' dos parâmetros da URL pelo termo escolhido (codificado)."""
    enc = quote_plus(term.strip())
    return re.sub(rf"(?<==){re.escape(old)}(?=&|$)", enc, url, flags=re.I)


def set_query_param(url: str, name: str, value: str | int) -> str:
    p = urlparse(url)
    q = parse_qsl(p.query, keep_blank_values=True)
    found = False
    out = []
    for k, v in q:
        if k == name:
            out.append((k, str(value)))
            found = True
        else:
            out.append((k, v))
    if not found:
        out.append((name, str(value)))
    return urlunparse(p._replace(query=urlencode(out)))


def has_query_param(url: str, name: str) -> bool:
    return any(k == name for k, _ in parse_qsl(urlparse(url).query, keep_blank_values=True))

from types import SimpleNamespace
net = SimpleNamespace(UA=UA, make_session=make_session, get=get, substitute_term=substitute_term,
                      set_query_param=set_query_param, has_query_param=has_query_param)


# ======================================================================
# EXTRAÇÃO (parsers)
# ======================================================================
"""Extração genérica de normas (título, ementa, link direto) a partir de HTML.

Duas estratégias, usando a que render mais resultados:
  A) tabelas: acha a coluna "Ementa" (ou a célula de texto mais longa) e o link da linha;
  B) âncora + contexto: acha links cujo texto parece título de norma
     ("Lei nº 123, de ...") e usa o texto ao redor como ementa.
Também fornece a descoberta do link "próxima página".
"""

import re
from dataclasses import dataclass
from urllib.parse import urljoin, urldefrag, urlparse

from bs4 import BeautifulSoup, Tag

TITLE_RE = re.compile(
    r"\b(Lei(?:\s+(?:Complementar|Ordin[aá]ria|Delegada|Estadual))?|Decreto(?:[\s-]+(?:Legislativo|Lei|Executivo|Estadual))?|"
    r"Resolu[cç][aã]o|Delibera[cç][aã]o|Emenda\s+(?:Constitucional|[aà]\s+Constitui[cç][aã]o)|Constitui[cç][aã]o|"
    r"Ato(?:\s+(?:Normativo|da\s+Mesa))?|Portaria|Medida\s+Provis[oó]ria|Provimento|Instru[cç][aã]o\s+Normativa)\b"
    r"[^\n]{0,70}?\d",
    re.I,
)
NOISE_LINE_RE = re.compile(
    r"^(fonte|d\.?o\.?|di[aá]rio|publica[cç][aã]o|publicad|natureza|situa[cç][aã]o|data|autor|vig[eê]ncia|"
    r"link|texto|ver |visualizar|detalhes?|download|pdf|iniciativa|origem|status)\b",
    re.I,
)
NEXT_TEXT_RE = re.compile(r"^\s*(pr[oó]xim[ao]|seguinte|next|avan[cç]ar|›|»|>|>>|&gt;)\s*$", re.I)
HEAD_EMENTA = re.compile(r"ementa|s[uú]mula|resumo|assunto|descri[cç][aã]o|objeto", re.I)
HEAD_TITLE = re.compile(r"tipo|esp[eé]cie|n[uú]mero|^n[ºo°.]*$|^ano$|norma|legisla|lei|ato", re.I)
PAGE_PARAMS = ("page", "pagina", "pg", "pag", "p", "pageNumber", "pageindex", "numeroPagina")


@dataclass
class Item:
    titulo: str
    ementa: str
    link: str


DATE_ONLY = re.compile(r"^\W*\d{1,2}\s*[/.-]\s*\d{1,2}\s*[/.-]\s*\d{2,4}\W*$")
DATE_PREFIX = re.compile(r"^\W*\d{1,2}\s*[/.-]\s*\d{1,2}\s*[/.-]\s*\d{2,4}\W*")
NUM_ONLY = re.compile(r"^[\d.,/\s-]+$")


def weak_title(t: str) -> bool:
    """Título que não é nome de norma: vazio, só data ('09/10/2025'), só número ou botão ('Texto')."""
    t = (t or "").strip()
    return not t or bool(DATE_ONLY.match(t) or NUM_ONLY.match(t) or JUNK_TITLE.match(t))


def title_from_lines(lines: list[str], ementa: str = "") -> str:
    """Procura nas linhas/células da entrada o nome da norma ('Decreto nº 10.321', 'Lei Complementar 12/2020')."""
    for ln in lines:
        ln = DATE_PREFIX.sub("", ln).strip(" -–:•·")
        if ln and ln != ementa and len(ln) <= 100 and TITLE_RE.match(ln):
            return ln
    tipo = num = ""
    for ln in lines:
        m = re.match(r"^(?:tipo|esp[eé]cie)\s*[:\-]\s*(.+)$", ln, re.I)
        if m:
            tipo = m.group(1).strip()
        m = re.match(r"^n[uú]mero\s*[:\-]\s*(.+)$", ln, re.I)
        if m:
            num = m.group(1).strip()
    return f"{tipo} nº {num}" if tipo and num else ""


def clean(s: str | None) -> str:
    return " ".join((s or "").split())


def valid_link(href: str | None, base: str) -> str | None:
    """Devolve URL absoluta se for um link navegável; None se for lixo."""
    if not href:
        return None
    href = href.strip()
    low = href.lower()
    if low.startswith(("javascript:", "mailto:", "tel:", "data:")):
        return None
    if href.startswith("#") and not href.startswith("#/"):
        return None
    url = urljoin(base, href)
    # link para a própria página de busca não é link de norma
    if urldefrag(url)[0] == urldefrag(base)[0] and "#/" not in url:
        return None
    return url


def _strip_page(soup: BeautifulSoup) -> BeautifulSoup:
    for t in soup(["script", "style", "noscript", "nav", "footer", "header", "svg"]):
        t.decompose()
    return soup


# ---------------------------------------------------------------- tabela
def _table_items(soup: BeautifulSoup, base: str) -> list[Item]:
    out: list[Item] = []
    for table in soup.find_all("table"):
        rows = table.find_all("tr")
        if len(rows) < 2:
            continue
        header: list[str] = []
        first = rows[0]
        hcells = first.find_all(["th", "td"])
        if first.find("th") or all(not c.find("a") for c in hcells):
            header = [clean(c.get_text(" ")) for c in hcells]
            body = rows[1:]
        else:
            body = rows
        eidx = next((i for i, h in enumerate(header) if HEAD_EMENTA.search(h)), None)
        tidx = [i for i, h in enumerate(header) if HEAD_TITLE.search(h) and not HEAD_EMENTA.search(h)]
        for tr in body:
            cells = tr.find_all(["td", "th"])
            if not cells:
                continue
            link = None
            link_text = ""
            for a in tr.find_all("a", href=True):
                u = valid_link(a["href"], base)
                if u:
                    link, link_text = u, clean(a.get_text(" "))
                    break
            if not link:
                continue
            texts = [clean(c.get_text(" ")) for c in cells]
            if eidx is not None and eidx < len(texts):
                ementa = texts[eidx]
            else:
                ementa = max(texts, key=len) if texts else ""
            if tidx and len(header) == len(texts):
                titulo = " ".join(texts[i] for i in tidx if texts[i] and i != eidx)
            else:
                titulo = link_text
            titulo = clean(titulo) or link_text
            if weak_title(titulo):
                cell_lines = []
                for c in cells:
                    cell_lines += [clean(x) for x in c.get_text("\n").split("\n") if clean(x)]
                    cell_lines.append(clean(c.get_text(" ")))
                titulo = title_from_lines(cell_lines, ementa) or ""
            if len(ementa) < 8 and not titulo:
                continue
            out.append(Item(titulo, ementa, link))
    return out


# ---------------------------------------------------------------- âncora + contexto
INLINE = ("b", "strong", "em", "i", "mark", "u", "font", "abbr", "sup", "sub")
HILITE = re.compile(r"high|mark|destaq|\bhl\b|found|match|keyword", re.I)


def _lines(el: Tag) -> list[str]:
    """Linhas de texto do contêiner; tags inline (ex.: <b>termo</b>) não quebram linha."""
    import copy

    c = copy.copy(el)
    for t in list(c.find_all(True)):
        if t.name in INLINE or (t.name == "span" and HILITE.search(" ".join(t.get("class", [])))):
            t.unwrap()
    for br in c.find_all("br"):
        br.replace_with("\n")
    c.smooth()
    return [clean(x) for x in c.get_text("\n").split("\n") if clean(x)]


NORM_HREF = re.compile(r"/norma/\d+", re.I)   # SAPL (Interlegis): /norma/3167?display
JUNK_TITLE = re.compile(
    r"^(pesquisar|pesquisa|nova pesquisa|limpar|buscar|busca|voltar|imprimir|in[ií]cio|home|texto( integral)?|"
    r"ver|visualizar|download|anterior|pr[oó]xim[ao]|primeir[ao]|[uú]ltim[ao]|\d{1,3})$",
    re.I,
)
SEARCH_PATH = re.compile(r"/(pesquisar?|pesquisa[-_a-z]*|busca(r)?|search|consulta[-_a-z]*)(\.aspx|\.php|\.html)?/?$", re.I)
SEARCHY_KEYS = re.compile(r"(^|[?&])(page|pagina|pg|ementa|q|termo|search|busca|salvar|o)=", re.I)


# Proposições em tramitação não são normas: "Projeto de Lei 150/2023", PEC etc.
PROPOSICAO_TITLE = re.compile(
    r"^\W*(projeto|proposta|proposi[cç][aã]o|substitutivo|requerimento|indica[cç][aã]o|mensagem)\b", re.I
)
PROPOSICAO_LINK = re.compile(r"/(processo|proposicao|proposicoes|tramitacao|projeto)[-_a-z]*(\.aspx|\.php)?(/|\?|$)", re.I)


def is_proposicao(it: "Item") -> bool:
    return bool(PROPOSICAO_TITLE.match(it.titulo or "") or PROPOSICAO_LINK.search(urlparse(it.link).path + "?"))


def is_search_link(url: str, base: str) -> bool:
    """Link que leva à própria busca/paginação (não a uma norma)."""
    u, b = urlparse(url), urlparse(base)
    if SEARCH_PATH.search(u.path):
        return True
    return u.netloc == b.netloc and u.path.rstrip("/") == b.path.rstrip("/") and bool(SEARCHY_KEYS.search("?" + u.query))


# como NOISE_LINE_RE, mas "do" só conta como Diário Oficial se vier "DO." ou "DO 123" (não "do Espectro")
JOIN_NOISE = re.compile(
    NOISE_LINE_RE.pattern.replace(r"d\.?o\.?", r"d\.?o(?=\.|\s+\d)").replace("publicad", r"publicad\w*"), re.I
)
TERMINAL = re.compile(r"(\.\.\.|[.;!?…])\s*$")
LABEL_RE = re.compile(r"^[^\W\d_][\w\s/]{0,24}:\s")


ABREV_END = re.compile(r"(?:^|\s)(?:art|arts|inc|incs|n|nº|no|dr|dra|sr|av|cap|p|pág|§)\.\s*$", re.I)


def ends_sentence(ln: str) -> bool:
    return bool(TERMINAL.search(ln)) and not ABREV_END.search(ln)


def _stream_fix(soup: BeautifulSoup, items: list["Item"]) -> None:
    """Minas (ALMG): lê o texto da página em ordem e atribui a cada norma as linhas que vêm depois do seu título,
    até o próximo título, um campo (Fonte, Publicação...) ou o fim da frase. Independe da estrutura dos blocos."""
    lines = _lines(soup.body or soup)
    titles = [it.titulo for it in items if it.titulo]

    def is_title(ln: str) -> bool:
        return any(t in ln for t in titles)

    pos = 0
    for it in items:
        if not it.titulo:
            continue
        idx = next((i for i in range(pos, len(lines)) if it.titulo in lines[i]), None)
        if idx is None:
            continue
        pos = idx + 1
        parts: list[str] = []
        j, skipped = idx + 1, 0
        while j < len(lines) and sum(map(len, parts)) < 1500:
            ln = lines[j]
            meta = ln.startswith("/") or bool(JOIN_NOISE.match(ln)) or bool(DATE_ONLY.match(ln)) or bool(LABEL_RE.match(ln))
            if is_title(ln):
                break
            if meta and not parts and skipped < 4:   # metadados logo após o título (data, "Publicado em...")
                skipped += 1
                j += 1
                continue
            if meta:
                break
            parts.append(ln)
            j += 1
            if ends_sentence(ln):
                break
        txt = clean(" ".join(parts))
        if txt and (len(txt) > len(it.ementa or "") or JOIN_NOISE.match(it.ementa or "")):
            it.ementa = txt


def _join_continuations(raw: list[str], ementa: str, title: str) -> str:
    """Reúne os pedaços de uma ementa que o HTML partiu em várias linhas.
    Começa na primeira linha abaixo do título (pulando datas) e segue até o ponto final ou um campo."""

    def stop(ln: str) -> bool:
        return (ln == title or ln.startswith("/") or bool(JOIN_NOISE.match(ln))
                or bool(DATE_ONLY.match(ln)) or bool(LABEL_RE.match(ln)))

    start = None
    for t_idx, ln in enumerate(raw):
        if ln == title or (title and title in ln):
            j = t_idx + 1
            while j < len(raw) and DATE_ONLY.match(raw[j]):
                j += 1
            if j < len(raw) and not stop(raw[j]) and len(raw[j]) >= 3:
                start = j
            break
    if start is None:
        try:
            start = raw.index(ementa)
        except ValueError:
            return ementa
    parts, j = [raw[start]], start + 1
    while j < len(raw) and not ends_sentence(parts[-1]) and sum(map(len, parts)) < 1500 and not stop(raw[j]):
        parts.append(raw[j]); j += 1
    k = start - 1
    while k >= 0 and parts[0][:1].islower() and not stop(raw[k]) and sum(map(len, parts)) < 1500:
        parts.insert(0, raw[k]); k -= 1
    joined = clean(" ".join(parts))
    return joined if len(joined) >= len(ementa) else ementa


def meta_descriptions(html: str) -> list[str]:
    soup = BeautifulSoup(html, "lxml")
    out = []
    for m in soup.find_all("meta"):
        key = (m.get("name") or m.get("property") or "").lower()
        if key in ("description", "og:description", "twitter:description") and m.get("content"):
            out.append(clean(m["content"]))
    return out


def full_ementa_from_html(html: str, partial: str) -> str:
    """Na página da norma, devolve a ementa completa: meta description (se começar como a parcial)
    ou o menor bloco de texto que começa com a parcial e é mais longo."""
    norm = lambda t: re.sub(r"\s+", " ", t or "").strip().lower()
    prefix = norm(partial)[:25]
    if len(prefix) < 12:
        return ""
    for d in meta_descriptions(html):
        if len(d) >= len(partial) - 3 and norm(d).startswith(prefix):
            return d
    soup = BeautifulSoup(html, "lxml")
    for t in soup(["script", "style", "noscript"]):
        t.decompose()
    for br in soup.find_all("br"):
        br.replace_with(" ")
    best = ""
    for el in soup.find_all(["p", "div", "span", "td", "li", "dd", "h1", "h2", "h3", "h4", "section", "article"]):
        txt = clean(el.get_text())
        if len(txt) > len(partial) + 5 and len(txt) < 2000 and norm(txt).startswith(prefix):
            if not best or len(txt) < len(best):
                best = txt
    return best


def _anchor_items(soup: BeautifulSoup, base: str, href_re: re.Pattern | None = None) -> list[Item]:
    best_by_link: dict[str, tuple[Tag, str]] = {}
    for a in soup.find_all("a", href=True):
        text = clean(a.get_text(" "))
        url = valid_link(a["href"], base)
        if not url or len(text) > 160:
            continue
        if href_re is not None:
            ok = bool(href_re.search(urlparse(url).path)) and not is_search_link(url, base)
        else:
            ok = 4 <= len(text) and bool(TITLE_RE.search(text))
        if not ok:
            continue
        if href_re is not None:
            m = href_re.search(urlparse(url).path)
            key = urlparse(url).netloc + m.group(0)           # uma norma = um id
            main = bool(re.fullmatch(href_re.pattern, urlparse(url).path.rstrip("/"), re.I))
        else:
            key, main = url, True
        score = (main, bool(TITLE_RE.search(text)), not JUNK_TITLE.match(text), len(text))
        cur = best_by_link.get(key)
        if cur is None or score > cur[2]:
            best_by_link[key] = (a, text, score)
    anchors = [(a, t) for a, t, _ in best_by_link.values()]
    if not anchors:
        return []
    title_ids = {id(a) for a, _ in anchors}
    out: list[Item] = []
    for a, text in anchors:
        container: Tag = a
        best: Tag | None = None
        for _ in range(6):
            parent = container.parent
            if not isinstance(parent, Tag) or parent.name in ("body", "html", "[document]"):
                break
            n_titles = sum(1 for x in parent.find_all("a", href=True) if id(x) in title_ids)
            if n_titles > 1:
                break
            container = parent
            if len(clean(container.get_text(" "))) > len(text) + 15:
                best = container
                break
        ementa = ""
        if best is not None:
            cands = []
            for ln in _lines(best):
                if ln == text or text in ln and len(ln) - len(text) < 6:
                    continue
                if ln.startswith("/") or NOISE_LINE_RE.match(ln) or len(ln) < 15:
                    continue
                cands.append(ln)
            if cands:
                ementa = max(cands, key=len)
            if ementa and "almg.gov.br" in base:   # Minas Gerais: ementa vem partida em vários pedaços
                ementa = _join_continuations(_lines(best), ementa, text)
        link = valid_link(a["href"], base)
        title = text
        if weak_title(title):
            title = title_from_lines(_lines(best), ementa) if best is not None else ""
        out.append(Item(title, ementa, link))  # type: ignore[arg-type]
    return out


def extract(html: str, base_url: str) -> list[Item]:
    soup = _strip_page(BeautifulSoup(html, "lxml"))
    items = _anchor_items(soup, base_url, NORM_HREF)  # SAPL: sinal forte, vale mais que heurísticas
    if not items:
        a = _table_items(soup, base_url)
        b = _anchor_items(soup, base_url)
        items = a if len(a) >= len(b) else b
    seen, uniq = set(), []
    for it in items:
        if it.link in seen or is_search_link(it.link, base_url) or is_proposicao(it):
            continue
        if JUNK_TITLE.match(it.titulo or "") and not it.ementa:
            continue
        seen.add(it.link)
        uniq.append(it)
    if uniq and "almg.gov.br" in base_url:
        _stream_fix(soup, uniq)
    return uniq


# ---------------------------------------------------------------- paginação
def find_next_url(html: str, base_url: str) -> str | None:
    soup = BeautifulSoup(html, "lxml")
    a = soup.find("a", rel=lambda v: v and "next" in (v if isinstance(v, list) else [v]), href=True)
    if a:
        return valid_link(a["href"], base_url)
    for a in soup.find_all("a", href=True):
        txt = clean(a.get_text(" "))
        label = clean(a.get("aria-label", "")) + " " + clean(a.get("title", ""))
        if NEXT_TEXT_RE.match(txt) or re.search(r"pr[oó]xim|next", label, re.I):
            parent_cls = " ".join(a.parent.get("class", [])) if a.parent else ""
            if "disabled" in parent_cls or "disabled" in " ".join(a.get("class", [])):
                continue
            u = valid_link(a["href"], base_url)
            if u:
                return u
    return None


def has_pager(html: str) -> bool:
    """A página tem controles de paginação (mesmo sem 'próxima' = última página)?"""
    return bool(re.search(r'class="[^"]*(pagination|paginacao|paginador|pager)[^"]*"', html, re.I))


def aspnet_next(html: str, current_page: int) -> dict | None:
    """Se a página usa paginação por postback ASP.NET, devolve {'form': Tag, 'target', 'arg'} da página seguinte."""
    soup = BeautifulSoup(html, "lxml")
    pat = re.compile(r"__doPostBack\('([^']+)','([^']*)'\)")
    want = str(current_page + 1)
    fallback = None
    for a in soup.find_all("a", href=True):
        m = pat.search(a["href"])
        if not m:
            continue
        txt = clean(a.get_text(" "))
        target, arg = m.groups()
        if re.search(rf"(Page\${want}|^{want}$)", arg) or txt == want:
            return {"target": target, "arg": arg, "soup": soup}
        if NEXT_TEXT_RE.match(txt) or re.search(r"Page\$Next", arg):
            fallback = {"target": target, "arg": arg, "soup": soup}
    return fallback


def aspnet_form_data(soup: BeautifulSoup, target: str, arg: str) -> tuple[str, dict]:
    form = soup.find("form")
    data: dict[str, str] = {}
    if form:
        for inp in form.find_all(["input", "select", "textarea"]):
            name = inp.get("name")
            if not name:
                continue
            typ = (inp.get("type") or "").lower()
            if typ in ("submit", "button", "image", "checkbox", "radio") and not inp.has_attr("checked"):
                continue
            if inp.name == "select":
                opt = inp.find("option", selected=True) or inp.find("option")
                data[name] = opt.get("value", "") if opt else ""
            else:
                data[name] = inp.get("value", "") or (inp.text if inp.name == "textarea" else "")
    data["__EVENTTARGET"] = target
    data["__EVENTARGUMENT"] = arg
    action = form.get("action", "") if form else ""
    return action, data

# ======================================================================
# NAVEGADOR HEADLESS
# ======================================================================
"""Coleta via navegador headless (Playwright) para sites dinâmicos (SPA/JavaScript).

Dois modos:
  - 'direct': abre a URL de busca já com o termo e lê o resultado renderizado;
  - 'form'  : abre a página de busca, localiza o campo de texto, digita o termo,
              envia o formulário e lê o resultado.
Em ambos, intercepta respostas JSON (APIs internas dos SPAs) como fonte extra de
ementas, e pagina clicando em "próxima".
"""

import hashlib
import re
from typing import Callable


try:  # Playwright é opcional
    from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout
    PLAYWRIGHT_OK = True
except Exception:  # pragma: no cover
    PLAYWRIGHT_OK = False
    PWTimeout = Exception  # type: ignore

URL_KEYS = ("url", "link", "href", "urldocumento", "linktexto", "urltexto", "linkdocumento", "caminho", "arquivo")
FIELD_HINT = re.compile(r"ementa|busca|pesq|search|termo|palavra|texto|assunto|^q$|query|filtro", re.I)
NEXT_SELECTORS = [
    "a[rel=next]",
    "li.next:not(.disabled) a",
    ".pagination .next:not(.disabled) a",
    "[aria-label*='róxim' i]:not([disabled])",
    "[aria-label*='next' i]:not([disabled])",
    "button:has-text('Próxima'):not([disabled])",
    "a:has-text('Próxima')",
    "a:has-text('Próximo')",
    "button:has-text('Próximo'):not([disabled])",
    "a:text-is('›')",
    "a:text-is('»')",
    "button:text-is('›')",
    "button:text-is('»')",
    "a:text-is('>')",
]


def _walk_json(obj, base: str, out: list[Item]):
    if isinstance(obj, list):
        if obj and all(isinstance(x, dict) for x in obj):
            for d in obj:
                low = {str(k).lower(): v for k, v in d.items()}
                ementa = next((clean(str(v)) for k, v in low.items() if "ementa" in k and isinstance(v, str)), "")
                link = next((v for k, v in low.items() if k in URL_KEYS and isinstance(v, str) and v.strip()), "")
                if ementa and link:
                    from urllib.parse import urljoin

                    titulo = next(
                        (clean(str(v)) for k, v in low.items() if k in ("titulo", "título", "nome", "descricao", "norma") and isinstance(v, str)),
                        "",
                    )
                    out.append(Item(titulo, ementa, urljoin(base, link)))
        for x in obj:
            _walk_json(x, base, out)
    elif isinstance(obj, dict):
        for v in obj.values():
            _walk_json(v, base, out)


def _system_chromium() -> str | None:
    import os, shutil
    for c in (os.environ.get("CHROMIUM_PATH"), shutil.which("chromium"), shutil.which("chromium-browser"), "/usr/bin/chromium"):
        if c and os.path.exists(c):
            return c
    return None


class Browser:
    def __init__(self, verify_ssl: bool = True, timeout_ms: int = 25000):
        self.verify_ssl = verify_ssl
        self.timeout_ms = timeout_ms
        self._pw = None
        self._browser = None

    def __enter__(self):
        if not PLAYWRIGHT_OK:
            raise RuntimeError("Playwright não está instalado (pip install playwright && playwright install chromium)")
        self._pw = sync_playwright().start()
        self._browser = self._pw.chromium.launch(
            headless=True, args=["--no-sandbox", "--disable-dev-shm-usage"], executable_path=_system_chromium()
        )
        return self

    def __exit__(self, *exc):
        try:
            if self._browser:
                self._browser.close()
        finally:
            if self._pw:
                self._pw.stop()

    # ------------------------------------------------------------ helpers
    def _settle(self, page):
        try:
            page.wait_for_load_state("networkidle", timeout=self.timeout_ms)
        except PWTimeout:
            pass
        page.wait_for_timeout(600)

    def _fill_form(self, page, term: str) -> bool:
        cands = page.locator(
            "input:visible:not([type=hidden]):not([type=checkbox]):not([type=radio]):not([type=submit]):not([type=button]):not([type=date])"
        )
        n = cands.count()
        best, best_score = None, -1
        for i in range(min(n, 30)):
            el = cands.nth(i)
            attrs = " ".join(
                (el.get_attribute(a) or "") for a in ("name", "id", "placeholder", "aria-label", "formcontrolname", "title")
            )
            score = 0
            if re.search("ementa", attrs, re.I):
                score += 6
            if FIELD_HINT.search(attrs):
                score += 3
            if (el.get_attribute("type") or "text") in ("search", "text"):
                score += 1
            if re.search(r"n[uú]mero|ano|data|autor|codigo|c[oó]digo", attrs, re.I) and not re.search("ementa", attrs, re.I):
                score -= 4
            if score > best_score:
                best, best_score = el, score
        if best is None:
            return False
        best.fill(term)
        before = page.url
        best.press("Enter")
        self._settle(page)
        return True

    def _click_search(self, page):
        for loc in (
            page.get_by_role("button", name=re.compile(r"pesquis|busc|consult|filtrar|procurar", re.I)),
            page.locator("input[type=submit]"),
            page.locator("button[type=submit]"),
        ):
            try:
                if loc.count():
                    loc.first.click(timeout=3000)
                    self._settle(page)
                    return True
            except Exception:
                continue
        return False

    def _click_next(self, page) -> bool:
        for sel in NEXT_SELECTORS:
            try:
                loc = page.locator(sel)
                if loc.count() and loc.first.is_visible():
                    loc.first.click(timeout=4000)
                    self._settle(page)
                    return True
            except Exception:
                continue
        return False

    # ------------------------------------------------------------ coleta
    def collect(
        self,
        url: str,
        term: str,
        mode: str = "direct",
        max_results: int = 50,
        max_pages: int = 10,
        log: Callable[[str], None] | None = None,
    ) -> tuple[list[Item], dict]:
        log = log or (lambda m: None)
        ctx = self._browser.new_context(ignore_https_errors=not self.verify_ssl, locale="pt-BR", user_agent=UA)
        page = ctx.new_page()
        page.set_default_timeout(self.timeout_ms)
        payloads: list = []

        def on_resp(resp):
            try:
                if "json" in (resp.headers.get("content-type") or "") and resp.status == 200:
                    payloads.append((resp.url, resp.json()))
            except Exception:
                pass

        page.on("response", on_resp)
        info = {"paginas": 0, "html_inicial": ""}
        items: list[Item] = []
        seen: set[str] = set()
        try:
            page.goto(url, wait_until="domcontentloaded", timeout=self.timeout_ms * 2)
            self._settle(page)
            if mode == "form":
                ok = self._fill_form(page, term)
                if not ok:
                    raise RuntimeError("campo de busca não encontrado na página")
                # se nada mudou ainda, tenta o botão
                if not extract(page.content(), page.url) and not payloads:
                    self._click_search(page)
            last_sig = ""
            for n in range(1, max_pages + 1):
                html = page.content()
                if n == 1:
                    info["html_inicial"] = html
                new = []
                for it in extract(html, page.url):
                    if it.link not in seen:
                        seen.add(it.link)
                        new.append(it)
                jitems: list[Item] = []
                for purl, data in payloads:
                    _walk_json(data, purl, jitems)
                payloads.clear()
                for it in jitems:
                    if it.link not in seen:
                        seen.add(it.link)
                        new.append(it)
                info["paginas"] = n
                items.extend(new)
                log(f"página {n}: +{len(new)}")
                if not new or len(items) >= max_results:
                    break
                sig = hashlib.md5(html.encode("utf-8", "ignore")).hexdigest()
                if sig == last_sig or not self._click_next(page):
                    break
                last_sig = sig
        finally:
            ctx.close()
        return items[:max_results], info

# ======================================================================
# MOTOR DE BUSCA
# ======================================================================
"""Orquestração da busca por estado.

Grupo 1 (estaduais.xlsx): o link já é a busca por "escola" -> troca-se o termo e
extraem-se os resultados (HTTP; se vier vazio, navegador headless).

Grupo 2 (estaduais2.xlsx): só há a página geral de legislação. Estratégia em camadas:
  1. SAPL: se a casa usa o SAPL (Interlegis), monta a URL de busca por ementa
     (host conhecido ou descoberto por sondagem em sapl.al.<uf>.leg.br);
  2. Navegador headless: abre a página, localiza o campo de busca, digita o termo,
     envia e lê os resultados (HTML renderizado + JSON das APIs internas);
  3. Se nada funcionar, o estado entra no relatório como "busca manual" com o link geral.
"""

import re
import time
from dataclasses import dataclass, field
from pathlib import Path
from typing import Callable
from urllib.parse import urljoin, quote_plus

import pandas as pd
import requests



UF = {
    "Alagoas": "al", "Distrito Federal": "df", "Maranhão": "ma", "Mato Grosso do Sul": "ms",
    "Paraíba": "pb", "Paraná": "pr", "Pernambuco": "pe", "Rio de Janeiro": "rj",
    "Rio Grande do Sul": "rs", "Tocantins": "to",
}
# Casas que sabidamente usam o SAPL (além do que a sondagem descobrir)
SAPL_KNOWN = {"Paraíba": "sapl.al.pb.leg.br"}
SAPL_PATH = (
    "/norma/pesquisar?tipo=&numero=&ano=&data_0=&data_1=&data_publicacao_0=&data_publicacao_1="
    "&ementa={q}&assuntos=&data_vigencia_0=&data_vigencia_1=&orgao=&o=&indexacao="
    "&autorianorma__autor=&autorianorma__primeiro_autor=unknown&autorianorma__autor__tipo="
    "&autorianorma__autor__parlamentar_set__filiacao__partido=&salvar=Pesquisar"
)
# parâmetro de página quando a URL não traz link "próxima"
PAGE_HINT = {"Goiás": "pagina"}


@dataclass
class Options:
    max_results: int = 30
    max_pages: int = 10
    timeout: int = 30
    verify_ssl: bool = False   # sites .gov costumam ter cadeia de certificados incompleta
    use_browser: bool = True


@dataclass
class StateResult:
    estado: str
    grupo: int
    items: list[Item] = field(default_factory=list)
    metodo: str = ""
    status: str = "erro"        # ok | vazio | erro
    mensagem: str = ""
    paginas: int = 0
    busca_url: str = ""         # URL usada na busca (ou página geral, no grupo 2)
    link_geral: str = ""
    debug_html: str = ""
    segundos: float = 0.0


# ------------------------------------------------------------------ planilhas
SOURCES = [
    {'estado': 'Alagoas', 'link': 'https://www.al.al.leg.br/leis/legislacao-estadual', 'grupo': 2},
    {'estado': 'Amazonas', 'link': 'https://sapl.al.am.leg.br/norma/pesquisar?tipo=&numero=&ano=&data_0=&data_1=&data_publicacao_0=&data_publicacao_1=&ementa=escola&assuntos=&data_vigencia_0=&data_vigencia_1=&orgao=&o=&indexacao=&autorianorma__autor=&autorianorma__primeiro_autor=unknown&autorianorma__autor__tipo=&autorianorma__autor__parlamentar_set__filiacao__partido=&salvar=Pesquisar', 'grupo': 1},
    {'estado': 'Distrito Federal', 'link': 'https://ple.cl.df.gov.br/#/norma/buscar', 'grupo': 2},
    {'estado': 'Espírito Santo', 'link': 'https://www3.al.es.gov.br/legislacao/consulta-legislacao.aspx?termo=escola', 'grupo': 1},
    {'estado': 'Goiás', 'link': 'https://legisla.casacivil.go.gov.br/pesquisa_legislacao?ementa=escola&periodo_inicial_legislacao&periodo_final_legislacao&periodo_inicial_diario&periodo_final_diario&pagina=1&qtd_por_pagina=10&ordenarPor=data&busca_avancada=1&semantico=false', 'grupo': 1},
    {'estado': 'Maranhão', 'link': 'https://legislacao.al.ma.leg.br/ged/busca.html?dswid=-7817', 'grupo': 2},
    {'estado': 'Mato Grosso', 'link': 'https://www.al.mt.gov.br/norma-juridica?almt_form_norma_juridica_ato_busca_avancada%5BatoTipo%5D%5Bautocomplete%5D=&almt_form_norma_juridica_ato_busca_avancada%5BconteudoDispositivo%5D=&almt_form_norma_juridica_ato_busca_avancada%5Bementa%5D=escola&almt_form_norma_juridica_ato_busca_avancada%5Bnumero%5D=&almt_form_norma_juridica_ato_busca_avancada%5Bano%5D=&almt_form_norma_juridica_ato_busca_avancada%5Bautor%5D%5Bautocomplete%5D=&almt_form_norma_juridica_ato_busca_avancada%5Bapelido%5D=&almt_form_norma_juridica_ato_busca_avancada%5BtagCondicao%5D=e&almt_form_norma_juridica_ato_busca_avancada%5BdataPublicacaoDe%5D=&almt_form_norma_juridica_ato_busca_avancada%5BdataPublicacaoAte%5D=&almt_form_norma_juridica_ato_busca_avancada%5BdataPromulgacaoDe%5D=&almt_form_norma_juridica_ato_busca_avancada%5BdataPromulgacaoAte%5D=&almt_form_norma_juridica_ato_busca_avancada%5BdataInicioVigenciaDe%5D=&almt_form_norma_juridica_ato_busca_avancada%5BdataInicioVigenciaAte%5D=&almt_form_norma_juridica_ato_busca_avancada%5BdataFimVigenciaDe%5D=&almt_form_norma_juridica_ato_busca_avancada%5BdataFimVigenciaAte%5D=&almt_form_norma_juridica_ato_busca_avancada%5Brevogado%5D=&almt_form_norma_juridica_ato_busca_avancada%5BpossuiVeto%5D=&almt_form_norma_juridica_ato_busca_avancada%5BpossuiRemissao%5D=', 'grupo': 1},
    {'estado': 'Mato Grosso do Sul', 'link': 'https://aacpdappls.net.ms.gov.br/appls/legislacao/secoge/govato.nsf/$$Search?OpenForm', 'grupo': 2},
    {'estado': 'Minas Gerais', 'link': 'https://www.almg.gov.br/atividade-parlamentar/leis/legislacao-mineira/?pesquisou=true&aba=pesquisa&q=escola&grupo=&num=&ano=&dataInicio=&dataFim=&sit=&ordem=2', 'grupo': 1},
    {'estado': 'Paraná', 'link': 'https://consultas.assembleia.pr.leg.br/#/norma-legal', 'grupo': 2},
    {'estado': 'Paraíba', 'link': 'https://www.al.pb.leg.br/leis-estaduais', 'grupo': 2},
    {'estado': 'Pará', 'link': 'https://bancodeleis.alepa.pa.gov.br/index.php?codigo=&ano=&tipo=&ementa=escola&ordem=ano_desc', 'grupo': 1},
    {'estado': 'Pernambuco', 'link': 'https://legis.alepe.pe.gov.br/', 'grupo': 2},
    {'estado': 'Piauí', 'link': 'https://sapl.al.pi.leg.br/norma/pesquisar?tipo=&numero=&ano=&data_0=&data_1=&data_publicacao_0=&data_publicacao_1=&ementa=escola&assuntos=&data_vigencia_0=&data_vigencia_1=&orgao=&o=&indexacao=&autorianorma__autor=&autorianorma__primeiro_autor=unknown&autorianorma__autor__tipo=&autorianorma__autor__parlamentar_set__filiacao__partido=&salvar=Pesquisar', 'grupo': 1},
    {'estado': 'Rio Grande do Norte', 'link': 'https://iframe.leisestaduais.com.br/rn?q=escola&origin=instituicao&date_start=&date_end=', 'grupo': 1},
    {'estado': 'Rio Grande do Sul', 'link': 'https://ww4.al.rs.gov.br/legislacao-estadual', 'grupo': 2},
    {'estado': 'Rio de Janeiro', 'link': 'https://www3.alerj.rj.gov.br/lotus_notes/default.asp?id=53', 'grupo': 2},
    {'estado': 'Roraima', 'link': 'https://sapl.al.rr.leg.br/norma/pesquisar?tipo=&numero=&ano=&data_0=&data_1=&data_publicacao_0=&data_publicacao_1=&ementa=escola&assuntos=&data_vigencia_0=&data_vigencia_1=&orgao=&o=&indexacao=&autorianorma__autor=&autorianorma__primeiro_autor=unknown&autorianorma__autor__tipo=&autorianorma__autor__parlamentar_set__filiacao__partido=&salvar=Pesquisar', 'grupo': 1},
    {'estado': 'Santa Catarina', 'link': 'https://leis.alesc.sc.gov.br/legislativo?search=escola', 'grupo': 1},
    {'estado': 'Sergipe', 'link': 'https://aleselegis.al.se.leg.br/consulta-legislacao.aspx?termo=escola', 'grupo': 1},
    {'estado': 'São Paulo', 'link': 'https://www.al.sp.gov.br/legislacao/pesquisa?busca=escola&tamanho=50', 'grupo': 1},
    {'estado': 'Tocantins', 'link': 'https://www.al.to.leg.br/legislacaoEstadual', 'grupo': 2},
]


def load_sources() -> list[dict]:
    return [dict(s) for s in SOURCES]


# ------------------------------------------------------------------ coleta HTTP
def collect_http(
    session: requests.Session, url: str, opt: Options, page_param: str | None = None
) -> tuple[list[Item], dict]:
    items: list[Item] = []
    seen: set[str] = set()
    visited: set[str] = {url}
    cur, post, pages = url, None, 0
    info = {"paginas": 0, "html_inicial": ""}
    while pages < opt.max_pages and len(items) < opt.max_results:
        try:
            if post:
                r = net.get(session, post["url"], opt.timeout, opt.verify_ssl, method="POST", data=post["data"])
            else:
                r = net.get(session, cur, opt.timeout, opt.verify_ssl)
        except requests.HTTPError:
            if pages == 0:
                raise
            break  # página além do fim
        html = r.text
        pages += 1
        if pages == 1:
            info["html_inicial"] = html
        new = [it for it in extract(html, r.url) if it.link not in seen]
        for it in new:
            seen.add(it.link)
        info["paginas"] = pages
        if not new:
            break
        items.extend(new)

        # próxima página: link > postback ASP.NET > incremento de parâmetro
        post = None
        nxt = find_next_url(html, r.url)
        if nxt and nxt not in visited:
            cur = nxt
        else:
            asp = aspnet_next(html, pages)
            if asp:
                action, data = aspnet_form_data(asp["soup"], asp["target"], asp["arg"])
                post = {"url": urljoin(r.url, action) if action else r.url, "data": data}
                continue
            if has_pager(html):
                break  # paginação visível sem "próxima" = última página
            param = page_param or next((p for p in PAGE_PARAMS if net.has_query_param(r.url, p)), None)
            if not param and "norma/pesquisar" in r.url:
                param = "page"
            if not param:
                break
            cur = net.set_query_param(r.url, param, pages + 1)
            if cur in visited:
                break
        visited.add(cur)
    return items[: opt.max_results], info


# ------------------------------------------------------------------ SAPL
def sapl_url(host: str, term: str) -> str:
    return f"https://{host}" + SAPL_PATH.format(q=quote_plus(term))


def probe_sapl(session: requests.Session, estado: str, opt: Options) -> str | None:
    hosts = [SAPL_KNOWN[estado]] if estado in SAPL_KNOWN else []
    uf = UF.get(estado)
    if uf and uf != "df":
        hosts.append(f"sapl.al.{uf}.leg.br")
    for host in dict.fromkeys(hosts):
        try:
            r = net.get(session, f"https://{host}/norma/pesquisar", min(opt.timeout, 15), opt.verify_ssl)
            if re.search(r'name=["\']ementa["\']', r.text, re.I):
                return host
        except Exception:
            continue
    return None


# ------------------------------------------------------------------ por estado
def complete_ementas(session: requests.Session, items: list[Item], opt: Options) -> dict:
    """Confere cada norma na sua própria página (meta description / texto) e completa a ementa.
    Devolve estatísticas para o diagnóstico: conferidas, completadas e falhas."""
    from concurrent.futures import ThreadPoolExecutor

    stats = {"total": len(items), "conferidas": 0, "completadas": 0, "falhas": 0, "erro": ""}

    def fix(it: Item) -> None:
        if not it.ementa:
            return
        try:
            r = net.get(session, it.link, min(opt.timeout, 20), opt.verify_ssl)
        except Exception as e:  # noqa: BLE001
            stats["falhas"] += 1
            stats["erro"] = stats["erro"] or f"{type(e).__name__}: {str(e)[:80]}"
            return
        stats["conferidas"] += 1
        full = full_ementa_from_html(r.text, it.ementa)
        if full and full != it.ementa:
            it.ementa = full
            stats["completadas"] += 1

    with ThreadPoolExecutor(max_workers=6) as ex:
        list(ex.map(fix, items))
    return stats


def _finish(res: StateResult, items: list[Item], metodo: str, info: dict, t0: float, opt: Options | None = None) -> StateResult:
    if "almg.gov.br" in (res.busca_url or ""):   # Minas Gerais: garante a ementa completa
        st = complete_ementas(net.make_session(), items, opt or Options())
        res.mensagem = (f"[busca_leg v4-almg] ALMG: ementa conferida na página da norma em {st['conferidas']}/{st['total']}, "
                        f"completadas {st['completadas']}, falhas {st['falhas']}" + (f" ({st['erro']})" if st["erro"] else ""))
    res.items, res.metodo = items, metodo
    res.paginas = info.get("paginas", 0)
    res.debug_html = info.get("html_inicial", "")[:400_000]
    res.status = "ok" if items else "vazio"
    res.segundos += time.time() - t0
    return res


def search_state(
    src: dict,
    term: str,
    opt: Options,
    session: requests.Session | None = None,
    browser: Browser | None = None,
    log: Callable[[str], None] | None = None,
    phase: str = "all",
    prev: StateResult | None = None,
) -> StateResult:
    """Busca um estado.

    phase="http"    -> só camadas HTTP (rápidas, podem rodar em paralelo);
    phase="browser" -> só camadas com navegador (continua de `prev`);
    phase="all"     -> as duas, em sequência.
    """
    log = log or (lambda m: None)
    session = session or net.make_session()
    estado, grupo = src["estado"], src["grupo"]
    t0 = time.time()
    res = prev or StateResult(estado=estado, grupo=grupo, link_geral=src["link"])
    errors: list[str] = [res.mensagem] if (prev and res.mensagem) else []
    had_exc = bool(prev and res.status == "erro")
    do_http = phase in ("all", "http")
    do_browser = phase in ("all", "browser") and opt.use_browser

    def done(msg_status: str | None = None) -> StateResult:
        res.mensagem = " | ".join(e for e in errors if e)
        res.status = msg_status or res.status
        res.segundos += time.time() - t0
        return res

    if grupo == 1:
        url = net.substitute_term(src["link"], term)
        res.busca_url = url
        if do_http:
            try:
                items, info = collect_http(session, url, opt, PAGE_HINT.get(estado))
                if items:
                    return _finish(res, items, "HTTP direto", info, t0, opt)
                errors.append("HTTP: nenhuma norma reconhecida no HTML")
                res.debug_html = info.get("html_inicial", "")[:400_000]
            except Exception as e:  # noqa: BLE001
                had_exc = True
                errors.append(f"HTTP: {type(e).__name__}: {str(e)[:150]}")
        if do_browser:
            if browser is None:
                errors.append("navegador indisponível")
            else:
                try:
                    items, info = browser.collect(url, term, "direct", opt.max_results, opt.max_pages, log)
                    if items:
                        return _finish(res, items, "Navegador (página renderizada)", info, t0, opt)
                    errors.append("Navegador: nenhuma norma reconhecida")
                    res.debug_html = info.get("html_inicial", "") or res.debug_html
                except Exception as e:  # noqa: BLE001
                    had_exc = True
                    errors.append(f"Navegador: {type(e).__name__}: {str(e)[:150]}")
        return done("erro" if had_exc and not res.debug_html else "vazio")

    # ---- grupo 2: estratégia em camadas
    res.busca_url = res.busca_url or src["link"]
    if do_http:
        host = probe_sapl(session, estado, opt)
        if host:
            url = sapl_url(host, term)
            res.busca_url = url
            try:
                items, info = collect_http(session, url, opt, "page")
                if items:
                    return _finish(res, items, f"SAPL ({host})", info, t0, opt)
                errors.append(f"SAPL {host}: sem resultados reconhecidos")
            except Exception as e:  # noqa: BLE001
                errors.append(f"SAPL: {type(e).__name__}: {str(e)[:150]}")
        else:
            errors.append("SAPL não detectado")
    if do_browser:
        if browser is None:
            errors.append("navegador indisponível")
        else:
            try:
                items, info = browser.collect(src["link"], term, "form", opt.max_results, opt.max_pages, log)
                if items:
                    return _finish(res, items, "Navegador (formulário preenchido)", info, t0, opt)
                errors.append("Navegador: formulário enviado, nenhuma norma com link direto reconhecida")
                res.debug_html = info.get("html_inicial", "")[:400_000]
            except Exception as e:  # noqa: BLE001
                errors.append(f"Navegador: {type(e).__name__}: {str(e)[:150]}")
    elif phase != "http":
        errors.append("navegador desativado")
    return done("erro")

# ======================================================================
# WORD
# ======================================================================
"""Geração do relatório Word: Estado | Ementa | Link da norma."""

import io
from datetime import datetime

from docx import Document
from docx.enum.section import WD_ORIENT
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor


AZUL = "1F3864"


def _hyperlink(paragraph, url: str, text: str, size: int = 9, color: str = "0563C1", bold: bool = False):
    part = paragraph.part
    r_id = part.relate_to(url, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink", is_external=True)
    h = OxmlElement("w:hyperlink")
    h.set(qn("r:id"), r_id)
    r = OxmlElement("w:r")
    rpr = OxmlElement("w:rPr")
    c = OxmlElement("w:color"); c.set(qn("w:val"), color); rpr.append(c)
    u = OxmlElement("w:u"); u.set(qn("w:val"), "single"); rpr.append(u)
    sz = OxmlElement("w:sz"); sz.set(qn("w:val"), str(size * 2)); rpr.append(sz)
    if bold:
        rpr.append(OxmlElement("w:b"))
    r.append(rpr)
    t = OxmlElement("w:t"); t.text = text; t.set(qn("xml:space"), "preserve"); r.append(t)
    h.append(r)
    paragraph._p.append(h)


def _shade(cell, hex_fill: str):
    tcpr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear"); shd.set(qn("w:color"), "auto"); shd.set(qn("w:fill"), hex_fill)
    tcpr.append(shd)


def _repeat_header(row):
    trpr = row._tr.get_or_add_trPr()
    el = OxmlElement("w:tblHeader"); el.set(qn("w:val"), "true"); trpr.append(el)


def _no_split(row):
    trpr = row._tr.get_or_add_trPr()
    trpr.append(OxmlElement("w:cantSplit"))


def _set_w(cell, cm: float):
    cell.width = Cm(cm)


def _run(p, text, size=9, bold=False, color=None, italic=False):
    r = p.add_run(text)
    r.font.size = Pt(size); r.bold = bold; r.italic = italic
    if color:
        r.font.color.rgb = RGBColor.from_string(color)
    return r


def _bottom_border(p):
    ppr = p._p.get_or_add_pPr()
    b = OxmlElement("w:pBdr")
    e = OxmlElement("w:bottom")
    for k, v in (("val", "single"), ("sz", "8"), ("space", "1"), ("color", AZUL)):
        e.set(qn(f"w:{k}"), v)
    b.append(e); ppr.append(b)


def build_docx(termo: str, results: list[StateResult]) -> bytes:
    doc = Document()
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Cm(21.0), Cm(29.7)
    for side in ("left_margin", "right_margin"):
        setattr(sec, side, Cm(1.8))
    sec.top_margin = sec.bottom_margin = Cm(1.8)
    st = doc.styles["Normal"]; st.font.name = "Calibri"; st.font.size = Pt(10)

    ok = sorted([r for r in results if r.items], key=lambda r: r.estado)
    sem = sorted([r for r in results if not r.items], key=lambda r: r.estado)
    total = sum(len(r.items) for r in ok)

    h = doc.add_paragraph(); _run(h, "Pesquisa de legislação estadual", 20, True, AZUL)
    p = doc.add_paragraph(); _run(p, "Termo pesquisado: ", 11, True); _run(p, f"“{termo}”", 11)
    p = doc.add_paragraph()
    _run(p, f"Gerado em {datetime.now():%d/%m/%Y às %H:%M} · {total} norma(s) em {len(ok)} estado(s)", 9, color="595959")
    if sem:
        p = doc.add_paragraph()
        _run(p, f"Sem resultados automáticos em {len(sem)} estado(s): " + ", ".join(r.estado for r in sem) + " (ver final do documento).", 9, color="595959")

    for r in ok:
        h = doc.add_paragraph()
        h.paragraph_format.space_before = Pt(16)
        h.paragraph_format.space_after = Pt(4)
        h.paragraph_format.keep_with_next = True
        _run(h, r.estado, 16, True, AZUL)
        _bottom_border(h)
        for it in r.items:
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(8)
            p.paragraph_format.space_after = Pt(0)
            p.paragraph_format.keep_with_next = True
            _run(p, it.titulo or "(nome da norma não identificado)", 11, True)
            p = doc.add_paragraph()
            p.paragraph_format.space_after = Pt(0)
            p.paragraph_format.keep_with_next = True
            _run(p, "Ementa: ", 10, True)
            _run(p, it.ementa or "(ementa não identificada na página)", 10, italic=not it.ementa)
            p = doc.add_paragraph()
            p.paragraph_format.space_after = Pt(2)
            _run(p, "Link: ", 10, True)
            _hyperlink(p, it.link, it.link, 9)

    if sem:
        doc.add_paragraph()
        h = doc.add_paragraph(); _run(h, "Estados sem resultado automático", 13, True, AZUL)
        p = doc.add_paragraph()
        _run(p, "Nestes estados não foi possível extrair normas com link direto. Abaixo, a página de busca para consulta manual (não é link de norma específica).", 9, color="595959")
        for r in sem:
            p = doc.add_paragraph(style="List Bullet")
            _run(p, f"{r.estado}: ", 9, True)
            _hyperlink(p, r.busca_url or r.link_geral, "abrir página de busca", 9)
    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()

# ======================================================================
# VISUAL
# ======================================================================
"""Camada visual do app (logo, CSS, hero). Não contém lógica de busca."""

import streamlit as st

VERSAO = "v5"
NAVY, NAVY2, ORANGE, GRAY = "#0A1F44", "#14346B", "#FF6B1A", "#6B7280"


def _mark(size: int = 44, tile: str = NAVY, glyph: str = "#FFFFFF") -> str:
    return (
        f'<svg width="{size}" height="{size}" viewBox="0 0 48 48" xmlns="http://www.w3.org/2000/svg" aria-hidden="true">'
        f'<rect width="48" height="48" rx="13" fill="{tile}"/>'
        f'<path d="M13 24H35A11 11 0 1 0 31.8 31.8" fill="none" stroke="{glyph}" stroke-width="5.2" stroke-linecap="round" stroke-linejoin="round"/>'
        f'<circle cx="38.5" cy="9.5" r="4.4" fill="{ORANGE}"/></svg>'
    )


def logo(on_dark: bool = False, size: int = 44) -> str:
    """Logomarca: monograma 'e' + wordmark minúsculo 'estaduais.'"""
    tile, glyph, word = (ORANGE, "#FFFFFF", "#FFFFFF") if on_dark else (NAVY, "#FFFFFF", NAVY)
    dot = "#FFFFFF" if on_dark else ORANGE
    return (
        f'<span class="eb-logo">{_mark(size, tile, glyph)}'
        f'<span class="eb-word" style="color:{word}">estaduais<i style="color:{dot}">.</i></span></span>'
    )


CSS = f"""
<style>
@import url('https://fonts.googleapis.com/css2?family=Sora:wght@500;700;800&family=Inter:wght@400;500;600&display=swap');
:root {{ --navy:{NAVY}; --navy2:{NAVY2}; --orange:{ORANGE}; --gray:{GRAY}; --line:#E5E7EB; --soft:#F3F4F6; }}
html, body, [class*="css"], .stApp {{ font-family:'Inter',system-ui,sans-serif; }}
.stApp {{ background:#fff; color:var(--navy); }}
h1,h2,h3,h4 {{ font-family:'Sora',sans-serif !important; color:var(--navy); letter-spacing:-.02em; }}
#MainMenu, footer {{ visibility:hidden; }}
[data-testid="stHeader"] {{ background:transparent; }}
.block-container {{ padding-top:1.4rem; max-width:1240px; }}

/* ---------- logo ---------- */
.eb-logo {{ display:inline-flex; align-items:center; gap:.7rem; }}
.eb-word {{ font-family:'Sora',sans-serif; font-weight:800; font-size:1.7rem; letter-spacing:-.045em; line-height:1; }}
.eb-word i {{ font-style:normal; }}

/* ---------- hero ---------- */
.eb-hero {{ position:relative; overflow:hidden; border-radius:28px; padding:2rem 2.6rem 2.4rem; margin-bottom:1.6rem;
  background:radial-gradient(1200px 400px at 85% -10%, rgba(255,107,26,.28), transparent 60%),
             linear-gradient(135deg,var(--navy) 0%, var(--navy2) 100%); color:#fff; }}
.eb-hero:before {{ content:""; position:absolute; right:-120px; top:-120px; width:300px; height:300px; border-radius:50%;
  border:50px solid rgba(255,107,26,.95); opacity:.95; }}
.eb-hero:after {{ content:""; position:absolute; right:150px; bottom:-120px; width:260px; height:260px; border-radius:50%;
  border:2px solid rgba(255,255,255,.18); box-shadow:0 0 0 38px rgba(255,255,255,.04); }}
.eb-hero > * {{ position:relative; z-index:2; }}
.eb-top {{ display:flex; justify-content:space-between; align-items:center; margin-bottom:2.2rem; }}
.eb-pill {{ white-space:nowrap; font-size:.72rem; letter-spacing:.14em; text-transform:uppercase; font-weight:600; color:#fff;
  border:1px solid rgba(255,255,255,.35); padding:.4rem .85rem; border-radius:999px; backdrop-filter:blur(4px); background:rgba(255,255,255,.07); }}
.eb-hero h1 {{ font-family:'Sora',sans-serif !important; color:#fff !important; font-weight:800; font-size:clamp(1.9rem,3.7vw,3rem);
  line-height:1.04; margin:0 0 1rem; max-width:760px; padding:0; }}
.eb-hero h1 em {{ font-style:normal; color:var(--orange); }}
.eb-hero p {{ color:#C9D3E6; font-size:1.06rem; max-width:620px; margin:0 0 1.6rem; line-height:1.55; }}
.eb-chips {{ display:flex; gap:.7rem; flex-wrap:wrap; }}
.eb-chip {{ background:rgba(255,255,255,.1); border:1px solid rgba(255,255,255,.18); border-radius:16px; padding:.7rem 1.05rem; min-width:128px; }}
.eb-chip b {{ display:block; font-family:'Sora',sans-serif; font-size:1.55rem; color:#fff; line-height:1.1; }}
.eb-chip span {{ font-size:.78rem; color:#C9D3E6; }}
.eb-chip.hot {{ background:var(--orange); border-color:var(--orange); }}
.eb-chip.hot span {{ color:#FFF1E8; }}

/* ---------- formulário de busca ---------- */
[data-testid="stForm"] {{ border:1px solid var(--line); border-radius:22px; padding:1.4rem 1.5rem 1.1rem;
  background:#fff; box-shadow:0 18px 40px -22px rgba(10,31,68,.35); border-top:5px solid var(--orange); }}
[data-testid="stForm"] label p {{ font-weight:600; color:var(--navy); font-size:.82rem; letter-spacing:.08em; text-transform:uppercase; }}
[data-testid="stTextInput"] input {{ height:3.2rem; border-radius:14px; font-size:1.08rem; background:var(--soft); border:1.5px solid transparent; }}
[data-testid="stTextInput"] input:focus {{ border-color:var(--orange); background:#fff; box-shadow:0 0 0 4px rgba(255,107,26,.15); }}

/* ---------- botões ---------- */
.stButton button, .stDownloadButton button, [data-testid="stFormSubmitButton"] button {{
  border-radius:14px; font-weight:700; letter-spacing:.01em; padding:.7rem 1.2rem; transition:all .18s ease; }}
button[kind="primary"], button[data-testid="stBaseButton-primary"], button[data-testid="stBaseButton-primaryFormSubmit"] {{
  background:linear-gradient(135deg,#FF8A3D,var(--orange)); border:none; color:#fff;
  box-shadow:0 10px 22px -10px rgba(255,107,26,.9); }}
button[kind="primary"]:hover, button[data-testid="stBaseButton-primary"]:hover, button[data-testid="stBaseButton-primaryFormSubmit"]:hover {{
  background:var(--navy); transform:translateY(-1px); box-shadow:0 12px 24px -10px rgba(10,31,68,.8); color:#fff; }}
button[data-testid="stBaseButton-secondary"] {{ border:1.5px solid var(--navy); color:var(--navy); background:#fff; }}
button[data-testid="stBaseButton-secondary"]:hover {{ background:var(--navy); color:#fff; border-color:var(--navy); }}

/* ---------- barra lateral ---------- */
[data-testid="stSidebar"] {{ background:var(--soft); border-right:1px solid var(--line); }}
[data-testid="stSidebar"] h2, [data-testid="stSidebar"] h3 {{ font-size:1.05rem; text-transform:uppercase; letter-spacing:.1em; color:var(--navy); }}
[data-testid="stSidebar"] h2:before {{ content:""; display:inline-block; width:10px; height:10px; border-radius:50%; background:var(--orange); margin-right:.5rem; }}
span[data-baseweb="tag"], [data-testid="stMultiSelect"] [data-baseweb="tag"] {{ background-color:var(--navy) !important; border-radius:8px !important; }}
span[data-baseweb="tag"] *, [data-testid="stMultiSelect"] [data-baseweb="tag"] * {{ color:#fff !important; }}

/* ---------- métricas ---------- */
[data-testid="stMetric"] {{ background:#fff; border:1px solid var(--line); border-left:6px solid var(--orange);
  border-radius:18px; padding:1rem 1.2rem; box-shadow:0 12px 28px -22px rgba(10,31,68,.6); }}
[data-testid="stMetricValue"] {{ font-family:'Sora',sans-serif; font-weight:800; color:var(--navy); font-size:2.3rem; }}
[data-testid="stMetricLabel"] p {{ color:var(--gray); text-transform:uppercase; letter-spacing:.09em; font-size:.74rem; font-weight:600; }}

/* ---------- abas, expanders, tabelas, avisos ---------- */
button[data-baseweb="tab"] {{ font-weight:600; color:var(--gray); font-size:1rem; padding:.7rem 1.1rem; }}
button[data-baseweb="tab"][aria-selected="true"] {{ color:var(--navy); }}
[data-baseweb="tab-highlight"] {{ background:var(--orange) !important; height:4px !important; border-radius:4px; }}
[data-testid="stExpander"] {{ border:1px solid var(--line); border-radius:16px; overflow:hidden; }}
[data-testid="stExpander"] summary:hover {{ background:var(--soft); }}
[data-testid="stDataFrame"] {{ border:1px solid var(--line); border-radius:16px; overflow:hidden; }}
[data-testid="stAlert"] {{ border-radius:16px; border:1px solid var(--line); background:var(--soft); color:var(--navy); }}
hr {{ border-color:var(--line) !important; }}
@media (max-width:760px) {{ .eb-pill {{ display:none; }} .eb-word {{ font-size:1.45rem; }} .eb-hero {{ padding:1.4rem 1.3rem 1.7rem; }} .eb-hero:before {{ width:200px; height:200px; border-width:38px; }} }}
</style>
"""


def inject() -> None:
    st.markdown(CSS, unsafe_allow_html=True)


def hero(n_estados: int) -> None:
    st.markdown(
        f'<div class="eb-hero"><div class="eb-top">{logo(on_dark=True)}<span class="eb-pill">Legislação estadual · {VERSAO}</span></div>'
        f'<h1>Toda a legislação estadual,<br><em>numa só busca.</em></h1>'
        f'<p>Digite um tema e o estaduais consulta as Assembleias Legislativas, reúne estado, ementa e link direto da norma e entrega tudo em um arquivo Word.</p>'
        f'<div class="eb-chips"><div class="eb-chip hot"><b>{n_estados}</b><span>estados consultados</span></div>'
        f'<div class="eb-chip"><b>.docx</b><span>relatório pronto</span></div>'
        f'<div class="eb-chip"><b>1 clique</b><span>link direto da norma</span></div></div></div>',
        unsafe_allow_html=True,
    )
def make_favicon():
    from PIL import Image, ImageDraw
    S = 256
    im = Image.new("RGBA", (S, S), (0, 0, 0, 0)); d = ImageDraw.Draw(im)
    d.rounded_rectangle((0, 0, S - 1, S - 1), radius=64, fill="#0A1F44")
    c, r, w = S // 2, int(S * 0.23), int(S * 0.115)
    d.arc((c - r, c - r, c + r, c + r), start=45, end=360, fill="white", width=w)
    d.line((c - r, c, c + r, c), fill="white", width=w)
    d.ellipse((S - 82, 34, S - 34, 82), fill="#FF6B1A")
    return im


# ======================================================================
# COMPATIBILIDADE: permite `from busca_leg.<submódulo> import ...` com um único arquivo
# ======================================================================
import sys as _sys
import types as _types

_self = _sys.modules[__name__]
_self.__path__ = []  # trata este módulo como pacote
_ns = {k: v for k, v in globals().items() if not k.startswith("__")}
for _name in ("browser", "docgen", "engine", "parsers", "ui"):
    _m = _types.ModuleType(f"{__name__}.{_name}")
    _m.__dict__.update(_ns)
    _sys.modules[_m.__name__] = _m
    setattr(_self, _name, _m)
