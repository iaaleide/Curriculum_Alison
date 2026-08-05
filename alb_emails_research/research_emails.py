#!/usr/bin/env python3
"""Research public company emails for ALB Brasil client list."""

from __future__ import annotations

import json
import re
import time
import urllib.parse
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path

import requests
from bs4 import BeautifulSoup

ROOT = Path("/workspace/alb_emails_research")
EMPRESAS = json.loads((ROOT / "empresas.json").read_text(encoding="utf-8"))
OUT = ROOT / "emails_encontrados.json"

EMAIL_RE = re.compile(
    r"(?<![A-Za-z0-9._%+-])([A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,})(?![A-Za-z0-9._%+-])"
)

BAD_DOMAINS = {
    "example.com",
    "email.com",
    "domain.com",
    "sentry.io",
    "wixpress.com",
    "cloudflare.com",
    "google.com",
    "gmail.com",  # keep personal? usually not company contact on corporate pages - filter later
    "schema.org",
    "w3.org",
    "wordpress.com",
    "godaddy.com",
    "squarespace.com",
    "facebook.com",
    "instagram.com",
    "linkedin.com",
    "youtube.com",
    "twitter.com",
    "x.com",
    "tiktok.com",
    "png",
    "jpg",
    "jpeg",
    "gif",
    "webp",
    "svg",
}

BAD_LOCAL = {
    "noreply",
    "no-reply",
    "donotreply",
    "mailer-daemon",
    "postmaster",
    "webmaster",
    "abuse",
    "spam",
    "privacy",
    "dpo",
    "lgpd",
}

PREFERRED_LOCAL = [
    "contato",
    "contact",
    "comercial",
    "vendas",
    "sales",
    "atendimento",
    "sac",
    "compras",
    "suprimentos",
    "fornecedor",
    "fornecedores",
    "ouvidoria",
    "financeiro",
    "adm",
    "administrativo",
    "recepcao",
    "info",
    "geral",
    "office",
    "hello",
    "faleconosco",
    "relacionamento",
    "export",
    "exportacao",
    "nutricao",
    "agro",
]

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36"
    ),
    "Accept-Language": "pt-BR,pt;q=0.9,en;q=0.8",
}

SESSION = requests.Session()
SESSION.headers.update(HEADERS)


def normalize_url(url: str) -> str | None:
    url = (url or "").strip()
    if not url:
        return None
    if url.startswith("mailto:"):
        return None
    if not re.match(r"^https?://", url, re.I):
        if "." in url and " " not in url:
            url = "https://" + url
        else:
            return None
    return url.rstrip("/")


def clean_email(e: str) -> str | None:
    e = e.strip().strip(".,;:<>\"'()[]{}").lower()
    e = e.replace("%20", "").replace(" ", "")
    if e.endswith((".", ",", ";")):
        e = e[:-1]
    # obfuscation leftovers
    e = e.replace("[at]", "@").replace("(at)", "@").replace(" at ", "@")
    if not EMAIL_RE.fullmatch(e):
        return None
    local, domain = e.split("@", 1)
    if any(domain.endswith(bad) or domain == bad for bad in BAD_DOMAINS):
        # allow gmail only if clearly company branded later; skip for now
        if domain not in {"gmail.com", "hotmail.com", "yahoo.com", "outlook.com", "uol.com.br", "terra.com.br"}:
            if domain in BAD_DOMAINS or any(domain.endswith("." + b) for b in BAD_DOMAINS):
                return None
    if any(x in local for x in BAD_LOCAL):
        return None
    if any(ext in domain for ext in (".png", ".jpg", ".jpeg", ".gif", ".webp", ".svg", ".css", ".js")):
        return None
    if len(local) < 2 or len(domain) < 4:
        return None
    # filter image/file false positives
    if re.search(r"\.(png|jpe?g|gif|webp|svg|css|js)$", e):
        return None
    return e


def score_email(email: str, company: str, website: str | None) -> int:
    score = 0
    local, domain = email.split("@", 1)
    for p in PREFERRED_LOCAL:
        if local == p or local.startswith(p + ".") or local.startswith(p + "_"):
            score += 50
            break
    # company name tokens in domain
    tokens = re.findall(r"[a-z0-9]{3,}", company.lower())
    for t in tokens[:5]:
        if t in domain:
            score += 30
            break
    if website:
        host = urllib.parse.urlparse(website).netloc.lower().replace("www.", "")
        if host and (domain == host or domain.endswith("." + host) or host.endswith(domain)):
            score += 40
    # personal mailbox lower priority unless only option
    if domain in {"gmail.com", "hotmail.com", "yahoo.com", "outlook.com", "uol.com.br", "terra.com.br"}:
        score -= 20
    if local in {"contato", "comercial", "vendas", "compras", "sac"}:
        score += 20
    return score


def extract_emails(text: str) -> set[str]:
    found = set()
    # deobfuscate common patterns
    text2 = text
    text2 = re.sub(r"\s*\[\s*at\s*\]\s*", "@", text2, flags=re.I)
    text2 = re.sub(r"\s*\(\s*at\s*\)\s*", "@", text2, flags=re.I)
    text2 = re.sub(r"\s+at\s+", "@", text2, flags=re.I)
    text2 = re.sub(r"\s*\[\s*dot\s*\]\s*", ".", text2, flags=re.I)
    text2 = re.sub(r"\s*\(\s*dot\s*\)\s*", ".", text2, flags=re.I)
    for m in EMAIL_RE.findall(text2):
        ce = clean_email(m)
        if ce:
            found.add(ce)
    # mailto
    for m in re.findall(r"mailto:([^\"'\s>?]+)", text, flags=re.I):
        ce = clean_email(urllib.parse.unquote(m.split("?")[0]))
        if ce:
            found.add(ce)
    return found


def fetch(url: str, timeout: float = 12) -> str:
    try:
        r = SESSION.get(url, timeout=timeout, allow_redirects=True)
        if r.status_code >= 400:
            return ""
        r.encoding = r.apparent_encoding or r.encoding
        return r.text or ""
    except Exception:
        return ""


def scrape_site(website: str) -> set[str]:
    url = normalize_url(website)
    if not url:
        return set()
    emails: set[str] = set()
    html = fetch(url)
    emails |= extract_emails(html)
    # follow contact-like links
    if html:
        soup = BeautifulSoup(html, "lxml")
        links = []
        for a in soup.find_all("a", href=True):
            href = a["href"].strip()
            txt = (a.get_text(" ") or "").lower()
            full = urllib.parse.urljoin(url + "/", href)
            key = (href + " " + txt).lower()
            if any(
                k in key
                for k in (
                    "contato",
                    "contact",
                    "fale",
                    "email",
                    "comercial",
                    "sobre",
                    "about",
                    "trabalhe",
                    "ouvidoria",
                    "sac",
                    "fornecedor",
                    "compras",
                )
            ):
                if full.startswith("http") and urllib.parse.urlparse(full).netloc:
                    links.append(full)
        # unique preserve order
        seen = set()
        for link in links:
            if link in seen:
                continue
            seen.add(link)
            if len(seen) > 6:
                break
            emails |= extract_emails(fetch(link))
    # try common contact paths
    base = url
    for path in ("/contato", "/contact", "/fale-conosco", "/contato/", "/contact-us", "/sobre", "/institucional"):
        if len(emails) >= 3:
            break
        emails |= extract_emails(fetch(base + path))
    return emails


def ddg_search(query: str, max_results: int = 8) -> list[str]:
    """DuckDuckGo HTML search; returns result snippets/urls text blob list."""
    url = "https://html.duckduckgo.com/html/"
    blobs: list[str] = []
    try:
        r = SESSION.post(url, data={"q": query}, timeout=20)
        if r.status_code != 200:
            return blobs
        soup = BeautifulSoup(r.text, "lxml")
        for res in soup.select(".result")[:max_results]:
            blobs.append(res.get_text(" ", strip=True))
            a = res.select_one("a.result__a")
            if a and a.get("href"):
                blobs.append(a["href"])
            sn = res.select_one(".result__snippet")
            if sn:
                blobs.append(sn.get_text(" ", strip=True))
    except Exception:
        pass
    return blobs


def bing_search(query: str, max_results: int = 8) -> list[str]:
    url = "https://www.bing.com/search?" + urllib.parse.urlencode({"q": query, "setlang": "pt-BR"})
    blobs: list[str] = []
    try:
        html = fetch(url, timeout=15)
        if not html:
            return blobs
        soup = BeautifulSoup(html, "lxml")
        for li in soup.select("li.b_algo")[:max_results]:
            blobs.append(li.get_text(" ", strip=True))
            a = li.select_one("h2 a")
            if a and a.get("href"):
                blobs.append(a["href"])
    except Exception:
        pass
    return blobs


def research_company(c: dict) -> dict:
    empresa = c["empresa"]
    website = c.get("website") or ""
    uf = c.get("uf") or ""
    mun = c.get("municipio") or ""
    emails: set[str] = set()
    sources: list[str] = []

    url = normalize_url(website)
    if url:
        site_emails = scrape_site(url)
        if site_emails:
            emails |= site_emails
            sources.append(f"website:{url}")

    queries = [
        f'"{empresa}" email contato',
        f'"{empresa}" "{uf}" contato@ OR comercial@ OR vendas@',
        f"{empresa} {mun} email",
    ]
    for q in queries:
        if len(emails) >= 5:
            break
        for blob in ddg_search(q) + bing_search(q):
            found = extract_emails(blob)
            if found:
                emails |= found
                sources.append(f"search:{q[:80]}")
        time.sleep(0.35)

    # if search found a likely corporate site without prior website, scrape top domains later via emails only

    ranked = sorted(
        emails,
        key=lambda e: score_email(e, empresa, url),
        reverse=True,
    )
    # keep top distinct useful
    best = ranked[:8]
    primary = best[0] if best else ""
    return {
        **c,
        "email_principal": primary,
        "emails_encontrados": best,
        "fonte": "; ".join(dict.fromkeys(sources))[:500],
        "status": "Encontrado" if primary else "Não encontrado",
        "confianca": (
            "Alta"
            if primary and score_email(primary, empresa, url) >= 60
            else ("Média" if primary and score_email(primary, empresa, url) >= 30 else ("Baixa" if primary else "-"))
        ),
    }


def main():
    results = []
    # resume support
    done = {}
    if OUT.exists():
        try:
            prev = json.loads(OUT.read_text(encoding="utf-8"))
            for p in prev:
                if p.get("email_principal") or p.get("status") == "Não encontrado":
                    done[p["id"]] = p
        except Exception:
            pass

    todo = [c for c in EMPRESAS if c["id"] not in done]
    print(f"Total {len(EMPRESAS)} | já feitos {len(done)} | restantes {len(todo)}")

    # process in small thread pools for site scraping, but keep search sequential-ish
    # Use 6 workers
    with ThreadPoolExecutor(max_workers=6) as ex:
        futs = {ex.submit(research_company, c): c for c in todo}
        for i, fut in enumerate(as_completed(futs), 1):
            c = futs[fut]
            try:
                res = fut.result()
            except Exception as e:
                res = {**c, "email_principal": "", "emails_encontrados": [], "fonte": "", "status": f"Erro: {e}", "confianca": "-"}
            done[res["id"]] = res
            results_sorted = [done[x["id"]] for x in EMPRESAS if x["id"] in done]
            OUT.write_text(json.dumps(results_sorted, ensure_ascii=False, indent=2), encoding="utf-8")
            print(
                f"[{i}/{len(todo)}] {res['empresa'][:40]:40} -> {res.get('email_principal') or res.get('status')}"
            )

    final = [done[x["id"]] for x in EMPRESAS]
    OUT.write_text(json.dumps(final, ensure_ascii=False, indent=2), encoding="utf-8")
    found = sum(1 for r in final if r.get("email_principal"))
    print(f"DONE: {found}/{len(final)} com email")


if __name__ == "__main__":
    main()
