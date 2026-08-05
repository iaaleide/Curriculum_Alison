#!/usr/bin/env python3
"""Faster two-pass email research for ALB Brasil clients."""

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
CURATED = json.loads((ROOT / "curated_emails.json").read_text(encoding="utf-8"))
OUT = ROOT / "emails_encontrados.json"

EMAIL_RE = re.compile(
    r"(?<![A-Za-z0-9._%+-])([A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,})(?![A-Za-z0-9._%+-])"
)

BAD_SUFFIX = (".png", ".jpg", ".jpeg", ".gif", ".webp", ".svg", ".css", ".js", ".woff")
BAD_LOCAL = ("noreply", "no-reply", "donotreply", "mailer-daemon", "postmaster", "abuse", "webpack", "sentry")
PREFERRED = (
    "contato", "contact", "comercial", "vendas", "sales", "atendimento", "sac",
    "compras", "suprimentos", "fornecedor", "fornecedores", "info", "adm",
    "administrativo", "export", "exportacao", "relacionamento", "ouvidoria",
)

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36"
    ),
    "Accept-Language": "pt-BR,pt;q=0.9,en;q=0.8",
}


def session():
    s = requests.Session()
    s.headers.update(HEADERS)
    return s


S = session()


def normalize_url(url: str) -> str | None:
    url = (url or "").strip()
    if not url or " " in url:
        return None
    if not re.match(r"^https?://", url, re.I):
        if "." in url:
            url = "https://" + url
        else:
            return None
    return url.rstrip("/")


def clean_email(e: str) -> str | None:
    e = urllib.parse.unquote(e).strip().strip(".,;:<>\"'()[]{}").lower()
    e = e.replace("%20", "").replace(" ", "")
    if not EMAIL_RE.fullmatch(e):
        return None
    local, domain = e.split("@", 1)
    if any(e.endswith(sfx) for sfx in BAD_SUFFIX):
        return None
    if any(x in local for x in BAD_LOCAL):
        return None
    if domain in {"example.com", "email.com", "domain.com", "sentry.io", "wixpress.com", "schema.org", "w3.org"}:
        return None
    if len(local) < 2:
        return None
    return e


def extract_emails(text: str) -> set[str]:
    found = set()
    text2 = re.sub(r"\s*\[\s*at\s*\]\s*", "@", text, flags=re.I)
    text2 = re.sub(r"\s*\(\s*at\s*\)\s*", "@", text2, flags=re.I)
    for m in EMAIL_RE.findall(text2):
        ce = clean_email(m)
        if ce:
            found.add(ce)
    for m in re.findall(r"mailto:([^\"'\s>?]+)", text, flags=re.I):
        ce = clean_email(m.split("?")[0])
        if ce:
            found.add(ce)
    return found


def fetch(url: str, timeout: float = 10) -> str:
    try:
        r = S.get(url, timeout=timeout, allow_redirects=True)
        if r.status_code >= 400:
            return ""
        r.encoding = r.apparent_encoding or r.encoding
        return r.text or ""
    except Exception:
        return ""


def score_email(email: str, company: str, website: str | None) -> int:
    score = 0
    local, domain = email.split("@", 1)
    if local in PREFERRED or any(local.startswith(p + x) for p in PREFERRED for x in (".", "_", "-")):
        score += 55
    tokens = re.findall(r"[a-z0-9]{3,}", company.lower())
    for t in tokens[:6]:
        if t in domain:
            score += 35
            break
    if website:
        host = urllib.parse.urlparse(website).netloc.lower().replace("www.", "")
        if host and (domain == host or host.endswith(domain) or domain.endswith(host.split(".")[-2] + "." + host.split(".")[-1] if "." in host else host)):
            score += 45
    if domain in {"gmail.com", "hotmail.com", "yahoo.com", "outlook.com", "uol.com.br", "terra.com.br", "icloud.com"}:
        score -= 15
    return score


def scrape_site(website: str) -> set[str]:
    url = normalize_url(website)
    if not url:
        return set()
    emails: set[str] = set()
    html = fetch(url)
    emails |= extract_emails(html)
    if html:
        soup = BeautifulSoup(html, "lxml")
        links = []
        for a in soup.find_all("a", href=True):
            href = a["href"].strip()
            txt = (a.get_text(" ") or "").lower()
            key = (href + " " + txt).lower()
            if any(k in key for k in ("contato", "contact", "fale", "comercial", "sac", "fornecedor", "compras")):
                full = urllib.parse.urljoin(url + "/", href)
                if full.startswith("http"):
                    links.append(full)
        for link in list(dict.fromkeys(links))[:5]:
            emails |= extract_emails(fetch(link))
    for path in ("/contato", "/fale-conosco", "/contact", "/contato/", "/institucional"):
        if len(emails) >= 4:
            break
        emails |= extract_emails(fetch(url + path))
    return emails


def bing_search_emails(query: str) -> set[str]:
    url = "https://www.bing.com/search?" + urllib.parse.urlencode({"q": query, "setlang": "pt-BR"})
    html = fetch(url, timeout=12)
    return extract_emails(html) if html else set()


def ddg_search_emails(query: str) -> set[str]:
    try:
        r = S.post("https://html.duckduckgo.com/html/", data={"q": query}, timeout=15)
        if r.status_code != 200:
            return set()
        return extract_emails(r.text)
    except Exception:
        return set()


def pick_best(emails: set[str], company: str, website: str | None) -> list[str]:
    return sorted(emails, key=lambda e: score_email(e, company, website), reverse=True)[:8]


def research_one(c: dict, do_search: bool = True) -> dict:
    key = c["empresa"].strip().lower()
    if key in CURATED:
        cur = CURATED[key]
        return {
            **c,
            "email_principal": cur.get("email_principal", ""),
            "emails_encontrados": cur.get("emails_encontrados", []),
            "fonte": cur.get("fonte", "curadoria"),
            "status": "Encontrado" if cur.get("email_principal") else "Não encontrado",
            "confianca": cur.get("confianca", "Alta"),
            "website_corrigido": cur.get("website_corrigido", ""),
            "obs_contato": cur.get("obs_contato", c.get("obs", "")),
        }

    website = c.get("website") or ""
    url = normalize_url(website)
    emails: set[str] = set()
    sources: list[str] = []

    if url:
        se = scrape_site(url)
        if se:
            emails |= se
            sources.append(f"website:{url}")

    if do_search and len(emails) < 2:
        q1 = f'"{c["empresa"]}" (contato OR comercial OR vendas) email'
        q2 = f'{c["empresa"]} {c.get("municipio","")} @'
        for q in (q1, q2):
            found = bing_search_emails(q)
            if not found:
                found = ddg_search_emails(q)
            if found:
                emails |= found
                sources.append(f"busca:{q[:70]}")
            time.sleep(0.2)
            if len(emails) >= 4:
                break

    best = pick_best(emails, c["empresa"], url)
    primary = best[0] if best else ""
    conf = "-"
    if primary:
        sc = score_email(primary, c["empresa"], url)
        conf = "Alta" if sc >= 60 else ("Média" if sc >= 25 else "Baixa")
    return {
        **c,
        "email_principal": primary,
        "emails_encontrados": best,
        "fonte": "; ".join(dict.fromkeys(sources))[:500],
        "status": "Encontrado" if primary else "Não encontrado",
        "confianca": conf,
        "website_corrigido": "",
        "obs_contato": c.get("obs", ""),
    }


def save(done: dict):
    final = [done[x["id"]] for x in EMPRESAS if x["id"] in done]
    OUT.write_text(json.dumps(final, ensure_ascii=False, indent=2), encoding="utf-8")


def main():
    done = {}
    if OUT.exists():
        try:
            for p in json.loads(OUT.read_text(encoding="utf-8")):
                done[p["id"]] = p
        except Exception:
            pass

    # Pass 1: curated + website-heavy, search only if no website
    todo = [c for c in EMPRESAS if c["id"] not in done or not done[c["id"]].get("status")]
    print(f"PASS1 researching {len(todo)} companies...", flush=True)

    with ThreadPoolExecutor(max_workers=8) as ex:
        futs = {}
        for c in todo:
            has_site = bool(normalize_url(c.get("website") or ""))
            key = c["empresa"].strip().lower()
            do_search = (not has_site) or key in CURATED
            # For curated, no need heavy search
            if key in CURATED:
                do_search = False
            futs[ex.submit(research_one, c, do_search)] = c
        for i, fut in enumerate(as_completed(futs), 1):
            res = fut.result()
            done[res["id"]] = res
            if i % 10 == 0 or i == len(futs):
                save(done)
            print(f"[{i}/{len(futs)}] {res['empresa'][:42]:42} -> {res.get('email_principal') or '-'}", flush=True)

    save(done)

    # Pass 2: retry missing with search
    todo2 = [c for c in EMPRESAS if not done[c["id"]].get("email_principal")]
    print(f"PASS2 retrying {len(todo2)} without email...", flush=True)
    with ThreadPoolExecutor(max_workers=5) as ex:
        futs = {ex.submit(research_one, c, True): c for c in todo2}
        for i, fut in enumerate(as_completed(futs), 1):
            res = fut.result()
            if res.get("email_principal") or not done[res["id"]].get("email_principal"):
                done[res["id"]] = res
            if i % 10 == 0 or i == len(futs):
                save(done)
            print(f"[P2 {i}/{len(futs)}] {res['empresa'][:42]:42} -> {res.get('email_principal') or '-'}", flush=True)

    save(done)
    final = [done[x["id"]] for x in EMPRESAS]
    found = sum(1 for r in final if r.get("email_principal"))
    print(f"DONE {found}/{len(final)}", flush=True)


if __name__ == "__main__":
    # clear previous partial if incomplete
    main()
