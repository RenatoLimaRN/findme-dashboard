#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
sincronizar_postos.py — Puxa as atividades avulsas cadastradas direto do
FindMe v2 (v2.findme.id) e (re)gera os postos/<slug>.json automaticamente.

Fluxo:
  1. login por sessão em v2.findme.id (POST /sessions)
  2. /locations → mapa {nome -> id} de todos os locais
  3. /filters/locations (API) → nomes dos locais MONITORADOS (do config.json)
  4. para cada local monitorado:
       /locations/{id}/edit → ids dos postos (stations)
       /stations/{sid}/single_activities → nome do posto + atividades (nome,dias,vezes)
  5. escreve postos/<slug>.json com local → postos → atividades

Uso:
    python sincronizar_postos.py                 # todos os locais monitorados
    python sincronizar_postos.py "ALTA VISTA"    # só locais cujo nome contém o texto
    python sincronizar_postos.py --todos         # todos os locais da conta (não só monitorados)
"""
from __future__ import annotations

import json
import re
import sys
import time
import unicodedata
from pathlib import Path

import requests
from bs4 import BeautifulSoup

import findme_programacao as fp

V2 = "https://v2.findme.id"
DIAS = ["Dom", "Seg", "Ter", "Qua", "Qui", "Sex", "Sab"]
ROOT = Path(__file__).resolve().parent


def _norm(s: str) -> str:
    s = unicodedata.normalize("NFD", s or "")
    s = "".join(c for c in s if unicodedata.category(c) != "Mn")
    return re.sub(r"\s+", " ", s).strip().upper()


def _slug(nome: str) -> str:
    s = unicodedata.normalize("NFD", nome or "")
    s = "".join(c for c in s if unicodedata.category(c) != "Mn")
    s = re.sub(r"^\s*CONDOMINIO\s+", "", s, flags=re.I).strip()
    return re.sub(r"[^A-Za-z0-9]+", "_", s.lower()).strip("_") or "local"


def login_v2(session, email, senha):
    r = session.get(V2 + "/login", timeout=30)
    m = re.search(r'name="authenticity_token"[^>]*value="([^"]+)"', r.text)
    csrf = m.group(1) if m else ""
    r = session.post(V2 + "/sessions",
                     data={"authenticity_token": csrf, "email": email,
                           "password": senha, "commit": "Entrar"},
                     timeout=30, allow_redirects=True)
    if "/login" in r.url:
        raise RuntimeError("login no v2.findme.id falhou (cheque email/senha do config.json)")
    return True


def listar_locais_v2(session):
    """{nome_norm: (id, nome_original)} a partir do dropdown da /locations."""
    r = session.get(V2 + "/locations", timeout=30)
    mapa = {}
    for m in re.finditer(r'<option value="(\d+)">([^<]+)</option>', r.text):
        idloc, nome = m.group(1), m.group(2).strip()
        if nome and not nome.isdigit():
            mapa[_norm(nome)] = (idloc, nome)
    return mapa


def stations_do_local(session, id_local):
    r = session.get(f"{V2}/locations/{id_local}/edit", timeout=30)
    return sorted(set(re.findall(r"/stations/(\d+)/single_activities", r.text)),
                  key=int)


def parse_station(session, sid):
    """Retorna (nome_posto, op_tipo, [{modelo,dias,vezes}])."""
    r = session.get(f"{V2}/stations/{sid}/single_activities", timeout=30)
    soup = BeautifulSoup(r.text, "html.parser")
    h1 = soup.find("h1")
    nome_posto = "?"
    if h1:
        t = re.sub(r"\s+", " ", h1.get_text()).strip()
        m = re.search(r"Atividades avulsas de (.+?)(?:\s+Atividade avulsa)?$", t)
        nome_posto = (m.group(1).strip() if m else t).strip()

    ativ = {}  # nome -> {dias:set, vezes:int}
    tabela = soup.find("table", id="table-header")
    if tabela and tabela.find("tbody"):
        for tr in tabela.find("tbody").find_all("tr"):
            tds = tr.find_all("td", recursive=False)
            for col, td in enumerate(tds[:7]):
                dia = DIAS[col]
                for div in td.find_all("div", class_="activity"):
                    a = div.find("a", href=re.compile(r"/single_activities/\d+/edit"))
                    if not a:
                        continue
                    nome = a.get_text().strip()
                    vezes = 1
                    mv = re.search(r"Realizar\s+(\d+)\s+vez", div.get_text())
                    if mv:
                        vezes = int(mv.group(1))
                    d = ativ.setdefault(nome, {"dias": set(), "vezes": vezes})
                    d["dias"].add(dia)

    op_tipo = "Vigilante" if "RONDA" in _norm(nome_posto) else "Limpeza"
    atividades = []
    for nome, d in ativ.items():
        atividades.append({
            "modelo": nome,
            "dias": [x for x in DIAS if x in d["dias"]],
            "vezes": d["vezes"],
        })
    return nome_posto, op_tipo, atividades


def main():
    args = [a for a in sys.argv[1:]]
    todos = "--todos" in args
    filtro = next((a for a in args if not a.startswith("--")), None)

    cfg = fp.load_config("config.json")
    token = fp.login(cfg["email"], cfg["password"])

    # nomes monitorados (do config) via /filters/locations
    monit = set()
    if not todos:
        alll = fp.api_get(token, "/filters/locations")
        uuids = set(cfg.get("locations", []))
        for l in alll:
            if isinstance(l, dict) and l.get("uuid") in uuids:
                monit.add(_norm(l.get("name", "")))

    s = requests.Session()
    s.headers.update({"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"})
    login_v2(s, cfg["email"], cfg["password"])
    print("login v2 OK")

    locais = listar_locais_v2(s)
    print(f"{len(locais)} locais na conta\n")

    alvos = []
    for nome_norm, (idloc, nome_orig) in locais.items():
        if filtro and _norm(filtro) not in nome_norm:
            continue
        if not todos and not filtro and nome_norm not in monit:
            continue
        alvos.append((idloc, nome_orig))

    print(f"=== sincronizando {len(alvos)} local(is) ===")
    gerados = 0
    for idloc, nome_orig in alvos:
        sids = stations_do_local(s, idloc)
        postos = []
        for sid in sids:
            nome_posto, op_tipo, atividades = parse_station(s, sid)
            if not atividades:
                continue
            postos.append({"posto": nome_posto, "op_tipo": op_tipo,
                           "atividades": atividades})
            time.sleep(0.2)
        if not postos:
            print(f"  - {nome_orig}: sem atividades (pulado)")
            continue
        out = {"local": nome_orig, "postos": postos}
        caminho = ROOT / "postos" / f"{_slug(nome_orig)}.json"
        caminho.write_text(json.dumps(out, ensure_ascii=False, indent=2) + "\n",
                           encoding="utf-8")
        n_ativ = sum(len(p["atividades"]) for p in postos)
        print(f"  ✔ {nome_orig}: {len(postos)} posto(s), {n_ativ} atividade(s) -> {caminho.name}")
        gerados += 1
    print(f"\n{gerados} arquivo(s) gerado(s) em postos/")


if __name__ == "__main__":
    main()
