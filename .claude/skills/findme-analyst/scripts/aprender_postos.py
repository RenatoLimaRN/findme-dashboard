#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
aprender_postos.py — Aprende novas atividades observadas no relatório, com
cautela (threshold de N=3 ocorrências antes de promover).

Lógica:
  1. Para cada modelo executado no dia (exceto avulsas — pontuais por
     natureza), incrementa contador em observados/<slug>.json keyed por
     (modelo, dia da semana).
  2. Quando um (modelo, dia) atinge o threshold:
       - Se o modelo é NOVO no postos/<slug>.json → adiciona como atividade
         recorrente (marca _auto=true, _aprendido_em=data).
       - Se o modelo JÁ existe mas o dia da semana não estava em "dias" →
         adiciona o dia.
  3. Cada incremento é idempotente por dia — rodar 2x no mesmo dia não
     duplica contagem.

Onde o modelo novo entra:
  - Se postos/<slug>.json tem só 1 posto → adiciona nesse posto
  - Se tem 2+ postos → cria/usa um posto "Outros (auto)" pra você
    recategorizar manualmente depois

Uso:
    python aprender_postos.py <dados.json> <postos-dir> <observados-dir> [--threshold N]

Imprime JSON com resumo.
"""
import argparse
import json
import re
import sys
import unicodedata
from datetime import date
from pathlib import Path

try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass

THRESHOLD_DEFAULT = 3
DIAS_PT_ORDEM = ["Seg", "Ter", "Qua", "Qui", "Sex", "Sab", "Dom"]
POSTO_OUTROS = "Outros (auto)"


def _slug(nome):
    s = unicodedata.normalize("NFD", nome or "")
    s = "".join(c for c in s if unicodedata.category(c) != "Mn")
    s = re.sub(r"^\s*CONDOMINIO\s+", "", s, flags=re.I).strip()
    return re.sub(r"[^A-Za-z0-9]+", "_", s.lower()).strip("_")


def _norm_modelo(m):
    s = unicodedata.normalize("NFD", m or "")
    s = "".join(c for c in s if unicodedata.category(c) != "Mn")
    return re.sub(r"\s+", " ", s.upper().strip())


def _weekday_pt(yyyy_mm_dd):
    try:
        return DIAS_PT_ORDEM[date.fromisoformat(yyyy_mm_dd).weekday()]
    except Exception:
        return None


def _carregar_json(fp: Path):
    if fp.exists():
        try:
            return json.loads(fp.read_text(encoding="utf-8"))
        except Exception:
            return None
    return None


def _salvar_json(fp: Path, data):
    fp.parent.mkdir(parents=True, exist_ok=True)
    fp.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


def _achar_atividade(postos_data, modelo_norm):
    """Procura uma atividade (modelo) já existente em qualquer posto.
    Retorna (posto_idx, ativ_idx) ou None."""
    for pi, p in enumerate(postos_data.get("postos", []) or []):
        for ai, a in enumerate(p.get("atividades", []) or []):
            if _norm_modelo(a.get("modelo", "")) == modelo_norm:
                return (pi, ai)
    return None


def _posto_destino(postos_data) -> int:
    """Decide em qual posto adicionar um modelo NOVO:
       - se há só 1 posto: usa esse
       - se há 2+: usa/cria 'Outros (auto)'"""
    postos = postos_data.setdefault("postos", [])
    if len(postos) == 1:
        return 0
    for pi, p in enumerate(postos):
        if p.get("posto") == POSTO_OUTROS:
            return pi
    postos.append({
        "posto": POSTO_OUTROS,
        "op_tipo": "N/C",
        "_instrucao": "Atividades aprendidas automaticamente. Mova para o posto correto quando souber.",
        "atividades": [],
    })
    return len(postos) - 1


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("dados", help="Caminho do dados.json (output do leitor)")
    ap.add_argument("postos_dir", help="Pasta postos/")
    ap.add_argument("observados_dir", help="Pasta observados/ (skill state)")
    ap.add_argument("--threshold", type=int, default=THRESHOLD_DEFAULT,
                    help=f"Quantas observações pra promover (default: {THRESHOLD_DEFAULT})")
    args = ap.parse_args()

    dados = _carregar_json(Path(args.dados))
    if dados is None:
        print(json.dumps({"erro": f"nao consegui ler dados.json: {args.dados}"}))
        sys.exit(1)

    data_alvo = dados.get("dia_alvo")
    if not data_alvo:
        print(json.dumps({"erro": "dados.json sem dia_alvo (rode o leitor com --data)"}))
        sys.exit(1)
    weekday = _weekday_pt(data_alvo)
    if weekday is None:
        print(json.dumps({"erro": f"data invalida: {data_alvo}"}))
        sys.exit(1)

    postos_dir = Path(args.postos_dir)
    obs_dir = Path(args.observados_dir)
    threshold = args.threshold

    n_promovidos = 0
    n_dias_expandidos = 0
    n_observados_inc = 0
    detalhe = []

    for agg in dados.get("atividades_agg", []) or []:
        slug = agg.get("slug")
        nome_local = agg.get("nome", slug)
        if not slug:
            continue

        obs = _carregar_json(obs_dir / f"{slug}.json") or {
            "local": nome_local, "slug": slug, "observados": {}
        }
        obs_mod = obs.setdefault("observados", {})

        postos_data = _carregar_json(postos_dir / f"{slug}.json")
        postos_existe = postos_data is not None

        modificou_obs = False
        modificou_postos = False

        for m in agg.get("por_modelo", []) or []:
            if m.get("avulsa"):
                continue
            modelo = (m.get("modelo") or "").strip()
            if not modelo or modelo == "(sem modelo)":
                continue
            modelo_norm = _norm_modelo(modelo)

            # incrementa observados (idempotente por dia)
            obs_modelo = obs_mod.setdefault(modelo_norm, {"_label": modelo, "dias": {}})
            # compat: se versão antiga (sem _label/dias), normaliza
            if "dias" not in obs_modelo:
                obs_modelo = {"_label": modelo, "dias": {k: v for k, v in obs_modelo.items() if isinstance(v, dict)}}
                obs_mod[modelo_norm] = obs_modelo
            obs_dia = obs_modelo["dias"].setdefault(weekday, {"qtd": 0, "ultima": ""})
            if obs_dia.get("ultima") != data_alvo:
                obs_dia["qtd"] += 1
                obs_dia["ultima"] = data_alvo
                modificou_obs = True
                n_observados_inc += 1

            qtd = obs_dia["qtd"]
            if qtd < threshold:
                continue
            if not postos_existe:
                continue  # sem template ainda — leitor sinaliza pra criar

            existente = _achar_atividade(postos_data, modelo_norm)
            if existente is None:
                pi = _posto_destino(postos_data)
                postos_data["postos"][pi]["atividades"].append({
                    "modelo": modelo,
                    "dias": [weekday],
                    "vezes": 1,
                    "_auto": True,
                    "_aprendido_em": data_alvo,
                    "_qtd_observada": qtd,
                })
                n_promovidos += 1
                modificou_postos = True
                detalhe.append({
                    "acao": "modelo_novo", "local": nome_local,
                    "modelo": modelo, "dia": weekday, "posto": postos_data["postos"][pi]["posto"],
                })
            else:
                pi, ai = existente
                ativ = postos_data["postos"][pi]["atividades"][ai]
                dias = ativ.setdefault("dias", [])
                if weekday not in dias:
                    dias.append(weekday)
                    ativ["_auto_dia_em"] = data_alvo
                    n_dias_expandidos += 1
                    modificou_postos = True
                    detalhe.append({
                        "acao": "dia_expandido", "local": nome_local,
                        "modelo": modelo, "dia": weekday,
                    })

        if modificou_obs:
            _salvar_json(obs_dir / f"{slug}.json", obs)
        if modificou_postos:
            _salvar_json(postos_dir / f"{slug}.json", postos_data)

    print(json.dumps({
        "ok": True,
        "threshold": threshold,
        "data_alvo": data_alvo,
        "dia_semana": weekday,
        "observacoes_incrementadas": n_observados_inc,
        "modelos_novos_promovidos": n_promovidos,
        "dias_expandidos": n_dias_expandidos,
        "detalhe": detalhe[:30],
    }, ensure_ascii=False))


if __name__ == "__main__":
    main()
