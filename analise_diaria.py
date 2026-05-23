#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
analise_diaria.py — Automação diária do FindMe.

Gera o relatório do dia anterior, enriquece a aba Atividades (skill
findme-analyst), salva snapshots no histórico, e envia por e-mail com o
.xlsx anexado.

Uso:
    python analise_diaria.py                   # ontem (D-1) — padrão
    python analise_diaria.py --data 2026-05-15 # dia específico
    python analise_diaria.py --sem-email       # roda tudo menos o envio

Pré-requisitos:
    - email_config.json preenchido (use email_config.json.template como base)
    - postos/*.json com os registros de avulsas esperadas

Para agendamento diário no Windows: rode `agendar_analise.bat` uma vez.
"""
from __future__ import annotations

import argparse
import json
import os
import smtplib
import subprocess
import sys
import traceback
from datetime import date, datetime, timedelta
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from pathlib import Path

try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass

SCRIPT_DIR = Path(__file__).resolve().parent
SKILL_DIR = SCRIPT_DIR / ".claude" / "skills" / "findme-analyst"
WORKSPACE = SKILL_DIR.parent / "findme-analyst-workspace"

ENV_UTF8 = {**os.environ, "PYTHONUTF8": "1", "PYTHONIOENCODING": "utf-8"}


def log(msg: str) -> None:
    print(f"[{datetime.now().strftime('%H:%M:%S')}] {msg}", flush=True)


# ─── 1) Gerar o relatório (chama findme_programacao.py) ──────────────────────

def gerar_relatorio(data_alvo: str) -> Path:
    """Roda findme_programacao.py para o dia-alvo (start=end=data_alvo).
    Retorna o caminho do .xlsx gerado em relatorios/<data>_<data>/."""
    os.chdir(SCRIPT_DIR)
    sys.path.insert(0, str(SCRIPT_DIR))
    import findme_programacao as fp

    # Monkey-patch ask_date pra não pedir input interativo
    fp.ask_date = lambda prompt, default: data_alvo

    try:
        fp.main()
    except SystemExit as e:
        if e.code not in (0, None):
            raise RuntimeError(f"findme_programacao.py saiu com código {e.code}")

    xlsx = SCRIPT_DIR / "relatorios" / f"{data_alvo}_{data_alvo}" / f"GERAL_{data_alvo}_{data_alvo}.xlsx"
    if not xlsx.exists():
        raise FileNotFoundError(f"Relatório esperado não foi gerado: {xlsx}")
    return xlsx


# ─── 2) Enriquecer a aba Atividades + snapshots (skill findme-analyst) ──────

def _run(cmd: list[str], stdout_to: Path | None = None) -> subprocess.CompletedProcess:
    if stdout_to:
        with open(stdout_to, "w", encoding="utf-8") as f:
            r = subprocess.run(cmd, stdout=f, stderr=subprocess.PIPE, text=True, env=ENV_UTF8)
    else:
        r = subprocess.run(cmd, capture_output=True, text=True, env=ENV_UTF8)
    if r.returncode != 0:
        err = r.stderr or "(sem stderr)"
        raise RuntimeError(f"Comando falhou ({cmd[0]}): {err}")
    return r


def enriquecer_e_snapshot(xlsx_path: Path, data_alvo: str) -> dict:
    """Roda ler_relatorio → enriquecer_atividades → snapshots por local.
    Retorna o dict de dados (cruzamento, KPIs, etc.) usado no e-mail."""
    workspace = WORKSPACE / f"dia-{data_alvo}"
    workspace.mkdir(parents=True, exist_ok=True)
    dados_path = workspace / "dados.json"
    snaps_dir = workspace / "snaps"
    snaps_dir.mkdir(exist_ok=True)

    log("  leitor...")
    _run(
        [sys.executable, str(SKILL_DIR / "scripts" / "ler_relatorio.py"),
         str(xlsx_path), "--data", data_alvo,
         "--postos-dir", str(SCRIPT_DIR / "postos"),
         "--historico-dir", str(SKILL_DIR / "historico")],
        stdout_to=dados_path,
    )

    log("  enriquecer...")
    _run([sys.executable, str(SKILL_DIR / "scripts" / "enriquecer_atividades.py"),
          str(xlsx_path), str(dados_path)])

    log("  snapshots...")
    dados = json.loads(dados_path.read_text(encoding="utf-8"))
    cruz_idx = dados.get("cruzamento_por_local", {})
    n = 0
    for agg in dados.get("atividades_agg", []):
        slug = agg.get("slug")
        if not slug:
            continue
        cruz = cruz_idx.get(slug, {})
        top = [m for m in agg.get("por_modelo", []) if m.get("nao_feita", 0) > 0][:5]
        snap = {
            "data": data_alvo, "local": agg["nome"],
            "ok": agg.get("ok", 0), "parcial": agg.get("parcial", 0),
            "nao_feita": agg.get("nao_feita", 0), "total": agg.get("total", 0),
        }
        if cruz:
            snap["esperadas_total"] = cruz.get("esperadas_total", 0)
            snap["feitas_das_esperadas"] = cruz.get("feitas_ok", 0)
            snap["perdidas_das_esperadas"] = cruz.get("perdidas", 0)
            snap["modelos_perdidos_hoje"] = [
                e["modelo"] for e in cruz.get("esperadas_detalhe", [])
                if e.get("status") == "nao_feita"
            ][:8]
        snap["top_falhas_modelos"] = [
            {"modelo": m["modelo"], "nao_feita": m["nao_feita"]} for m in top
        ]
        snap_path = snaps_dir / f"{slug}.json"
        snap_path.write_text(json.dumps(snap, ensure_ascii=False, indent=2), encoding="utf-8")
        _run([sys.executable, str(SKILL_DIR / "scripts" / "snapshot.py"),
              str(SKILL_DIR / "historico"), str(snap_path)])
        n += 1
    log(f"  {n} snapshots salvos em historico/{data_alvo}/")
    return dados


# ─── 3) Montar HTML do e-mail ──────────────────────────────────────────────

CSS = """
body{font-family:-apple-system,Segoe UI,Roboto,sans-serif;color:#1a1a1a;line-height:1.5;max-width:720px;padding:8px}
h1{font-size:18px;color:#1F4E79;border-bottom:2px solid #1F4E79;padding-bottom:6px;margin:0 0 4px}
h2{font-size:14px;color:#1F4E79;margin:20px 0 8px;text-transform:uppercase;letter-spacing:.5px}
.sub{color:#6B7280;font-size:13px;margin:0 0 18px}
.kpi{display:inline-block;padding:5px 12px;background:#F4F8FC;border-left:3px solid #2E75B6;margin:0 6px 6px 0;font-size:13px}
table{border-collapse:collapse;width:100%;font-size:13px;margin:6px 0}
th{text-align:left;background:#1F4E79;color:#fff;padding:6px 10px;font-weight:600}
td{padding:5px 10px;border-bottom:1px solid #E5E7EB}
.ok{background:#E8F5E9}.warn{background:#FFF8E1}.crit{background:#FFEBEE}.incon{background:#F3E5F5}
.muted{color:#6B7280;font-size:12px;margin-top:18px;padding-top:10px;border-top:1px solid #E5E7EB}
"""

DIAS_PT = ["segunda", "terça", "quarta", "quinta", "sexta", "sábado", "domingo"]


def _categoria(agg: dict) -> str:
    total = agg.get("total", 0)
    ok = agg.get("ok", 0)
    parcial = agg.get("parcial", 0)
    nf = agg.get("nao_feita", 0)
    if total >= 5 and ok == 0 and parcial == 0:
        return "incon"  # inconclusivo (artefato de coleta)
    pct = (100.0 * ok / total) if total else 0
    if pct >= 90:
        return "ok"
    if pct >= 70:
        return "ok"  # ainda bom
    if pct >= 40:
        return "warn"
    return "crit"


def montar_html(data_alvo: str, dados: dict) -> str:
    d = datetime.fromisoformat(data_alvo)
    dia_label = f"{DIAS_PT[d.weekday()]}, {d.strftime('%d/%m/%Y')}"

    kpis = dados.get("kpis_gerais") or {}
    pct = kpis.get("pct_eficiencia_geral", "—")
    n_locais = kpis.get("locais_monitorados", "—")
    n_crit = kpis.get("locais_criticos", "—")
    n_nf = kpis.get("atividades_nao_feitas", "—")

    locais = []
    for a in dados.get("atividades_agg", []):
        total = a.get("total", 0)
        if total == 0:
            continue
        cat = _categoria(a)
        ok = a.get("ok", 0)
        pct_l = (100.0 * ok / total) if total else 0
        locais.append({
            "nome": a["nome"], "ok": ok, "parcial": a.get("parcial", 0),
            "nao_feita": a.get("nao_feita", 0), "total": total,
            "pct": pct_l, "cat": cat, "slug": a.get("slug", ""),
        })
    locais.sort(key=lambda x: -x["pct"])

    cruz_idx = dados.get("cruzamento_por_local", {})

    def linha(l):
        info = ""
        cruz = cruz_idx.get(l["slug"], {})
        if cruz.get("esperadas_total", 0) > 0:
            info = f' <span class="muted">· 📋 {cruz.get("feitas_ok",0)}/{cruz["esperadas_total"]} avulsas esperadas</span>'
        return (f'<tr class="{l["cat"]}"><td>{l["nome"]}</td>'
                f'<td>{l["ok"]} OK · {l["parcial"]} parc · {l["nao_feita"]} nf</td>'
                f'<td><b>{l["pct"]:.0f}%</b>{info}</td></tr>')

    bom = [l for l in locais if l["cat"] == "ok"]
    medio = [l for l in locais if l["cat"] == "warn"]
    crit = [l for l in locais if l["cat"] == "crit"]
    incon = [l for l in locais if l["cat"] == "incon"]

    # Nova seção: Top justificativas do dia
    just_top = dados.get("justificativas_top") or []
    just_html = ""
    if just_top:
        linhas_just = []
        for j in just_top[:8]:
            cat = j["categoria"]
            qtd = j["qtd"]
            linhas_just.append(f'<tr><td>{cat}</td><td style="text-align:right;width:60px"><b>{qtd}x</b></td></tr>')
        # Pega amostras de texto livre nas justificativas dos postos críticos
        amostras = []
        for slug in [l["slug"] for l in crit + medio][:5]:
            agg = next((a for a in dados.get("atividades_agg", []) if a.get("slug") == slug), None)
            if not agg:
                continue
            for j in (agg.get("justificativas") or [])[:2]:
                texto = j.get("texto", "")
                # remove o "Categoria — " do começo se for redundante; mantém só descrição
                if " — " in texto:
                    desc = texto.split(" — ", 1)[1].strip()
                else:
                    desc = texto
                if desc and desc != "-" and len(desc) > 3:
                    amostras.append(f'<li><b>{agg["nome"][:35]}</b> · {j["modelo"][:25]}: <i>"{desc[:90]}"</i></li>')
                    if len(amostras) >= 8:
                        break
            if len(amostras) >= 8:
                break
        amostras_html = ""
        if amostras:
            amostras_html = (
                '<p class="sub" style="margin-top:10px"><b>Frases reais da equipe:</b></p>'
                '<ul style="font-size:12px;margin:4px 0;padding-left:22px">'
                + "".join(amostras) + "</ul>")
        just_html = (
            '<h2>Por que falhou (justificativas)</h2>'
            '<table><tr><th>Categoria</th><th style="text-align:right">Ocorrências</th></tr>'
            + "".join(linhas_just) + "</table>" + amostras_html)

    historico = dados.get("historico_por_local", {})
    padroes_html = ""
    if historico:
        padroes_msgs = []
        for slug, snaps in historico.items():
            if len(snaps) >= 2:
                pct_hoje = next((l["pct"] for l in locais if l["slug"] == slug), None)
                pct_ant = snaps[0].get("pct_cumprimento")
                if pct_hoje is not None and pct_ant is not None:
                    if pct_hoje < 30 and pct_ant < 30:
                        nome = next((l["nome"] for l in locais if l["slug"] == slug), slug)
                        padroes_msgs.append(f"{nome} segue crítico ({pct_hoje:.0f}% hoje, {pct_ant:.0f}% antes)")
                    elif pct_hoje >= 80 and pct_ant < 30:
                        nome = next((l["nome"] for l in locais if l["slug"] == slug), slug)
                        padroes_msgs.append(f"{nome} recuperou ({pct_ant:.0f}% → {pct_hoje:.0f}%)")
        if padroes_msgs:
            padroes_html = "<h2>Padrões no histórico</h2><ul>" + "".join(
                f"<li>{m}</li>" for m in padroes_msgs[:8]) + "</ul>"

    n_template_vazio = len(dados.get("postos_template_vazio") or [])
    avisos = []
    if n_template_vazio:
        avisos.append(f"{n_template_vazio} posto(s) com template vazio em <code>postos/</code> aguardam preenchimento — sem isso, o cruzamento esperado×feito fica off pra eles.")

    avisos_html = ""
    if avisos:
        avisos_html = '<div class="muted" style="border-top:none;color:#B45309">⚠ ' + " · ".join(avisos) + "</div>"

    # Aprendizado automático — resumo no rodapé
    aprend = dados.get("aprendizado") or {}
    aprend_html = ""
    n_novos = aprend.get("modelos_novos_promovidos", 0)
    n_dias = aprend.get("dias_expandidos", 0)
    if n_novos or n_dias:
        bits = []
        if n_novos:
            bits.append(f"<b>{n_novos}</b> atividade(s) nova(s) inclusa(s) no registro")
        if n_dias:
            bits.append(f"<b>{n_dias}</b> dia(s) da semana expandido(s)")
        det = aprend.get("detalhe") or []
        lista = ""
        if det:
            itens = []
            for d_ in det[:10]:
                if d_.get("acao") == "modelo_novo":
                    itens.append(f"<li>{d_['local']} → <b>{d_['modelo']}</b> ({d_['dia']}, em <i>{d_.get('posto','?')}</i>)</li>")
                elif d_.get("acao") == "dia_expandido":
                    itens.append(f"<li>{d_['local']} · {d_['modelo']} → +{d_['dia']}</li>")
            if itens:
                lista = "<ul style='margin:6px 0;padding-left:22px;font-size:12px'>" + "".join(itens) + "</ul>"
        threshold = aprend.get("threshold", 3)
        aprend_html = (f'<div class="muted" style="color:#1B5E20">🎓 <b>Aprendizado automático</b> '
                       f'(após {threshold}+ aparições): ' + " · ".join(bits) + ".</div>"
                       + lista)

    return f"""<!DOCTYPE html>
<html><head><meta charset="utf-8"><style>{CSS}</style></head><body>
<h1>FindMe — Fechamento Operacional</h1>
<p class="sub">{dia_label} · gerado automaticamente</p>
<div>
  <span class="kpi"><b>{pct}%</b> eficiência geral</span>
  <span class="kpi"><b>{n_locais}</b> locais</span>
  <span class="kpi"><b>{n_crit}</b> críticos</span>
  <span class="kpi"><b>{n_nf}</b> não feitas</span>
</div>

{f'<h2>Crítico — ação urgente ({len(crit)})</h2><table><tr><th>Posto</th><th>Status</th><th>Cumprimento</th></tr>{"".join(linha(l) for l in crit)}</table>' if crit else ""}

{f'<h2>Atenção ({len(medio)})</h2><table><tr><th>Posto</th><th>Status</th><th>Cumprimento</th></tr>{"".join(linha(l) for l in medio)}</table>' if medio else ""}

{f'<h2>Inconclusivos — provável artefato de coleta ({len(incon)})</h2><p class="sub">Locais sem nenhum OK/parcial em volume alto. Escalar pro suporte FindMe.</p><table><tr><th>Posto</th><th>Status</th><th>Cumprimento</th></tr>{"".join(linha(l) for l in incon)}</table>' if incon else ""}

{f'<h2>Bom desempenho ({len(bom)})</h2><table><tr><th>Posto</th><th>Status</th><th>Cumprimento</th></tr>{"".join(linha(l) for l in bom)}</table>' if bom else ""}

{just_html}

{padroes_html}

<div class="muted">
<p>Detalhe linha-a-linha na aba <b>Atividades</b> do arquivo anexado: cores por status (verde=Completa, amarelo=Parcial, vermelho=Não Feita, vermelho-escuro=Esperada não registrada). Cabeçalho de cada local mostra cumprimento % e avulsas esperadas feitas.</p>
</div>
{avisos_html}
{aprend_html}
</body></html>"""


# ─── 4) Enviar e-mail via SMTP ──────────────────────────────────────────────

def carregar_email_config() -> dict:
    fp = SCRIPT_DIR / "email_config.json"
    if not fp.exists():
        raise SystemExit(
            "Falta email_config.json na raiz do projeto.\n"
            "Copie email_config.json.template, preencha e salve como email_config.json.\n"
            "Pra gerar a app password do Gmail: https://myaccount.google.com/apppasswords"
        )
    return json.loads(fp.read_text(encoding="utf-8"))


def enviar_email(cfg: dict, assunto: str, html: str, anexo: Path) -> None:
    msg = MIMEMultipart()
    msg["Subject"] = assunto
    from_name = cfg.get("from_name") or "FindMe Analyst"
    msg["From"] = f"{from_name} <{cfg['user']}>"
    destinos = cfg["to"] if isinstance(cfg["to"], list) else [cfg["to"]]
    msg["To"] = ", ".join(destinos)
    msg.attach(MIMEText(html, "html", "utf-8"))

    if anexo and anexo.exists():
        with open(anexo, "rb") as f:
            part = MIMEBase("application",
                            "vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            part.set_payload(f.read())
        encoders.encode_base64(part)
        part.add_header("Content-Disposition",
                        f'attachment; filename="{anexo.name}"')
        msg.attach(part)

    host = cfg.get("smtp_host", "smtp.gmail.com")
    port = int(cfg.get("smtp_port", 587))
    with smtplib.SMTP(host, port, timeout=30) as s:
        s.starttls()
        s.login(cfg["user"], cfg["password"])
        s.send_message(msg, from_addr=cfg["user"], to_addrs=destinos)


def _achar_relatorio_que_inclui(data_alvo: str) -> Path | None:
    """Procura em relatorios/ uma pasta YYYY-MM-DD_YYYY-MM-DD cujo range
    inclua a data_alvo. Prefere range exato de 1 dia."""
    rel_dir = SCRIPT_DIR / "relatorios"
    if not rel_dir.exists():
        return None
    candidatos = []
    alvo = date.fromisoformat(data_alvo)
    for folder in rel_dir.iterdir():
        if not folder.is_dir():
            continue
        parts = folder.name.split("_")
        if len(parts) != 2:
            continue
        try:
            ini = date.fromisoformat(parts[0])
            fim = date.fromisoformat(parts[1])
        except ValueError:
            continue
        if ini <= alvo <= fim:
            xlsx = folder / f"GERAL_{folder.name}.xlsx"
            if xlsx.exists():
                # prioridade: range de 1 dia > range maior
                duracao = (fim - ini).days
                candidatos.append((duracao, xlsx))
    if not candidatos:
        return None
    candidatos.sort(key=lambda x: x[0])
    return candidatos[0][1]


# ─── Main ───────────────────────────────────────────────────────────────────

def main():
    p = argparse.ArgumentParser()
    p.add_argument("--data", help="YYYY-MM-DD (default: ontem)")
    p.add_argument("--sem-email", action="store_true",
                   help="Gera e enriquece, mas não envia e-mail.")
    p.add_argument("--pular-gerar", action="store_true",
                   help="Pula a geração (usa o relatorio que já existe).")
    p.add_argument("--sem-aprender", action="store_true",
                   help="Desativa o auto-aprendizado de atividades novas.")
    p.add_argument("--threshold-aprender", type=int, default=3,
                   help="Quantas vezes uma atividade precisa aparecer pra ser promovida (default: 3).")
    args = p.parse_args()

    data_alvo = args.data or (date.today() - timedelta(days=1)).isoformat()
    log(f"=== ANÁLISE DIÁRIA — {data_alvo} ===")

    try:
        if args.pular_gerar:
            xlsx = _achar_relatorio_que_inclui(data_alvo)
            if xlsx is None:
                log(f"[1/3] --pular-gerar pedido, mas nenhum relatório cobre {data_alvo}.")
                log(f"      Caindo para geração via findme_programacao.py...")
                xlsx = gerar_relatorio(data_alvo)
                log(f"  -> {xlsx}")
            else:
                log(f"[1/3] Pulando geração (usando {xlsx.relative_to(SCRIPT_DIR)})")
        else:
            log("[1/3] Gerando relatório (chama findme_programacao.py)...")
            xlsx = gerar_relatorio(data_alvo)
            log(f"  -> {xlsx}")

        log("[2/3] Enriquecendo + snapshots...")
        dados = enriquecer_e_snapshot(xlsx, data_alvo)

        # Aprendizado automático (após enriquecer, antes do email)
        if not args.sem_aprender:
            log("  aprendendo atividades novas...")
            try:
                dados_path = WORKSPACE / f"dia-{data_alvo}" / "dados.json"
                obs_dir = SKILL_DIR / "observados"
                r = subprocess.run(
                    [sys.executable, str(SKILL_DIR / "scripts" / "aprender_postos.py"),
                     str(dados_path), str(SCRIPT_DIR / "postos"), str(obs_dir),
                     "--threshold", str(args.threshold_aprender)],
                    capture_output=True, text=True, env=ENV_UTF8,
                )
                if r.returncode == 0:
                    aprend = json.loads(r.stdout.strip() or "{}")
                    dados["aprendizado"] = aprend
                    n_novos = aprend.get("modelos_novos_promovidos", 0)
                    n_dias = aprend.get("dias_expandidos", 0)
                    log(f"  aprendi: {n_novos} modelo(s) novo(s), {n_dias} dia(s) expandido(s)")
                else:
                    log(f"  WARNING: aprender_postos falhou: {r.stderr.strip()[:200]}")
            except Exception as e:
                log(f"  WARNING: aprender_postos exception: {e}")

        if args.sem_email:
            log("OK (--sem-email, pulando envio)")
            return 0

        log("[3/3] Enviando e-mail...")
        cfg = carregar_email_config()
        html = montar_html(data_alvo, dados)
        d_br = datetime.fromisoformat(data_alvo).strftime("%d/%m/%Y")
        enviar_email(cfg, f"FindMe — Fechamento {d_br}", html, xlsx)
        destinos = cfg["to"] if isinstance(cfg["to"], list) else [cfg["to"]]
        log(f"  -> enviado para: {', '.join(destinos)}")
        log("CONCLUÍDO COM SUCESSO")
        return 0

    except Exception as e:
        log(f"ERRO: {e}")
        traceback.print_exc()
        # Tentar avisar por e-mail mesmo em caso de erro
        try:
            cfg = carregar_email_config()
            err_html = f"<h1>FindMe — Erro na análise de {data_alvo}</h1><pre>{traceback.format_exc()}</pre>"
            enviar_email(cfg, f"FindMe — ERRO {data_alvo}", err_html, None)
        except Exception:
            pass
        return 1


if __name__ == "__main__":
    sys.exit(main())
