#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
gerar_pdf_3dias.py — Resumo de tendência dos últimos N dias (default 3).

Lê os snapshots diários em historico/YYYY-MM-DD/<slug>.json e gera um PDF A4
com:
  • KPIs gerais dia a dia (eficiência, críticos, não feitas) — a curva
  • eficiência por local em cada dia + seta de tendência (quem piora/melhora)
  • locais piorando (queda consistente) em destaque
  • atividades que falham em TODOS os dias (problema crônico, não pontual)

"Conforme vai tendo os dados": usa os até N dias mais recentes (≤ dia alvo)
que têm snapshot — se só houver 1 ou 2, mostra esses.

Uso como módulo:
    from gerar_pdf_3dias import gerar
    gerar(Path("historico"), "2026-06-13", Path("resumo3.pdf"), n_dias=3)

Uso por linha de comando:
    python gerar_pdf_3dias.py <historico_dir> <YYYY-MM-DD> <saida.pdf> [n_dias]
"""
from __future__ import annotations

import json
import sys
from collections import Counter, defaultdict
from datetime import datetime
from pathlib import Path

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.platypus import (SimpleDocTemplate, Table, TableStyle,
                                Paragraph, Spacer)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER

AZUL      = colors.HexColor("#1F4E79")
VERDE     = colors.HexColor("#1E6B3C")
VERDE_CLR = colors.HexColor("#C6EFCE")
VERM      = colors.HexColor("#C0504D")
VERM_CLR  = colors.HexColor("#FFCCCC")
AMR_CLR   = colors.HexColor("#FFEB9C")
AMR_TXT   = colors.HexColor("#7F6000")
CINZA_CLR = colors.HexColor("#EDEAE2")
CINZA_TXT = colors.HexColor("#5F5E5A")
BRANCO    = colors.white
TXT       = colors.HexColor("#222222")


def _efic(ok, parcial, total):
    return round((ok + parcial * 0.5) / total * 100, 1) if total else 0.0


def _coletar(historico_dir: Path, dia_alvo: str, n_dias: int):
    """Retorna (dias_ordenados, {dia: {slug: snap}})."""
    historico_dir = Path(historico_dir)
    dias = []
    if historico_dir.exists():
        for d in historico_dir.iterdir():
            if d.is_dir() and len(d.name) == 10 and d.name <= dia_alvo:
                dias.append(d.name)
    dias = sorted(dias, reverse=True)[:n_dias]
    dias = sorted(dias)  # cronológico
    por_dia = {}
    for dia in dias:
        snaps = {}
        for f in (historico_dir / dia).glob("*.json"):
            try:
                s = json.loads(f.read_text(encoding="utf-8"))
                snaps[s.get("slug", f.stem)] = s
            except Exception:
                pass
        por_dia[dia] = snaps
    return dias, por_dia


def _cor_efic(pct, inconclusivo=False):
    if inconclusivo:
        return CINZA_CLR, CINZA_TXT
    if pct >= 70:
        return VERDE_CLR, VERDE
    if pct >= 40:
        return AMR_CLR, AMR_TXT
    return VERM_CLR, VERM


def gerar(historico_dir, dia_alvo: str, saida: Path, n_dias: int = 3) -> Path:
    saida = Path(saida)
    saida.parent.mkdir(parents=True, exist_ok=True)
    dias, por_dia = _coletar(historico_dir, dia_alvo, n_dias)

    styles = getSampleStyleSheet()
    st_titulo = ParagraphStyle("t", parent=styles["Normal"],
                               fontName="Helvetica-Bold", fontSize=16,
                               textColor=AZUL, leading=19)
    st_sub = ParagraphStyle("s", parent=styles["Normal"], fontName="Helvetica",
                            fontSize=9, textColor=CINZA_TXT, leading=12)
    st_secao = ParagraphStyle("sec", parent=styles["Normal"],
                              fontName="Helvetica-Bold", fontSize=11,
                              textColor=AZUL, spaceBefore=8, spaceAfter=4)
    st_cel = ParagraphStyle("c", parent=styles["Normal"], fontName="Helvetica",
                            fontSize=8.5, leading=11, textColor=TXT)
    st_txt = ParagraphStyle("x", parent=styles["Normal"], fontName="Helvetica",
                            fontSize=9, leading=14, textColor=TXT)
    st_rodape = ParagraphStyle("r", parent=styles["Normal"],
                               fontName="Helvetica-Oblique", fontSize=7.5,
                               textColor=CINZA_TXT, alignment=2)

    doc = SimpleDocTemplate(str(saida), pagesize=A4,
                            leftMargin=14 * mm, rightMargin=14 * mm,
                            topMargin=12 * mm, bottomMargin=12 * mm,
                            title=f"Resumo {len(dias)} dias até {dia_alvo}")
    largura = doc.width
    elems = []

    dias_br = [datetime.fromisoformat(d).strftime("%d/%m") for d in dias]
    intervalo = f"{dias_br[0]} a {dias_br[-1]}" if len(dias) > 1 else (dias_br[0] if dias_br else "—")

    elems.append(Paragraph(f"Resumo de tendência — {len(dias)} dia(s)", st_titulo))
    elems.append(Paragraph(f"FindMe · {intervalo}", st_sub))
    elems.append(Spacer(1, 10))

    if not dias:
        elems.append(Paragraph("Sem snapshots no histórico ainda.", st_txt))
        doc.build(elems)
        return saida

    # ── KPIs gerais dia a dia ────────────────────────────────────────────────
    elems.append(Paragraph("KPIs gerais por dia", st_secao))
    linhas = [["", *dias_br]]
    def _linha_kpi(rotulo, fn):
        return [rotulo, *[fn(por_dia[d]) for d in dias]]
    def _efic_geral(snaps):
        ok = sum(s.get("ok", 0) for s in snaps.values())
        pa = sum(s.get("parcial", 0) for s in snaps.values())
        tot = sum(s.get("total", 0) for s in snaps.values())
        return f"{_efic(ok, pa, tot):.0f}%"
    def _criticos(snaps):
        return str(sum(1 for s in snaps.values()
                       if _efic(s.get("ok", 0), s.get("parcial", 0),
                                s.get("total", 0)) < 70))
    def _naofeitas(snaps):
        return str(sum(s.get("nao_feita", 0) for s in snaps.values()))
    linhas.append(_linha_kpi("Eficiência geral", _efic_geral))
    linhas.append(_linha_kpi("Locais críticos (<70%)", _criticos))
    linhas.append(_linha_kpi("Atividades não feitas", _naofeitas))
    cw = [largura * 0.4] + [largura * 0.6 / len(dias)] * len(dias)
    tk = Table(linhas, colWidths=cw)
    tk.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), AZUL),
        ("TEXTCOLOR", (0, 0), (-1, 0), BRANCO),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTNAME", (0, 1), (0, -1), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), 9),
        ("ALIGN", (1, 0), (-1, -1), "CENTER"),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [BRANCO, colors.HexColor("#F7F7F5")]),
        ("LINEBELOW", (0, 0), (-1, -1), 0.4, colors.HexColor("#D9D9D9")),
    ]))
    elems.append(tk)
    elems.append(Spacer(1, 8))

    # ── Eficiência por local × dias + tendência ──────────────────────────────
    elems.append(Paragraph("Eficiência por local (com tendência)", st_secao))
    slugs = set()
    nomes = {}
    for d in dias:
        for slug, s in por_dia[d].items():
            slugs.add(slug)
            nomes[slug] = s.get("local", slug)

    def efic_dia(slug, dia):
        s = por_dia[dia].get(slug)
        if not s:
            return None, False
        inc = s.get("dados_inconclusivos") or (
            s.get("total", 0) >= 5 and s.get("ok", 0) == 0 and s.get("parcial", 0) == 0)
        return _efic(s.get("ok", 0), s.get("parcial", 0), s.get("total", 0)), inc

    linhas_loc = [["Local", *dias_br, "Tend."]]
    cor_cmds = []
    dados_loc = []
    for slug in slugs:
        vals = [efic_dia(slug, d) for d in dias]
        presentes = [v[0] for v in vals if v[0] is not None]
        # tendência: compara primeiro e último valor presente
        tend = "—"
        if len(presentes) >= 2:
            delta = presentes[-1] - presentes[0]
            if delta <= -10:
                tend = "▼ caindo"
            elif delta >= 10:
                tend = "▲ subindo"
            else:
                tend = "= estável"
        media = sum(presentes) / len(presentes) if presentes else -1
        dados_loc.append((media, slug, nomes[slug], vals, tend, presentes))

    # piores primeiro
    dados_loc.sort(key=lambda x: x[0])
    for ri, (media, slug, nome, vals, tend, presentes) in enumerate(dados_loc, start=1):
        cels = [Paragraph(nome, st_cel)]
        for ci, (pct, inc) in enumerate(vals, start=1):
            if pct is None:
                cels.append("·")
            else:
                cels.append("?" if inc else f"{pct:.0f}%")
                bg, fg = _cor_efic(pct, inc)
                cor_cmds.append(("BACKGROUND", (ci, ri), (ci, ri), bg))
                cor_cmds.append(("TEXTCOLOR", (ci, ri), (ci, ri), fg))
        cels.append(tend)
        linhas_loc.append(cels)

    cw2 = [largura * 0.46] + [largura * 0.36 / len(dias)] * len(dias) + [largura * 0.18]
    tl = Table(linhas_loc, colWidths=cw2, repeatRows=1)
    tl.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), AZUL),
        ("TEXTCOLOR", (0, 0), (-1, 0), BRANCO),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), 8.5),
        ("FONTNAME", (1, 1), (-2, -1), "Helvetica-Bold"),
        ("ALIGN", (1, 0), (-1, -1), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), 2.5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 2.5),
        ("LINEBELOW", (0, 0), (-1, -1), 0.3, colors.HexColor("#E0E0E0")),
    ] + cor_cmds))
    elems.append(tl)
    elems.append(Spacer(1, 8))

    # ── Atividades que falham SEMPRE (nos dias com dados) ─────────────────────
    falhas_por_modelo = defaultdict(set)
    for d in dias:
        for s in por_dia[d].values():
            for m in s.get("top_falhas_modelos", []) or []:
                if m.get("nao_feita", 0) > 0:
                    falhas_por_modelo[(s.get("local", ""), m.get("modelo", ""))].add(d)
    cronicas = [(loc, mod) for (loc, mod), ds in falhas_por_modelo.items()
                if len(ds) >= max(2, len(dias))]
    if cronicas:
        elems.append(Paragraph(
            f"Atividades que falharam nos {len(dias)} dias (crônicas)", st_secao))
        txt = "<br/>".join(
            f'<b>{loc}</b> — {mod}' for loc, mod in sorted(cronicas)[:14])
        elems.append(Paragraph(txt, st_txt))
        elems.append(Spacer(1, 6))

    # ── Legenda + rodapé ──────────────────────────────────────────────────────
    elems.append(Paragraph(
        '<font color="#1E6B3C">■</font> ≥70%&nbsp;&nbsp;'
        '<font color="#7F6000">■</font> 40-70%&nbsp;&nbsp;'
        '<font color="#C0504D">■</font> &lt;40%&nbsp;&nbsp;'
        '<font color="#5F5E5A">■</font> ? = verificar&nbsp;&nbsp;· = sem dado no dia',
        ParagraphStyle("leg", fontName="Helvetica", fontSize=7.5,
                       textColor=CINZA_TXT, leading=11)))
    elems.append(Spacer(1, 4))
    gerado = datetime.now().strftime("%d/%m/%Y %H:%M")
    elems.append(Paragraph(f"Gerado automaticamente em {gerado}", st_rodape))

    doc.build(elems)
    return saida


def main():
    if len(sys.argv) < 4:
        print("uso: python gerar_pdf_3dias.py <historico_dir> <YYYY-MM-DD> <saida.pdf> [n_dias]",
              file=sys.stderr)
        sys.exit(1)
    n = int(sys.argv[4]) if len(sys.argv) > 4 else 3
    out = gerar(Path(sys.argv[1]), sys.argv[2], Path(sys.argv[3]), n_dias=n)
    print(json.dumps({"ok": True, "pdf": str(out)}, ensure_ascii=False))


if __name__ == "__main__":
    main()
