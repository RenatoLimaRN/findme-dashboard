#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
gerar_pdf_resumo.py — Resumo executivo (PDF) do fechamento diário.

Lê o dict de dados produzido pelo pipeline (ler_relatorio.py / analise_diaria)
e gera um PDF A4 enxuto com:
  • faixa de 4 KPIs (eficiência, locais, críticos, não feitas)
  • tabela de TODOS os locais, ordenada do pior pro melhor, colorida por
    faixa de eficiência (com destaque pros prováveis artefatos de coleta)
  • as justificativas mais usadas no dia

Eficiência por local = (feitas + 0,5 × parciais) / total — mesma fórmula do
KPI geral, pra não confundir com o "cumprimento" (só feitas) do Excel.

Uso como módulo:
    from gerar_pdf_resumo import gerar
    gerar(dados, "2026-06-11", Path("resumo.pdf"))

Uso por linha de comando:
    python gerar_pdf_resumo.py dados.json 2026-06-11 resumo.pdf
"""
from __future__ import annotations

import json
import sys
from datetime import datetime
from pathlib import Path

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.platypus import (SimpleDocTemplate, Table, TableStyle,
                                Paragraph, Spacer)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT

# ─── Paleta (mesma do Excel) ─────────────────────────────────────────────────
AZUL      = colors.HexColor("#1F4E79")
AZUL2     = colors.HexColor("#4472C4")
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

DIAS_PT = ["segunda-feira", "terça-feira", "quarta-feira", "quinta-feira",
           "sexta-feira", "sábado", "domingo"]


def _efic(ok: int, parcial: int, total: int) -> float:
    return round((ok + parcial * 0.5) / total * 100, 1) if total else 0.0


def _faixa(ok: int, parcial: int, nf: int, total: int):
    """Retorna (cor_fundo, cor_texto, rotulo) da linha do local."""
    # artefato de coleta: muita atividade, zero feita E zero parcial
    if total >= 5 and ok == 0 and parcial == 0:
        return CINZA_CLR, CINZA_TXT, "verificar"
    pct = _efic(ok, parcial, total)
    if pct >= 90:
        return VERDE_CLR, VERDE, "ótimo"
    if pct >= 70:
        return VERDE_CLR, VERDE, "bom"
    if pct >= 40:
        return AMR_CLR, AMR_TXT, "atenção"
    return VERM_CLR, VERM, "crítico"


def _kpis(dados: dict) -> dict:
    k = dados.get("kpis_gerais", {}) or {}
    locais = dados.get("locais", []) or []
    if not k:
        ok = sum(l.get("ok", 0) for l in locais)
        pa = sum(l.get("parcial", 0) for l in locais)
        nf = sum(l.get("nao_feita", 0) for l in locais)
        tot = sum(l.get("total", 0) for l in locais)
        k = {
            "pct_eficiencia_geral": _efic(ok, pa, tot),
            "locais_monitorados": len(locais),
            "locais_criticos": sum(1 for l in locais
                                   if _efic(l.get("ok", 0), l.get("parcial", 0),
                                            l.get("total", 0)) < 70),
            "atividades_nao_feitas": nf,
        }
    return k


def gerar(dados: dict, data_alvo: str, saida: Path) -> Path:
    saida = Path(saida)
    saida.parent.mkdir(parents=True, exist_ok=True)

    dt = datetime.fromisoformat(data_alvo)
    data_br = dt.strftime("%d/%m/%Y")
    dia_sem = DIAS_PT[dt.weekday()]

    styles = getSampleStyleSheet()
    st_titulo = ParagraphStyle("titulo", parent=styles["Normal"],
                               fontName="Helvetica-Bold", fontSize=16,
                               textColor=AZUL, leading=19)
    st_sub = ParagraphStyle("sub", parent=styles["Normal"],
                            fontName="Helvetica", fontSize=9,
                            textColor=CINZA_TXT, leading=12)
    st_data = ParagraphStyle("data", parent=styles["Normal"],
                             fontName="Helvetica-Bold", fontSize=13,
                             textColor=TXT, alignment=2, leading=16)
    st_data_sub = ParagraphStyle("datasub", parent=styles["Normal"],
                                 fontName="Helvetica", fontSize=9,
                                 textColor=CINZA_TXT, alignment=2, leading=12)
    st_secao = ParagraphStyle("secao", parent=styles["Normal"],
                              fontName="Helvetica-Bold", fontSize=11,
                              textColor=AZUL, spaceBefore=8, spaceAfter=4)
    st_cel = ParagraphStyle("cel", parent=styles["Normal"],
                            fontName="Helvetica", fontSize=8.5, leading=11,
                            textColor=TXT)
    st_just = ParagraphStyle("just", parent=styles["Normal"],
                             fontName="Helvetica", fontSize=9, leading=15,
                             textColor=TXT)
    st_rodape = ParagraphStyle("rodape", parent=styles["Normal"],
                               fontName="Helvetica-Oblique", fontSize=7.5,
                               textColor=CINZA_TXT, alignment=2)

    doc = SimpleDocTemplate(
        str(saida), pagesize=A4,
        leftMargin=14 * mm, rightMargin=14 * mm,
        topMargin=12 * mm, bottomMargin=12 * mm,
        title=f"Resumo FindMe {data_br}",
    )
    largura = doc.width
    elems = []

    # ── Cabeçalho ────────────────────────────────────────────────────────────
    cab = Table([[
        [Paragraph("Relatório diário de operações", st_titulo),
         Paragraph("FindMe · rondas e limpeza", st_sub)],
        [Paragraph(data_br, st_data),
         Paragraph(dia_sem, st_data_sub)],
    ]], colWidths=[largura * 0.62, largura * 0.38])
    cab.setStyle(TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LINEBELOW", (0, 0), (-1, -1), 1.2, AZUL),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
    ]))
    elems.append(cab)
    elems.append(Spacer(1, 10))

    # ── KPIs ──────────────────────────────────────────────────────────────────
    k = _kpis(dados)
    efic = k.get("pct_eficiencia_geral", 0)
    n_loc = k.get("locais_monitorados", 0)
    n_crit = k.get("locais_criticos", 0)
    n_nf = k.get("atividades_nao_feitas", 0)

    def kpi_cell(valor, label, fundo, txt_cor):
        return [Paragraph(f'<font size=22><b>{valor}</b></font>',
                          ParagraphStyle("kv", fontName="Helvetica-Bold",
                                         fontSize=22, textColor=txt_cor,
                                         alignment=TA_CENTER, leading=24)),
                Paragraph(label,
                          ParagraphStyle("kl", fontName="Helvetica", fontSize=8,
                                         textColor=txt_cor, alignment=TA_CENTER,
                                         leading=10))]

    efic_fundo = VERM_CLR if efic < 70 else (AMR_CLR if efic < 90 else VERDE_CLR)
    efic_txt = VERM if efic < 70 else (AMR_TXT if efic < 90 else VERDE)
    crit_fundo = VERM_CLR if n_crit > 0 else CINZA_CLR
    crit_txt = VERM if n_crit > 0 else CINZA_TXT

    kpis = Table([[
        kpi_cell(f"{efic:.0f}%", "eficiência geral", efic_fundo, efic_txt),
        kpi_cell(n_loc, "locais monitorados", CINZA_CLR, CINZA_TXT),
        kpi_cell(n_crit, "críticos (<70%)", crit_fundo, crit_txt),
        kpi_cell(n_nf, "atividades não feitas", CINZA_CLR, CINZA_TXT),
    ]], colWidths=[largura / 4.0] * 4)
    kpis.setStyle(TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), 8),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
        ("BACKGROUND", (0, 0), (0, 0), efic_fundo),
        ("BACKGROUND", (1, 0), (1, 0), CINZA_CLR),
        ("BACKGROUND", (2, 0), (2, 0), crit_fundo),
        ("BACKGROUND", (3, 0), (3, 0), CINZA_CLR),
        ("LEFTPADDING", (0, 0), (-1, -1), 4),
        ("RIGHTPADDING", (0, 0), (-1, -1), 4),
        ("INNERGRID", (0, 0), (-1, -1), 3, BRANCO),
    ]))
    elems.append(kpis)
    elems.append(Spacer(1, 6))

    # ── Tabela de TODOS os locais ─────────────────────────────────────────────
    elems.append(Paragraph("Todos os locais — ordenados do pior pro melhor",
                           st_secao))

    locais = list(dados.get("locais", []) or [])
    for l in locais:
        l["_efic"] = _efic(l.get("ok", 0), l.get("parcial", 0), l.get("total", 0))
    locais.sort(key=lambda l: (l["_efic"], -l.get("total", 0)))

    header = ["#", "Local", "Feitas", "Parc.", "Não feitas",
              "Total", "Efic.", "Situação"]
    linhas = [header]
    estilos_linha = []
    for i, l in enumerate(locais):
        ok = l.get("ok", 0); pa = l.get("parcial", 0)
        nf = l.get("nao_feita", 0); tot = l.get("total", 0)
        fundo, cor_txt, rotulo = _faixa(ok, pa, nf, tot)
        linhas.append([
            str(i + 1),
            Paragraph(str(l.get("nome", "-")), st_cel),
            str(ok), str(pa), str(nf), str(tot),
            f"{l['_efic']:.0f}%", rotulo,
        ])
        r = i + 1
        estilos_linha.append(("BACKGROUND", (6, r), (7, r), fundo))
        estilos_linha.append(("TEXTCOLOR", (6, r), (7, r), cor_txt))

    col_w = [largura * w for w in
             (0.04, 0.40, 0.085, 0.075, 0.11, 0.075, 0.085, 0.13)]
    tab = Table(linhas, colWidths=col_w, repeatRows=1)
    base = [
        ("BACKGROUND", (0, 0), (-1, 0), AZUL),
        ("TEXTCOLOR", (0, 0), (-1, 0), BRANCO),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, 0), 8),
        ("FONTNAME", (0, 1), (-1, -1), "Helvetica"),
        ("FONTSIZE", (0, 1), (-1, -1), 8.5),
        ("ALIGN", (2, 0), (-1, -1), "CENTER"),
        ("ALIGN", (0, 0), (0, -1), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("FONTNAME", (6, 1), (6, -1), "Helvetica-Bold"),
        ("TOPPADDING", (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
        ("LINEBELOW", (0, 0), (-1, -1), 0.4, colors.HexColor("#D9D9D9")),
        ("ROWBACKGROUNDS", (0, 1), (5, -1), [BRANCO, colors.HexColor("#F7F7F5")]),
    ]
    tab.setStyle(TableStyle(base + estilos_linha))
    elems.append(tab)
    elems.append(Spacer(1, 8))

    # ── Justificativas ─────────────────────────────────────────────────────────
    jt = dados.get("justificativas_top", []) or []
    if jt:
        elems.append(Paragraph("Por que falhou — justificativas do dia",
                               st_secao))
        chips = "&nbsp;&nbsp;".join(
            f'<font backColor="#E6F1FB" color="#0C447C"> '
            f'{j.get("categoria", "-")} · {j.get("qtd", 0)} </font>'
            for j in jt[:8]
        )
        elems.append(Paragraph(chips, st_just))
        elems.append(Spacer(1, 6))

    # ── Legenda + rodapé ───────────────────────────────────────────────────────
    legenda = ('<font color="#1E6B3C">■</font> bom/ótimo&nbsp;&nbsp;'
               '<font color="#7F6000">■</font> atenção (40-70%)&nbsp;&nbsp;'
               '<font color="#C0504D">■</font> crítico (&lt;40%)&nbsp;&nbsp;'
               '<font color="#5F5E5A">■</font> verificar (0 feitas — '
               'possível falha de cadastro/coleta)')
    elems.append(Paragraph(legenda, ParagraphStyle(
        "leg", fontName="Helvetica", fontSize=7.5, textColor=CINZA_TXT,
        leading=11)))
    elems.append(Spacer(1, 4))
    gerado = datetime.now().strftime("%d/%m/%Y %H:%M")
    elems.append(Paragraph(
        f"Gerado automaticamente em {gerado} · detalhe completo no Excel anexo",
        st_rodape))

    doc.build(elems)
    return saida


def main():
    if len(sys.argv) < 4:
        print("uso: python gerar_pdf_resumo.py <dados.json> <YYYY-MM-DD> <saida.pdf>",
              file=sys.stderr)
        sys.exit(1)
    dados = json.loads(Path(sys.argv[1]).read_text(encoding="utf-8"))
    out = gerar(dados, sys.argv[2], Path(sys.argv[3]))
    print(json.dumps({"ok": True, "pdf": str(out)}, ensure_ascii=False))


if __name__ == "__main__":
    main()
