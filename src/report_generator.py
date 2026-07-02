from __future__ import annotations

from dataclasses import dataclass, field
from io import BytesIO
from pathlib import Path
from tempfile import TemporaryDirectory
from typing import Any, Dict, List, Optional

import matplotlib.pyplot as plt
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_AUTO_SIZE
from pptx.util import Inches, Pt

MESES = {
    1: "Janeiro", 2: "Fevereiro", 3: "Março", 4: "Abril",
    5: "Maio", 6: "Junho", 7: "Julho", 8: "Agosto",
    9: "Setembro", 10: "Outubro", 11: "Novembro", 12: "Dezembro",
}

BLUE = RGBColor(25, 53, 91)
GOLD = RGBColor(201, 151, 64)
LIGHT = RGBColor(246, 248, 252)
TEXT = RGBColor(38, 38, 38)
MUTED = RGBColor(92, 101, 115)
WHITE = RGBColor(255, 255, 255)

@dataclass
class ReportInputs:
    mes: int
    ano: int
    tickets_abertos: int = 0
    tickets_fechados: int = 0
    imagem_tickets_mensal: Optional[bytes] = None
    reducao_fevereiro: str = ""
    fator_contexto_mes: str = ""
    comentario_tickets_mensal: str = ""
    rodape_tickets_mensal: str = ""

    tickets_diarios: List[Dict[str, Any]] = field(default_factory=list)
    imagem_tickets_diarios: Optional[bytes] = None
    analise_periodo: str = ""
    queda_dias: str = ""
    retomada_aumento_dias: str = ""
    observacao_diarios: str = ""

    imagem_top_analistas: Optional[bytes] = None
    top_analistas: List[Dict[str, Any]] = field(default_factory=list)
    destaque_analistas: str = ""
    observacao_analistas: str = ""

    imagem_top_categorias: Optional[bytes] = None
    top_categorias: List[Dict[str, Any]] = field(default_factory=list)
    destaque_categorias: str = ""
    chamados_destaque: str = ""

    imagem_teiti: Optional[bytes] = None
    comentario_teiti: str = ""
    imagem_infotech: Optional[bytes] = None
    comentario_infotech: str = ""

    tel_img_af: Optional[bytes] = None
    tel_valor_af: float = 0.0
    tel_img_affit: Optional[bytes] = None
    tel_valor_affit: float = 0.0
    tel_img_afsw: Optional[bytes] = None
    tel_valor_afsw: float = 0.0

    sharepoint_img: Optional[bytes] = None
    sharepoint_comentario: str = ""
    capacidade_total: str = ""
    espaco_livre: str = ""
    dias_restantes: str = ""
    crescimento: str = ""
    custo_por_gb: str = ""
    valor_sharepoint: str = ""

    inventario_df: List[Dict[str, Any]] = field(default_factory=list)
    inventario_comentario: str = ""

    seguranca_img: Optional[bytes] = None
    sumario_seguranca: str = ""
    endpoints_gerenciados: str = ""
    endpoints_ativos: str = ""
    ameacas_bloqueadas: str = ""
    risco_empresa: str = ""

    licencas: List[Dict[str, Any]] = field(default_factory=list)

    o365_img_af: Optional[bytes] = None
    o365_img_affit: Optional[bytes] = None
    o365_img_afsw: Optional[bytes] = None
    o365_observacoes: str = ""

    sol_texto_af_1: str = ""
    sol_img_af: Optional[bytes] = None
    sol_texto_af_2: str = ""
    sol_texto_affit_1: str = ""
    sol_img_affit: Optional[bytes] = None
    sol_texto_affit_2: str = ""
    sol_texto_afsw_1: str = ""
    sol_img_afsw: Optional[bytes] = None
    sol_texto_afsw_2: str = ""


def _safe_int(v) -> int:
    try:
        if v is None or v == "":
            return 0
        return int(float(str(v).replace(",", ".")))
    except Exception:
        return 0


def _safe_float(v) -> float:
    try:
        if v is None or v == "":
            return 0.0
        return float(str(v).replace("R$", "").replace(".", "").replace(",", ".").strip())
    except Exception:
        return 0.0


def _money(v: float) -> str:
    return f"R$ {v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def _clear_slide(slide) -> None:
    spTree = slide.shapes._spTree
    for shp in list(slide.shapes):
        spTree.remove(shp._element)


def _ensure_slides(prs: Presentation, qty: int):
    blank = prs.slide_layouts[6] if len(prs.slide_layouts) > 6 else prs.slide_layouts[0]
    while len(prs.slides) < qty:
        prs.slides.add_slide(blank)


def _background(slide):
    bg = slide.background
    fill = bg.fill
    fill.solid()
    fill.fore_color.rgb = RGBColor(250, 250, 250)


def _header(slide, title: str, subtitle: str = ""):
    x, y, w, h = Inches(0), Inches(0), Inches(13.333), Inches(0.75)
    rect = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, w, h)
    rect.fill.solid(); rect.fill.fore_color.rgb = BLUE
    rect.line.fill.background()
    box = slide.shapes.add_textbox(Inches(0.45), Inches(0.16), Inches(9.8), Inches(0.45))
    p = box.text_frame.paragraphs[0]
    p.text = title
    p.font.name = "Nunito Sans"
    p.font.bold = True
    p.font.size = Pt(22)
    p.font.color.rgb = WHITE
    if subtitle:
        b = slide.shapes.add_textbox(Inches(10.4), Inches(0.2), Inches(2.5), Inches(0.35))
        bp = b.text_frame.paragraphs[0]
        bp.text = subtitle
        bp.font.name = "Nunito Sans"
        bp.font.bold = True
        bp.font.size = Pt(12)
        bp.font.color.rgb = WHITE
        bp.alignment = PP_ALIGN.RIGHT


def _footer(slide, text: str = "Relatório Mensal de Tecnologia da Informação"):
    line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(7.23), Inches(13.333), Inches(0.03))
    line.fill.solid(); line.fill.fore_color.rgb = GOLD; line.line.fill.background()
    box = slide.shapes.add_textbox(Inches(0.45), Inches(7.28), Inches(12.4), Inches(0.22))
    p = box.text_frame.paragraphs[0]
    p.text = text
    p.font.name = "Nunito Sans"
    p.font.size = Pt(8.5)
    p.font.color.rgb = MUTED


def _panel(slide, x, y, w, h, fill=WHITE, line=RGBColor(220,220,220)):
    sh = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, h)
    sh.fill.solid(); sh.fill.fore_color.rgb = fill
    sh.line.color.rgb = line
    return sh


def _textbox(slide, text: str, x, y, w, h, size=13, bold=False, color=TEXT, align=None, name="Nunito Sans"):
    box = slide.shapes.add_textbox(x, y, w, h)
    tf = box.text_frame
    tf.word_wrap = True
    tf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
    p = tf.paragraphs[0]
    p.text = str(text or "")
    p.font.name = name
    p.font.size = Pt(size)
    p.font.bold = bold
    p.font.color.rgb = color
    if align is not None:
        p.alignment = align
    return box


def _label_value(slide, label, value, x, y, w, h, value_size=24):
    _panel(slide, x, y, w, h, fill=LIGHT)
    _textbox(slide, label, x+Inches(0.12), y+Inches(0.08), w-Inches(0.24), Inches(0.22), size=8.5, bold=True, color=MUTED)
    _textbox(slide, value, x+Inches(0.12), y+Inches(0.28), w-Inches(0.24), h-Inches(0.35), size=value_size, bold=True, color=BLUE, align=PP_ALIGN.CENTER)


def _add_image(slide, img_bytes: Optional[bytes], x, y, w, h, label="Imagem"):
    _panel(slide, x, y, w, h, fill=RGBColor(245,247,250))
    if img_bytes:
        try:
            slide.shapes.add_picture(BytesIO(img_bytes), x+Inches(0.05), y+Inches(0.05), width=w-Inches(0.1), height=h-Inches(0.1))
            return
        except Exception:
            pass
    _textbox(slide, label, x+Inches(0.15), y+h/2-Inches(0.1), w-Inches(0.3), Inches(0.3), size=12, bold=True, color=MUTED, align=PP_ALIGN.CENTER)


def _save_bar_chart(labels, values, title, out_path: Path, horizontal=False):
    plt.figure(figsize=(8.5, 4.4))
    if horizontal:
        plt.barh(labels, values)
        plt.xlabel("Quantidade")
    else:
        plt.bar(labels, values)
        plt.ylabel("Quantidade")
        plt.xticks(rotation=20, ha="right")
    plt.title(title)
    plt.tight_layout()
    plt.savefig(out_path, dpi=180, transparent=True)
    plt.close()


def _save_compare_chart(abertos, fechados, out_path: Path):
    plt.figure(figsize=(6.4, 3.7))
    plt.bar(["Abertos", "Fechados"], [abertos, fechados])
    plt.title("Tickets abertos VS fechados")
    plt.ylabel("Quantidade")
    plt.tight_layout()
    plt.savefig(out_path, dpi=180, transparent=True)
    plt.close()


def _save_daily_chart(rows, out_path: Path):
    labels = [str(r.get("Dia", "")) for r in rows]
    abertos = [_safe_int(r.get("Abertos", 0)) for r in rows]
    fechados = [_safe_int(r.get("Fechados", 0)) for r in rows]
    plt.figure(figsize=(8.5, 4.2))
    plt.plot(labels, abertos, marker="o", label="Abertos")
    plt.plot(labels, fechados, marker="o", label="Fechados")
    plt.title("Abertos e fechados diários")
    plt.ylabel("Quantidade")
    plt.xticks(rotation=25, ha="right")
    plt.grid(True, alpha=0.25)
    plt.legend()
    plt.tight_layout()
    plt.savefig(out_path, dpi=180, transparent=True)
    plt.close()


def _bullets(slide, text: str, x, y, w, h, size=12):
    items = [i.strip() for i in str(text or "").splitlines() if i.strip()]
    if not items:
        items = [""]
    box = slide.shapes.add_textbox(x, y, w, h)
    tf = box.text_frame
    tf.clear(); tf.word_wrap = True
    for idx, item in enumerate(items):
        p = tf.paragraphs[0] if idx == 0 else tf.add_paragraph()
        p.text = item
        p.level = 0
        p.font.name = "Nunito Sans"
        p.font.size = Pt(size)
        p.font.color.rgb = TEXT
    return box


def _inventory_summary(rows: List[Dict[str, Any]]) -> List[str]:
    if not rows:
        return ["Nenhuma planilha anexada."]
    total = len(rows)
    cols = list(rows[0].keys()) if rows else []
    result = [f"Total de registros analisados: {total}", f"Colunas identificadas: {', '.join(map(str, cols[:8]))}"]
    # try common columns
    common_groups = ["Tipo", "Categoria", "Status", "Situação", "Unidade", "Empresa", "Local", "Modelo"]
    for col in common_groups:
        match = next((c for c in cols if str(c).strip().lower() == col.lower()), None)
        if match:
            counts = {}
            for r in rows:
                key = str(r.get(match, "")).strip() or "Não informado"
                counts[key] = counts.get(key, 0) + 1
            top = sorted(counts.items(), key=lambda kv: kv[1], reverse=True)[:5]
            result.append(f"{match}: " + "; ".join(f"{k} ({v})" for k, v in top))
    return result[:7]


def _month_label(data: ReportInputs) -> str:
    return f"{MESES.get(int(data.mes), str(data.mes))} / {data.ano}"


def _prep_slide(prs, index: int, title: str, data: ReportInputs):
    slide = prs.slides[index]
    _clear_slide(slide)
    _background(slide)
    _header(slide, title, _month_label(data))
    _footer(slide)
    return slide


def build_pptx(template_path: str | Path, data: ReportInputs) -> bytes:
    prs = Presentation(str(template_path))
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    _ensure_slides(prs, 13)

    with TemporaryDirectory() as tmp:
        tmp_path = Path(tmp)

        # 1 Tickets mensais
        s = _prep_slide(prs, 0, "Comparativo mensal de tickets: abertos VS fechados", data)
        chart = tmp_path / "tickets_mensal.png"
        _save_compare_chart(data.tickets_abertos, data.tickets_fechados, chart)
        s.shapes.add_picture(str(chart), Inches(0.55), Inches(1.0), width=Inches(5.6), height=Inches(3.25))
        _add_image(s, data.imagem_tickets_mensal, Inches(6.35), Inches(1.0), Inches(6.45), Inches(3.25), "Imagem do mês")
        _label_value(s, "Abertos", str(data.tickets_abertos), Inches(0.65), Inches(4.45), Inches(1.8), Inches(0.85))
        _label_value(s, "Fechados", str(data.tickets_fechados), Inches(2.65), Inches(4.45), Inches(1.8), Inches(0.85))
        _panel(s, Inches(4.65), Inches(4.45), Inches(3.9), Inches(1.4))
        _textbox(s, "Redução de fevereiro", Inches(4.82), Inches(4.57), Inches(3.5), Inches(0.22), size=9, bold=True, color=BLUE)
        _textbox(s, data.reducao_fevereiro, Inches(4.82), Inches(4.85), Inches(3.5), Inches(0.85), size=11)
        _panel(s, Inches(8.75), Inches(4.45), Inches(4.05), Inches(1.4))
        _textbox(s, "Fator de contexto do mês", Inches(8.92), Inches(4.57), Inches(3.7), Inches(0.22), size=9, bold=True, color=BLUE)
        _textbox(s, data.fator_contexto_mes, Inches(8.92), Inches(4.85), Inches(3.7), Inches(0.85), size=11)
        _panel(s, Inches(0.65), Inches(5.95), Inches(12.15), Inches(0.95))
        _textbox(s, data.comentario_tickets_mensal or data.rodape_tickets_mensal, Inches(0.85), Inches(6.1), Inches(11.75), Inches(0.6), size=12)

        # 2 Diarios
        s = _prep_slide(prs, 1, "Abertos e fechados diários - Análise do período", data)
        chart = tmp_path / "tickets_diarios.png"
        _save_daily_chart(data.tickets_diarios, chart)
        s.shapes.add_picture(str(chart), Inches(0.55), Inches(1.0), width=Inches(6.5), height=Inches(3.45))
        _add_image(s, data.imagem_tickets_diarios, Inches(7.25), Inches(1.0), Inches(5.55), Inches(3.45), "Imagem do período")
        labels = [("Análise do período", data.analise_periodo), ("Queda nos dias", data.queda_dias), ("Retomada de aumento nos dias", data.retomada_aumento_dias), ("Observação", data.observacao_diarios)]
        x_positions = [0.55, 3.75, 6.95, 10.15]
        for (label, text), x in zip(labels, x_positions):
            _panel(s, Inches(x), Inches(4.68), Inches(3.0), Inches(2.12))
            _textbox(s, label, Inches(x+0.15), Inches(4.82), Inches(2.7), Inches(0.25), size=9, bold=True, color=BLUE)
            _textbox(s, text, Inches(x+0.15), Inches(5.12), Inches(2.7), Inches(1.45), size=10.5)

        # 3 Top analistas
        s = _prep_slide(prs, 2, "Top analistas com mais chamados", data)
        analistas = sorted(data.top_analistas, key=lambda r: _safe_int(r.get("Quantidade", 0)), reverse=True)[:10]
        chart = tmp_path / "top_analistas.png"
        _save_bar_chart([str(r.get("Analista", "")) for r in analistas], [_safe_int(r.get("Quantidade", 0)) for r in analistas], "Top analistas", chart, horizontal=True)
        s.shapes.add_picture(str(chart), Inches(0.55), Inches(1.0), width=Inches(6.2), height=Inches(3.7))
        _add_image(s, data.imagem_top_analistas, Inches(6.95), Inches(1.0), Inches(5.85), Inches(2.65), "Imagem")
        _panel(s, Inches(6.95), Inches(3.88), Inches(5.85), Inches(1.15))
        _textbox(s, "Destaque dos analistas", Inches(7.12), Inches(4.02), Inches(5.5), Inches(0.25), size=9, bold=True, color=BLUE)
        _textbox(s, data.destaque_analistas, Inches(7.12), Inches(4.32), Inches(5.5), Inches(0.5), size=11)
        _panel(s, Inches(0.55), Inches(5.18), Inches(12.25), Inches(1.45))
        _textbox(s, "Observação / comentário", Inches(0.75), Inches(5.35), Inches(11.85), Inches(0.25), size=10, bold=True, color=BLUE)
        _textbox(s, data.observacao_analistas, Inches(0.75), Inches(5.65), Inches(11.85), Inches(0.65), size=11)

        # 4 Top categorias
        s = _prep_slide(prs, 3, "Top 10 categorias de chamados", data)
        cats = sorted(data.top_categorias, key=lambda r: _safe_int(r.get("Quantidade", 0)), reverse=True)[:10]
        chart = tmp_path / "top_categorias.png"
        _save_bar_chart([str(r.get("Categoria", "")) for r in cats], [_safe_int(r.get("Quantidade", 0)) for r in cats], "Top 10 categorias", chart, horizontal=True)
        s.shapes.add_picture(str(chart), Inches(0.55), Inches(1.0), width=Inches(6.2), height=Inches(3.7))
        _add_image(s, data.imagem_top_categorias, Inches(6.95), Inches(1.0), Inches(5.85), Inches(2.55), "Imagem")
        _panel(s, Inches(6.95), Inches(3.75), Inches(5.85), Inches(1.0))
        _textbox(s, "Destaque", Inches(7.12), Inches(3.88), Inches(5.5), Inches(0.25), size=9, bold=True, color=BLUE)
        _textbox(s, data.destaque_categorias, Inches(7.12), Inches(4.15), Inches(5.5), Inches(0.45), size=11)
        _panel(s, Inches(0.55), Inches(4.95), Inches(12.25), Inches(1.85))
        _textbox(s, "Chamados em destaque", Inches(0.75), Inches(5.1), Inches(11.85), Inches(0.25), size=10, bold=True, color=BLUE)
        _bullets(s, data.chamados_destaque, Inches(0.75), Inches(5.42), Inches(11.85), Inches(1.15), size=11)

        # 4 Teiti
        s = _prep_slide(prs, 4, "Notebooks e monitores alugados Teiti", data)
        _add_image(s, data.imagem_teiti, Inches(0.75), Inches(1.05), Inches(12.0), Inches(4.85), "Imagem Teiti")
        _panel(s, Inches(0.75), Inches(6.05), Inches(12.0), Inches(0.8))
        _textbox(s, data.comentario_teiti, Inches(0.95), Inches(6.22), Inches(11.6), Inches(0.42), size=12)

        # 5 Infotech
        s = _prep_slide(prs, 5, "Notebooks e monitores alugados Infotech", data)
        _add_image(s, data.imagem_infotech, Inches(0.75), Inches(1.05), Inches(12.0), Inches(4.85), "Imagem Infotech")
        _panel(s, Inches(0.75), Inches(6.05), Inches(12.0), Inches(0.8))
        _textbox(s, data.comentario_infotech, Inches(0.95), Inches(6.22), Inches(11.6), Inches(0.42), size=12)

        # 6 Telefonia
        s = _prep_slide(prs, 6, "Telefonia Móvel", data)
        blocks = [("Afonso França", data.tel_img_af, data.tel_valor_af), ("AFFIT", data.tel_img_affit, data.tel_valor_affit), ("AFSW", data.tel_img_afsw, data.tel_valor_afsw)]
        for i, (empresa, img, valor) in enumerate(blocks):
            x = Inches(0.65 + i*4.25)
            _textbox(s, empresa, x, Inches(1.0), Inches(3.75), Inches(0.35), size=15, bold=True, color=BLUE, align=PP_ALIGN.CENTER)
            _add_image(s, img, x, Inches(1.45), Inches(3.75), Inches(3.4), f"Imagem {empresa}")
            _label_value(s, "Valor mensal", _money(valor), x, Inches(5.1), Inches(3.75), Inches(1.05), value_size=18)

        # 7 SharePoint
        s = _prep_slide(prs, 7, "SharePoint", data)
        _add_image(s, data.sharepoint_img, Inches(0.55), Inches(1.0), Inches(5.9), Inches(3.8), "Imagem SharePoint")
        _panel(s, Inches(6.65), Inches(1.0), Inches(6.15), Inches(1.15))
        _textbox(s, data.sharepoint_comentario, Inches(6.85), Inches(1.2), Inches(5.75), Inches(0.7), size=12)
        _label_value(s, "Capacidade total", data.capacidade_total, Inches(6.65), Inches(2.35), Inches(1.95), Inches(1.35), value_size=33.5)
        _label_value(s, "Espaço livre", data.espaco_livre, Inches(8.75), Inches(2.35), Inches(1.95), Inches(1.35), value_size=33.5)
        _label_value(s, "Dias restantes", data.dias_restantes, Inches(10.85), Inches(2.35), Inches(1.95), Inches(1.35), value_size=33.5)
        _label_value(s, "Crescimento", data.crescimento, Inches(6.65), Inches(4.05), Inches(1.95), Inches(0.95), value_size=12.5)
        _label_value(s, "Custo por GB", data.custo_por_gb, Inches(8.75), Inches(4.05), Inches(1.95), Inches(0.95), value_size=12.5)
        _label_value(s, "Valor R$", data.valor_sharepoint, Inches(10.85), Inches(4.05), Inches(1.95), Inches(0.95), value_size=12.5)
        _panel(s, Inches(0.55), Inches(5.15), Inches(12.25), Inches(1.35))
        _textbox(s, "Indicadores SharePoint", Inches(0.75), Inches(5.35), Inches(11.85), Inches(0.25), size=11, bold=True, color=BLUE)

        # 8 Inventario
        s = _prep_slide(prs, 8, "Inventário", data)
        _panel(s, Inches(0.65), Inches(1.05), Inches(12.0), Inches(5.65))
        summary = _inventory_summary(data.inventario_df)
        _textbox(s, "Análise automática da planilha", Inches(0.95), Inches(1.3), Inches(11.4), Inches(0.35), size=16, bold=True, color=BLUE)
        _bullets(s, "\n".join(summary), Inches(0.95), Inches(1.85), Inches(11.4), Inches(2.4), size=13)
        _textbox(s, "Comentário", Inches(0.95), Inches(4.65), Inches(11.4), Inches(0.25), size=11, bold=True, color=BLUE)
        _textbox(s, data.inventario_comentario, Inches(0.95), Inches(4.95), Inches(11.4), Inches(1.2), size=12)

        # 9 Segurança
        s = _prep_slide(prs, 9, "Sumário executivo de segurança", data)
        _add_image(s, data.seguranca_img, Inches(0.65), Inches(1.05), Inches(6.1), Inches(3.55), "Imagem de segurança")
        _panel(s, Inches(6.95), Inches(1.05), Inches(5.85), Inches(3.55))
        _textbox(s, "Sumário executivo de segurança", Inches(7.15), Inches(1.25), Inches(5.45), Inches(0.3), size=12.5, bold=True, color=BLUE)
        _textbox(s, data.sumario_seguranca, Inches(7.15), Inches(1.65), Inches(5.45), Inches(2.55), size=11.5)
        kpis = [("Endpoints gerenciados", data.endpoints_gerenciados), ("Endpoints ativos", data.endpoints_ativos), ("Ameaças bloqueadas", data.ameacas_bloqueadas), ("Risco da empresa", data.risco_empresa)]
        for i, (label, value) in enumerate(kpis):
            _label_value(s, label, value, Inches(0.65 + i*3.05), Inches(4.95), Inches(2.75), Inches(1.05), value_size=19)

        # 10 Licenças
        s = _prep_slide(prs, 10, "Gestão de Licenças de Software - Afonso França", data)
        for i, empresa in enumerate(["Afonso França", "AFFIT", "AFSW"]):
            x = Inches(0.65 + i*4.25)
            _panel(s, x, Inches(1.05), Inches(3.75), Inches(5.55))
            _textbox(s, empresa, x+Inches(0.15), Inches(1.25), Inches(3.45), Inches(0.35), size=15, bold=True, color=BLUE, align=PP_ALIGN.CENTER)
            lines = []
            for r in data.licencas:
                if str(r.get("Empresa", "")).strip().lower() == empresa.lower():
                    lines.append(f"{r.get('Software','')} - {r.get('Vencimento','')}")
            _bullets(s, "\n".join(lines) if lines else "Sem licenças informadas.", x+Inches(0.25), Inches(1.8), Inches(3.25), Inches(4.3), size=11)

        # 11 Office 365
        s = _prep_slide(prs, 11, "Fatura VIVO Office 365 - Resumo consolidado", data)
        imgs = [("Afonso França", data.o365_img_af), ("AFFIT", data.o365_img_affit), ("AFSW", data.o365_img_afsw)]
        for i, (empresa, img) in enumerate(imgs):
            x = Inches(0.65 + i*4.25)
            _textbox(s, empresa, x, Inches(1.0), Inches(3.75), Inches(0.35), size=14, bold=True, color=BLUE, align=PP_ALIGN.CENTER)
            _add_image(s, img, x, Inches(1.45), Inches(3.75), Inches(3.55), f"Imagem {empresa}")
        _panel(s, Inches(0.65), Inches(5.35), Inches(12.0), Inches(1.15))
        _textbox(s, "Observações", Inches(0.85), Inches(5.52), Inches(11.6), Inches(0.25), size=10, bold=True, color=BLUE)
        _textbox(s, data.o365_observacoes, Inches(0.85), Inches(5.82), Inches(11.6), Inches(0.45), size=12)

        # 12 Soluções
        s = _prep_slide(prs, 12, "Análise de soluções", data)
        solutions = [
            ("Afonso França", data.sol_texto_af_1, data.sol_img_af, data.sol_texto_af_2),
            ("AFFIT", data.sol_texto_affit_1, data.sol_img_affit, data.sol_texto_affit_2),
            ("AFSW", data.sol_texto_afsw_1, data.sol_img_afsw, data.sol_texto_afsw_2),
        ]
        for i, (empresa, t1, img, t2) in enumerate(solutions):
            x = Inches(0.65 + i*4.25)
            _textbox(s, empresa, x, Inches(1.0), Inches(3.75), Inches(0.35), size=14, bold=True, color=BLUE, align=PP_ALIGN.CENTER)
            _panel(s, x, Inches(1.42), Inches(3.75), Inches(0.75))
            _textbox(s, t1, x+Inches(0.15), Inches(1.55), Inches(3.45), Inches(0.43), size=10)
            _add_image(s, img, x, Inches(2.35), Inches(3.75), Inches(2.35), f"Imagem {empresa}")
            _panel(s, x, Inches(4.95), Inches(3.75), Inches(1.25))
            _textbox(s, t2, x+Inches(0.15), Inches(5.12), Inches(3.45), Inches(0.82), size=10)

        output = BytesIO()
        prs.save(output)
        return output.getvalue()
