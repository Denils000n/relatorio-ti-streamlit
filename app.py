from __future__ import annotations

from datetime import date
from pathlib import Path

import pandas as pd
import streamlit as st

from src.report_generator import ReportInputs, build_pptx, MESES

BASE_DIR = Path(__file__).resolve().parent
DEFAULT_TEMPLATE = BASE_DIR / "assets" / "template_relatorio_ti.pptx"

st.set_page_config(page_title="Gerador de Relatório Mensal TI", layout="wide")
st.title("Gerador de Relatório Mensal TI")
st.caption("Preencha os dados do mês, anexe imagens/planilhas e gere o PPTX mensal no padrão do modelo.")


def uploaded_bytes(file):
    return file.getvalue() if file is not None else None


def money_input(label: str, value: float = 0.0):
    return st.number_input(label, min_value=0.0, value=float(value), step=100.0, format="%.2f")


def read_inventory(file):
    if file is None:
        return pd.DataFrame()
    name = file.name.lower()
    try:
        if name.endswith(".csv"):
            return pd.read_csv(file)
        return pd.read_excel(file)
    except Exception as exc:
        st.error(f"Não foi possível ler a planilha de inventário: {exc}")
        return pd.DataFrame()


with st.sidebar:
    st.header("Período")
    hoje = date.today()
    mes = st.selectbox("Mês", options=list(MESES.keys()), index=max(hoje.month - 1, 0), format_func=lambda x: MESES[x])
    ano = st.number_input("Ano", min_value=2020, max_value=2100, value=hoje.year, step=1)
    uploaded_template = st.file_uploader("Modelo PPTX opcional", type=["pptx"], help="Use apenas se quiser substituir o modelo padrão.")

abas = st.tabs([
    "1 Tickets Mensais",
    "2 Tickets Diários",
    "3 Top Analistas",
    "4 Top Categorias",
    "5 Teiti",
    "6 Infotech",
    "7 Telefonia",
    "8 SharePoint",
    "9 Inventário",
    "10 Segurança",
    "11 Licenças",
    "12 Office 365",
    "13 Soluções",
    "Gerar PPTX",
])

with abas[0]:
    st.header("Comparativo mensal de tickets: abertos VS fechados")
    c1, c2, c3 = st.columns(3)
    with c1:
        tickets_abertos = st.number_input("Chamados abertos", min_value=0, value=120, step=1)
    with c2:
        tickets_fechados = st.number_input("Chamados fechados", min_value=0, value=105, step=1)
    with c3:
        imagem_tickets_mensal = st.file_uploader("Imagem do slide", type=["png", "jpg", "jpeg"], key="img_tickets_mensal")
    reducao_fevereiro = st.text_area("Redução de fevereiro", height=80)
    fator_contexto_mes = st.text_area("Fator de contexto do mês", height=80)
    comentario_tickets_mensal = st.text_area("Comentário", height=90)
    rodape_tickets_mensal = st.text_area("Rodapé", height=68)

with abas[1]:
    st.header("Abertos e fechados diários - Análise do período")
    imagem_tickets_diarios = st.file_uploader("Imagem do slide", type=["png", "jpg", "jpeg"], key="img_tickets_diarios")
    dias_padrao = pd.DataFrame({"Dia": list(range(1, 8)), "Abertos": [12, 15, 9, 18, 20, 7, 11], "Fechados": [10, 13, 12, 16, 17, 9, 12]})
    tickets_diarios = st.data_editor(dias_padrao, num_rows="dynamic", use_container_width=True, key="tickets_diarios")
    analise_periodo = st.text_area("Análise do período", height=80)
    queda_dias = st.text_area("Queda nos dias", height=70)
    retomada_aumento_dias = st.text_area("Retomada de aumento nos dias", height=70)
    observacao_diarios = st.text_area("Observação", height=80)

with abas[2]:
    st.header("Top analistas com mais chamados")
    imagem_top_analistas = st.file_uploader("Imagem do slide", type=["png", "jpg", "jpeg"], key="img_top_analistas")
    top_analistas = st.data_editor(pd.DataFrame([
        {"Analista": "Analista 1", "Quantidade": 25},
        {"Analista": "Analista 2", "Quantidade": 18},
        {"Analista": "Analista 3", "Quantidade": 14},
    ]), num_rows="dynamic", use_container_width=True, key="top_analistas")
    destaque_analistas = st.text_area("Destaque dos analistas", height=80)
    observacao_analistas = st.text_area("Observação / comentário", height=100)

with abas[3]:
    st.header("Top 10 categorias de chamados")
    imagem_top_categorias = st.file_uploader("Imagem do slide", type=["png", "jpg", "jpeg"], key="img_top_categorias")
    top_categorias = st.data_editor(pd.DataFrame([
        {"Categoria": "Acessos", "Quantidade": 30},
        {"Categoria": "Hardware", "Quantidade": 24},
        {"Categoria": "Sistemas", "Quantidade": 20},
        {"Categoria": "Rede", "Quantidade": 14},
        {"Categoria": "E-mail", "Quantidade": 12},
        {"Categoria": "Impressoras", "Quantidade": 9},
        {"Categoria": "Telefonia", "Quantidade": 7},
        {"Categoria": "Office 365", "Quantidade": 6},
        {"Categoria": "Backup", "Quantidade": 5},
        {"Categoria": "Segurança", "Quantidade": 4},
    ]), num_rows="dynamic", use_container_width=True, key="top_categorias")
    destaque_categorias = st.text_area("Destaque das categorias", height=70)
    chamados_destaque = st.text_area("Chamados em destaque", height=130, help="Digite um item por linha.")

with abas[4]:
    st.header("Notebooks e monitores alugados Teiti")
    imagem_teiti = st.file_uploader("Imagem Teiti", type=["png", "jpg", "jpeg"], key="img_teiti")
    comentario_teiti = st.text_area("Comentário Teiti", height=90)

with abas[5]:
    st.header("Notebooks e monitores alugados Infotech")
    imagem_infotech = st.file_uploader("Imagem Infotech", type=["png", "jpg", "jpeg"], key="img_infotech")
    comentario_infotech = st.text_area("Comentário Infotech", height=90)

with abas[6]:
    st.header("Telefonia Móvel")
    c1, c2, c3 = st.columns(3)
    with c1:
        st.subheader("Afonso França")
        tel_img_af = st.file_uploader("Imagem Afonso França", type=["png", "jpg", "jpeg"], key="tel_af")
        tel_valor_af = money_input("Valor mensal Afonso França")
    with c2:
        st.subheader("AFFIT")
        tel_img_affit = st.file_uploader("Imagem AFFIT", type=["png", "jpg", "jpeg"], key="tel_affit")
        tel_valor_affit = money_input("Valor mensal AFFIT")
    with c3:
        st.subheader("AFSW")
        tel_img_afsw = st.file_uploader("Imagem AFSW", type=["png", "jpg", "jpeg"], key="tel_afsw")
        tel_valor_afsw = money_input("Valor mensal AFSW")

with abas[7]:
    st.header("SharePoint")
    sharepoint_img = st.file_uploader("Imagem SharePoint", type=["png", "jpg", "jpeg"], key="sharepoint_img")
    sharepoint_comentario = st.text_area("Comentário SharePoint", height=90)
    c1, c2, c3 = st.columns(3)
    capacidade_total = c1.text_input("Capacidade total", value="")
    espaco_livre = c2.text_input("Espaço livre", value="")
    dias_restantes = c3.text_input("Dias restantes", value="")
    c4, c5, c6 = st.columns(3)
    crescimento = c4.text_input("Crescimento", value="")
    custo_por_gb = c5.text_input("Custo por GB", value="")
    valor_sharepoint = c6.text_input("Valor R$", value="")

with abas[8]:
    st.header("Inventário")
    inventario_file = st.file_uploader("Anexar planilha de inventário", type=["xlsx", "xls", "csv"], key="inventario_file")
    inventario_df = read_inventory(inventario_file)
    if not inventario_df.empty:
        st.success(f"Planilha carregada: {inventario_df.shape[0]} linhas e {inventario_df.shape[1]} colunas.")
        st.dataframe(inventario_df.head(100), use_container_width=True)
    inventario_comentario = st.text_area("Comentário do inventário", height=100)

with abas[9]:
    st.header("Sumário executivo de segurança")
    seguranca_img = st.file_uploader("Imagem de segurança", type=["png", "jpg", "jpeg"], key="seg_img")
    sumario_seguranca = st.text_area("Sumário executivo de segurança", height=100)
    c1, c2, c3, c4 = st.columns(4)
    endpoints_gerenciados = c1.text_input("Endpoints gerenciados", value="")
    endpoints_ativos = c2.text_input("Endpoints ativos", value="")
    ameacas_bloqueadas = c3.text_input("Ameaças bloqueadas", value="")
    risco_empresa = c4.text_input("Risco da empresa", value="")

with abas[10]:
    st.header("Gestão de Licenças de Software - Afonso França")
    licencas = st.data_editor(pd.DataFrame([
        {"Empresa": "Afonso França", "Software": "", "Vencimento": ""},
        {"Empresa": "AFFIT", "Software": "", "Vencimento": ""},
        {"Empresa": "AFSW", "Software": "", "Vencimento": ""},
    ]), num_rows="dynamic", use_container_width=True, key="licencas")

with abas[11]:
    st.header("Fatura VIVO Office 365 - Resumo consolidado")
    c1, c2, c3 = st.columns(3)
    with c1:
        o365_img_af = st.file_uploader("Imagem Afonso França", type=["png", "jpg", "jpeg"], key="o365_af")
    with c2:
        o365_img_affit = st.file_uploader("Imagem AFFIT", type=["png", "jpg", "jpeg"], key="o365_affit")
    with c3:
        o365_img_afsw = st.file_uploader("Imagem AFSW", type=["png", "jpg", "jpeg"], key="o365_afsw")
    o365_observacoes = st.text_area("Observações Office 365", height=100)

with abas[12]:
    st.header("Análise de soluções")
    st.subheader("Afonso França")
    sol_texto_af_1 = st.text_area("Texto inicial Afonso França", height=70)
    sol_img_af = st.file_uploader("Imagem Afonso França", type=["png", "jpg", "jpeg"], key="sol_af")
    sol_texto_af_2 = st.text_area("Texto final Afonso França", height=70)
    st.subheader("AFFIT")
    sol_texto_affit_1 = st.text_area("Texto inicial AFFIT", height=70)
    sol_img_affit = st.file_uploader("Imagem AFFIT", type=["png", "jpg", "jpeg"], key="sol_affit")
    sol_texto_affit_2 = st.text_area("Texto final AFFIT", height=70)
    st.subheader("AFSW")
    sol_texto_afsw_1 = st.text_area("Texto inicial AFSW", height=70)
    sol_img_afsw = st.file_uploader("Imagem AFSW", type=["png", "jpg", "jpeg"], key="sol_afsw")
    sol_texto_afsw_2 = st.text_area("Texto final AFSW", height=70)

with abas[13]:
    st.header("Gerar apresentação")
    st.write("Quando todos os campos estiverem preenchidos, clique para gerar o arquivo PowerPoint.")
    if st.button("Gerar PPTX", type="primary"):
        template_path = DEFAULT_TEMPLATE
        temp_template = None
        if uploaded_template is not None:
            temp_template = BASE_DIR / "template_upload_temp.pptx"
            temp_template.write_bytes(uploaded_template.getvalue())
            template_path = temp_template

        data = ReportInputs(
            mes=int(mes), ano=int(ano),
            tickets_abertos=int(tickets_abertos), tickets_fechados=int(tickets_fechados),
            imagem_tickets_mensal=uploaded_bytes(imagem_tickets_mensal),
            reducao_fevereiro=reducao_fevereiro, fator_contexto_mes=fator_contexto_mes,
            comentario_tickets_mensal=comentario_tickets_mensal, rodape_tickets_mensal=rodape_tickets_mensal,
            tickets_diarios=tickets_diarios.fillna("").to_dict("records"),
            imagem_tickets_diarios=uploaded_bytes(imagem_tickets_diarios),
            analise_periodo=analise_periodo, queda_dias=queda_dias,
            retomada_aumento_dias=retomada_aumento_dias, observacao_diarios=observacao_diarios,
            imagem_top_analistas=uploaded_bytes(imagem_top_analistas),
            top_analistas=top_analistas.fillna("").to_dict("records"),
            destaque_analistas=destaque_analistas, observacao_analistas=observacao_analistas,
            imagem_top_categorias=uploaded_bytes(imagem_top_categorias),
            top_categorias=top_categorias.fillna("").to_dict("records"),
            destaque_categorias=destaque_categorias, chamados_destaque=chamados_destaque,
            imagem_teiti=uploaded_bytes(imagem_teiti), comentario_teiti=comentario_teiti,
            imagem_infotech=uploaded_bytes(imagem_infotech), comentario_infotech=comentario_infotech,
            tel_img_af=uploaded_bytes(tel_img_af), tel_valor_af=float(tel_valor_af),
            tel_img_affit=uploaded_bytes(tel_img_affit), tel_valor_affit=float(tel_valor_affit),
            tel_img_afsw=uploaded_bytes(tel_img_afsw), tel_valor_afsw=float(tel_valor_afsw),
            sharepoint_img=uploaded_bytes(sharepoint_img), sharepoint_comentario=sharepoint_comentario,
            capacidade_total=capacidade_total, espaco_livre=espaco_livre, dias_restantes=dias_restantes,
            crescimento=crescimento, custo_por_gb=custo_por_gb, valor_sharepoint=valor_sharepoint,
            inventario_df=inventario_df.fillna("").to_dict("records"), inventario_comentario=inventario_comentario,
            seguranca_img=uploaded_bytes(seguranca_img), sumario_seguranca=sumario_seguranca,
            endpoints_gerenciados=endpoints_gerenciados, endpoints_ativos=endpoints_ativos,
            ameacas_bloqueadas=ameacas_bloqueadas, risco_empresa=risco_empresa,
            licencas=licencas.fillna("").to_dict("records"),
            o365_img_af=uploaded_bytes(o365_img_af), o365_img_affit=uploaded_bytes(o365_img_affit), o365_img_afsw=uploaded_bytes(o365_img_afsw),
            o365_observacoes=o365_observacoes,
            sol_texto_af_1=sol_texto_af_1, sol_img_af=uploaded_bytes(sol_img_af), sol_texto_af_2=sol_texto_af_2,
            sol_texto_affit_1=sol_texto_affit_1, sol_img_affit=uploaded_bytes(sol_img_affit), sol_texto_affit_2=sol_texto_affit_2,
            sol_texto_afsw_1=sol_texto_afsw_1, sol_img_afsw=uploaded_bytes(sol_img_afsw), sol_texto_afsw_2=sol_texto_afsw_2,
        )
        pptx_bytes = build_pptx(template_path, data)
        nome = f"Relatorio_Mensal_TI_{MESES[int(mes)]}_{int(ano)}.pptx"
        st.success("PPTX gerado com sucesso.")
        st.download_button("Baixar PPTX", data=pptx_bytes, file_name=nome, mime="application/vnd.openxmlformats-officedocument.presentationml.presentation")
        if temp_template and temp_template.exists():
            temp_template.unlink()
