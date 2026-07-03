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
 streamlit as st
import pandas as pd
import sqlite3
import os
import io
import calendar
import unicodedata
import requests
import json
from datetime import date, datetime

# ============================================================
# CONFIGURACOES & CONSTANTES
# ============================================================

st.set_page_config(
    page_title="Gerenciador de Licencas",
    page_icon="=",
    layout="wide",
    initial_sidebar_state="expanded"
)

DB_PATH = os.environ.get("LICENCAS_DB_PATH",
    os.path.join(os.path.dirname(os.path.abspath(__file__)), "licencas.db"))

CREDS_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), ".ms365_creds.json")

MESES_PT = ["","Janeiro","Fevereiro","Marco","Abril","Maio","Junho",
            "Julho","Agosto","Setembro","Outubro","Novembro","Dezembro"]

EMPRESA_PREFIXOS = {"01":"Afonso Franca","02":"AFFIT","03":"AFDI","04":"AFSW"}

EMPRESA_MAP = {
    "AF":"Afonso Franca","af":"Afonso Franca",
    "Afonso Franca":"Afonso Franca","Afonso Franca":"Afonso Franca",
    "AFFIT":"AFFIT","Affit":"AFFIT",
    "AFDI":"AFDI","AFSW":"AFSW",
}

COLUNAS_SISTEMA = ["Empresa","Colaborador","Centro de Custo",
                   "Tipo de Licenca","Valor da Licenca","Vencimento","Status"]

SINONIMOS = {
    "Empresa":         ["empresa","company","companhia","entidade","filial","unidade"],
    "Colaborador":     ["colaborador","funcionario","employee","nome","name",
                        "usuario","user","login","upn","username"],
    "Centro de Custo": ["centro de custo","cc","cost center","departamento","setor","area","depto","centrocusto"],
    "Tipo de Licenca": ["tipo de licenca","tipo licenca","tipo","licenca",
                        "license","produto","product","software","sku","aplicacao"],
    "Valor da Licenca":["valor da licenca","valor licenca","valor","value",
                        "preco","price","custo","cost","valor por licenca"],
    "Vencimento":      ["vencimento","expiration","expiry","data vencimento",
                        "data de vencimento","validade","valid until","expires"],
    "Status":          ["status","situacao","state","estado"],
}

STATUS_VALIDOS = ["Pendente","Em andamento","Renovada","Cancelada"]

ALERTA_OPCOES = {
    "30 dias":30,"60 dias":60,"90 dias":90,
    "6 meses":180,"12 meses":365,"18 meses":540,
    "24 meses":730,"36 meses":1095,"48 meses":1460,"60 meses":1825,
}

COR_ALERTA = {
    "Vencida":"#FF5252","Critica":"#FF9800",
    "Atencao":"#FFC107","Ok":"#4CAF50","Sem data":"#9E9E9E",
}

# Nomes amigaveis para SKUs Microsoft 365
MS365_SKU_NAMES = {
    "ENTERPRISEPACK":               "Office 365 E3",
    "ENTERPRISEPREMIUM":            "Office 365 E5",
    "STANDARDPACK":                 "Office 365 E1",
    "DESKLESSPACK":                 "Office 365 F3",
    "O365_BUSINESS_PREMIUM":        "Microsoft 365 Business Premium",
    "SPB":                          "Microsoft 365 Business Premium",
    "SMB_BUSINESS":                 "Microsoft 365 Business Basic",
    "O365_BUSINESS":                "Microsoft 365 Apps for Business",
    "OFFICESUBSCRIPTION":           "Microsoft 365 Apps for Enterprise",
    "TEAMS_EXPLORATORY":            "Microsoft Teams Exploratory",
    "TEAMS1":                       "Microsoft Teams Enterprise",
    "MCOSTANDARD":                  "Skype for Business Online (Plan 2)",
    "POWER_BI_STANDARD":            "Power BI (Free)",
    "POWER_BI_PRO":                 "Power BI Pro",
    "POWER_BI_PREMIUM_P1":          "Power BI Premium P1",
    "PBI_PREMIUM_PER_USER":         "Power BI Premium Per User",
    "FLOW_FREE":                    "Power Automate Free",
    "POWERFLOW_P1":                 "Power Automate Plan 1",
    "POWERFLOW_P2":                 "Power Automate Plan 2",
    "POWERAPPS_VIRAL":              "Power Apps Developer Plan",
    "POWERAPPS_DEV":                "Power Apps Developer Plan",
    "PROJECTPROFESSIONAL":          "Project Online Professional",
    "PROJECTPREMIUM":               "Project Online Premium",
    "PROJECT_PLAN1":                "Project Plan 1",
    "PROJECT_PLAN3":                "Project Plan 3",
    "VISIOCLIENT":                  "Visio Plan 2",
    "VISIOONLINE_PLAN1":            "Visio Plan 1",
    "WINDOWS_STORE":                "Windows Store",
    "WIN10_PRO_ENT_SUB":            "Windows 10/11 Enterprise E3",
    "WIN_DEF_ATP":                  "Microsoft Defender for Endpoint",
    "INTUNE_A":                     "Microsoft Intune",
    "EMS":                          "Enterprise Mobility + Security E3",
    "EMSPREMIUM":                   "Enterprise Mobility + Security E5",
    "AAD_PREMIUM":                  "Azure Active Directory Premium P1",
    "AAD_PREMIUM_P2":               "Azure Active Directory Premium P2",
    "RIGHTSMANAGEMENT":             "Azure Information Protection Plan 1",
    "STREAM":                       "Microsoft Stream",
    "ONEDRIVE_BASIC":               "OneDrive for Business (Basic)",
    "SHAREPOINT_S_DEVELOPER":       "SharePoint Online (Plan 1)",
    "SHAREPOINTENTERPRISE":         "SharePoint Online (Plan 2)",
    "EXCHANGESTANDARD":             "Exchange Online (Plan 1)",
    "EXCHANGEENTERPRISE":           "Exchange Online (Plan 2)",
    "MCOMEETADV":                   "Microsoft 365 Audio Conferencing",
    "PHONESYSTEM_VIRTUALUSER":      "Microsoft Teams Phone Resource Account",
    "MCOEV":                        "Microsoft Teams Phone Standard",
    "TEAMS_PREMIUM":                "Microsoft Teams Premium",
    "COPILOT_STUDIO_VIRAL_TRIAL":   "Power Virtual Agents (Trial)",
    "Microsoft_365_Copilot":        "Microsoft 365 Copilot",
    "DYN365_ENTERPRISE_SALES":      "Dynamics 365 Sales Premium",
    "DYN365_ENTERPRISE_P1":         "Dynamics 365 Customer Engagement Plan",
    "Dynamics_365_Sales_Premium":   "Dynamics 365 Sales Premium",
    "CRM_ONLINE_ENTERPRISE":        "Dynamics 365 Customer Voice",
    "FORMS_PRO":                    "Dynamics 365 Customer Voice",
    "ENTERPRISEPACK_B_PILOT":       "Office 365 E3 (sem Teams)",
    "STANDARDPACK_STUDENT":         "Office 365 A1 for Students",
    "POWER_BI_ADDON":               "Power BI for Office 365 Add-On",
}

# ============================================================
# BANCO DE DADOS
# ============================================================

def get_conn():
    return sqlite3.connect(DB_PATH, check_same_thread=False)


def init_db():
    conn = get_conn()
    conn.executescript("""
        CREATE TABLE IF NOT EXISTS licencas (
            id               INTEGER PRIMARY KEY AUTOINCREMENT,
            empresa          TEXT    DEFAULT 'Nao informada',
            colaborador      TEXT    NOT NULL,
            centro_custo     TEXT,
            tipo_licenca     TEXT    NOT NULL,
            valor_licenca    REAL,
            vencimento       TEXT,
            status           TEXT    DEFAULT 'Pendente',
            alerta           TEXT    DEFAULT 'Sem data',
            dias_para_vencer INTEGER,
            fonte            TEXT    DEFAULT 'planilha',
            criado_em        TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            atualizado_em    TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        );
        CREATE UNIQUE INDEX IF NOT EXISTS idx_unico
            ON licencas(colaborador, tipo_licenca, COALESCE(empresa,''));
        CREATE TABLE IF NOT EXISTS importacoes (
            id                    INTEGER PRIMARY KEY AUTOINCREMENT,
            arquivo               TEXT,
            aba                   TEXT,
            data_importacao       TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            registros_novos       INTEGER DEFAULT 0,
            registros_atualizados INTEGER DEFAULT 0,
            registros_total       INTEGER DEFAULT 0
        );
    """)
    conn.commit()
    conn.close()


def carregar_licencas(filtros=None):
    conn = get_conn()
    q = "SELECT * FROM licencas WHERE 1=1"
    params = []
    if filtros:
        if filtros.get("empresa"):   q += " AND empresa=?";      params.append(filtros["empresa"])
        if filtros.get("tipo"):      q += " AND tipo_licenca=?";  params.append(filtros["tipo"])
        if filtros.get("alerta"):    q += " AND alerta=?";        params.append(filtros["alerta"])
        if filtros.get("status"):    q += " AND status=?";        params.append(filtros["status"])
    q += " ORDER BY vencimento ASC NULLS LAST"
    df = pd.read_sql_query(q, conn, params=params)
    conn.close()
    return df


def atualizar_registro(id_, campos):
    conn = get_conn()
    sets = [f"{k}=?" for k in campos] + ["atualizado_em=CURRENT_TIMESTAMP"]
    conn.execute(f"UPDATE licencas SET {','.join(sets)} WHERE id=?",
                 list(campos.values()) + [id_])
    conn.commit()
    conn.close()


def upsert_licencas(df, fonte="planilha"):
    conn = get_conn()
    novos = atualizados = 0
    for _, row in df.iterrows():
        colab   = str(row.get("Colaborador","") or "").strip()
        tipo    = str(row.get("Tipo de Licenca","") or "").strip()
        empresa = str(row.get("Empresa","") or "Nao informada").strip()
        if not colab or not tipo:
            continue
        vencimento = None
        v = row.get("Vencimento")
        if v and pd.notna(v) and str(v).strip() not in ("","None","nan","NaT"):
            try: vencimento = pd.to_datetime(v).strftime("%Y-%m-%d")
            except Exception: pass
        valor = None
        val_raw = row.get("Valor da Licenca")
        if val_raw and pd.notna(val_raw):
            try:
                s = str(val_raw).replace("R$","").replace(".","").replace(",",".").strip()
                valor = float(s)
            except Exception: pass
        status = row.get("Status","Pendente")
        if status not in STATUS_VALIDOS: status = "Pendente"
        centro = str(row.get("Centro de Custo","") or "").strip()
        cur = conn.execute(
            "SELECT id FROM licencas WHERE colaborador=? AND tipo_licenca=? AND COALESCE(empresa,'')=?",
            (colab, tipo, empresa))
        ex = cur.fetchone()
        if ex:
            conn.execute(
                "UPDATE licencas SET empresa=?,centro_custo=?,valor_licenca=?,"
                "vencimento=?,status=?,fonte=?,atualizado_em=CURRENT_TIMESTAMP WHERE id=?",
                (empresa,centro,valor,vencimento,status,fonte,ex[0]))
            atualizados += 1
        else:
            conn.execute(
                "INSERT INTO licencas (empresa,colaborador,centro_custo,tipo_licenca,"
                "valor_licenca,vencimento,status,fonte) VALUES(?,?,?,?,?,?,?,?)",
                (empresa,colab,centro,tipo,valor,vencimento,status,fonte))
            novos += 1
    conn.commit()
    conn.close()
    return novos, atualizados


def log_importacao(arquivo, aba, novos, atualizados, total):
    conn = get_conn()
    conn.execute(
        "INSERT INTO importacoes (arquivo,aba,registros_novos,registros_atualizados,registros_total)"
        " VALUES(?,?,?,?,?)", (arquivo,aba,novos,atualizados,total))
    conn.commit()
    conn.close()


def get_historico():
    conn = get_conn()
    df = pd.read_sql_query(
        "SELECT arquivo,aba,data_importacao,registros_novos,registros_atualizados,registros_total "
        "FROM importacoes ORDER BY data_importacao DESC LIMIT 50", conn)
    conn.close()
    return df


def recalcular_alertas(dias_alerta):
    conn = get_conn()
    rows = conn.execute("SELECT id,vencimento FROM licencas").fetchall()
    hoje = date.today()
    lim_atencao = int(dias_alerta)
    lim_critica  = max(7, int(dias_alerta * 0.33))
    updates = []
    for id_, venc_str in rows:
        if not venc_str:
            updates.append(("Sem data", None, id_))
            continue
        try:
            venc = datetime.strptime(venc_str, "%Y-%m-%d").date()
            dias = (venc - hoje).days
            if dias < 0:            alerta = "Vencida"
            elif dias <= lim_critica:  alerta = "Critica"
            elif dias <= lim_atencao:  alerta = "Atencao"
            else:                      alerta = "Ok"
            updates.append((alerta, dias, id_))
        except Exception:
            updates.append(("Sem data", None, id_))
    conn.executemany("UPDATE licencas SET alerta=?,dias_para_vencer=? WHERE id=?", updates)
    conn.commit()
    conn.close()

# ============================================================
# UTILITARIOS DE IMPORTACAO
# ============================================================

def normalizar(txt):
    txt = str(txt).lower().strip()
    return unicodedata.normalize("NFKD", txt).encode("ascii","ignore").decode()


def fix_enc(s):
    if isinstance(s, str):
        try: return s.encode('latin1').decode('utf-8')
        except Exception: return s
    return s


def norm_empresa(v):
    if pd.isna(v): return "Nao informada"
    s = fix_enc(str(v).strip())
    return EMPRESA_MAP.get(s, s)


def dedup_columns(df):
    """Remove colunas duplicadas mantendo a primeira ocorrencia."""
    cols = pd.Series(df.columns)
    seen = {}
    new_cols = []
    for c in cols:
        if c in seen:
            seen[c] += 1
            new_cols.append(f"{c}_{seen[c]}")
        else:
            seen[c] = 0
            new_cols.append(c)
    df.columns = new_cols
    return df


def listar_abas(arquivo):
    try:
        arquivo.seek(0)
        return pd.ExcelFile(arquivo).sheet_names
    except Exception:
        return []


def ler_arquivo(arquivo, aba=None):
    nome = arquivo.name.lower()
    if nome.endswith(".csv"):
        for sep, enc in [(",","utf-8"),(";","utf-8"),(";","latin1"),(",","latin1")]:
            try:
                arquivo.seek(0)
                df = pd.read_csv(arquivo, sep=sep, encoding=enc)
                return dedup_columns(df)
            except Exception:
                continue
        raise ValueError("Nao foi possivel ler o CSV.")
    else:
        arquivo.seek(0)
        df = pd.read_excel(arquivo, sheet_name=aba)
        return dedup_columns(df)


def detectar_mapeamento(df):
    mapa = {}
    colunas_norm = [normalizar(c) for c in df.columns]
    colunas_orig = list(df.columns)
    for col_sis, sinonimos in SINONIMOS.items():
        for i, col_n in enumerate(colunas_norm):
            if any(s in col_n or col_n in s for s in sinonimos):
                if col_sis not in mapa:
                    mapa[col_sis] = colunas_orig[i]
                break
    return mapa


def aplicar_mapeamento(df, mapa):
    # Construir rename sem conflitos
    rename = {}
    destinos_usados = set()
    for sys_col, src_col in mapa.items():
        if src_col in df.columns and sys_col not in destinos_usados:
            if src_col != sys_col:  # apenas renomear se diferente
                rename[src_col] = sys_col
            destinos_usados.add(sys_col)

    # Verificar que o rename nao vai criar duplicatas
    rename_final = {}
    for src, dst in rename.items():
        if dst not in df.columns or src == dst:
            rename_final[src] = dst
        # se dst ja existe como coluna E src != dst → pular (evita duplicata)

    df = df.rename(columns=rename_final)
    df = dedup_columns(df)  # seguranca extra

    for col in COLUNAS_SISTEMA:
        if col not in df.columns:
            df[col] = None

    # Auto-detectar empresa pelo CC se vazio
    mask = df["Empresa"].isna() | (df["Empresa"].astype(str).str.strip() == "")
    if mask.all():
        df["Empresa"] = df["Centro de Custo"].apply(
            lambda x: EMPRESA_PREFIXOS.get(str(x).strip()[:2], "Nao informada") if pd.notna(x) else "Nao informada")
    else:
        df.loc[mask, "Empresa"] = df.loc[mask, "Centro de Custo"].apply(
            lambda x: EMPRESA_PREFIXOS.get(str(x).strip()[:2], "Nao informada") if pd.notna(x) else "Nao informada")

    df["Status"] = df["Status"].apply(lambda x: x if x in STATUS_VALIDOS else "Pendente")
    return df[COLUNAS_SISTEMA]


def eh_mapeamento_exato(mapa):
    """Retorna True se todas as 7 colunas foram mapeadas."""
    return all(c in mapa for c in COLUNAS_SISTEMA)

# ============================================================
# MICROSOFT 365 GRAPH API
# ============================================================

def salvar_creds_ms365(tenant, client_id, client_secret):
    with open(CREDS_PATH, "w") as f:
        json.dump({"tenant_id": tenant, "client_id": client_id,
                   "client_secret": client_secret}, f)


def carregar_creds_ms365():
    if os.path.exists(CREDS_PATH):
        try:
            with open(CREDS_PATH) as f:
                return json.load(f)
        except Exception:
            pass
    return {}


def obter_token_ms365(tenant_id, client_id, client_secret):
    url = f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"
    resp = requests.post(url, data={
        "grant_type":    "client_credentials",
        "client_id":     client_id,
        "client_secret": client_secret,
        "scope":         "https://graph.microsoft.com/.default",
    }, timeout=30)
    data = resp.json()
    if "access_token" not in data:
        raise ValueError(f"Erro ao obter token: {data.get('error_description','Verifique as credenciais.')}")
    return data["access_token"]


def buscar_skus_ms365(token):
    headers = {"Authorization": f"Bearer {token}"}
    resp = requests.get("https://graph.microsoft.com/v1.0/subscribedSkus",
                        headers=headers, timeout=30).json()
    mapa = {}
    for sku in resp.get("value", []):
        part = sku.get("skuPartNumber","")
        nome = MS365_SKU_NAMES.get(part, part)
        mapa[sku["skuId"]] = nome
    return mapa


def buscar_usuarios_ms365(token, sku_map):
    headers = {"Authorization": f"Bearer {token}"}
    url = ("https://graph.microsoft.com/v1.0/users"
           "?$select=displayName,department,companyName,userPrincipalName,assignedLicenses"
           "&$top=999")
    registros = []
    paginas = 0
    while url and paginas < 50:  # max 50 paginas = 49.999 usuarios
        resp = requests.get(url, headers=headers, timeout=60).json()
        for u in resp.get("value", []):
            empresa = u.get("companyName") or "Afonso Franca"
            empresa = EMPRESA_MAP.get(empresa, empresa)
            depto   = u.get("department") or ""
            nome    = u.get("displayName") or u.get("userPrincipalName","")
            for lic in u.get("assignedLicenses", []):
                if lic.get("disabledPlans"):
                    continue  # licenca com planos desativados = provavel nao-usada
                sku_id = lic["skuId"]
                tipo   = sku_map.get(sku_id, sku_id)
                registros.append({
                    "Empresa":          empresa,
                    "Colaborador":      nome,
                    "Centro de Custo":  depto,
                    "Tipo de Licenca":  tipo,
                    "Valor da Licenca": None,
                    "Vencimento":       "",
                    "Status":           "Pendente",
                })
        url = resp.get("@odata.nextLink")
        paginas += 1
    return pd.DataFrame(registros) if registros else pd.DataFrame(columns=COLUNAS_SISTEMA)

# ============================================================
# UTILITARIOS GERAIS
# ============================================================

def formatar_brl(valor):
    try:
        return "R$ {:,.2f}".format(float(valor)).replace(",","X").replace(".",",").replace("X",".")
    except Exception:
        return "-"


def adicionar_meses(dt, n):
    import calendar as _cal
    mes = ((dt.month - 1 + n) % 12) + 1
    ano = dt.year + ((dt.month - 1 + n) // 12)
    dia = min(dt.day, _cal.monthrange(ano, mes)[1])
    return dt.replace(year=ano, month=mes, day=dia)


def gerar_excel(df):
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as w:
        df.to_excel(w, index=False, sheet_name="Licencas")
    return out.getvalue()

# ============================================================
# INICIALIZACAO
# ============================================================

init_db()

for k, v in {"pagina":"Painel","mes_sel":date.today().month,"ano_sel":date.today().year,
             "data_sel":None,"dias_alerta":30,"importado":False}.items():
    if k not in st.session_state:
        st.session_state[k] = v

# ============================================================
# SIDEBAR
# ============================================================

with st.sidebar:
    st.markdown("## Licencas - Afonso Franca")
    st.markdown("---")
    pagina = st.radio("Menu", ["Painel","Importar","MS365 Sync","Licencas","Exportar"],
                      index=["Painel","Importar","MS365 Sync","Licencas","Exportar"]
                      .index(st.session_state.pagina), key="radio_pagina")
    st.session_state.pagina = pagina

    st.markdown("---")
    alerta_idx = list(ALERTA_OPCOES.keys()).index("30 dias")
    alerta_sel = st.selectbox("Alerta de vencimento", list(ALERTA_OPCOES.keys()),
                              index=alerta_idx, key="alerta_sel")
    dias_alerta = ALERTA_OPCOES[alerta_sel]
    if dias_alerta != st.session_state.dias_alerta:
        st.session_state.dias_alerta = dias_alerta

    recalcular_alertas(dias_alerta)

    st.markdown("---")
    df_all = carregar_licencas()
    if len(df_all) > 0:
        venc  = int((df_all["alerta"]=="Vencida").sum())
        crit  = int((df_all["alerta"]=="Critica").sum())
        atenc = int((df_all["alerta"]=="Atencao").sum())
        total_lic = len(df_all)
        st.metric("Total licencas", total_lic)
        c1,c2 = st.columns(2)
        c1.metric("Vencidas", venc, delta=None)
        c2.metric("Criticas", crit, delta=None)
        if atenc > 0:
            st.warning(f"{atenc} licencas vencem em {alerta_sel}")
        if venc > 0:
            st.error(f"{venc} licencas JA VENCIDAS")
    else:
        st.info("Sem licencas. Importe uma planilha.")

# ============================================================
# PAGINA: PAINEL
# ============================================================

if st.session_state.pagina == "Painel":
    st.title("Painel de Vencimentos")

    df = carregar_licencas()
    if len(df) == 0:
        st.info("Nenhuma licenca cadastrada ainda.")
        st.markdown("**Como comecar:**")
        st.markdown("1. Va em **Importar** e faca upload da planilha `licencas_padronizado.xlsx`")
        st.markdown("2. Ou va em **MS365 Sync** para buscar automaticamente as licencas do Microsoft 365")
        st.stop()

    # Filtros
    with st.expander("Filtros", expanded=False):
        fc1,fc2,fc3,fc4 = st.columns(4)
        emp_f   = fc1.selectbox("Empresa",        ["Todas"]+sorted(df["empresa"].dropna().unique().tolist()),   key="pan_emp")
        tipo_f  = fc2.selectbox("Tipo",           ["Todos"]+sorted(df["tipo_licenca"].dropna().unique().tolist()),key="pan_tipo")
        alerta_f= fc3.selectbox("Alerta",         ["Todos","Vencida","Critica","Atencao","Ok","Sem data"],      key="pan_al")
        fonte_f = fc4.selectbox("Fonte",          ["Todos"]+sorted(df["fonte"].dropna().unique().tolist()),     key="pan_fonte")

    df_fil = df.copy()
    if emp_f    != "Todas": df_fil = df_fil[df_fil["empresa"]      == emp_f]
    if tipo_f   != "Todos": df_fil = df_fil[df_fil["tipo_licenca"] == tipo_f]
    if alerta_f != "Todos": df_fil = df_fil[df_fil["alerta"]       == alerta_f]
    if fonte_f  != "Todos": df_fil = df_fil[df_fil["fonte"]        == fonte_f]

    m1,m2,m3,m4,m5 = st.columns(5)
    m1.metric("Registros", len(df_fil))
    m2.metric("Vencidas",  int((df_fil["alerta"]=="Vencida").sum()))
    m3.metric("Criticas",  int((df_fil["alerta"]=="Critica").sum()))
    m4.metric("Atencao",   int((df_fil["alerta"]=="Atencao").sum()))
    val_tot = df_fil["valor_licenca"].dropna().sum()
    m5.metric("Valor Total", formatar_brl(val_tot) if val_tot > 0 else "-")

    st.markdown("---")
    col_cal, col_det = st.columns([3, 2])

    with col_cal:
        nav_a, nav_b, nav_c = st.columns([1,4,1])
        if nav_a.button("<<", key="prev_m"):
            if st.session_state.mes_sel == 1: st.session_state.mes_sel=12; st.session_state.ano_sel-=1
            else: st.session_state.mes_sel -= 1
            st.session_state.data_sel = None; st.rerun()

        nav_b.markdown(
            f"<h3 style='text-align:center;margin:0'>"
            f"{MESES_PT[st.session_state.mes_sel]} / {st.session_state.ano_sel}</h3>",
            unsafe_allow_html=True)

        if nav_c.button(">>", key="next_m"):
            if st.session_state.mes_sel == 12: st.session_state.mes_sel=1; st.session_state.ano_sel+=1
            else: st.session_state.mes_sel += 1
            st.session_state.data_sel = None; st.rerun()

        df_fil["_vd"] = pd.to_datetime(df_fil["vencimento"], errors="coerce").dt.date
        mes = st.session_state.mes_sel
        ano = st.session_state.ano_sel
        hoje = date.today()

        def info_dia(dia):
            dt = date(ano, mes, dia)
            rows = df_fil[df_fil["_vd"] == dt]
            if len(rows) == 0: return None, 0
            for a in ["Vencida","Critica","Atencao","Ok"]:
                if a in rows["alerta"].tolist(): return a, len(rows)
            return "Sem data", len(rows)

        for i, d in enumerate(["Seg","Ter","Qua","Qui","Sex","Sab","Dom"]):
            st.columns(7)[i].markdown(f"<center><small><b>{d}</b></small></center>", unsafe_allow_html=True)

        for semana in calendar.monthcalendar(ano, mes):
            scols = st.columns(7)
            for i, dia in enumerate(semana):
                if dia == 0: scols[i].write(" "); continue
                al, n_lic = info_dia(dia)
                eh_hoje = date(ano,mes,dia) == hoje
                if al:       lbl = f"**{dia}**\n{n_lic}"; tp = "primary"
                elif eh_hoje: lbl = f"**{dia}**"; tp = "secondary"
                else:         lbl = str(dia); tp = "secondary"
                if scols[i].button(lbl, key=f"d_{ano}_{mes}_{dia}",
                                   use_container_width=True, type=tp):
                    st.session_state.data_sel = date(ano,mes,dia); st.rerun()

    with col_det:
        if st.session_state.data_sel:
            dt_sel = st.session_state.data_sel
            st.markdown(f"### {dt_sel.strftime('%d/%m/%Y')}")
            df_dia = df_fil[df_fil["_vd"] == dt_sel]
            if len(df_dia) == 0:
                st.info("Nenhuma licenca vence nesta data.")
            else:
                for _, row in df_dia.iterrows():
                    cor = COR_ALERTA.get(row["alerta"],"#9E9E9E")
                    dias_txt = f"{int(row['dias_para_vencer'])}d" if pd.notna(row.get("dias_para_vencer")) else "?"
                    st.markdown(f"""
<div style="border-left:4px solid {cor};padding:8px 12px;margin:6px 0;background:#FAFAFA;border-radius:4px">
  <b>{row['colaborador']}</b><br>
  <small>{row['tipo_licenca']}</small><br>
  <small>Empresa: {row['empresa']} | CC: {row['centro_custo'] or '-'}</small><br>
  <small>Valor: {formatar_brl(row['valor_licenca'])} | <b style="color:{cor}">{row['alerta']} ({dias_txt})</b></small>
</div>""", unsafe_allow_html=True)
                    with st.expander(f"Renovar #{row['id']}"):
                        try: vd_atual = datetime.strptime(row["vencimento"],"%Y-%m-%d").date()
                        except Exception: vd_atual = date.today()
                        e1,e2 = st.columns(2)
                        ns = e1.selectbox("Status", STATUS_VALIDOS,
                            index=STATUS_VALIDOS.index(row["status"]) if row["status"] in STATUS_VALIDOS else 0,
                            key=f"st_{row['id']}")
                        nd = e2.date_input("Vencimento", value=vd_atual, key=f"dt_{row['id']}")
                        rc = st.columns(4)
                        for j, mn in enumerate([1,3,6,12]):
                            if rc[j].button(f"+{mn}m", key=f"r1_{row['id']}_{mn}"):
                                nova = adicionar_meses(vd_atual, mn)
                                atualizar_registro(row["id"],{"vencimento":nova.strftime("%Y-%m-%d"),"status":"Renovada"})
                                recalcular_alertas(dias_alerta)
                                st.success(f"Renovado para {nova.strftime('%d/%m/%Y')}"); st.rerun()
                        rc2 = st.columns(4)
                        for j, mn in enumerate([24,36,48,60]):
                            if rc2[j].button(f"+{mn}m", key=f"r2_{row['id']}_{mn}"):
                                nova = adicionar_meses(vd_atual, mn)
                                atualizar_registro(row["id"],{"vencimento":nova.strftime("%Y-%m-%d"),"status":"Renovada"})
                                recalcular_alertas(dias_alerta)
                                st.success(f"Renovado para {nova.strftime('%d/%m/%Y')}"); st.rerun()
                        if st.button("Salvar", key=f"sv_{row['id']}", type="primary"):
                            atualizar_registro(row["id"],{"status":ns,"vencimento":nd.strftime("%Y-%m-%d")})
                            recalcular_alertas(dias_alerta); st.success("Salvo!"); st.rerun()
        else:
            st.markdown("### Proximas a vencer")
            df_urg = df_fil[df_fil["alerta"].isin(["Vencida","Critica","Atencao"])].sort_values("dias_para_vencer")
            if len(df_urg) > 0:
                for _, row in df_urg.head(12).iterrows():
                    cor = COR_ALERTA.get(row["alerta"],"#9E9E9E")
                    dias_txt = f"{int(row['dias_para_vencer'])}d" if pd.notna(row.get("dias_para_vencer")) else "vencida"
                    st.markdown(f"""
<div style="border-left:4px solid {cor};padding:5px 10px;margin:3px 0;font-size:13px">
  <b>{row['colaborador']}</b> - {row['tipo_licenca']}<br>
  <small>{row['empresa']} | <b style="color:{cor}">{dias_txt}</b></small>
</div>""", unsafe_allow_html=True)
            else:
                st.success("Nenhuma licenca critica no periodo filtrado.")

# ============================================================
# PAGINA: IMPORTAR
# ============================================================

elif st.session_state.pagina == "Importar":
    st.title("Importar Planilha")
    st.markdown(
        "Faca upload da planilha. O sistema detecta as colunas automaticamente e salva no banco.  \n"
        "Registros com mesmo **Colaborador + Tipo + Empresa** serao **atualizados** automaticamente.")

    arquivo = st.file_uploader("Selecione a planilha (.xlsx ou .csv)", type=["xlsx","xls","csv"])

    if arquivo:
        abas = listar_abas(arquivo)
        aba_sel = None
        juntar = False

        if abas:
            if len(abas) > 1:
                opcao = st.radio("Qual aba?", ["Importar todas as abas"] + abas, key="aba_sel")
                juntar = opcao == "Importar todas as abas"
                aba_sel = None if juntar else opcao
            else:
                aba_sel = abas[0]
                st.caption(f"Aba: {aba_sel}")

        try:
            if juntar:
                dfs = []
                for aba in abas:
                    try:
                        arquivo.seek(0)
                        df_aba = pd.read_excel(arquivo, sheet_name=aba)
                        df_aba = dedup_columns(df_aba)
                        df_aba = df_aba.dropna(how="all")
                        if len(df_aba) > 0 and len(df_aba.columns) > 0:
                            dfs.append(df_aba)
                    except Exception:
                        pass
                if not dfs:
                    st.error("Nenhuma aba com dados validos encontrada.")
                    st.stop()
                df_raw = pd.concat(dfs, ignore_index=True)
                df_raw = dedup_columns(df_raw)
                abas_label = "todas"
            else:
                df_raw = ler_arquivo(arquivo, aba_sel)
                abas_label = aba_sel or "csv"

            df_raw = df_raw.dropna(how="all")
            st.caption(f"{len(df_raw)} linhas | {len(df_raw.columns)} colunas: {', '.join(df_raw.columns.tolist()[:8])}")

            # Detectar mapeamento automatico
            mapa_auto = detectar_mapeamento(df_raw)
            mapeado_completo = eh_mapeamento_exato(mapa_auto)

            if mapeado_completo:
                # IMPORTACAO DIRETA - sem necessidade de ajuste manual
                st.success("Colunas detectadas automaticamente. Pronto para importar!")
                df_prev = aplicar_mapeamento(df_raw.copy(), mapa_auto)
                df_prev = df_prev.dropna(subset=["Colaborador","Tipo de Licenca"])

                col_info1, col_info2, col_info3 = st.columns(3)
                col_info1.metric("Registros", len(df_prev))
                col_info2.metric("Com vencimento", df_prev["Vencimento"].notna().sum())
                col_info3.metric("Com valor", df_prev["Valor da Licenca"].notna().sum())

                with st.expander("Preview (5 primeiros registros)"):
                    st.dataframe(df_prev.head(5), use_container_width=True)

                if st.button(f"Importar {len(df_prev)} registros", type="primary", key="btn_import_direto"):
                    novos, atualizados = upsert_licencas(df_prev, fonte="planilha")
                    log_importacao(arquivo.name, abas_label, novos, atualizados, len(df_prev))
                    recalcular_alertas(st.session_state.dias_alerta)
                    st.success(f"Concluido: {novos} novos + {atualizados} atualizados")
                    st.session_state.pagina = "Painel"
                    st.rerun()
            else:
                # Mapeamento parcial - mostrar UI de ajuste
                faltando = [c for c in COLUNAS_SISTEMA if c not in mapa_auto]
                st.warning(f"Ajuste o mapeamento para: {', '.join(faltando)}")

                cols_disp = ["(nao mapear)"] + df_raw.columns.tolist()
                mapa_user = {}
                grid = st.columns(2)
                for idx, col_sis in enumerate(COLUNAS_SISTEMA):
                    container = grid[idx % 2]
                    default = mapa_auto.get(col_sis, "(nao mapear)")
                    if default not in cols_disp: default = "(nao mapear)"
                    sel = container.selectbox(col_sis, cols_disp,
                        index=cols_disp.index(default), key=f"map_{col_sis}")
                    if sel != "(nao mapear)":
                        mapa_user[col_sis] = sel

                if "Colaborador" in mapa_user and "Tipo de Licenca" in mapa_user:
                    df_prev = aplicar_mapeamento(df_raw.copy(), mapa_user)
                    df_prev = df_prev.dropna(subset=["Colaborador","Tipo de Licenca"])
                    with st.expander("Preview"):
                        st.dataframe(df_prev.head(5), use_container_width=True)
                    if st.button(f"Importar {len(df_prev)} registros", type="primary", key="btn_import_manual"):
                        novos, atualizados = upsert_licencas(df_prev, fonte="planilha")
                        log_importacao(arquivo.name, abas_label, novos, atualizados, len(df_prev))
                        recalcular_alertas(st.session_state.dias_alerta)
                        st.success(f"Concluido: {novos} novos + {atualizados} atualizados")
                        st.session_state.pagina = "Painel"
                        st.rerun()
                else:
                    st.error("Mapeie pelo menos Colaborador e Tipo de Licenca.")

        except Exception as e:
            st.error(f"Erro: {e}")

    st.markdown("---")
    st.subheader("Historico de importacoes")
    hist = get_historico()
    if len(hist) > 0:
        hist.columns = ["Arquivo","Aba","Data","Novos","Atualizados","Total"]
        st.dataframe(hist, use_container_width=True, hide_index=True)
    else:
        st.info("Nenhuma importacao realizada ainda.")

# ============================================================
# PAGINA: MS365 SYNC
# ============================================================

elif st.session_state.pagina == "MS365 Sync":
    st.title("Sincronizar Microsoft 365")
    st.markdown(
        "Busca automaticamente todos os usuarios e suas licencas atribuidas no seu tenant MS365.  \n"
        "Os dados sao inseridos/atualizados no banco com a mesma logica das planilhas importadas.")

    with st.expander("Como obter as credenciais?", expanded=False):
        st.markdown("""
**Passos no portal Azure (portal.azure.com):**

1. Va em **Azure Active Directory** → **Registros de app** → **Novo registro**
2. Nome: ex. `Gerenciador Licencas`; Conta: *Este diretorio apenas*
3. Apos criar: copie o **ID do aplicativo (client_id)** e o **ID do diretorio (tenant_id)**
4. Va em **Certificados e segredos** → **Novo segredo** → copie o valor (client_secret)
5. Va em **Permissoes de API** → **Adicionar permissao** → **Microsoft Graph** → **Permissoes de aplicativo**
6. Adicione: `User.Read.All` e `Organization.Read.All`
7. Clique em **Conceder consentimento do administrador**

As credenciais sao salvas localmente no arquivo `.ms365_creds.json` ao lado do app.
        """)

    creds = carregar_creds_ms365()
    with st.form("form_creds"):
        st.subheader("Credenciais Azure AD")
        tenant_id     = st.text_input("Tenant ID (ID do diretorio)", value=creds.get("tenant_id",""))
        client_id     = st.text_input("Client ID (ID do aplicativo)", value=creds.get("client_id",""))
        client_secret = st.text_input("Client Secret", value=creds.get("client_secret",""), type="password")
        col_save, col_sync = st.columns(2)
        salvar_btn = col_save.form_submit_button("Salvar credenciais")
        sincronizar_btn = col_sync.form_submit_button("Sincronizar agora", type="primary")

    if salvar_btn:
        if tenant_id and client_id and client_secret:
            salvar_creds_ms365(tenant_id, client_id, client_secret)
            st.success("Credenciais salvas.")
        else:
            st.error("Preencha todos os campos.")

    if sincronizar_btn:
        if not (tenant_id and client_id and client_secret):
            st.error("Preencha e salve as credenciais antes de sincronizar.")
        else:
            salvar_creds_ms365(tenant_id, client_id, client_secret)
            with st.spinner("Conectando ao Microsoft Graph..."):
                try:
                    token = obter_token_ms365(tenant_id, client_id, client_secret)
                    st.info("Token obtido. Buscando SKUs...")
                    sku_map = buscar_skus_ms365(token)
                    st.info(f"{len(sku_map)} tipos de licenca encontrados. Buscando usuarios...")
                    df_ms = buscar_usuarios_ms365(token, sku_map)

                    if len(df_ms) == 0:
                        st.warning("Nenhum usuario com licenca atribuida encontrado.")
                    else:
                        st.info(f"{len(df_ms)} registros encontrados. Importando...")
                        novos, atualizados = upsert_licencas(df_ms, fonte="ms365_sync")
                        log_importacao("MS365 Graph API", "sync", novos, atualizados, len(df_ms))
                        recalcular_alertas(st.session_state.dias_alerta)
                        st.success(f"Sincronizacao concluida: {novos} novos + {atualizados} atualizados")

                        col1, col2, col3 = st.columns(3)
                        col1.metric("Total registros", len(df_ms))
                        col2.metric("Usuarios unicos", df_ms["Colaborador"].nunique())
                        col3.metric("Tipos de licenca", df_ms["Tipo de Licenca"].nunique())

                        with st.expander("Preview"):
                            st.dataframe(df_ms.head(10), use_container_width=True)

                        if st.button("Ir para o Painel", type="primary"):
                            st.session_state.pagina = "Painel"; st.rerun()
                except Exception as e:
                    st.error(f"Erro na sincronizacao: {e}")

    st.markdown("---")
    st.subheader("Status da ultima sincronizacao")
    conn = get_conn()
    last_sync = pd.read_sql_query(
        "SELECT data_importacao,registros_novos,registros_atualizados,registros_total "
        "FROM importacoes WHERE arquivo='MS365 Graph API' ORDER BY data_importacao DESC LIMIT 1", conn)
    conn.close()
    if len(last_sync) > 0:
        r = last_sync.iloc[0]
        st.info(f"Ultima sync: {r['data_importacao']} | {r['registros_total']} registros "
                f"({r['registros_novos']} novos, {r['registros_atualizados']} atualizados)")
    else:
        st.info("Nenhuma sincronizacao realizada ainda.")

# ============================================================
# PAGINA: LICENCAS
# ============================================================

elif st.session_state.pagina == "Licencas":
    st.title("Gerenciar Licencas")
    df = carregar_licencas()
    if len(df) == 0:
        st.info("Nenhuma licenca cadastrada."); st.stop()

    fc1,fc2,fc3,fc4 = st.columns(4)
    emp_f2    = fc1.selectbox("Empresa",  ["Todas"]+sorted(df["empresa"].dropna().unique().tolist()),    key="lic_emp")
    tipo_f2   = fc2.selectbox("Tipo",     ["Todos"]+sorted(df["tipo_licenca"].dropna().unique().tolist()),key="lic_tipo")
    alerta_f2 = fc3.selectbox("Alerta",   ["Todos","Vencida","Critica","Atencao","Ok","Sem data"],        key="lic_al")
    status_f2 = fc4.selectbox("Status",   ["Todos"]+STATUS_VALIDOS,                                      key="lic_st")

    df_fil2 = df.copy()
    if emp_f2    != "Todas": df_fil2 = df_fil2[df_fil2["empresa"]      == emp_f2]
    if tipo_f2   != "Todos": df_fil2 = df_fil2[df_fil2["tipo_licenca"] == tipo_f2]
    if alerta_f2 != "Todos": df_fil2 = df_fil2[df_fil2["alerta"]       == alerta_f2]
    if status_f2 != "Todos": df_fil2 = df_fil2[df_fil2["status"]       == status_f2]

    st.caption(f"{len(df_fil2)} de {len(df)} registros")
    COLS_ED = ["id","colaborador","empresa","centro_custo","tipo_licenca",
               "valor_licenca","vencimento","status","alerta","dias_para_vencer"]

    edited = st.data_editor(
        df_fil2[COLS_ED].reset_index(drop=True),
        column_config={
            "id":               st.column_config.NumberColumn("ID",       disabled=True),
            "colaborador":      st.column_config.TextColumn("Colaborador"),
            "empresa":          st.column_config.SelectboxColumn("Empresa", options=list(EMPRESA_PREFIXOS.values())+["Nao informada"]),
            "centro_custo":     st.column_config.TextColumn("CC"),
            "tipo_licenca":     st.column_config.TextColumn("Tipo"),
            "valor_licenca":    st.column_config.NumberColumn("Valor R$", format="%.2f"),
            "vencimento":       st.column_config.TextColumn("Vencimento (AAAA-MM-DD)"),
            "status":           st.column_config.SelectboxColumn("Status", options=STATUS_VALIDOS),
            "alerta":           st.column_config.TextColumn("Alerta",    disabled=True),
            "dias_para_vencer": st.column_config.NumberColumn("Dias",    disabled=True),
        },
        use_container_width=True, height=420, num_rows="fixed", key="tabela_ed")

    if st.button("Salvar alteracoes", type="primary"):
        orig = df_fil2[COLS_ED].reset_index(drop=True)
        alt = 0
        for i in range(len(edited)):
            diffs = {c: edited.iloc[i][c] for c in ["colaborador","empresa","centro_custo",
                     "tipo_licenca","valor_licenca","vencimento","status"]
                     if str(edited.iloc[i][c]) != str(orig.iloc[i][c])}
            if diffs:
                atualizar_registro(int(edited.iloc[i]["id"]), diffs); alt += 1
        if alt > 0:
            recalcular_alertas(st.session_state.dias_alerta)
            st.success(f"{alt} registro(s) atualizado(s)."); st.rerun()
        else:
            st.info("Sem alteracoes.")

# ============================================================
# PAGINA: EXPORTAR
# ============================================================

elif st.session_state.pagina == "Exportar":
    st.title("Exportar Dados")
    df = carregar_licencas()
    if len(df) == 0:
        st.info("Nenhuma licenca cadastrada."); st.stop()

    fc1,fc2,fc3,fc4 = st.columns(4)
    emp_e    = fc1.selectbox("Empresa", ["Todas"]+sorted(df["empresa"].dropna().unique().tolist()), key="exp_e")
    tipo_e   = fc2.selectbox("Tipo",    ["Todos"]+sorted(df["tipo_licenca"].dropna().unique().tolist()), key="exp_t")
    alerta_e = fc3.selectbox("Alerta",  ["Todos","Vencida","Critica","Atencao","Ok","Sem data"],   key="exp_a")
    status_e = fc4.selectbox("Status",  ["Todos"]+STATUS_VALIDOS, key="exp_s")

    df_exp = df.copy()
    if emp_e    != "Todas": df_exp = df_exp[df_exp["empresa"]      == emp_e]
    if tipo_e   != "Todos": df_exp = df_exp[df_exp["tipo_licenca"] == tipo_e]
    if alerta_e != "Todos": df_exp = df_exp[df_exp["alerta"]       == alerta_e]
    if status_e != "Todos": df_exp = df_exp[df_exp["status"]       == status_e]

    st.metric("Registros a exportar", len(df_exp))
    st.dataframe(df_exp.head(10), use_container_width=True, hide_index=True)

    if len(df_exp) > 0:
        st.download_button(
            label=f"Baixar Excel ({len(df_exp)} registros)",
            data=gerar_excel(df_exp),
            file_name=f"licencas_{date.today().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary")

    st.markdown("---")
    st.subheader("Estatisticas")
    s1,s2,s3 = st.columns(3)
    with s1:
        st.markdown("**Por empresa**")
        st.dataframe(df.groupby("empresa").size().reset_index(name="Licencas")
                     .sort_values("Licencas",ascending=False), use_container_width=True, hide_index=True)
    with s2:
        st.markdown("**Por alerta**")
        ord_al = {"Vencida":0,"Critica":1,"Atencao":2,"Ok":3,"Sem data":4}
        al = df.groupby("alerta").size().reset_index(name="Registros")
        al["_o"] = al["alerta"].map(ord_al)
        st.dataframe(al.sort_values("_o").drop(columns="_o"), use_container_width=True, hide_index=True)
    with s3:
        st.markdown("**Top 10 por valor**")
        top = df.groupby("tipo_licenca")["valor_licenca"].sum().sort_values(ascending=False).head(10).reset_index()
        top.columns = ["Tipo","Valor Total R$"]
        top["Valor Total R$"] = top["Valor Total R$"].apply(lambda x: formatar_brl(x) if x else "-")
        st.dataframe(top, use_container_width=True, hide_index=True)

