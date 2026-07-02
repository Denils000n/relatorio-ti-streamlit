from pathlib import Path
from src.report_generator import ReportInputs, build_pptx

BASE_DIR = Path(__file__).resolve().parent
TEMPLATE = BASE_DIR / "assets" / "template_relatorio_ti.pptx"

if __name__ == "__main__":
    data = ReportInputs(
        mes=2,
        ano=2026,
        tickets_abertos=120,
        tickets_fechados=105,
        reducao_fevereiro="Redução registrada no volume de chamados em relação ao período anterior.",
        fator_contexto_mes="Período com ajustes operacionais e acompanhamento de demandas críticas.",
        comentario_tickets_mensal="Comentário de teste.",
        tickets_diarios=[{"Dia": 1, "Abertos": 10, "Fechados": 8}, {"Dia": 2, "Abertos": 15, "Fechados": 12}],
        top_categorias=[{"Categoria": "Acessos", "Quantidade": 30}, {"Categoria": "Hardware", "Quantidade": 20}],
        chamados_destaque="Chamado 001 - Exemplo\nChamado 002 - Exemplo",
        capacidade_total="10 TB",
        espaco_livre="3 TB",
        dias_restantes="90",
        crescimento="5%",
        custo_por_gb="R$ 0,15",
        valor_sharepoint="R$ 1.500,00",
        inventario_df=[{"Tipo": "Notebook", "Status": "Ativo", "Empresa": "Afonso França"}],
        licencas=[{"Empresa": "Afonso França", "Software": "Microsoft 365", "Vencimento": "31/12/2026"}],
    )
    output = build_pptx(TEMPLATE, data)
    out_path = BASE_DIR / "teste_relatorio_ti.pptx"
    out_path.write_bytes(output)
    print(f"Arquivo gerado: {out_path}")
