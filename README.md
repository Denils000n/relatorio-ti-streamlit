# Gerador de Relatório Mensal TI

Aplicativo em Python + Streamlit para preencher dados mensais de TI, anexar imagens/planilhas e gerar um PowerPoint mensal no padrão do modelo.

## Slides gerados

1. Comparativo mensal de tickets: abertos VS fechados
2. Abertos e fechados diários - análise do período
3. Top analistas com mais chamados
4. Top 10 categorias de chamados
5. Notebooks e monitores alugados Teiti
6. Notebooks e monitores alugados Infotech
7. Telefonia Móvel
8. SharePoint
9. Inventário com análise de planilha
10. Sumário executivo de segurança
11. Gestão de Licenças de Software - Afonso França
12. Fatura VIVO Office 365 - Resumo consolidado
13. Análise de soluções

## Rodar localmente

```powershell
cd "C:\Relatorio\relatorio_ti_streamlit_v3\relatorio_ti_streamlit_v2"
python -m venv .venv
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
python -m streamlit run app.py
```

## Teste rápido

```powershell
python smoke_test.py
```

## Publicar no Streamlit Community Cloud

Suba estes arquivos para um repositório no GitHub e, no Streamlit Community Cloud, selecione `app.py` como arquivo principal.
