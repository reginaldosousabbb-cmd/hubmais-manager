"""
HUBMAIS MANAGER — Painel de Métricas
=======================================
Lê a Base Histórica (gerada pelo atualizar_base.py) e calcula as métricas
derivadas: individual (colaborador) e supervisor (visão do time).

Gera dashboard.xlsx com 2 abas: "Painel Individual" e "Painel Supervisor".

⚠️ PONTOS QUE DEPENDEM DE VOCÊ CONFIRMAR (ajustar no bloco CONFIG abaixo):
  - META_SEMANAL_RECEBIDO: hoje é um valor-suposição (5 dias úteis x R$ 2.200,
    que é o piso da faixa "Meta cumprida" no Resultado Parcial). Se a meta
    semanal real for outra, troca aqui.
  - LIMIAR_TMA_FORA_PADRAO: % de desvio da média do turno pra marcar alguém
    com TMA fora do padrão (hoje 30% pra mais ou pra menos).
"""

from pathlib import Path
from datetime import timedelta

import pandas as pd
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

# ==========================================
# CONFIG (ajustável)
# ==========================================
PASTA_PROJETO = Path(__file__).resolve().parent.parent
ARQUIVO_BASE = PASTA_PROJETO / "dados" / "base_historica.xlsx"
ARQUIVO_DASHBOARD = PASTA_PROJETO / "dados" / "dashboard.xlsx"

META_SEMANAL_RECEBIDO = 1923 * 6  # R$ 11.538 — meta diária R$1.923 x 6 dias úteis (seg a sáb)
LIMIAR_TMA_FORA_PADRAO = 0.30      # 30% de desvio da média do turno


# ==========================================
# HELPERS
# ==========================================
def tma_para_segundos(valor) -> float:
    """Converte 'HH:MM:SS' (string salva pelo atualizar_base.py) em segundos."""
    if pd.isna(valor) or valor in (None, ""):
        return None
    h, m, s = str(valor).split(":")
    return int(h) * 3600 + int(m) * 60 + int(s)


def segundos_para_tma(segundos) -> str:
    if pd.isna(segundos) or segundos is None:
        return "-"
    total = int(round(segundos))
    h, resto = divmod(total, 3600)
    m, s = divmod(resto, 60)
    return f"{h:02d}:{m:02d}:{s:02d}"


def carregar_base() -> pd.DataFrame:
    df = pd.read_excel(ARQUIVO_BASE, sheet_name="Base Histórica")
    df["Data"] = pd.to_datetime(df["Data"])
    df["TMA_segundos"] = df["TMA"].apply(tma_para_segundos)
    df["Semana"] = df["Data"].dt.isocalendar().week
    return df


# ==========================================
# MÉTRICAS INDIVIDUAIS
# ==========================================
def calcular_individual(df: pd.DataFrame) -> pd.DataFrame:
    # "Hoje" = último dia com Recebido fechado (não o último dia do TMA, que pode
    # já ter chegado mas ainda estar pendente de fechamento no Resultado Parcial).
    dias_fechados = df.dropna(subset=["Recebido"])["Data"]
    data_hoje = dias_fechados.max()
    data_ontem_util = dias_fechados[dias_fechados < data_hoje].max()
    semana_atual = df.loc[df["Data"] == data_hoje, "Semana"].iloc[0]

    linhas = []
    for colaborador, grupo in df.groupby("Colaborador"):
        hoje = grupo[grupo["Data"] == data_hoje]
        ontem = grupo[grupo["Data"] == data_ontem_util] if pd.notna(data_ontem_util) else grupo.iloc[0:0]
        semana = grupo[grupo["Semana"] == semana_atual]

        recebido_hoje = hoje["Recebido"].sum() if not hoje.empty else 0.0
        negociado_hoje = hoje["Negociado"].sum() if not hoje.empty else 0.0
        recebido_ontem = ontem["Recebido"].sum() if not ontem.empty else None
        recebido_semana = semana["Recebido"].sum()
        tma_hoje = hoje["TMA_segundos"].iloc[0] if not hoje.empty else None

        pct_recebimento = (recebido_hoje / negociado_hoje) if negociado_hoje else None
        progresso_meta = recebido_semana / META_SEMANAL_RECEBIDO
        if recebido_ontem:
            evolucao = (recebido_hoje - recebido_ontem) / recebido_ontem
        else:
            evolucao = None

        linhas.append({
            "Colaborador": colaborador,
            "Data": data_hoje,
            "TMA do dia": segundos_para_tma(tma_hoje),
            "Recebido hoje": recebido_hoje,
            "% Recebimento (recebido/negociado)": pct_recebimento,
            "Recebido na semana": recebido_semana,
            "Meta semanal": META_SEMANAL_RECEBIDO,
            "Progresso da meta semanal": progresso_meta,
            "Evolução vs. dia anterior": evolucao,
        })

    resultado = pd.DataFrame(linhas)
    resultado["Ranking do turno (recebido hoje)"] = (
        resultado["Recebido hoje"].rank(ascending=False, method="min").astype(int)
    )
    resultado = resultado.sort_values("Ranking do turno (recebido hoje)").reset_index(drop=True)
    return resultado


# ==========================================
# MÉTRICAS DE SUPERVISOR
# ==========================================
def calcular_supervisor(individual: pd.DataFrame, df: pd.DataFrame) -> pd.DataFrame:
    v = individual.copy()

    v["Ranking recebimento"] = v["Recebido hoje"].rank(ascending=False, method="min").astype(int)
    v["Ranking conversão"] = v["% Recebimento (recebido/negociado)"].rank(ascending=False, method="min").astype(int)

    media_tma = v["TMA do dia"].apply(tma_para_segundos).mean()
    v["TMA fora do padrão?"] = v["TMA do dia"].apply(tma_para_segundos).apply(
        lambda s: "⚠️ Sim" if pd.notna(s) and media_tma and abs(s - media_tma) / media_tma > LIMIAR_TMA_FORA_PADRAO
        else "-"
    )

    v["Melhorando / Caindo"] = v["Evolução vs. dia anterior"].apply(
        lambda e: "📈 Melhorando" if pd.notna(e) and e > 0
        else ("📉 Caindo" if pd.notna(e) and e < 0 else "-")
    )

    # Negocia muito x recebe pouco: acima da mediana em negociado, abaixo da mediana em conversão
    mediana_negociado = individual["Recebido na semana"].median()  # proxy de volume de atividade
    mediana_conv = individual["% Recebimento (recebido/negociado)"].median()
    negociado_hoje_map = df.groupby("Colaborador")["Negociado"].sum()
    v["Negocia muito x recebe pouco?"] = v.apply(
        lambda r: "⚠️ Sim" if (negociado_hoje_map.get(r["Colaborador"], 0) > negociado_hoje_map.median())
        and pd.notna(r["% Recebimento (recebido/negociado)"])
        and r["% Recebimento (recebido/negociado)"] < mediana_conv
        else "-",
        axis=1,
    )

    negociacoes_map = df.groupby("Colaborador")["Negociações"].sum()
    v["Recebe bem x baixo volume?"] = v.apply(
        lambda r: "⚠️ Sim" if r["Recebido hoje"] > v["Recebido hoje"].median()
        and negociacoes_map.get(r["Colaborador"], 0) < negociacoes_map.median()
        else "-",
        axis=1,
    )

    colunas = ["Colaborador", "Recebido hoje", "Ranking recebimento",
               "% Recebimento (recebido/negociado)", "Ranking conversão",
               "TMA do dia", "TMA fora do padrão?", "Melhorando / Caindo",
               "Negocia muito x recebe pouco?", "Recebe bem x baixo volume?"]
    return v[colunas].sort_values("Ranking recebimento").reset_index(drop=True)


# ==========================================
# ESCRITA DO EXCEL
# ==========================================
def formatar_aba(ws, df: pd.DataFrame, colunas_pct=(), colunas_moeda=(), larguras=None):
    fonte_padrao = Font(name="Arial", size=10)
    fonte_header = Font(name="Arial", size=10, bold=True, color="FFFFFF")
    fill_header = PatternFill("solid", fgColor="1F4E78")
    borda_fina = Border(*(Side(style="thin", color="D9D9D9"),) * 4)

    ws.append(list(df.columns))
    for cel in ws[1]:
        cel.font = fonte_header
        cel.fill = fill_header
        cel.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[1].height = 32

    for _, linha in df.iterrows():
        ws.append(list(linha))

    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        for idx, cel in enumerate(row):
            nome_col = df.columns[idx]
            cel.font = fonte_padrao
            cel.border = borda_fina
            if nome_col in colunas_pct:
                cel.number_format = "0.0%"
            elif nome_col in colunas_moeda:
                cel.number_format = 'R$ #,##0.00;(R$ #,##0.00);"-"'

    larguras = larguras or [18] * len(df.columns)
    for i, largura in enumerate(larguras, start=1):
        ws.column_dimensions[get_column_letter(i)].width = largura

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions


def salvar_dashboard(individual: pd.DataFrame, supervisor: pd.DataFrame, caminho: Path):
    wb = openpyxl.Workbook()

    ws1 = wb.active
    ws1.title = "Painel Individual"
    formatar_aba(
        ws1, individual,
        colunas_pct=("% Recebimento (recebido/negociado)", "Progresso da meta semanal", "Evolução vs. dia anterior"),
        colunas_moeda=("Recebido hoje", "Recebido na semana", "Meta semanal"),
        larguras=[30, 12, 12, 14, 20, 16, 14, 22, 22, 14],
    )

    ws2 = wb.create_sheet("Painel Supervisor")
    formatar_aba(
        ws2, supervisor,
        colunas_pct=("% Recebimento (recebido/negociado)",),
        colunas_moeda=("Recebido hoje",),
        larguras=[30, 14, 16, 24, 16, 12, 16, 18, 20, 20],
    )

    wb.save(caminho)


# ==========================================
# EXECUÇÃO
# ==========================================
if __name__ == "__main__":
    print("=" * 60)
    print("📊 HUBMAIS MANAGER — PAINEL DE MÉTRICAS")
    print("=" * 60)

    df = carregar_base()
    individual = calcular_individual(df)
    supervisor = calcular_supervisor(individual, df)

    print(f"\n✅ Métricas calculadas para {len(individual)} colaboradores")
    print(f"   Data de referência: {individual['Data'].iloc[0].date()}")

    try:
        salvar_dashboard(individual, supervisor, ARQUIVO_DASHBOARD)
    except PermissionError:
        print(f"\n❌ Não consegui salvar: '{ARQUIVO_DASHBOARD.name}' parece estar aberto no Excel.")
        raise SystemExit(1)

    print(f"\n💾 Dashboard salvo em: {ARQUIVO_DASHBOARD}")