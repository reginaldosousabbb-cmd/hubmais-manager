"""
HUBMAIS MANAGER — Exportar dados pro Dashboard (GitHub Pages)
=================================================================
Reaproveita os cálculos do metricas.py e exporta um docs/data.json com:
  - "por_data": métricas completas (individual + supervisor) para CADA dia
    fechado do mês, não só o mais recente — alimenta o seletor de data.
  - "series_colaborador": série diária completa (recebido/negociado/tma) de
    cada colaborador — alimenta o gráfico individual ao abrir o card.
"""

import json
import sys
from datetime import datetime
from pathlib import Path

import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parent))
from metricas import (  # noqa: E402
    carregar_base, calcular_individual, calcular_supervisor, dias_fechados_ordenados, gerar_sugestao,
    META_DIARIA_RECEBIDO, META_SEMANAL_RECEBIDO, LIMIAR_TMA_FORA_PADRAO, LIMIAR_CONVERSAO_PREMIACAO,
    FAIXAS_PREMIACAO,
)

PASTA_PROJETO = Path(__file__).resolve().parent.parent
ARQUIVO_JSON = PASTA_PROJETO / "docs" / "data.json"

DIAS_TENDENCIA_EQUIPE = 10  # janela do gráfico de tendência da equipe por data


def nn(v):
    """None-safe: converte NaN/NaT em None."""
    return None if pd.isna(v) else v


def montar_metricas_do_dia(df: pd.DataFrame, data_ref: pd.Timestamp) -> dict:
    individual = calcular_individual(df, data_referencia=data_ref)
    supervisor = calcular_supervisor(individual, df[df["Data"] <= data_ref])
    sup_por_nome = supervisor.set_index("Colaborador").to_dict(orient="index")

    colaboradores = []
    for _, linha in individual.iterrows():
        nome = linha["Colaborador"]
        s = sup_por_nome[nome]
        l = linha.to_dict()
        pct = nn(linha["% Recebimento (recebido/negociado)"])
        rec_sem_ant = nn(linha["Recebido na semana anterior"])
        evo_sem = nn(linha["Evolução vs. semana anterior"])
        evo_dia = nn(linha["Evolução vs. dia anterior"])
        conv_mes = nn(linha["Premiação: pct_conversao_mes"])
        falta_prox = linha["Premiação: falta_proxima_faixa"]
        dias_restantes = int(linha["Dias úteis restantes no mês"])
        ritmo_necessario = (falta_prox / dias_restantes) if (falta_prox and dias_restantes > 0) else None

        colaboradores.append({
            "nome": nome,
            "tma": linha["TMA do dia"],
            "recebido_hoje": round(nn(linha["Recebido hoje"]) or 0, 2),
            "pct_recebimento": round(pct, 4) if pct is not None else None,
            "recebido_semana": round(linha["Recebido na semana"], 2),
            "meta_semana": META_SEMANAL_RECEBIDO,
            "progresso_semana": round(linha["Progresso da meta semanal"], 4),
            "recebido_semana_anterior": round(rec_sem_ant, 2) if rec_sem_ant is not None else None,
            "evolucao_semana": round(evo_sem, 4) if evo_sem is not None else None,
            "recebido_mes": round(linha["Recebido no mês"], 2),
            "meta_mes": round(linha["Meta mensal"], 2),
            "progresso_mes": round(linha["Progresso da meta mensal"], 4),
            "evolucao_dia": round(evo_dia, 4) if evo_dia is not None else None,
            "ranking": int(linha["Ranking do turno (recebido hoje)"]),
            "tma_fora_padrao": s["TMA fora do padrão?"] != "-",
            "tendencia_texto": s["Melhorando / Caindo"],
            "negocia_muito_recebe_pouco": s["Negocia muito x recebe pouco?"] != "-",
            "recebe_bem_baixo_volume": s["Recebe bem x baixo volume?"] != "-",
            "premiacao": {
                "conversao_mes_pct": round(conv_mes, 4) if conv_mes is not None else None,
                "elegivel": bool(linha["Premiação: elegivel_conversao"]),
                "limiar_conversao": LIMIAR_CONVERSAO_PREMIACAO,
                "valor_bonus_atual": int(linha["Premiação: valor_bonus_atual"]),
                "proximo_valor_bonus": (
                    int(linha["Premiação: proxima_faixa_valor"])
                    if nn(linha["Premiação: proxima_faixa_valor"]) is not None else None
                ),
                "falta_proxima_faixa": round(falta_prox, 2),
                "dias_uteis_restantes_mes": dias_restantes,
                "ritmo_diario_necessario": round(ritmo_necessario, 2) if ritmo_necessario is not None else None,
            },
            "sugestao": gerar_sugestao(l, s),
        })
    colaboradores.sort(key=lambda c: c["ranking"])

    equipe_recebido_semana = sum(c["recebido_semana"] for c in colaboradores)
    equipe_meta_semana = META_SEMANAL_RECEBIDO * len(colaboradores)
    equipe_recebido_mes = sum(c["recebido_mes"] for c in colaboradores)
    equipe_meta_mes = colaboradores[0]["meta_mes"] * len(colaboradores) if colaboradores else 0
    equipe_recebido_hoje = sum(c["recebido_hoje"] for c in colaboradores)

    janela = (
        df[df["Data"] <= data_ref]
        .dropna(subset=["Recebido"])
        .groupby("Data")["Recebido"].sum()
        .sort_index()
        .tail(DIAS_TENDENCIA_EQUIPE)
    )
    tendencia_equipe = [{"data": d.strftime("%d/%m"), "recebido": round(v, 2)} for d, v in janela.items()]

    return {
        "equipe": {
            "recebido_hoje": round(equipe_recebido_hoje, 2),
            "recebido_semana": round(equipe_recebido_semana, 2),
            "meta_semana": equipe_meta_semana,
            "progresso_semana": round(equipe_recebido_semana / equipe_meta_semana, 4),
            "recebido_mes": round(equipe_recebido_mes, 2),
            "meta_mes": round(equipe_meta_mes, 2),
            "progresso_mes": round(equipe_recebido_mes / equipe_meta_mes, 4) if equipe_meta_mes else None,
            "tendencia": tendencia_equipe,
        },
        "colaboradores": colaboradores,
    }


def montar_series_colaborador(df: pd.DataFrame) -> dict:
    """Série diária completa (mês inteiro) de cada colaborador, pro gráfico individual."""
    series = {}
    fechado = df.dropna(subset=["Recebido"]).sort_values("Data")
    for nome, grupo in fechado.groupby("Colaborador"):
        series[nome] = [
            {
                "data": row["Data"].strftime("%d/%m"),
                "recebido": round(row["Recebido"], 2),
                "negociado": round(row["Negociado"], 2) if pd.notna(row["Negociado"]) else 0,
                "tma_segundos": row["TMA_segundos"] if pd.notna(row["TMA_segundos"]) else None,
            }
            for _, row in grupo.iterrows()
        ]
    return series


def exportar():
    df = carregar_base()
    dias = dias_fechados_ordenados(df)
    if not dias:
        raise SystemExit("Nenhum dia fechado (com Recebido) na base — nada pra exportar.")

    por_data = {}
    for dia in dias:
        chave = pd.Timestamp(dia).strftime("%d/%m/%Y")
        por_data[chave] = montar_metricas_do_dia(df, pd.Timestamp(dia))

    datas_disponiveis = [pd.Timestamp(d).strftime("%d/%m/%Y") for d in dias]

    payload = {
        "gerado_em": datetime.now().strftime("%d/%m/%Y %H:%M"),
        "supervisor": "Reginaldo Sousa Barbosa",
        "turno": "Vespertino",
        "config": {
            "meta_diaria": META_DIARIA_RECEBIDO,
            "meta_semanal": META_SEMANAL_RECEBIDO,
            "limiar_tma_pct": LIMIAR_TMA_FORA_PADRAO,
            "limiar_conversao_premiacao": LIMIAR_CONVERSAO_PREMIACAO,
            "faixas_premiacao": [{"pct": p, "valor": v} for p, v in FAIXAS_PREMIACAO],
        },
        "datas_disponiveis": datas_disponiveis,
        "data_mais_recente": datas_disponiveis[-1],
        "por_data": por_data,
        "series_colaborador": montar_series_colaborador(df),
    }

    ARQUIVO_JSON.parent.mkdir(exist_ok=True)
    with open(ARQUIVO_JSON, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)

    print(f"💾 Dados do site exportados em: {ARQUIVO_JSON}")
    print(f"   {len(datas_disponiveis)} dias navegáveis, de {datas_disponiveis[0]} a {datas_disponiveis[-1]}")


if __name__ == "__main__":
    print("=" * 60)
    print("🌐 HUBMAIS MANAGER — EXPORTANDO DADOS PRO DASHBOARD")
    print("=" * 60)
    exportar()
