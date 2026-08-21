"""
HUBMAIS MANAGER — Exportar dados pro Dashboard (GitHub Pages)
=================================================================
Reaproveita os cálculos do metricas.py e exporta um docs/data.json
que o docs/index.html lê pra montar o painel visual.
"""

import json
import sys
from datetime import datetime
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from metricas import (  # noqa: E402
    carregar_base, calcular_individual, calcular_supervisor,
    tma_para_segundos, META_SEMANAL_RECEBIDO, LIMIAR_TMA_FORA_PADRAO,
)

PASTA_PROJETO = Path(__file__).resolve().parent.parent
ARQUIVO_JSON = PASTA_PROJETO / "docs" / "data.json"

DIAS_TENDENCIA = 10  # quantos dias mostrar no gráfico de tendência


def montar_tendencia_colaborador(df, colaborador: str):
    dados = (
        df[df["Colaborador"] == colaborador]
        .dropna(subset=["Recebido"])
        .sort_values("Data")
        .tail(DIAS_TENDENCIA)
    )
    return [
        {"data": d.strftime("%d/%m"), "recebido": round(r, 2)}
        for d, r in zip(dados["Data"], dados["Recebido"])
    ]


def montar_tendencia_equipe(df):
    dados = (
        df.dropna(subset=["Recebido"])
        .groupby("Data")["Recebido"].sum()
        .sort_index()
        .tail(DIAS_TENDENCIA)
    )
    return [{"data": d.strftime("%d/%m"), "recebido": round(v, 2)} for d, v in dados.items()]


def exportar():
    df = carregar_base()
    individual = calcular_individual(df)
    supervisor = calcular_supervisor(individual, df)
    sup_por_nome = supervisor.set_index("Colaborador").to_dict(orient="index")

    data_ref = individual["Data"].iloc[0]

    colaboradores = []
    for _, linha in individual.iterrows():
        nome = linha["Colaborador"]
        s = sup_por_nome[nome]
        colaboradores.append({
            "nome": nome,
            "tma": linha["TMA do dia"],
            "recebido_hoje": round(linha["Recebido hoje"], 2),
            "pct_recebimento": (
                round(linha["% Recebimento (recebido/negociado)"], 4)
                if pd_notna(linha["% Recebimento (recebido/negociado)"]) else None
            ),
            "recebido_semana": round(linha["Recebido na semana"], 2),
            "meta_semana": META_SEMANAL_RECEBIDO,
            "progresso": round(linha["Progresso da meta semanal"], 4),
            "evolucao": (
                round(linha["Evolução vs. dia anterior"], 4)
                if pd_notna(linha["Evolução vs. dia anterior"]) else None
            ),
            "ranking": int(linha["Ranking do turno (recebido hoje)"]),
            "tma_fora_padrao": s["TMA fora do padrão?"] != "-",
            "tendencia_texto": s["Melhorando / Caindo"],
            "negocia_muito_recebe_pouco": s["Negocia muito x recebe pouco?"] != "-",
            "recebe_bem_baixo_volume": s["Recebe bem x baixo volume?"] != "-",
            "tendencia": montar_tendencia_colaborador(df, nome),
        })

    colaboradores.sort(key=lambda c: c["ranking"])

    equipe_recebido_semana = sum(c["recebido_semana"] for c in colaboradores)
    equipe_meta_semana = META_SEMANAL_RECEBIDO * len(colaboradores)
    equipe_recebido_hoje = sum(c["recebido_hoje"] for c in colaboradores)

    payload = {
        "gerado_em": datetime.now().strftime("%d/%m/%Y %H:%M"),
        "data_referencia": data_ref.strftime("%d/%m/%Y"),
        "supervisor": "Reginaldo Sousa Barbosa",
        "turno": "Vespertino",
        "config": {
            "meta_diaria": META_SEMANAL_RECEBIDO / 6,
            "meta_semanal": META_SEMANAL_RECEBIDO,
            "limiar_tma_pct": LIMIAR_TMA_FORA_PADRAO,
        },
        "equipe": {
            "recebido_semana": round(equipe_recebido_semana, 2),
            "meta_semana": equipe_meta_semana,
            "progresso": round(equipe_recebido_semana / equipe_meta_semana, 4),
            "recebido_hoje": round(equipe_recebido_hoje, 2),
            "tendencia": montar_tendencia_equipe(df),
        },
        "colaboradores": colaboradores,
    }

    ARQUIVO_JSON.parent.mkdir(exist_ok=True)
    with open(ARQUIVO_JSON, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)

    print(f"💾 Dados do site exportados em: {ARQUIVO_JSON}")


def pd_notna(v):
    import pandas as pd
    return pd.notna(v)


if __name__ == "__main__":
    print("=" * 60)
    print("🌐 HUBMAIS MANAGER — EXPORTANDO DADOS PRO DASHBOARD")
    print("=" * 60)
    exportar()
