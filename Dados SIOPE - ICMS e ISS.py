"""
Dados SIOPE - ICMS e ISS
=========================
Lê os CSVs de cache extraídos pelo script principal e gera a planilha
"Dados SIOPE - ICMS e ISS.xlsx" com duas abas:
  - Municipios : Cota-Parte do ICMS e ISS (todas as sub-métricas)
  - Estados    : ICMS (todas as sub-métricas)

Regra de valor: vlr_receita_realizada - vlr_outras_deducoes (ambas em abs)
Fonte dos dados: pasta cache/ gerada pelo script "Extrair Todas Receitas - Todos os Entes.py"
Anos cobertos: 2019–2025 (regras do Word); 2026 aplica as mesmas regras de 2023+
"""

from pathlib import Path
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

# ─────────────────────────────────────────────
# CAMINHOS
# ─────────────────────────────────────────────
BASE_DIR   = Path(__file__).parent
CACHE_DIR  = BASE_DIR / "Registro Receita - Todos Entes" / "cache"
LISTA_ENTES = BASE_DIR / "Lista de Entes.xlsx"
OUTPUT     = BASE_DIR / "Dados SIOPE - ICMS e ISS.xlsx"

# ─────────────────────────────────────────────
# MAPEAMENTO DE CONTAS POR ANO
# Estrutura: { (ano_inicio, ano_fim): { metrica: cod_conta } }
# ─────────────────────────────────────────────

# Municípios — ISS
ISS = {
    (2019, 2020): {
        "ISS Principal":           "41113050100",
        "ISS Multas e Juros Mora": "41911400000",
        "ISS DA - Principal":      "41931130000",
        "ISS DA - Multas e Juros": "41913130000",
        "ISS FECOP":               "41113050200",
    },
    (2021, 2022): {
        "ISS Principal":           "11180231",
        "ISS Multas e Juros Mora": "11180232",
        "ISS DA - Principal":      "11180233",
        "ISS DA - Multas e Juros": "11180234",
        "ISS FECOP":               "11180240",
    },
    (2023, 2026): {
        "ISS Principal":           "11145111",
        "ISS Multas e Juros Mora": "11145112",
        "ISS DA - Principal":      "11145113",
        "ISS DA - Multas e Juros": "11145114",
        "ISS FECOP":               "11145120",
    },
}

# Municípios — ISS específico para o DF (2019-2020)
# O DF usa a conta agregada 41113050000 no lugar de 41113050100,
# e não possui ISS FECOP nesse período (ISS = soma métricas 2 a 5).
# Em 2021+ o DF usa as mesmas contas dos demais municípios.
ISS_DF = {
    (2019, 2020): {
        "ISS Principal":           "41113050000",   # Exclusivo DF (demais municípios: 41113050100)
        "ISS Multas e Juros Mora": "41911400000",
        "ISS DA - Principal":      "41931130000",
        "ISS DA - Multas e Juros": "41913130000",
        # sem ISS FECOP para o DF em 2019-2020
    },
    (2021, 2022): {
        "ISS Principal":           "11180231",
        "ISS Multas e Juros Mora": "11180232",
        "ISS DA - Principal":      "11180233",
        "ISS DA - Multas e Juros": "11180234",
        "ISS FECOP":               "11180240",
    },
    (2023, 2026): {
        "ISS Principal":           "11145111",
        "ISS Multas e Juros Mora": "11145112",
        "ISS DA - Principal":      "11145113",
        "ISS DA - Multas e Juros": "11145114",
        "ISS FECOP":               "11145120",
    },
}

# Municípios — Cota-Parte do ICMS
COTAPARTE = {
    (2019, 2020): {"Cota-Parte": "41722010100"},
    (2021, 2022): {"Cota-Parte": "17280110"},
    (2023, 2026): {"Cota-Parte": "17215000"},
}

# Estados — ICMS
ICMS = {
    (2019, 2020): {
        "ICMS Principal":           "41113020100",
        "ICMS Multas e Juros Mora": "41911420000",
        "ICMS DA - Principal":      "41931150000",
        "ICMS DA - Multas e Juros": "41913150000",
        "ICMS FECOP":               "41113020200",
    },
    (2021, 2022): {
        "ICMS Principal":           "11180211",
        "ICMS Multas e Juros Mora": "11180212",
        "ICMS DA - Principal":      "11180213",
        "ICMS DA - Multas e Juros": "11180215",
        "ICMS FECOP":               "11180220",
    },
    (2023, 2026): {
        "ICMS Principal":           "11145011",
        "ICMS Multas e Juros Mora": "11145012",
        "ICMS DA - Principal":      "11145013",
        "ICMS DA - Multas e Juros": "11145014",
        "ICMS FECOP":               "11145020",
    },
}

# Ordem das colunas métricas nas abas
COLUNAS_MUN = ["Cota-Parte", "ISS", "ISS Principal", "ISS Multas e Juros Mora",
               "ISS DA - Principal", "ISS DA - Multas e Juros", "ISS FECOP"]
COLUNAS_EST = ["ICMS", "ICMS Principal", "ICMS Multas e Juros Mora",
               "ICMS DA - Principal", "ICMS DA - Multas e Juros", "ICMS FECOP"]

# ─────────────────────────────────────────────
# FUNÇÕES AUXILIARES
# ─────────────────────────────────────────────

def build_lookup(mapping_dict):
    """
    Retorna dict: { (cod_conta, ano): metrica }
    para lookup rápido ao processar cada linha.
    """
    lookup = {}
    for (ano_ini, ano_fim), contas in mapping_dict.items():
        for metrica, cod in contas.items():
            for ano in range(ano_ini, ano_fim + 1):
                lookup[(cod, ano)] = metrica
    return lookup

def net_value(row):
    """vlr_receita_realizada - vlr_outras_deducoes, ambas em abs."""
    r = abs(row["vlr_receita_realizada"] or 0)
    d = abs(row["vlr_outras_deducoes"]   or 0)
    return r - d

def processar_csv(csv_path, lookups, chave_grupo):
    """
    Lê um CSV, aplica os lookups e retorna DataFrame pivotado:
    colunas = [chave_grupo..., metrica, valor]
    """
    dtype = {
        "sgl_uf": str, "cod_municipio": str, "num_ano": str, "cod_conta": str,
    }
    df = pd.read_csv(csv_path, dtype=dtype)
    df = df.drop_duplicates()

    df["num_ano"] = pd.to_numeric(df["num_ano"], errors="coerce").astype("Int64")
    df["vlr_receita_realizada"] = pd.to_numeric(df["vlr_receita_realizada"], errors="coerce").fillna(0)
    df["vlr_outras_deducoes"]   = pd.to_numeric(df["vlr_outras_deducoes"],   errors="coerce").fillna(0)

    # Montar lookup unificado para todos os mappings passados
    lookup_total = {}
    for lk in lookups:
        lookup_total.update(build_lookup(lk))

    contas_relevantes = {cod for (cod, _) in lookup_total}
    df = df[df["cod_conta"].isin(contas_relevantes)].copy()
    if df.empty:
        return pd.DataFrame()

    df["_metrica"] = df.apply(
        lambda r: lookup_total.get((r["cod_conta"], r["num_ano"])), axis=1
    )
    df = df[df["_metrica"].notna()].copy()
    if df.empty:
        return pd.DataFrame()

    df["_valor"] = df.apply(net_value, axis=1)

    # Agrupa e pivota
    grp = df.groupby(chave_grupo + ["_metrica"])["_valor"].sum().reset_index()
    pivot = grp.pivot_table(
        index=chave_grupo,
        columns="_metrica",
        values="_valor",
        aggfunc="sum"
    ).reset_index()
    pivot.columns.name = None
    return pivot

# ─────────────────────────────────────────────
# CARREGAR LISTA DE ENTES (cod_ibge dos estados)
# ─────────────────────────────────────────────

def carregar_cod_ibge_estados():
    wb = openpyxl.load_workbook(LISTA_ENTES, read_only=True)
    ws = wb["Estados"]
    mapa = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row[1] and row[2]:
            mapa[str(row[2]).strip()] = str(int(row[1]))
    wb.close()
    return mapa

# ─────────────────────────────────────────────
# PROCESSAR MUNICÍPIOS
# ─────────────────────────────────────────────

def processar_df_como_municipio():
    """
    O DF declara como Estado, então seus dados estão em est_todos.csv.
    Extrai as linhas do DF e aplica lookups de ISS e Cota-Parte,
    retornando pivot com as mesmas colunas dos demais municípios.
    """
    csv_path = CACHE_DIR / "est_todos.csv"
    if not csv_path.exists():
        return pd.DataFrame()

    dtype = {
        "sgl_uf": str, "cod_municipio": str, "num_ano": str, "cod_conta": str,
    }
    df = pd.read_csv(csv_path, dtype=dtype)
    df = df.drop_duplicates()
    df = df[df["sgl_uf"] == "DF"].copy()
    if df.empty:
        return pd.DataFrame()

    # Injeta cod_municipio do DF para reutilizar processar_csv via arquivo temporário
    df["cod_municipio"] = "530010"

    # Usa ISS_DF: conta principal diferente em 2019-2020 e sem FECOP nesse período
    lookup_total = {}
    for lk in [ISS_DF, COTAPARTE]:
        lookup_total.update(build_lookup(lk))

    df["num_ano"] = pd.to_numeric(df["num_ano"], errors="coerce").astype("Int64")
    df["vlr_receita_realizada"] = pd.to_numeric(df["vlr_receita_realizada"], errors="coerce").fillna(0)
    df["vlr_outras_deducoes"]   = pd.to_numeric(df["vlr_outras_deducoes"],   errors="coerce").fillna(0)

    contas_relevantes = {cod for (cod, _) in lookup_total}
    df = df[df["cod_conta"].isin(contas_relevantes)].copy()
    if df.empty:
        return pd.DataFrame()

    df["_metrica"] = df.apply(
        lambda r: lookup_total.get((r["cod_conta"], r["num_ano"])), axis=1
    )
    df = df[df["_metrica"].notna()].copy()
    if df.empty:
        return pd.DataFrame()

    df["_valor"] = df.apply(net_value, axis=1)

    chave_grupo = ["sgl_uf", "cod_municipio", "num_ano"]
    grp = df.groupby(chave_grupo + ["_metrica"])["_valor"].sum().reset_index()
    pivot = grp.pivot_table(
        index=chave_grupo,
        columns="_metrica",
        values="_valor",
        aggfunc="sum"
    ).reset_index()
    pivot.columns.name = None
    return pivot


def processar_municipios():
    csvs = sorted(CACHE_DIR.glob("mun_*.csv"))
    partes = []
    total = len(csvs)

    for i, csv_path in enumerate(csvs, 1):
        uf = csv_path.stem.replace("mun_", "")
        print(f"  [{i:02d}/{total}] Municipios {uf}...", end=" ")

        pivot = processar_csv(
            csv_path,
            lookups=[ISS, COTAPARTE],
            chave_grupo=["sgl_uf", "cod_municipio", "num_ano"]
        )

        if pivot.empty:
            print("sem dados relevantes")
            continue

        print(f"{len(pivot):,} linhas")
        partes.append(pivot)

    # DF: dados municipais estão no CSV de estados (DF é Município e Estado)
    print(f"  [DF] Municipio DF (via est_todos.csv)...", end=" ")
    pivot_df = processar_df_como_municipio()
    if not pivot_df.empty:
        print(f"{len(pivot_df):,} linhas")
        partes.append(pivot_df)
    else:
        print("sem dados relevantes")

    if not partes:
        return pd.DataFrame()

    df = pd.concat(partes, ignore_index=True)

    # ISS = soma das métricas 2 a 6 (se não vier calculado)
    iss_sub = ["ISS Principal", "ISS Multas e Juros Mora",
               "ISS DA - Principal", "ISS DA - Multas e Juros", "ISS FECOP"]
    for col in iss_sub:
        if col not in df.columns:
            df[col] = 0.0
    df["ISS"] = df[iss_sub].fillna(0).sum(axis=1)

    # Garantir todas as colunas métricas
    for col in COLUNAS_MUN:
        if col not in df.columns:
            df[col] = None

    df = df.rename(columns={"sgl_uf": "UF", "cod_municipio": "COD_IBGE", "num_ano": "ANO"})
    df.insert(3, "BASE", "SIOPE")
    df = df[["UF", "COD_IBGE", "ANO", "BASE"] + COLUNAS_MUN]
    df = df.sort_values(["UF", "COD_IBGE", "ANO"]).reset_index(drop=True)
    return df

# ─────────────────────────────────────────────
# PROCESSAR ESTADOS
# ─────────────────────────────────────────────

def processar_estados():
    csv_path = CACHE_DIR / "est_todos.csv"
    print(f"  Estados...", end=" ")

    pivot = processar_csv(
        csv_path,
        lookups=[ICMS],
        chave_grupo=["sgl_uf", "num_ano"]
    )

    if pivot.empty:
        print("sem dados relevantes")
        return pd.DataFrame()

    print(f"{len(pivot):,} linhas")

    # ICMS = soma das métricas 2 a 6
    icms_sub = ["ICMS Principal", "ICMS Multas e Juros Mora",
                "ICMS DA - Principal", "ICMS DA - Multas e Juros", "ICMS FECOP"]
    for col in icms_sub:
        if col not in pivot.columns:
            pivot[col] = 0.0
    pivot["ICMS"] = pivot[icms_sub].fillna(0).sum(axis=1)

    for col in COLUNAS_EST:
        if col not in pivot.columns:
            pivot[col] = None

    cod_ibge_map = carregar_cod_ibge_estados()
    pivot = pivot.rename(columns={"sgl_uf": "UF", "num_ano": "ANO"})
    pivot.insert(1, "COD_IBGE", pivot["UF"].map(cod_ibge_map))
    pivot.insert(3, "BASE", "SIOPE")
    pivot = pivot[["UF", "COD_IBGE", "ANO", "BASE"] + COLUNAS_EST]
    pivot = pivot.sort_values(["UF", "ANO"]).reset_index(drop=True)
    return pivot

# ─────────────────────────────────────────────
# ESCREVER EXCEL
# ─────────────────────────────────────────────

HEADER_FILL = PatternFill("solid", fgColor="1F4E79")
HEADER_FONT = Font(bold=True, color="FFFFFF")

def _escrever_aba(wb, nome_aba, df):
    ws = wb.create_sheet(nome_aba)

    # Cabeçalho
    for col_idx, col_name in enumerate(df.columns, 1):
        cell = ws.cell(row=1, column=col_idx, value=col_name)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center")

    # Dados
    for row_idx, row in enumerate(df.itertuples(index=False), 2):
        for col_idx, val in enumerate(row, 1):
            ws.cell(row=row_idx, column=col_idx, value=val)

    # Larguras
    larguras = {"UF": 6, "COD_IBGE": 12, "ANO": 7, "BASE": 8}
    for col_idx, col_name in enumerate(df.columns, 1):
        w = larguras.get(col_name, 24)
        ws.column_dimensions[get_column_letter(col_idx)].width = w

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

# ─────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────

def main():
    print("=== Dados SIOPE - ICMS e ISS ===\n")

    print("Processando municipios...")
    df_mun = processar_municipios()
    print(f"  Total municipios: {len(df_mun):,} linhas\n")

    print("Processando estados...")
    df_est = processar_estados()
    print(f"  Total estados: {len(df_est):,} linhas\n")

    print(f"Gerando {OUTPUT.name}...")
    wb = openpyxl.Workbook()
    wb.remove(wb.active)  # remove aba padrão

    if not df_mun.empty:
        _escrever_aba(wb, "Municipios", df_mun)
    if not df_est.empty:
        _escrever_aba(wb, "Estados", df_est)

    wb.save(OUTPUT)
    print(f"Arquivo salvo: {OUTPUT}")

if __name__ == "__main__":
    main()
