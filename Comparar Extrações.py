"""
Comparar Extrações
==================
Compara duas versões da planilha consolidada `Dados SIOPE - ICMS e ISS.xlsx`
e gera um relatório de diferenças em texto.

Serve para documentar, a cada atualização da base, o que mudou em relação à
extração anterior: entes que passaram a declarar, retificações de valores e
eventuais remoções.

Uso:
  python "Comparar Extrações.py" <planilha_anterior> <planilha_nova> <relatorio_saida>

Exemplo:
  python "Comparar Extrações.py" ^
      "Dados Antigos/Extração 05-06-26/Dados SIOPE - ICMS e ISS (05-06-26).xlsx" ^
      "Dados SIOPE - ICMS e ISS.xlsx" ^
      "Dados Antigos/Extração 08-09-26/Comparação 05-06-26 vs 08-09-26.txt"
"""

import sys
from pathlib import Path

import pandas as pd

CHAVE = ["UF", "COD_IBGE", "ANO"]
TOLERANCIA = 0.005   # diferenças abaixo de meio centavo são ruído de ponto flutuante


def comparar_aba(dn, dv, aba, out):
    """Compara uma aba e escreve o resultado em `out` (lista de linhas)."""
    metricas = [c for c in dn.columns if c not in CHAVE + ["BASE"]]

    kn = dn.set_index(CHAVE)
    kv = dv.set_index(CHAVE)

    novas     = kn.index.difference(kv.index)
    removidas = kv.index.difference(kn.index)
    comuns    = kn.index.intersection(kv.index)

    out.append("")
    out.append("=" * 78)
    out.append(f"ABA: {aba}")
    out.append("=" * 78)
    out.append(f"  linhas anterior : {len(dv):>8,}")
    out.append(f"  linhas nova     : {len(dn):>8,}")
    out.append(f"  linhas novas    : {len(novas):>8,}")
    out.append(f"  linhas removidas: {len(removidas):>8,}")
    out.append(f"  linhas em comum : {len(comuns):>8,}")

    a = kn.loc[comuns, metricas].fillna(0)
    b = kv.loc[comuns, metricas].fillna(0)
    dif = (a - b).abs() > TOLERANCIA
    alteradas = dif.any(axis=1)
    out.append(f"  linhas ALTERADAS: {int(alteradas.sum()):>8,}")

    # ── Novas linhas ────────────────────────────────────────────
    if len(novas):
        out.append("")
        out.append(f"  --- {len(novas):,} LINHAS NOVAS (entes que passaram a declarar) ---")
        por_ano = pd.Series([i[2] for i in novas]).value_counts().sort_index()
        out.append("    por ANO: " + ", ".join(f"{ano}={n}" for ano, n in por_ano.items()))
        por_uf = pd.Series([i[0] for i in novas]).value_counts()
        out.append("    por UF : " + ", ".join(f"{uf}={n}" for uf, n in por_uf.items()))
        out.append("    detalhe (UF / COD_IBGE / ANO):")
        for uf, cod, ano in sorted(novas):
            out.append(f"      {uf}  {cod}  {ano}")

    # ── Linhas removidas ────────────────────────────────────────
    if len(removidas):
        out.append("")
        out.append(f"  --- {len(removidas):,} LINHAS REMOVIDAS (deixaram de constar na fonte) ---")
        for uf, cod, ano in sorted(removidas):
            out.append(f"      {uf}  {cod}  {ano}")

    # ── Retificações ────────────────────────────────────────────
    if int(alteradas.sum()):
        out.append("")
        out.append("  --- RESUMO DAS ALTERAÇÕES POR MÉTRICA ---")
        for m in metricas:
            n = int(dif[m].sum())
            if n:
                delta = (a[m] - b[m]).sum()
                out.append(f"      {m:<26} {n:>5} linhas   delta R$ {delta:+,.2f}")

        por_ano = pd.Series([i[2] for i in comuns[alteradas]]).value_counts().sort_index()
        out.append("    alterações por ANO: " + ", ".join(f"{ano}={n}" for ano, n in por_ano.items()))

        out.append("")
        out.append(f"  --- {int(alteradas.sum()):,} RETIFICAÇÕES (detalhe valor a valor) ---")
        for idx in sorted(comuns[alteradas]):
            out.append(f"      {idx[0]} / {idx[1]} / {idx[2]}")
            for m in metricas:
                va, vb = a.loc[idx, m], b.loc[idx, m]
                if abs(va - vb) > TOLERANCIA:
                    out.append(f"          {m:<26} {vb:>18,.2f}  ->  {va:>18,.2f}"
                               f"   ({va - vb:+,.2f})")

    # ── Totais por métrica ──────────────────────────────────────
    out.append("")
    out.append("  --- TOTAIS POR MÉTRICA (base inteira) ---")
    for m in metricas:
        ta, tb = dn[m].sum(), dv[m].sum()
        out.append(f"      {m:<26} anterior R$ {tb:>20,.2f}   nova R$ {ta:>20,.2f}"
                   f"   ({ta - tb:+,.2f})")


def main():
    if len(sys.argv) != 4:
        print(__doc__)
        sys.exit(1)

    p_velho, p_novo, p_saida = (Path(x) for x in sys.argv[1:4])
    for p in (p_velho, p_novo):
        if not p.exists():
            print(f"ERRO: arquivo não encontrado: {p}")
            sys.exit(1)

    out = []
    out.append("RELATÓRIO DE COMPARAÇÃO — Dados SIOPE (ICMS e ISS)")
    out.append("")
    out.append(f"  extração anterior: {p_velho}")
    out.append(f"  extração nova    : {p_novo}")
    out.append(f"  tolerância       : R$ {TOLERANCIA}")

    for aba in ("Municipios", "Estados"):
        dn = pd.read_excel(p_novo,  sheet_name=aba)
        dv = pd.read_excel(p_velho, sheet_name=aba)
        comparar_aba(dn, dv, aba, out)

    texto = "\n".join(out) + "\n"
    p_saida.parent.mkdir(parents=True, exist_ok=True)
    p_saida.write_text(texto, encoding="utf-8")

    # Ecoa no console tudo menos os blocos longos de detalhe
    for linha in out:
        if not linha.startswith("      "):
            print(linha)
    print(f"\nRelatório salvo em: {p_saida}")


if __name__ == "__main__":
    main()
