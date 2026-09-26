"""
Extrair Todas Receitas - Todos os Entes
========================================
Baixa os registros de receita (período 6) do SIOPE para todos os municípios
e estados, anos 2019-2025, e salva em planilhas Excel organizadas por UF.

Estrutura de saída:
  Registro Receita - Todos Entes/
    cache/          → CSVs intermediários (fonte de verdade para resume)
    Municipios/     → Municipios_XX.xlsx  (um por UF, divide se >500k linhas)
    Estados.xlsx    → todos os estados (divide se >500k linhas)

Uso:
  python "Extrair Todas Receitas - Todos os Entes.py"            # baixa + compila
  python "Extrair Todas Receitas - Todos os Entes.py" --baixar   # só baixa
  python "Extrair Todas Receitas - Todos os Entes.py" --compilar # só compila Excel
  python "Extrair Todas Receitas - Todos os Entes.py" --forcar   # redownload tudo (faz backup do cache anterior)
"""

import os
import sys
import csv
import json
import time
import shutil
import threading
import argparse
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path

import requests
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

# ─────────────────────────────────────────────
# CONFIGURAÇÕES
# ─────────────────────────────────────────────
BASE_DIR       = Path(__file__).parent
LISTA_ENTES    = BASE_DIR / "Lista de Entes.xlsx"
OUTPUT_DIR     = BASE_DIR / "Registro Receita - Todos Entes"
HISTORICO_DIR  = BASE_DIR / "Dados Antigos"   # arquivo morto das extrações anteriores
CACHE_DIR      = OUTPUT_DIR / "cache"
MUN_DIR        = OUTPUT_DIR / "Municipios"
EST_DIR        = OUTPUT_DIR
PROGRESSO_FILE = OUTPUT_DIR / "progresso.json"

API_URL        = "https://www.fnde.gov.br/webservices/siope/tc/receita"
ANOS           = list(range(2019, 2026))   # 2019 a 2025
PERIODO        = 6
MAX_LINHAS     = 500_000
MAX_WORKERS    = 4
DELAY_ENTRE_CHAMADAS = 0.20   # segundos entre requisições por thread
MAX_TENTATIVAS = 3

COLUNAS = [
    "sgl_uf", "cod_municipio", "num_ano",
    "cod_conta", "des_natureza",
    "vlr_receita_realizada", "vlr_intraorcamentaria",
    "vlr_deducao_fundeb", "vlr_outras_deducoes",
]

# ─────────────────────────────────────────────
# ARQUIVAMENTO DA EXTRAÇÃO ANTERIOR
# ─────────────────────────────────────────────
def fazer_backup_output():
    """
    Move OUTPUT_DIR para 'Dados Antigos/Extração DD-MM-AA' antes de iniciar
    uma nova extração completa, preservando o histórico da última extração.
    A data usada é a do arquivo mais antigo no cache (data inicial da extração).

    A planilha consolidada vigente (`Dados SIOPE - ICMS e ISS.xlsx`) é copiada
    junto, para que cada pasta do histórico contenha o cache, os Excel por UF e
    a consolidação correspondente àquela extração.
    """
    if not OUTPUT_DIR.exists():
        return

    csvs = list(CACHE_DIR.glob("*.csv")) if CACHE_DIR.exists() else []
    if csvs:
        data_ts  = min(p.stat().st_mtime for p in csvs)
        data_str = time.strftime("%d-%m-%y", time.localtime(data_ts))
    else:
        data_str = time.strftime("%d-%m-%y")

    HISTORICO_DIR.mkdir(exist_ok=True)

    nome_backup = HISTORICO_DIR / f"Extração {data_str}"
    if nome_backup.exists():
        sufixo = 2
        while True:
            candidato = HISTORICO_DIR / f"Extração {data_str} - {sufixo}"
            if not candidato.exists():
                nome_backup = candidato
                break
            sufixo += 1

    print(f"Arquivando extração anterior em: {HISTORICO_DIR.name}/{nome_backup.name}")
    OUTPUT_DIR.rename(nome_backup)

    # Leva a planilha consolidada vigente para dentro do histórico
    consolidada = BASE_DIR / "Dados SIOPE - ICMS e ISS.xlsx"
    if consolidada.exists():
        destino = nome_backup / f"Dados SIOPE - ICMS e ISS ({data_str}).xlsx"
        shutil.copy2(consolidada, destino)
        print(f"  Planilha consolidada copiada: {destino.name}")

    print("Arquivamento concluído.\n")


# ─────────────────────────────────────────────
# PROGRESSO (thread-safe)
# ─────────────────────────────────────────────
_lock_progresso = threading.Lock()
_lock_csv       = threading.Lock()

def carregar_progresso():
    if PROGRESSO_FILE.exists():
        with open(PROGRESSO_FILE, "r") as f:
            return set(json.load(f))
    return set()

def salvar_progresso(concluidos):
    """
    Grava o progresso de forma atômica (arquivo temporário + os.replace) e
    tolerante a falhas: se a escrita falhar (ex.: OSError [Errno 22] por lock
    transitório de sincronização em nuvem), apenas avisa e segue — os CSVs em
    cache/ são a fonte de verdade e o resume é reconstruído a partir deles.
    Antes esta falha propagava pela thread e abortava a extração inteira.
    """
    tmp = PROGRESSO_FILE.with_suffix(".json.tmp")
    try:
        with open(tmp, "w") as f:
            json.dump(list(concluidos), f)
        os.replace(tmp, PROGRESSO_FILE)
    except OSError as e:
        print(f"  [AVISO] Não foi possível salvar o progresso ({e}); extração continua.")
        try:
            if tmp.exists():
                tmp.unlink()
        except OSError:
            pass

def marcar_concluido(concluidos, chave):
    with _lock_progresso:
        concluidos.add(chave)
        if len(concluidos) % 10 == 0:   # salva a cada 10 (antes: 50)
            salvar_progresso(concluidos)

def reconstruir_do_cache():
    """
    Escaneia os CSVs existentes e retorna o conjunto de chaves já gravadas.
    Garante que ao retomar, tarefas já no CSV não sejam refeitas mesmo que
    o progresso.json esteja desatualizado (ex.: interrupção entre saves).

    Usa csv.reader com acesso por índice (mais rápido que DictReader).
    Índices: 0=sgl_uf, 1=cod_municipio, 2=num_ano
    """
    concluidos = set()

    # Municípios
    for csv_path in sorted(CACHE_DIR.glob("mun_*.csv")):
        try:
            with open(csv_path, "r", encoding="utf-8", newline="") as f:
                reader = csv.reader(f)
                next(reader, None)   # pula cabeçalho
                for row in reader:
                    if len(row) >= 3 and row[1] and row[2]:
                        concluidos.add(f"mun_{row[1].strip()}_{row[2].strip()}")
        except Exception as e:
            print(f"  [AVISO] Erro ao escanear {csv_path.name}: {e}")

    # Estados
    est_path = CACHE_DIR / "est_todos.csv"
    if est_path.exists():
        try:
            with open(est_path, "r", encoding="utf-8", newline="") as f:
                reader = csv.reader(f)
                next(reader, None)
                for row in reader:
                    if len(row) >= 3 and row[0] and row[2]:
                        concluidos.add(f"est_{row[0].strip()}_{row[2].strip()}")
        except Exception as e:
            print(f"  [AVISO] Erro ao escanear est_todos.csv: {e}")

    return concluidos

# ─────────────────────────────────────────────
# LEITURA DA LISTA DE ENTES
# ─────────────────────────────────────────────
def ler_entes():
    wb = openpyxl.load_workbook(LISTA_ENTES, read_only=True)

    municipios = []
    for row in wb["Municipios"].iter_rows(min_row=2, values_only=True):
        if row[1] and row[2]:
            municipios.append({"cod_municipio": str(int(row[1])), "sgl_uf": str(row[2]).strip()})

    estados = []
    for row in wb["Estados"].iter_rows(min_row=2, values_only=True):
        if row[2]:
            estados.append({"sgl_uf": str(row[2]).strip()})

    wb.close()
    return municipios, estados

# ─────────────────────────────────────────────
# CHAMADA À API
# ─────────────────────────────────────────────
def fetch_receita(session, params):
    for tentativa in range(1, MAX_TENTATIVAS + 1):
        try:
            r = session.get(API_URL, params=params, timeout=30)
            r.raise_for_status()
            return r.json().get("ente", [])
        except Exception as e:
            if tentativa == MAX_TENTATIVAS:
                print(f"  [ERRO] {params} após {MAX_TENTATIVAS} tentativas: {e}")
                return None
            time.sleep(2 * tentativa)
    return None

def extrair_linhas(entes, sgl_uf, cod_municipio, ano):
    linhas = []
    for ente in entes:
        for rec in ente.get("receita", []):
            linhas.append([
                sgl_uf,
                cod_municipio or "",
                ano,
                rec.get("cod_conta", ""),
                rec.get("des_natureza", ""),
                rec.get("vlr_receita_realizada"),
                rec.get("vlr_intraorcamentaria"),
                rec.get("vlr_deducao_fundeb"),
                rec.get("vlr_outras_deducoes"),
            ])
    return linhas

# ─────────────────────────────────────────────
# ESCRITA NO CACHE CSV
# ─────────────────────────────────────────────
def csv_path_municipio(uf):
    return CACHE_DIR / f"mun_{uf}.csv"

def csv_path_estados():
    return CACHE_DIR / "est_todos.csv"

def _append_csv(caminho, linhas):
    existe = caminho.exists()
    with _lock_csv:
        with open(caminho, "a", newline="", encoding="utf-8") as f:
            w = csv.writer(f)
            if not existe:
                w.writerow(COLUNAS)
            w.writerows(linhas)

# ─────────────────────────────────────────────
# TAREFA INDIVIDUAL (executada por thread)
# ─────────────────────────────────────────────
def tarefa_municipio(session, ente, ano, concluidos):
    cod  = ente["cod_municipio"]
    uf   = ente["sgl_uf"]
    chave = f"mun_{cod}_{ano}"

    with _lock_progresso:
        ja_feito = chave in concluidos
    if ja_feito:
        return chave, 0, True  # pulado

    params = {"num_ano": ano, "num_periodo": PERIODO, "cod_municipio": cod}
    entes_ret = fetch_receita(session, params)
    time.sleep(DELAY_ENTRE_CHAMADAS)

    if entes_ret is None:
        return chave, 0, False  # erro

    linhas = extrair_linhas(entes_ret, uf, cod, ano)
    if linhas:
        _append_csv(csv_path_municipio(uf), linhas)

    marcar_concluido(concluidos, chave)
    return chave, len(linhas), True

def tarefa_estado(session, estado, ano, concluidos):
    uf    = estado["sgl_uf"]
    chave = f"est_{uf}_{ano}"

    with _lock_progresso:
        ja_feito = chave in concluidos
    if ja_feito:
        return chave, 0, True

    params = {"num_ano": ano, "num_periodo": PERIODO, "sgl_uf": uf}
    entes_ret = fetch_receita(session, params)
    time.sleep(DELAY_ENTRE_CHAMADAS)

    if entes_ret is None:
        return chave, 0, False

    linhas = extrair_linhas(entes_ret, uf, None, ano)
    if linhas:
        _append_csv(csv_path_estados(), linhas)

    marcar_concluido(concluidos, chave)
    return chave, len(linhas), True

# ─────────────────────────────────────────────
# DOWNLOAD
# ─────────────────────────────────────────────
def baixar(forcar=False):
    municipios, estados = ler_entes()

    if forcar:
        print("Modo --forcar: fazendo backup do cache anterior e iniciando extração limpa...")
        fazer_backup_output()
        OUTPUT_DIR.mkdir(exist_ok=True)
        CACHE_DIR.mkdir(exist_ok=True)
        MUN_DIR.mkdir(exist_ok=True)
        concluidos = set()
    else:
        OUTPUT_DIR.mkdir(exist_ok=True)
        CACHE_DIR.mkdir(exist_ok=True)
        MUN_DIR.mkdir(exist_ok=True)
        concluidos = carregar_progresso()
        if CACHE_DIR.exists() and any(CACHE_DIR.glob("*.csv")):
            print("Verificando cache para resume seguro (pode levar alguns instantes)...")
            concluidos_csv = reconstruir_do_cache()
            novos = concluidos_csv - concluidos
            if novos:
                print(f"  {len(novos):,} tarefas encontradas no cache mas ausentes do progresso.json — incorporadas.")
                concluidos |= concluidos_csv
                salvar_progresso(concluidos)
            else:
                print(f"  Cache e progresso.json sincronizados ({len(concluidos):,} tarefas já concluídas).")

    # --- Contagem total de tarefas ---
    tarefas_mun  = [(e, a) for e in municipios for a in ANOS]
    tarefas_est  = [(e, a) for e in estados    for a in ANOS]
    total        = len(tarefas_mun) + len(tarefas_est)
    pendentes_m  = sum(1 for e, a in tarefas_mun if f"mun_{e['cod_municipio']}_{a}" not in concluidos)
    pendentes_e  = sum(1 for e, a in tarefas_est if f"est_{e['sgl_uf']}_{a}" not in concluidos)
    pendentes    = pendentes_m + pendentes_e

    print(f"\nTotal de tarefas : {total:,}  |  Pendentes: {pendentes:,}  |  Já concluídas: {total - pendentes:,}")
    print(f"Municípios: {len(municipios):,}  |  Estados: {len(estados)}  |  Anos: {ANOS[0]}-{ANOS[-1]}  |  Threads: {MAX_WORKERS}\n")

    concluidas = total - pendentes
    erros = 0

    session = requests.Session()

    def _progresso(chave, n_linhas, ok):
        nonlocal concluidas, erros
        concluidas += 1
        if not ok:
            erros += 1
        pct = concluidas / total * 100
        tipo = "MUN" if chave.startswith("mun") else "EST"
        status = f"  [{tipo}] {chave:35s} -> {n_linhas:4d} linhas  | {concluidas:6,}/{total:,} ({pct:5.1f}%)  erros={erros}"
        print(status, end="\r")

    # --- Municípios ---
    print("Baixando municípios...")
    try:
        with ThreadPoolExecutor(max_workers=MAX_WORKERS) as ex:
            futures = {
                ex.submit(tarefa_municipio, session, e, a, concluidos): (e, a)
                for e, a in tarefas_mun
            }
            for fut in as_completed(futures):
                chave, n, ok = fut.result()
                _progresso(chave, n, ok)
    except KeyboardInterrupt:
        print("\n\n[INTERROMPIDO] Salvando progresso...")
        salvar_progresso(concluidos)
        print(f"Progresso salvo ({len(concluidos):,} tarefas). Retome sem --forcar para continuar.")
        return

    salvar_progresso(concluidos)
    print(f"\n\nMunicípios concluídos. Erros acumulados: {erros}")

    # --- Estados ---
    print("\nBaixando estados...")
    try:
        with ThreadPoolExecutor(max_workers=MAX_WORKERS) as ex:
            futures = {
                ex.submit(tarefa_estado, session, e, a, concluidos): (e, a)
                for e, a in tarefas_est
            }
            for fut in as_completed(futures):
                chave, n, ok = fut.result()
                _progresso(chave, n, ok)
    except KeyboardInterrupt:
        print("\n\n[INTERROMPIDO] Salvando progresso...")
        salvar_progresso(concluidos)
        print(f"Progresso salvo ({len(concluidos):,} tarefas). Retome sem --forcar para continuar.")
        return

    salvar_progresso(concluidos)
    print(f"\n\nEstados concluídos. Erros totais: {erros}\n")

# ─────────────────────────────────────────────
# COMPILAÇÃO EXCEL
# ─────────────────────────────────────────────
HEADER_FILL = PatternFill("solid", fgColor="1F4E79")
HEADER_FONT = Font(bold=True, color="FFFFFF")
COL_WIDTHS  = {
    "sgl_uf": 7, "cod_municipio": 13, "num_ano": 7,
    "cod_conta": 15, "des_natureza": 65,
    "vlr_receita_realizada": 22, "vlr_intraorcamentaria": 20,
    "vlr_deducao_fundeb": 18, "vlr_outras_deducoes": 18,
}
NUMERIC_COLS = {
    "vlr_receita_realizada", "vlr_intraorcamentaria",
    "vlr_deducao_fundeb", "vlr_outras_deducoes"
}

def _novo_workbook(caminho):
    wb = openpyxl.Workbook(write_only=True)
    ws = wb.create_sheet("Receitas SIOPE")
    # cabeçalho
    header_cells = []
    for col_idx, col in enumerate(COLUNAS, start=1):
        cell = openpyxl.cell.WriteOnlyCell(ws, value=col)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center")
        header_cells.append(cell)
    ws.append(header_cells)
    # largura das colunas
    for i, col in enumerate(COLUNAS, start=1):
        ws.column_dimensions[get_column_letter(i)].width = COL_WIDTHS.get(col, 15)
    return wb, ws

def _gravar_csv_para_excel(csv_file, base_path, base_name):
    """Le um CSV e grava em um ou mais xlsx (divide em >MAX_LINHAS)."""
    if not csv_file.exists():
        print(f"  [AVISO] Cache nao encontrado: {csv_file.name} - pulando.")
        return 0

    def _caminho(p):
        if p == 1:
            return base_path / f"{base_name}.xlsx"
        return base_path / f"{base_name}_parte{p}.xlsx"

    part       = 1
    row_count  = 0
    total_rows = 0
    wb, ws     = _novo_workbook(_caminho(part))

    with open(csv_file, "r", encoding="utf-8", newline="") as f:
        reader = csv.DictReader(f)
        for linha in reader:
            if row_count >= MAX_LINHAS:
                wb.save(_caminho(part))
                wb.close()
                print(f"    -> {_caminho(part).name} ({row_count:,} linhas)")
                part += 1
                row_count = 0
                wb, ws = _novo_workbook(_caminho(part))

            row_data = []
            for col in COLUNAS:
                val = linha.get(col, "")
                if col in NUMERIC_COLS:
                    try:
                        val = float(val) if val not in ("", "None", None) else None
                    except ValueError:
                        val = None
                row_data.append(val)
            ws.append(row_data)
            row_count  += 1
            total_rows += 1

    wb.save(_caminho(part))
    wb.close()
    print(f"    -> {_caminho(part).name} ({row_count:,} linhas)")
    return total_rows

def compilar():
    MUN_DIR.mkdir(exist_ok=True)

    # — Municípios por UF —
    print("Compilando municípios por UF...")
    csvs_mun = sorted(CACHE_DIR.glob("mun_*.csv"))
    if not csvs_mun:
        print("  Nenhum cache de município encontrado. Execute com --baixar primeiro.")
    for csv_file in csvs_mun:
        uf = csv_file.stem.replace("mun_", "")
        print(f"  UF {uf}...")
        n = _gravar_csv_para_excel(csv_file, MUN_DIR, f"Municipios_{uf}")
        print(f"    Total: {n:,} linhas")

    # — Estados —
    print("\nCompilando estados...")
    n = _gravar_csv_para_excel(csv_path_estados(), EST_DIR, "Estados")
    print(f"  Total: {n:,} linhas")

    print("\nCompilação concluída!")

# ─────────────────────────────────────────────
# ENTRY POINT
# ─────────────────────────────────────────────
def main():
    parser = argparse.ArgumentParser(description="Extração de receitas SIOPE - todos os entes")
    parser.add_argument("--baixar",   action="store_true", help="Somente baixar dados da API")
    parser.add_argument("--compilar", action="store_true", help="Somente compilar CSVs em Excel")
    parser.add_argument("--forcar",   action="store_true", help="Redownload completo ignorando cache")
    args = parser.parse_args()

    apenas_baixar   = args.baixar
    apenas_compilar = args.compilar
    forcar          = args.forcar

    if not apenas_compilar:
        baixar(forcar=forcar)
    if not apenas_baixar:
        compilar()

if __name__ == "__main__":
    main()
