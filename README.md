# SIOPE — Extração de Dados

Projeto de extração automatizada de dados de receitas públicas do **SIOPE** (Sistema de Informações sobre Orçamentos Públicos em Educação), operacionalizado pelo FNDE/MEC.

---

## Visão Geral

A API REST do SIOPE disponibiliza dados orçamentários declarados pelos entes públicos (Estados, DF e Municípios). Este projeto consome essa API para extrair registros de receitas e organizá-los em planilhas Excel para análise.

**Referência técnica da API:** `SIOPE-TCs_Especificacao_WS_2.1.pdf`

---

## Estrutura do Projeto

```
SIOPE - Extração de Dados/
│
├── README.md                                        ← este arquivo
├── SIOPE-TCs_Especificacao_WS_2.1.pdf               ← manual oficial da API
├── Regra de extração.docx                           ← regras de negócio e mapeamento de contas
├── Lista de Entes.xlsx                              ← cadastro de municípios e estados (fonte dos códigos IBGE)
│
├── Extrair Todas Receitas - Todos os Entes.py       ← script de extração completa via API
├── Dados SIOPE - ICMS e ISS.py                      ← script de consolidação (gera a planilha final)
├── Dados SIOPE - ICMS e ISS.xlsx                    ← planilha consolidada (saída)
│
└── Registro Receita - Todos Entes/                  ← pasta de saída da extração
    ├── progresso.json                               ← controle de resume (tarefas concluídas)
    ├── cache/                                       ← CSVs intermediários por UF (fonte de verdade)
    │   ├── mun_XX.csv                               ← municípios do estado XX
    │   └── est_todos.csv                            ← todos os estados
    ├── Municipios/
    │   ├── Municipios_XX.xlsx                       ← planilha por UF (divide em _parte2, _parte3... se >500k linhas)
    │   └── ...
    └── Estados.xlsx                                 ← todos os 27 estados (divide se >500k linhas)
```

---

## API SIOPE — Informações Técnicas

**Base URL:**
```
https://www.fnde.gov.br/webservices/siope/tc/receita
```

**Parâmetros de consulta:**

| Parâmetro       | Tipo | Descrição                                      | Exemplo   |
|-----------------|------|------------------------------------------------|-----------|
| `num_ano`       | int  | Ano do exercício                               | `2024`    |
| `num_periodo`   | int  | Bimestre (1 a 6). Período 6 = acumulado anual  | `6`       |
| `cod_municipio` | str  | Código IBGE de 6 dígitos do município          | `292740`  |
| `sgl_uf`        | str  | Sigla da UF (para dados estaduais)             | `BA`      |

> **Atenção:** Para municípios, usar `cod_municipio` (6 dígitos — sem o dígito verificador do IBGE).
> Para estados, usar apenas `sgl_uf` sem `cod_municipio`.

**Exemplo de chamada — Município:**
```
GET https://www.fnde.gov.br/webservices/siope/tc/receita?num_ano=2024&num_periodo=6&cod_municipio=292740
```

**Exemplo de chamada — Estado:**
```
GET https://www.fnde.gov.br/webservices/siope/tc/receita?num_ano=2024&num_periodo=6&sgl_uf=BA
```

**Estrutura da resposta (JSON):**
```json
{
  "dat_consulta": "2024-01-01 10:00:00",
  "ente": [{
    "sgl_uf": "BA",
    "cod_municipio": "292740",
    "num_ano": "2024",
    "num_periodo": "6",
    "receita": [{
      "cod_conta": "11145110",
      "des_natureza": "Imposto sobre Serviços de Qualquer Natureza - ISSQN",
      "vlr_previsao_atualizada": 1950000000.00,
      "vlr_receita_realizada": 1979102342.31,
      "vlr_intraorcamentaria": 0,
      "vlr_deducao_fundeb": 0,
      "vlr_outras_deducoes": 0
    }]
  }]
}
```

---

## Mudança de Classificação de Contas (PCASP)

Entre os anos de exercício, o SIOPE migrou para uma nova estrutura do PCASP. Os códigos de conta mudaram:

| Receita                     | Até ~2021 (código antigo) | A partir de ~2022 (código novo) |
|-----------------------------|---------------------------|---------------------------------|
| ISS — Município             | `41113050100`             | `11145110` (total) / `11145111` (principal) |
| Cota-Parte ICMS — Município | `41722010100`             | `17215000`                      |
| ICMS — Estado               | `41113020100`             | `11145010` (total) / `11145011` (principal) |

No formato novo (2022+), cada receita vem detalhada em sub-contas:
- `xxx0` = total (principal + multas + dívida ativa)
- `xxx1` = principal
- `xxx2` = multas e juros de mora
- `xxx3` = dívida ativa
- `xxx4` = multas e juros de mora da dívida ativa

---

## Lista de Entes (`Lista de Entes.xlsx`)

| Aba          | Colunas                           | Registros |
|--------------|-----------------------------------|-----------|
| `Municipios` | `tipo_ente`, `cod_ibge`, `UF`     | 5.571     |
| `Estados`    | `tipo_ente`, `cod_ibge`, `UF`     | 27        |

Os maiores estados em número de municípios: MG (853), SP (645), RS (497), BA (417), PR (399).

---

## Script Principal: `Extrair Todas Receitas - Todos os Entes.py`

### O que faz
Baixa todos os registros de receita (período 6 = acumulado anual) para todos os 5.571 municípios e 27 estados, para os anos de 2019 a 2026, e salva em planilhas Excel organizadas por UF.

### Colunas nas planilhas de saída

| Coluna                    | Descrição                        |
|---------------------------|----------------------------------|
| `sgl_uf`                  | Sigla da UF                      |
| `cod_municipio`           | Código IBGE (vazio para estados) |
| `num_ano`                 | Ano do exercício                 |
| `cod_conta`               | Código da conta de receita       |
| `des_natureza`            | Descrição da receita             |
| `vlr_receita_realizada`   | Valor realizado (R$)             |
| `vlr_intraorcamentaria`   | Valor intra-orçamentário (R$)    |
| `vlr_deducao_fundeb`      | Dedução FUNDEB (R$)              |
| `vlr_outras_deducoes`     | Outras deduções (R$)             |

### Como usar

```bash
# Uso padrão: baixa tudo + compila Excel
python "Extrair Todas Receitas - Todos os Entes.py"

# Apenas baixar dados da API (sem gerar Excel)
python "Extrair Todas Receitas - Todos os Entes.py" --baixar

# Apenas compilar Excel a partir do cache já baixado
python "Extrair Todas Receitas - Todos os Entes.py" --compilar

# Redownload completo, ignorando todo o cache anterior
python "Extrair Todas Receitas - Todos os Entes.py" --forcar
```

### Parâmetros configuráveis (topo do script)

| Variável                  | Padrão | Descrição                                         |
|---------------------------|--------|---------------------------------------------------|
| `ANOS`                    | 2019–2026 | Intervalo de anos a extrair                    |
| `PERIODO`                 | `6`    | Bimestre (6 = acumulado anual)                    |
| `MAX_LINHAS`              | 500.000 | Limite de linhas por arquivo Excel               |
| `MAX_WORKERS`             | `4`    | Threads paralelas de download                     |
| `DELAY_ENTRE_CHAMADAS`    | `0.20s`| Pausa entre requisições por thread                |
| `MAX_TENTATIVAS`          | `3`    | Tentativas em caso de erro na API                 |

### Mecanismo de resume

O script salva o progresso em `Registro Receita - Todos Entes/progresso.json`. Se interrompido (inclusive com Ctrl+C), ao re-executar ele retoma automaticamente de onde parou.

**Proteção contra duplicatas no resume:**
Ao retomar, o script escaneia os CSVs existentes em `cache/` e reconstrói o conjunto de tarefas já gravadas — independente do que estiver no `progresso.json`. Isso garante que, mesmo que o JSON esteja desatualizado (ex.: interrupção entre saves), nenhuma tarefa já presente no CSV será re-executada e gravada em duplicata. O progresso é salvo a cada 10 tarefas concluídas (além de sempre ao final ou ao interromper com Ctrl+C).

Os dados baixados ficam nos CSVs em `cache/` — que são a fonte de verdade. O Excel pode ser recompilado a qualquer momento com `--compilar` sem precisar re-baixar.

### Estimativa de volume e tempo

| Item                        | Quantidade         |
|-----------------------------|--------------------|
| Chamadas à API              | ~45.800            |
| Tempo estimado (4 threads)  | 40–60 minutos      |
| Linhas estimadas — Municípios | ~10–15 milhões   |
| Linhas estimadas — Estados  | ~200–300 mil       |
| Estados com divisão de arquivo | MG, SP, RS (estimado) |

### Dependências Python

```
requests
openpyxl
```
Ambas disponíveis no Anaconda por padrão. Python 3.13+ compatível.

---

---

## Script de Consolidação: `Dados SIOPE - ICMS e ISS.py`

### O que faz
Lê os CSVs de cache gerados pelo script principal e produz a planilha `Dados SIOPE - ICMS e ISS.xlsx` com duas abas consolidadas, aplicando as regras de extração por faixa de ano (mudança de classificação PCASP).

### Regra de valor
Para todas as métricas: `valor = abs(vlr_receita_realizada) - abs(vlr_outras_deducoes)`

### Aba Municípios — colunas

| Coluna | Descrição |
|---|---|
| `UF` | Sigla do estado |
| `COD_IBGE` | Código IBGE de 6 dígitos do município |
| `ANO` | Ano do exercício |
| `BASE` | Sempre "SIOPE" |
| `Cota-Parte` | Cota-Parte do ICMS recebida pelo município |
| `ISS` | Total ISS (soma das 5 sub-métricas abaixo) |
| `ISS Principal` | ISS arrecadado (principal) |
| `ISS Multas e Juros Mora` | Multas e juros sobre ISS |
| `ISS DA - Principal` | Dívida ativa do ISS |
| `ISS DA - Multas e Juros` | Multas e juros da dívida ativa do ISS |
| `ISS FECOP` | Adicional ISS - Fundo Municipal de Combate à Pobreza |

### Aba Estados — colunas

| Coluna | Descrição |
|---|---|
| `UF` | Sigla do estado |
| `COD_IBGE` | Código IBGE do estado (2 dígitos) |
| `ANO` | Ano do exercício |
| `BASE` | Sempre "SIOPE" |
| `ICMS` | Total ICMS (soma das 5 sub-métricas abaixo) |
| `ICMS Principal` | ICMS arrecadado (principal) |
| `ICMS Multas e Juros Mora` | Multas e juros sobre ICMS |
| `ICMS DA - Principal` | Dívida ativa do ICMS |
| `ICMS DA - Multas e Juros` | Multas e juros da dívida ativa do ICMS |
| `ICMS FECOP` | Adicional ICMS - Fundo Estadual de Combate à Pobreza |

### Mapeamento de códigos de conta por faixa de ano

| Métrica | 2019–2020 | 2021–2022 | 2023–2026 |
|---|---|---|---|
| **ISS Principal** | `41113050100` | `11180231` | `11145111` |
| **ISS Multas e Juros** | `41911400000` | `11180232` | `11145112` |
| **ISS DA - Principal** | `41931130000` | `11180233` | `11145113` |
| **ISS DA - Multas e Juros** | `41913130000` | `11180234` | `11145114` |
| **ISS FECOP** | `41113050200` | `11180240` | `11145120` |
| **Cota-Parte** | `41722010100` | `17280110` | `17215000` |
| **ICMS Principal** | `41113020100` | `11180211` | `11145011` |
| **ICMS Multas e Juros** | `41911420000` | `11180212` | `11145012` |
| **ICMS DA - Principal** | `41931150000` | `11180213` | `11145013` |
| **ICMS DA - Multas e Juros** | `41913150000` | `11180215` | `11145014` |
| **ICMS FECOP** | `41113020200` | `11180220` | `11145020` |

### Tratamento especial do DF

O Distrito Federal é ao mesmo tempo **Estado e Município**, portanto:
- Seus dados de **ISS** e **Cota-Parte do ICMS** (receitas municipais) estão em `est_todos.csv`, não em um `mun_DF.csv`.
- O script extrai essas linhas diretamente de `est_todos.csv`, filtrando `sgl_uf == "DF"`, e as inclui na **aba Municipios** com `COD_IBGE = 530010`.
- Os dados de **ICMS** do DF continuam na **aba Estados** normalmente.
- O DF **não recebe Cota-Parte do ICMS** (é o próprio estado arrecadador), portanto essa coluna ficará vazia para o DF.

#### Contas de ISS exclusivas do DF em 2019-2020

Em 2019-2020 o DF usa contas diferentes dos demais municípios e **não possui ISS FECOP**:

| Métrica | Conta DF | Conta demais municípios |
|---|---|---|
| ISS Principal | `41113050000` | `41113050100` |
| ISS Multas e Juros Mora | `41911400000` | `41911400000` (igual) |
| ISS DA – Principal | `41931130000` | `41931130000` (igual) |
| ISS DA – Multas e Juros | `41913130000` | `41913130000` (igual) |
| ISS FECOP | — (sem FECOP) | `41113050200` |

> A partir de 2021 o DF passa a usar os mesmos códigos que os demais municípios.

O script utiliza o dicionário `ISS_DF` para aplicar essas regras exclusivamente às linhas do DF.

### Deduplicação automática

O script aplica `drop_duplicates()` em cada CSV ao lê-lo, eliminando automaticamente linhas repetidas antes de qualquer cálculo. Isso garante resultados corretos mesmo que o CSV de origem contenha duplicatas geradas por interrupções no script de extração.

### Como usar
```bash
python "Dados SIOPE - ICMS e ISS.py"
```
Requer que os CSVs de cache já estejam gerados. Saída: `Dados SIOPE - ICMS e ISS.xlsx` na raiz do projeto.

---

## Histórico de Alterações

| Data       | Descrição                                                                 |
|------------|---------------------------------------------------------------------------|
| 2026-05-13 | Criação do projeto. Fase exploratória com testes para Salvador e Bahia (2019/2024) para identificar os códigos de conta e a mudança de classificação PCASP entre períodos. |
| 2026-05-13 | Criação do script `Extrair Todas Receitas - Todos os Entes.py`. Extração completa para todos os 5.571 municípios e 27 estados, anos 2019–2026, período 6. |
| 2026-05-13 | Criação do script `Dados SIOPE - ICMS e ISS.py`. Consolidação de ISS (municípios), Cota-Parte do ICMS (municípios) e ICMS (estados) com regras de extração por faixa de ano. Resultado: 38.482 linhas (municípios) + 181 linhas (estados). |
| 2026-05-14 | Remoção dos arquivos de teste (`extrair_receitas.py` e `receitas_siope_teste.xlsx`), utilizados apenas na fase exploratória inicial. |
| 2026-05-19 | Correção (`Dados SIOPE - ICMS e ISS.py`): DF incluído na aba Municipios com COD_IBGE 530010, extraindo ISS e Cota-Parte de `est_todos.csv`. Em 2019-2020 o DF usa contas exclusivas (`41113050000` para ISS Principal, sem FECOP) — tratadas pelo dicionário `ISS_DF`. |
| 2026-05-19 | Correção (`Dados SIOPE - ICMS e ISS.py`): conta de ICMS Principal 2019-2020 corrigida de `41113020000` para `41113020100`. |
| 2026-05-19 | Correção (CSVs de cache): removidas linhas duplicadas dos arquivos `mun_GO.csv` (2.280), `mun_RO.csv` (580) e `est_todos.csv` (456), geradas por interrupções e retomadas da extração. |
| 2026-05-19 | Melhoria (`Dados SIOPE - ICMS e ISS.py`): adicionado `drop_duplicates()` na leitura de todos os CSVs para proteção automática contra duplicatas residuais. |
| 2026-05-19 | Melhoria (`Extrair Todas Receitas - Todos os Entes.py`): resume reconstrói o progresso diretamente dos CSVs ao iniciar (evita re-download de dados já gravados); progresso salvo a cada 10 tarefas (antes: 50); Ctrl+C salva o progresso antes de encerrar. |
