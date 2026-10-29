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
├── Regra de extração.docx                           ← regras de negócio e mapeamento de contas (versão vigente)
├── Regras de extração - Histórico/                  ← versões anteriores da regra, com data e versão no nome
├── Lista de Entes.xlsx                              ← cadastro de municípios e estados (fonte dos códigos IBGE)
│
├── Extrair Todas Receitas - Todos os Entes.py       ← script de extração completa via API
├── Dados SIOPE - ICMS e ISS.py                      ← script de consolidação (gera a planilha final)
├── Comparar Extrações.py                            ← script de comparação entre duas extrações
├── Dados SIOPE - ICMS e ISS.xlsx                    ← planilha consolidada (saída)
│
├── Registro Receita - Todos Entes/                  ← pasta de saída da extração VIGENTE
│   ├── progresso.json                               ← controle de resume (tarefas concluídas)
│   ├── cache/                                       ← CSVs intermediários por UF (fonte de verdade)
│   │   ├── mun_XX.csv                               ← municípios do estado XX
│   │   └── est_todos.csv                            ← todos os estados
│   ├── Municipios/
│   │   ├── Municipios_XX.xlsx                       ← planilha por UF (divide em _parte2, _parte3... se >500k linhas)
│   │   └── ...
│   └── Estados.xlsx                                 ← todos os 27 estados (divide se >500k linhas)
│
└── Dados Antigos/                                   ← HISTÓRICO das extrações anteriores
    ├── Extração DD-MM-AA/                           ← uma pasta por extração arquivada
    │   ├── cache/                                   ← CSVs daquela extração
    │   ├── Municipios/                              ← Excel por UF daquela extração
    │   ├── Estados.xlsx
    │   ├── progresso.json
    │   └── Dados SIOPE - ICMS e ISS (DD-MM-AA).xlsx ← consolidação vigente naquela data
    ├── Compilações anteriores/                      ← planilhas substituídas por mudança de REGRA (mesma extração)
    │   └── Extração DD-MM-AA - regra vN/            ← planilha + script que a gerou
    └── Comparações/                                 ← relatórios de diferença entre extrações ou regras
        └── Comparação <anterior> vs <nova>.txt
```

> A pasta `Dados Antigos/` é o arquivo morto do projeto: guarda cada extração
> passada por inteiro (cache, Excel por UF e a planilha consolidada daquela data),
> as compilações substituídas por mudança de regra e os relatórios de comparação.
> Cada extração ocupa ~700 MB de cache.

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
Baixa todos os registros de receita (período 6 = acumulado anual) para todos os 5.571 municípios e 27 estados, para os anos de 2019 a 2025, e salva em planilhas Excel organizadas por UF.

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

# Redownload completo: faz backup do cache anterior e inicia extração limpa
python "Extrair Todas Receitas - Todos os Entes.py" --forcar
```

### Parâmetros configuráveis (topo do script)

| Variável                  | Padrão | Descrição                                         |
|---------------------------|--------|---------------------------------------------------|
| `ANOS`                    | 2019–2025 | Intervalo de anos a extrair                    |
| `PERIODO`                 | `6`    | Bimestre (6 = acumulado anual)                    |
| `MAX_LINHAS`              | 500.000 | Limite de linhas por arquivo Excel               |
| `MAX_WORKERS`             | `4`    | Threads paralelas de download                     |
| `DELAY_ENTRE_CHAMADAS`    | `0.20s`| Pausa entre requisições por thread                |
| `MAX_TENTATIVAS`          | `3`    | Tentativas em caso de erro na API                 |

### Arquivamento automático da extração anterior (`--forcar`)

Ao executar com `--forcar`, o script **move** a pasta `Registro Receita - Todos Entes` para `Dados Antigos/Extração DD-MM-AA` — onde a data é a do arquivo mais antigo no cache (data de início da extração anterior) — e então cria uma pasta limpa para a nova extração. Nada é deletado.

Junto com a pasta, o script **copia a planilha consolidada vigente** (`Dados SIOPE - ICMS e ISS.xlsx`) para dentro do arquivo, renomeada como `Dados SIOPE - ICMS e ISS (DD-MM-AA).xlsx`. Assim cada pasta do histórico é autocontida: cache, Excel por UF e a consolidação daquela data.

Se já existir um arquivamento com a mesma data, é adicionado sufixo numérico (`- 2`, `- 3`, ...).

**Rotina recomendada a cada atualização da base:**

```bash
# 1. Extrai tudo de novo (arquiva a extração anterior automaticamente)
python "Extrair Todas Receitas - Todos os Entes.py" --forcar

# 2. Regera a planilha consolidada
python "Dados SIOPE - ICMS e ISS.py"

# 3. Documenta o que mudou em relação à extração anterior
python "Comparar Extrações.py" ^
    "Dados Antigos/Extração <data anterior>/Dados SIOPE - ICMS e ISS (<data anterior>).xlsx" ^
    "Dados SIOPE - ICMS e ISS.xlsx" ^
    "Dados Antigos/Comparações/Comparação <data anterior> vs <data nova>.txt"
```

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

O script implementa a **`Regra de extração.docx` versão 4** (28/10/2026). O documento é a referência completa das regras; esta seção resume como elas estão no código.

### Regra de valor
- **Geral:** `valor = abs(vlr_receita_realizada) - abs(vlr_outras_deducoes)`, conta a conta.
- **Limiar de 10%** (`LIMIAR_DEDUCAO`), só no **ICMS dos Estados** e nas **Compensacoes**: se a dedução passar de 10% da receita realizada da mesma conta, ela é ignorada e vale só `abs(vlr_receita_realizada)`. Motivo: alguns entes lançam como "dedução" o repasse de 25% aos municípios ou o FUNDEB.
- **2019–2020, deduções em conta própria** (nesses anos a coluna `vlr_outras_deducoes` vem zerada):
  - ISS: `49111305000` é subtraída do ISS Principal, sem limiar.
  - ICMS: `49111302000` é subtraída do ICMS Principal e `49111302001` do ICMS FECOP, com o limiar de 10%.
- A dedução do FUNDEB (`vlr_deducao_fundeb`) não é descontada em nenhuma métrica.

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
| `Compensacoes` | Cota-parte da compensação das perdas de ICMS (LC 194/2022), 2023+. Não entra no ISS nem na Cota-Parte |

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
| `Compensacoes` | Compensação das perdas de ICMS (LC 194/2022), 2023+. Não entra no ICMS |

### Mapeamento de códigos de conta por faixa de ano

Contas separadas por "+" são somadas; "−" indica dedução.

| Métrica | 2019–2020 | 2021–2022 | 2023–2026 |
|---|---|---|---|
| **ISS Principal** | `41113050100` − `49111305000` | `11180231` | `11145111` |
| **ISS Multas e Juros** | `41911400000` | `11180232` + `11180235` + `11180236` | `11145112` + `11145115` + `11145116` |
| **ISS DA - Principal** | `41931130000` | `11180233` | `11145113` |
| **ISS DA - Multas e Juros** | `41913130000` | `11180234` + `11180237` + `11180238` | `11145114` + `11145117` + `11145118` |
| **ISS FECOP** | `41113050200` | `11180240` | `11145120` |
| **Cota-Parte** | `41722010100` | `17280110` | `17215000` |
| **ICMS Principal** | `41113020100` − `49111302000`* | `11180211` | `11145011` |
| **ICMS Multas e Juros** | `41911420000` | `11180212` + `11180215` + `11180216` | `11145012` + `11145015` + `11145016` |
| **ICMS DA - Principal** | `41931150000` | `11180213` | `11145013` |
| **ICMS DA - Multas e Juros** | `41913150000` | `11180214` + `11180217` + `11180218` | `11145014` + `11145017` + `11145018` |
| **ICMS FECOP** | `41113020200` − `49111302001`* | `11180220` | `11145020` |
| **Compensacoes (Municípios)** | — | — | `17295300` + `17196200` + `17196300` |
| **Compensacoes (Estados)** | — | — | `17196200` + `17196300` (+ `17295300`) |

\* No ICMS dos Estados e nas Compensacoes, toda dedução (coluna ou conta própria) está sujeita ao limiar de 10%.

Nas contas de 8 dígitos, o último dígito indica o desdobramento: 2 e 4 são a forma **agrupada** de multas e juros; 5 a 8, a forma **separada**. Cada ente usa uma das duas, por isso as duas são somadas.

### Compensacoes — regra de duplicidade
Se, no mesmo ente-ano, a soma da rubrica dos Estados (`17196200` + `17196300`) for **igual** à rubrica dos Municípios (`17295300`), é o mesmo valor lançado duas vezes: o script mantém só a rubrica dos Municípios (função `aplicar_duplicidade_compensacao`). Valores diferentes são tratados como parcelas distintas e somados. Os municípios que lançaram na rubrica dos Estados entram normalmente, porque município só recebe essa compensação como cota-parte.

### Tratamento especial do DF

O Distrito Federal é ao mesmo tempo **Estado e Município**, portanto:
- Seus dados de **ISS** e **Cota-Parte do ICMS** (receitas municipais) estão em `est_todos.csv`, não em um `mun_DF.csv`.
- O script extrai essas linhas diretamente de `est_todos.csv`, filtrando `sgl_uf == "DF"`, e as inclui na **aba Municipios** com `COD_IBGE = 530010`.
- Os dados de **ICMS** do DF continuam na **aba Estados** normalmente.
- O DF **não recebe Cota-Parte do ICMS** (é o próprio estado arrecadador), portanto essa coluna ficará vazia para o DF.
- A **compensação** da LC 194 o DF recebe como Estado: ela entra na **aba Estados**, e na aba Municipios a coluna `Compensacoes` do DF fica vazia.
- Em 2019–2020 a dedução `49111305000` também é subtraída do ISS Principal do DF (`41113050000`).

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

## Script de Comparação: `Comparar Extrações.py`

### O que faz
Compara duas versões da planilha consolidada e gera um relatório em texto com tudo que mudou entre uma extração e outra. Serve para documentar cada atualização da base — os entes podem retificar valores ou passar a declarar exercícios que antes estavam em branco.

### Como usar
```bash
python "Comparar Extrações.py" <planilha_anterior> <planilha_nova> <relatorio_saida>
```

Os relatórios ficam em `Dados Antigos/Comparações/`.

### O que o relatório traz (por aba)

| Seção | Conteúdo |
|---|---|
| Contagens | linhas antes/depois, novas, removidas, em comum, alteradas |
| Linhas novas | entes/anos que passaram a declarar — totais por ANO e por UF, mais o detalhe UF/COD_IBGE/ANO |
| Linhas removidas | entes/anos que deixaram de constar na fonte |
| Resumo por métrica | quantas linhas mudaram em cada métrica e o delta somado em R$ |
| Retificações | cada linha alterada, valor a valor: antes → depois, com a diferença |
| Totais por métrica | soma da base inteira, antes e depois |

A tolerância de comparação é de R$ 0,005 (meio centavo), para ignorar ruído de ponto flutuante.

Colunas que existem em só uma das versões (por exemplo, uma métrica nova) são comparadas contra vazio e listadas no cabeçalho de cada aba. O script também serve para comparar duas compilações da **mesma extração** com regras diferentes. Nesse caso, onde o relatório diz "retificações", leia "valores alterados pela regra".

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
| 2026-06-05 | Ajuste (`Extrair Todas Receitas - Todos os Entes.py`): faixa de anos alterada de 2019–2026 para 2019–2025 (período 6 de 2026 não disponível até o final do ano). `--forcar` passa a fazer backup automático do cache anterior (renomeia a pasta com data da extração original) ao invés de deletar, preservando o histórico da última extração. |
| 2026-09-25 | Reestruturação do histórico: criada a pasta `Dados Antigos/`, com uma subpasta por extração arquivada (cache + Excel por UF + consolidação daquela data) e `Comparações/` para os relatórios de diferença. Migrados para lá o backup de junho (`Extração 05-06-26`), o cache duplicado de junho que existia na pasta viva (`Extração 05-06-26 - 2`), a consolidação de 08-09-26 (`Extração 08-09-26 (apenas consolidação)` — só o resultado, o cache dessa data não está nesta pasta) e o comparativo junho×setembro feito manualmente. |
| 2026-09-25 | Ajuste (`Extrair Todas Receitas - Todos os Entes.py`): `--forcar` passa a **mover** a extração anterior para `Dados Antigos/Extração DD-MM-AA` (antes criava pastas irmãs na raiz) e copia junto a planilha consolidada vigente, tornando cada pasta do histórico autocontida. |
| 2026-09-25 | Correção (`Extrair Todas Receitas - Todos os Entes.py`): `salvar_progresso()` grava de forma atômica (arquivo temporário + `os.replace`) e não aborta mais a extração em caso de falha de escrita. Relevante nesta pasta por estar sob sincronização do OneDrive, onde locks transitórios podem causar `OSError [Errno 22]` — antes essa falha propagava pela thread e derrubava a extração inteira. |
| 2026-09-25 | Criação do script `Comparar Extrações.py`: gera relatório de diferenças entre duas versões da planilha consolidada (linhas novas, removidas, retificações valor a valor e totais por métrica). Relatórios salvos em `Dados Antigos/Comparações/`. |
| 2026-09-25 | Nova extração completa. 39.179 tarefas, 0 erros de API, 0 avisos de escrita, 7.039.603 linhas de cache (16:14–18:29). Resultado: 38.929 linhas (municípios) + 182 linhas (estados). Diferenças vs. 08-09-26: **MG passou a declarar ICMS estadual de 2024** (antes MG não tinha nenhum ano na aba Estados) — impacto de +R$ 81,53 bilhões, elevando o total de ICMS de R$ 4,231 tri para R$ 4,313 tri; 1 município novo (TO/170190/2025); 1 retificação de ISS (BA/292540/2024, de R$ 977.993,00 para R$ 136.765,77). |
| 2026-09-25 | Observação sobre a fonte: o SIOPE segue sem declaração estadual de MG para 2019-2023 e 2025, e sem RS/2019 — a API retorna `ente: []`. Por isso a aba Estados tem 182 linhas, e não 189. |
| 2026-10-28 | Regra de extração revisada (versões 2, 3 e 4 do documento, todas em 28/10/2026). Versões anteriores do documento arquivadas na nova pasta `Regras de extração - Histórico/`. Resumo: ISS e ICMS passam a somar as contas separadas de multas e juros (último dígito 5 a 8); a `11180215` ("Multas") sai de ICMS DA – Multas e Juros e vai para ICMS Multas e Juros; deduções de 2019–2020 em conta própria (`49111305000` no ISS; `49111302000`/`49111302001` no ICMS); limiar de 10% nas deduções do ICMS dos Estados e das Compensacoes; nova métrica de compensação das perdas de ICMS (LC 194/2022). |
| 2026-10-28 | Script `Dados SIOPE - ICMS e ISS.py` reescrito para a regra v4: mapeamento com várias contas por métrica, deduções em conta própria, limiar de 10% (`LIMIAR_DEDUCAO`), regra de duplicidade das compensações e nova coluna **`Compensacoes`** (última coluna das duas abas). Script anterior guardado junto com a planilha anterior. |
| 2026-10-28 | **Nova compilação sobre a MESMA extração de 25/09/2026**: o cache não mudou, só as regras de negócio. Compilação anterior (regra v1) arquivada em `Dados Antigos/Compilações anteriores/Extração 25-09-26 - regra v1/`, junto com o script que a gerou. Resultado: 38.929 linhas (municípios) + 182 (estados), as mesmas chaves. Diferenças pela regra (relatório em `Dados Antigos/Comparações/Comparação 25-09-26 regra v1 vs regra v4 (mesma extração).txt`): ISS −R$ 233,2 mi em 2019–2020 e +R$ 721,6 mi em 2021–2025; ICMS −R$ 61,7 mi em 2019–2020 e +R$ 117,7 bi em 2021–2025 (sobretudo deduções acima de 10%, que deixaram de ser descontadas; ex.: MG/2024 passa de R$ 81,5 bi para R$ 101,7 bi); Cota-Parte sem alteração; Compensacoes: R$ 2,52 bi nos municípios e R$ 21,86 bi nos estados (2023–2025). Alteradas 5.278 linhas na aba Municipios e 107 na aba Estados. |
| 2026-10-28 | Ajuste (`Comparar Extrações.py`): passa a aceitar colunas presentes em só uma das versões (comparadas contra vazio), necessário para a nova coluna `Compensacoes`. |
| 2026-10-28 | Observação: 4 municípios-ano ficam com ISS negativo. Três (PI/220375/2020, PI/220300/2020, RS/430690/2020) vêm da dedução `49111305000` maior que o próprio ISS e estão registrados na seção 2.5 da regra. O quarto, RS/430471/2021, **já era negativo na regra anterior**: o município lançou na coluna de deduções R$ 161.572,55, valor maior que o ISS Principal de R$ 139.466,74. |
