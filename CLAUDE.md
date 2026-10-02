# PH-Sheets — Precificadora Inteligente

Ferramenta local de precificação automatizada para varejo. Processa arquivos **XML (NFe)** e **CSV (Sistema)** e gera planilhas Excel estilizadas com cálculos completos de custo, impostos e margem, além de um dashboard visual de análise da nota.

---

## Tecnologias

| Camada | Tecnologia | Uso |
|--------|-----------|-----|
| Backend | Python 3.x + Flask | Servidor local na porta 8080 |
| Processamento | Pandas | Merge/fallback EAN → REF entre XML e CSV |
| XML | xml.etree.ElementTree | Extração de itens da NFe |
| Excel | openpyxl (puro) | Geração com fórmulas reais — sem pandas.to_excel para garantir compatibilidade Mac/Windows |
| Frontend | HTML5 + Bootstrap 5 + Vanilla JS | Interface com Drag & Drop, preview da tabela e dashboard |
| Impostos externos | API SEFAZ AL | Consulta automática de ST e Antecipado via chave da NFe |

---

## Estrutura do Projeto

```
PH-Sheets/
├── app.py                  # Servidor Flask (porta 8080) + rotas HTTP
├── core/
│   ├── __init__.py
│   └── processador.py      # Motor de cálculo, geração Excel e dashboard HTML
├── static/
│   └── logo.png
├── templates/
│   └── index.html          # Interface web (form, drop zones, preview, dashboard)
├── tests/
│   ├── conftest.py          # Fixtures e fábrica make_row()
│   ├── test_embalagem.py    # Testes de extrair_qtd_embalagem
│   ├── test_calcular.py     # Testes de _calcular (regras fiscais)
│   ├── test_credito.py      # Testes de extrair_credito_icms e faixas de cor
│   ├── test_excel.py        # Testes de salvar_excel_estilizado
│   ├── test_sefaz_api.py    # Testes de pareamento ST/ANT com a API SEFAZ
│   └── test_helpers.py      # Testes de limpar_preco, limpar_str, _arredondar_x9
├── requirements.txt
├── requirements-dev.txt     # pytest + pytest-cov
├── pytest.ini
└── README.md
```

---

## Fluxo Geral do Sistema

```
Usuário preenche formulário
        │
        ▼
[index.html] POST /processar (FormData)
        │
        ▼
[app.py] Recebe XML + CSV + params
        │
        ├──► Salva XML e CSV em /tmp/
        │
        ▼
[processador.py] gerar_tabela()
        │
        ├──► 1. Parse XML da NFe (ElementTree)
        │        └── Extrai: nNF, chaveNFe, itens (desc, EAN, ref, qtd, vProd, IPI, ST)
        │
        ├──► 2. Consulta API SEFAZ AL (ST + Antecipado por chave da NFe)
        │        └── Timeout 10s — se falhar, continua sem os valores externos
        │
        ├──► 3. Lê CSV do sistema (Pandas)
        │        └── Colunas: BARRA (EAN), REFERÊNCIA, PREÇO
        │
        ├──► 4. Merge XML × CSV
        │        ├── Tentativa 1: join por EAN (ean_xml == ean_sys)
        │        └── Fallback: join por REF (ref_xml == ref_sys) se preço = 0 ou NaN
        │
        ├──► 5. Detecção de embalagem
        │        └── Regex na descrição XML: "CAIXA COM N", "PCT C/N", etc.
        │            Anti-absurdo: se p_sys × N > nf_u × mult × 3, ignora multiplicação
        │
        ├──► 6. Monta lista de rows (dicionários por produto)
        │
        └──► Retorna (rows, params, num_nf)
                │
                ▼
[app.py] Chama salvar_excel_estilizado() → arquivo .xlsx em /tmp/
        │
        ▼
[app.py] Chama gerar_dashboard_html(rows) → HTML inline
        │
        ▼
[app.py] Retorna JSON {tabela HTML (20 produtos), dashboard HTML, download_url}
        │
        ▼
[index.html] Injeta tabela + dashboard no DOM, exibe botão de download
```

---

## Regras de Cálculo da Planilha

### Mapeamento de Colunas

```
A    B        C    D    E    F       G      H       I       J       K       L
NF   DESC     REF  SKU  QTD  NF_U    ST_U   ANT_U   IPI_U   C_REAL  FRETE   DESP

M         N    O       P       Q        R         S        T      U          V        W         X
CRED_ICMS CST  C_ENT   FED     CARTÃO   ICMS_S    C_SAÍDA  META%  PREÇO_MÍN  P_ATUAL  P_VAREJO  MARGEM

Y       Z        AA       AB         AC       AD            AE          AF         AG
NF_ATC  PREÇO_ATC FED_ATC CARTÃO_ATC ICM_ATC  C_SAÍDA_ATC  MARGEM_ATC  AUDIT_SYS  AUDIT_EMB
```

### Linha 2 — Parâmetros Globais (editáveis)

| Coluna | Parâmetro | Padrão |
|--------|-----------|--------|
| J (C_REAL) | Multiplicador varejo | 2.0 |
| K (FRETE) | % Frete/logística | 10% |
| L (DESP) | % Despesas operacionais | 10% |
| P (FED) | % Imposto Federal | 9.13% |
| Q (CARTÃO) | % Taxa cartão | 4% |
| R (ICMS_S) | % ICMS na venda | 21% |
| Y (NF_ATC) | Multiplicador atacado | 1.3 |
| Z (P_ATC) | % Desconto atacado | 15% |

### Fórmulas por Produto (linha 3+) — VAREJO

```
Passo 1 — CUSTO REAL (col J)
    = ROUND(NF_U × MULT$2, 2)
    ← Multiplicador aplicado PRIMEIRO sobre o preço de NF

Passo 2 — FRETE (col K)
    = ROUND(C_REAL × FRETE$2, 2)

Passo 3 — DESPESA (col L)
    = ROUND(C_REAL × DESP$2, 2)

Passo 4 — CRÉDITO ICMS (col M)
    = IF(ST>0 OR CST IN CST_SEM_CREDITO,
        0,
        -ROUND(NF_U × TAXA_CRED, 2))
    ← TAXA_CRED (col AUDIT_CRED) = crédito DESTACADO na NF ÷ vProd
      (vICMS no regime normal, vCredICMSSN no Simples; sem destaque → 0)
    ← Zerado se há ST ou CST isento
    ← ANT NÃO zera o crédito — ver RULES.md seção 3

Passo 5 — CUSTO ENTRADA (col O)
    = ROUND(C_REAL + ST + ANT + IPI + FRETE + DESP + CRED, 2)
    (CRED já é negativo)

Passo 6 — PREÇO VAREJO (col W)
    = IF(P_ATUAL > 0,
        P_ATUAL,
        arredondar_99(C_REAL × 2))
    arredondar_99 = INT(x) - IF(x - INT(x) <= 0.5, 1, 0) + 0.99

Passo 7 — FEDERAL (col P)
    = ROUND(P_VAREJO × FED$2, 2)

Passo 8 — CARTÃO (col Q)
    = ROUND(P_VAREJO × CART$2, 2)

Passo 9 — ICMS SAÍDA (col R)
    = ROUND(P_VAREJO × ICMS$2, 2)

Passo 10 — CUSTO SAÍDA (col S)
    = ROUND(C_ENT + FEDERAL + CARTÃO + ICMS_S, 2)

Passo 11 — META % (col T)
    Padrão 15% — EDITÁVEL por produto individualmente

Passo 12 — PREÇO MÍN VIÁVEL (col U)
    = IF(META > 0, ROUND(C_SAÍDA / (1 - META%), 2), 0)

Passo 13 — MARGEM REAL (col X)
    = IF(P_VAREJO > 0, ROUND((P_VAREJO - C_SAÍDA) / P_VAREJO, 4), 0)
```

### Fórmulas por Produto (linha 3+) — ATACADO

```
Passo 14 — NF ATC (col Y)
    = ROUND(NF_U × MULT_ATC$2, 2)
    ← Multiplicador atacado separado (padrão 1.3 vs 2.0 do varejo)

Passo 15 — PREÇO ATC (col Z)
    = ROUND(P_VAREJO × (1 − DESC_ATC$2), 2)
    ← Preço atacado = Preço varejo com desconto (padrão 15%)

Passo 16 — FEDERAL ATC (col AA)
    = ROUND(NF_ATC × FED$2, 2)
    ← Base de cálculo é NF_ATC (custo atacado), não o preço de venda

Passo 17 — CARTÃO ATC (col AB)
    = ROUND(P_ATC × CART$2, 2)
    ← Base de cálculo é P_ATC (preço de venda atacado)

Passo 18 — ICMS ATC (col AC)
    = ROUND(NF_ATC × ICM$2, 2)
    ← Base de cálculo é NF_ATC

Passo 19 — CUSTO SAÍDA ATC (col AD)
    = ROUND(C_ENT + FED_ATC + CART_ATC + ICM_ATC, 2)
    ← Usa o mesmo C_ENT do varejo (custo de entrada é igual)

Passo 20 — MARGEM ATC PED (col AE)
    = IF(P_ATC_PED > 0, ROUND((P_ATC_PED − C_SAÍDA_ATC) / P_ATC_PED, 4), 0)
    ← usa C_SAÍDA_ATC (custo atacado) como base

Passo 21 — MARGEM ATC PDV (col AF)
    = IF(P_ATC_PDV > 0, ROUND((P_ATC_PDV − C_SAÍDA) / P_ATC_PDV, 4), 0)
    ← usa C_SAÍDA do varejo — regra de negócio: margem PDV reflete os impostos
       de saída do varejo para que o chefe possa substituir P_VAREJO e ver a
       margem real do PDV atacado imediatamente
```

### Crédito de ICMS — valor destacado na NF

O crédito **sempre** vem do ICMS destacado no XML para o item (`extrair_credito_icms()` em `core/processador.py`):

| Fornecedor | Campo do XML | `cred_pct` |
|------------|--------------|------------|
| Regime normal (tag `CST`) | `vICMS` | `vICMS ÷ vProd` |
| Simples Nacional (tag `CSOSN`) | `vCredICMSSN` | `vCredICMSSN ÷ vProd` (`vICMS` de CSOSN 900 não conta) |
| Nada destacado | — | `0` |

- Nunca usar taxa padrão/fallback — não existe mais parâmetro de crédito no formulário.
- ST > 0.005 e `CST_SEM_CREDITO` continuam zerando o crédito como proteção extra.
- As faixas Importado 4% / SP 7% / PE 12% / AL 19% (`CRED_CORES`) servem só para **cor/legenda**; a taxa efetiva é associada à faixa mais próxima (±0,5 p.p.) por `faixa_cred()`.
- Regras completas e histórico (bug NF 875, CSOSN 103): `RULES.md` §3.

### Detecção de Embalagem

```
Regex aplicada na descrição XML:
    (?:CAIXA COM|PACOTE COM|KIT COM|PCT\s*C/|CX\s*C/|C/)\s*(\d+)

Se match e qtd > 1:
    Proteção anti-absurdo:
        IF p_sys × qtd > nf_u × mult × 3
            → qtd_emb = 1  (preço da caixa já no sistema)
        ELSE
            → p_atual = p_sys × qtd_emb
```

---

## Modo Marketplace (Shopee / Mercado Livre)

Fluxo **independente** do varejo/atacado — não alterar `_calcular`, `salvar_excel_estilizado` nem `gerar_dashboard_html` por causa dele. Código em `core/marketplace.py`, tela em `templates/marketplace.html` (`/marketplace`), testes em `tests/test_marketplace.py`. O `index.html` só tem o botão de troca de modo.

Fluxo em 3 passos:
1. `POST /marketplace/produtos` (XML + CSV + variáveis) → reaproveita `gerar_tabela()` e devolve os produtos da nota em JSON
2. Usuário marca na tela os produtos que vão para o marketplace
3. `POST /marketplace/gerar` (JSON com os selecionados) → prévia + dashboard + Excel (abas `Marketplace` e `Faixas`)

Cálculo por produto (`calcular_marketplace`):
```
C. ENTRADA  = mesma conta do varejo, com multiplicador próprio (mult_mkt)
CUSTO BASE  = C. ENTRADA + embalagem de envio
PREÇO BASE  = preço que dá a margem alvo só com Federal + ICMS (sem taxas de marketplace)
PREÇO CANAL = P tal que  P − custo_base − taxas(P) − impostos(P) = margem·P
              taxas dependem da faixa em que o próprio P cai → testa faixa a faixa
              arredonda p/ cima até X,X9 sem pular para a faixa seguinte
```
- Impostos de venda: Federal sobre o preço; ICMS = 0 com ST, e ANT abatido do ICMS (igual ao varejo). Sem taxa de cartão.
- Faixas em `SHOPEE_FAIXAS` / `ML_FAIXAS` (anotação 2024). ML ≥ R$ 79: sem tarifa fixa, vendedor paga frete = `frete_ml × (1 − desconto da reputação)` (verde 50%, amarela 40%, vermelha 0%).
- Colunas definidas uma vez em `COLUNAS_MKT` (prévia e Excel usam a mesma lista). Custo em ordem crescente com prefixos `(×) (+) (−) (=)` até `(=) CUSTO ENTRADA` / `(=) CUSTO BASE`; cada taxa/imposto de venda tem coluna própria.
- Variáveis no **topo da coluna** em que são usadas (Excel: linha 1 grupos, 2 cabeçalhos, **3 variáveis**, dados a partir da 4). Fórmulas apontam para `$COL$3`; Federal/ICMS de Shopee e ML espelham as variáveis da venda base.
- Excel: preços sugeridos são **valores editáveis**; taxas, impostos, lucro e margem são fórmulas (VLOOKUP nas faixas da aba Faixas).
- Fora da v1: Amazon e comparação com o preço da concorrência.

---

## Estrutura da Prévia HTML (primeiros 20 produtos)

A prévia é gerada em `app.py` como HTML inline com **cabeçalho duplo**:
- Linha 1: agrupadores "VAREJO" (`#dbeafe`) e "ATACADO" (`#fef9c3`) com `colspan`
- Linha 2: nomes individuais das 31 colunas (índices 0-based)

### Colunas da Prévia (índice → nome)

| Idx | Nome | Cor (linhas normais) |
|-----|------|----------------------|
| 0–14 | NF, DESC, REF, SKU, QTD, NF UNIT, ST, ANT, IPI, FRETE, DESP, CRED ICMS, C.REAL, C.ENTRADA, CST | sem cor |
| 15–18 | FEDERAL, CARTÃO, ICMS S., C. SAÍDA | Azul `#BDD7EE` bold |
| 19 | META % | sem cor |
| 20 | PREÇO MÍN VRJ | Azul `#BDD7EE` bold |
| 21 | PREÇO ATUAL | sem cor |
| 22 | PREÇO VAREJO | Azul `#BDD7EE` bold |
| 23 | MARGEM | Azul `#BDD7EE` bold |
| 24–33 | NF ATC, PREÇO ATC PEDIDO, PREÇO ATC PDV, FEDERAL ATC, CARTÃO ATC, ICMS ATC, C. SAÍDA ATC, MARGEM ATC, PREÇO PCT ATC, P. COMPRA PCT | Amarelo `#FFF2CC` bold |

> Linhas ST → pêssego `#FCE4D6` em tudo. Linhas ANT → verde menta `#D1FAE5` em tudo. As cores acima só se aplicam a linhas normais.

---

## Regras do Dashboard

O dashboard é gerado em Python como HTML inline e injetado via JSON no frontend.
**Atenção:** `<script>` tags injetadas via `innerHTML` são silenciadas pelo browser — o frontend re-executa via `document.createElement('script')`.

### Métricas exibidas

| Card | Cálculo |
|------|---------|
| Total Itens | `len(rows)` |
| Com ST | `count(st_u > 0.005)` |
| Com ANT | `count(ant_u > 0.005 and st_u <= 0.005)` |
| Normal | `total - com_st - com_ant` |
| Sem Preço no Sistema | `count(p_atual <= 0)` — usarão preço calculado |
| Valor Total NF | `sum(nf_u × qtd)` |
| Margem Estimada | `mean((p_atual - nf_u) / p_atual)` para itens com preço |
| Lucro Estimado | `sum((p_var - c_saida) × qtd)` — requer `metricas` calculadas |

### Barra de Distribuição (3 segmentos)

```
[=== Normal (X%) ===][=== ST (Y%) ===][=== ANT (Z%) ===]
    azul (#BDD7EE)     pêssego (#FCE4D6) verde menta (#D1FAE5)
```

### Gráfico de Lucro por Produto

- Barras CSS animadas por produto, ordenadas por lucro decrescente
- Slider JS "% Vendido" (0–100%) escala proporcionalmente os valores
- Script injetado via `document.createElement('script')` — não via innerHTML

### Alertas

Lista todos os produtos com `p_atual <= 0` — esses usarão o preço calculado automaticamente (`arredondar_99(C_REAL × 2)`).

---

## Cores da Planilha Excel

| Cor | Hex | Aplicação |
|-----|-----|-----------|
| Azul | `#BDD7EE` | Bloco VAREJO: FEDERAL, CARTÃO, ICMS SAÍDA, CUSTO SAÍDA, PREÇO MÍN VIÁVEL VRJ, PREÇO VAREJO, MARGEM REAL |
| Amarelo claro | `#FFF2CC` | Bloco ATACADO completo: NF ATC, PREÇO ATC PEDIDO, PREÇO ATC PDV, FEDERAL ATC, CARTÃO ATC, ICMS ATC, CUSTO SAÍDA ATC, MARGEM ATC, PREÇO PCT ATC, P. COMPRA PCT |
| Verde menta | `#D1FAE5` | Linha inteira de produto com ANT > 0 e ST = 0 |
| Verde | `#E2EFDA` | META % (editável por produto) |
| Pêssego | `#FCE4D6` | Linha inteira de produto com ST > 0 |
| Cinza | `#D9D9D9` | Colunas de auditoria (P.UNIT SISTEMA / QTD EMB / TAXA CRED quando ST/isento) |
| Param | `#F2F2F2` | Cabeçalhos e linha de parâmetros |

> ST e ANT têm cores distintas. Threshold: sempre usar `> 0.005` (não `> 0`) para evitar falsos positivos de ponto flutuante.
> Em linhas ST e ANT, as cores azul/amarelo dos blocos varejo/atacado prevalecem por cima da cor de linha.

---

## Comportamento do Excel

- Todas as colunas calculadas são gravadas como **fórmulas Excel reais** (não valores estáticos).
- A linha 2 contém os parâmetros globais — alterar qualquer célula recalcula toda a planilha.
- `fullCalcOnLoad=True` garante recálculo imediato ao abrir no Mac.
- Painéis congelados em `A3` (cabeçalhos e parâmetros sempre visíveis).
- META % (col T) é editável por linha — padrão 15%, em verde bold.

---

## Executável (PyInstaller)

```bash
# macOS
pyinstaller --noconfirm --onefile --windowed \
    --add-data "templates:templates" \
    --add-data "static:static" \
    --icon "static/logo.png" \
    --name "PrecificadoraDaOnda" \
    app.py

# Windows (PowerShell) — usa ; no --add-data
pyinstaller --noconfirm --onefile --windowed `
    --add-data "templates;templates" `
    --add-data "static;static" `
    --icon "static/logo.png" `
    --name "PrecificadoraDaOnda" `
    app.py
```

Para detectar ambiente frozen: `if getattr(sys, 'frozen', False)` em `app.py` — ajusta os caminhos de template e static via `sys._MEIPASS`.

---

## Testes

### Rodar os testes

```bash
source venv/bin/activate
python -m pytest tests/ -v
```

### Estrutura

```
tests/
├── conftest.py          # fixtures e fábrica make_row()
├── test_embalagem.py    # extrair_qtd_embalagem — 4 padrões + anti-absurdo
├── test_calcular.py     # _calcular — regras fiscais e fórmulas de precificação
├── test_credito.py      # extrair_credito_icms — vICMS / vCredICMSSN por CST/CSOSN + faixas de cor
├── test_excel.py        # salvar_excel_estilizado — geração sem KeyError, fórmulas, cores
├── test_sefaz_api.py    # parear_impostos_api — ST/ANT, duplicidade, rateio
└── test_helpers.py      # limpar_preco, limpar_str, _arredondar_x9
```

### Padrões obrigatórios

**1. Use `make_row()` com overrides mínimos**
Cada teste deve sobrescrever apenas os campos relevantes para o comportamento testado. Isso evita que mudanças no default quebrem testes não relacionados.

```python
# BOM — deixa claro o que está sendo testado
row = make_row(nf_u=10.0, st_u=3.0)

# RUIM — noise: por que todos esses campos?
row = {'nf_u': 10.0, 'st_u': 3.0, 'ant_u': 0.0, 'ipi_u': 0.0, ...}
```

**2. Zere `cred_pct` quando for isolar outra variável**
`_calcular` lê o crédito só de `row['cred_pct']` (sem ele, crédito = 0). O `make_row` tem `cred_pct=0.04` como default. Se o teste não está testando crédito ICMS, passe `cred_pct=0.0` para evitar que o crédito contamine o expected value.

```python
# Testando c_ent sem ruído do crédito
row = make_row(nf_u=10.0, ant_u=2.0, p_atual=25.0, cred_pct=0.0)
assert m['c_ent'] == pytest.approx(12.0)  # 10 + 2
```

**3. Use `P_zero` para isolar, `P` para testar com parâmetros reais**
- `P_zero` (mult=1.0, todos os % zerados): isola um componente sem interferência fiscal
- `P` (defaults do formulário): valida comportamento com valores reais de produção

**4. `pytest.approx()` em todo float**
Nunca comparar floats com `==` direto. Use `pytest.approx()`.

```python
assert m['margem'] == pytest.approx(0.5)       # OK
assert m['margem'] == 0.5                       # NUNCA
```

**5. `@pytest.mark.parametrize` para table-driven tests**
Quando o mesmo comportamento vale para múltiplos inputs (ex: vários CSTs isentos), use parametrize em vez de repetir o teste.

```python
@pytest.mark.parametrize("cst", ['40', '41', '50', '60', '102', '103', '500'])
def test_cst_isento_zera_credito(self, P_zero, cst):
    row = make_row(cst=cst, cred_pct=0.10, p_atual=20.0)
    assert _calcular(row, P_zero)['cred'] == 0.0
```

### O que testar ao adicionar nova feature

| Mudança | Testes obrigatórios |
|---------|-------------------|
| Novo padrão regex de embalagem | Positivo (detecta) + negativo (não detecta falso positivo) + anti-absurdo |
| Nova regra de zeragem de imposto | Testa que zera + testa que outros casos NÃO zeram |
| Novo CST isento | Adiciona a `CST_SEM_CREDITO` (`core/processador.py`) + ao `@parametrize` de `test_cst_isento_zera_credito` |
| Novo grupo/campo de ICMS no XML | Caso em `test_credito.py` com o trecho do XML real |
| Nova fórmula de custo | Testa com valores onde o cálculo manual é trivial (`cred_pct=0.0`, `P_zero`) |

### Comportamentos documentados pelos testes

- `limpar_preco("12.50")` retorna `1250.0` — ponto é tratado como separador de milhar (formato BR). Valor decimal deve usar vírgula: `"12,50"`.
- ANT sozinho (`ant_u > 0`, `st_u = 0`) **não** zera crédito ICMS — apenas ST e CSTs isentos zeram.
- `_calcular` usa só `row['cred_pct']`; sem ele o crédito é 0 (nunca taxa padrão).
- Crédito vem do valor destacado: `vICMS` (CST) ou `vCredICMSSN` (CSOSN). CSOSN 900 com só `vICMS` → crédito 0.

---

## Convenções de Código

- **Fórmulas no Excel:** usar `get_column_letter()` para referenciar colunas pelo nome lógico (ex: `F`, `G`, `N`). Nunca hardcode `"A3"` em fórmulas — sempre derivar da constante `COL`.
- **Merge CSV:** sempre tentar EAN primeiro, depois REF. Nunca assumir que o CSV terá todas as colunas — checar `if 'COLUNA' in df_sys.columns` antes de acessar.
- **SEFAZ AL:** toda chamada à API deve estar em try/except — o processamento não pode depender dela.
- **openpyxl puro:** não usar `pandas.to_excel()` — quebra fórmulas em alguns ambientes Mac.
- **Preço de venda:** nunca usar 0 ou None — aplicar `arredondar_99(C_REAL × 2)` como fallback.
