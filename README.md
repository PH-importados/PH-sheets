# 🧾 Precificadora Inteligente

Ferramenta local de precificação automatizada para varejo. Processa arquivos **XML (NFe)** e **CSV (Sistema)** e gera planilhas Excel com cálculos completos de custo, impostos e margem, além de um dashboard visual de análise da nota.

---

## 🚀 Tecnologias Utilizadas

| Camada | Tecnologia |
|--------|-----------|
| Backend | Python 3.x + Flask |
| Processamento | Pandas (merge/fallback EAN → REF) |
| XML | ElementTree (extração de NFe) |
| Excel | **Openpyxl puro** (sem pandas.to_excel — garante fórmulas no Mac/Windows) |
| Frontend | HTML5 + Bootstrap + JavaScript nativo (Drag & Drop) |
| Impostos externos | API SEFAZ AL (ST e Antecipado) |

---

## 🛠️ Instalação e Configuração

### macOS (M1/M2/M3)

```bash
# 1. Clonar o repositório
git clone https://github.com/seu-usuario/ph-SheetGenerator.git
cd ph-SheetGenerator

# 2. Criar e ativar ambiente virtual
python3 -m venv venv
source venv/bin/activate

# 3. Instalar dependências
pip install Flask pandas openpyxl requests
```

> **Nota:** o pacote `requests` é necessário para a consulta à API da SEFAZ AL (ST/Antecipado).

---

## 📖 Como Usar

1. Execute o servidor Flask:
```bash
python3 app.py
```

2. O navegador abrirá automaticamente em `http://127.0.0.1:8080`.
3. Preencha o **Nome do Fornecedor** e o **Número da Nota**.
4. Arraste o **XML da NFe** e o **CSV do sistema** para as áreas indicadas.
5. Ajuste os parâmetros financeiros (multiplicador, impostos, frete). O crédito de ICMS vem da própria nota — não há campo para ele.
6. Clique em **Gerar Prévia** para visualizar os dados e o dashboard na tela.
7. Clique em **Baixar Excel** para salvar a planilha estilizada com fórmulas.

---

## 🧠 Lógica de Processamento

### Busca de Preço (Fallback)
1. Tenta cruzar pelo **EAN** (código de barras)
2. Se não encontrar preço válido (ou preço = R$ 0,01), tenta pela **Referência/Código**
3. Se ainda não encontrar, o preço de venda é calculado automaticamente via `arredondar_99(Custo Real × 2)`

### Detecção de Embalagem
O sistema lê a descrição do XML (ex: *"CAIXA COM 12"*, *"CAIXA COM 24"*) e multiplica o preço unitário do sistema pela quantidade da embalagem para calcular o **Preço Atual** correto.

**Proteção anti-absurdo:** se o resultado da multiplicação for mais que 3× o custo esperado, o sistema assume que o cadastro já tem o preço da caixa e não multiplica (ex: grampos com 5000 unidades).

### Fluxos de Cálculo

#### Produto SEM Substituição Tributária
```
Custo Entrada = NF unit + IPI + Frete(10%) + Despesa(10%) - Crédito ICMS(variável)
Custo Real    = Custo Entrada × Multiplicador
Custo Saída   = Custo Real + Federal(9,13%) + Cartão(4%) + ICMS Saída(21%)
Preço Mín.    = Custo Saída ÷ (1 - Meta%)
```

> O **Crédito de ICMS** é o valor **destacado na nota** para cada item: `vICMS` (fornecedor no regime normal) ou `vCredICMSSN` (fornecedor do Simples). Se a nota não destaca ICMS — ex: Simples CSOSN 102/103 — o crédito é zero.
> As faixas Importado 4% | SP 7% | PE 12% | AL 19% servem só para colorir a coluna. Detalhes em `RULES.md` §3.

#### Produto COM Substituição Tributária (ST)
```
Custo Entrada = NF unit + ST + ANT + IPI + Frete(10%) + Despesa(10%)
Custo Real    = Custo Entrada × Multiplicador
Custo Saída   = Custo Real + Federal(9,13%) + Cartão(4%) + ICMS Saída(21%)
Preço Mín.    = Custo Saída ÷ (1 - Meta%)
```

> ST e Antecipado são consultados automaticamente na **API da SEFAZ AL**.  
> O campo **Crédito de ICMS** é zerado automaticamente quando há ST.

---

## 🎨 Formatação da Planilha (Excel)

### Estrutura de Colunas
| Col | Campo | Tipo |
|-----|-------|------|
| A–E | NF, Descrição, REF, SKU, QTD | Fixo (XML) |
| F–I | NF Unit, ST, ANT, IPI | Fixo (XML) |
| J–L | Frete, Despesa, Cred. ICMS | **Fórmula** |
| M–N | Custo Entrada, Custo Real | **Fórmula** |
| O | CST | Fixo (XML) |
| P–S | Federal, Cartão, ICMS Saída, Custo Saída | **Fórmula** |
| T | Meta % | **Editável** (padrão 15%, por produto) |
| U | Preço Mín. Viável | **Fórmula** |
| V | Preço Atual (sistema × emb.) | Fixo (CSV) |
| W | Preço Varejo | **Fórmula** |
| X | Margem Real | **Fórmula** |
| Y–Z | P.Unit Sistema, Qtd Emb. | Auditoria (cinza) |

### Cores
- 🔵 **Azul:** Preço Varejo e Margem Real
- 🟡 **Amarelo:** Preço Mínimo Viável
- 🟢 **Verde:** Coluna META % (editável por produto)
- 🍑 **Pêssego:** Linhas com ST > 0
- ⚫ **Cinza:** Colunas de auditoria (P.Unit Sistema / Qtd Emb.)
- 🔴 **Vermelho (só nas células ST/ANT):** valor de ST/ANT **estimado por rateio** — ver seção [Conferência de ST/ANT rateado](#-conferência-de-stant-rateado) abaixo

### ⚠️ Conferência de ST/ANT rateado

Quando uma nota tem vários produtos diferentes com a **mesma descrição** (ex: variações de cor/tamanho do mesmo item), o sistema casa cada produto do XML com o imposto correspondente retornado pela SEFAZ AL, na mesma ordem em que aparecem na nota — isso resolve corretamente o caso mais comum.

Em casos raros, se a quantidade de produtos com aquela descrição na nota **não bater** com a quantidade de valores que a SEFAZ retornou pra ela, o sistema não arrisca adivinhar qual valor pertence a qual produto: ele **divide o total proporcionalmente** entre os produtos e sinaliza a linha para conferência manual:

- A **descrição do produto** ganha o prefixo `⚠ CONFERIR IMPOSTO —`.
- Só as **células de ST e ANT** dessa linha ficam **vermelhas** (o resto da linha mantém a cor normal — pêssego, verde menta, azul, etc).
- No rodapé da planilha, abaixo da legenda de crédito ICMS, aparece uma legenda vermelha explicando o alerta.

Nesses casos, confira manualmente o valor de ST/ANT desses produtos antes de fechar o preço.

### Comportamento das Fórmulas
Todos os campos calculados são gravados como **fórmulas Excel reais** — alterar qualquer parâmetro na linha 2 recalcula toda a planilha automaticamente. O arquivo é gerado com `fullCalcOnLoad=True` para garantir recálculo imediato ao abrir no Mac.

---

## 📊 Dashboard

Após processar, um painel é exibido abaixo da tabela com:
- Total de itens, quantos com ST e sem ST
- Produtos sem preço no sistema (usarão preço calculado)
- Valor total da NF
- Margem estimada média
- Barra de distribuição ST vs. Sem ST
- Lista de alertas de produtos sem preço
- Lista de alertas de produtos com ST/ANT estimado por rateio (ver [Conferência de ST/ANT rateado](#-conferência-de-stant-rateado))

---

## 📂 Estrutura do Projeto

```
ph-SheetGenerator/
├── app.py                  # Servidor Flask (porta 8080)
├── core/
│   ├── __init__.py
│   └── processador.py      # Motor de cálculo, geração Excel e dashboard HTML
├── static/
│   └── logo.png
└── templates/
    └── index.html          # Interface web
```

---

## 📦 Gerar Executável

### macOS (.app)

```bash
# Instalar PyInstaller (uma vez)
pip install pyinstaller

# Gerar o .app
pyinstaller --noconfirm --onefile --windowed \
    --add-data "templates:templates" \
    --add-data "static:static" \
    --icon "static/logo.png" \
    --name "PrecificadoraDaOnda" \
    app.py
```

### Windows (.exe)

```powershell
pip install pyinstaller

pyinstaller --noconfirm --onefile --windowed `
    --add-data "templates;templates" `
    --add-data "static;static" `
    --icon "static/logo.png" `
    --name "PrecificadoraDaOnda" `
    app.py
```

> **Nota Windows:** use `;` para separar caminhos no `--add-data`. No Mac/Linux use `:`.

O executável gerado estará em `dist/PrecificadoraDaOnda` (Mac) ou `dist/PrecificadoraDaOnda.exe` (Windows).

---

## 🔄 Como Atualizar

1. Altere o código e teste localmente
2. Rode novamente o comando do PyInstaller (sobrescreve o `dist/` automaticamente)
3. Compacte o `.app` ou `.exe` (`botão direito → Comprimir`)
4. Envie para os usuários via WhatsApp, Drive ou Slack
5. O usuário substitui o ícone antigo pelo novo

> O programa não se atualiza pela internet — a distribuição é manual via arquivo compactado.

---

## ⚠️ Observações Importantes

- Na primeira execução no **Windows**, o SmartScreen pode exibir alerta. Clique em "Mais informações" → "Executar assim mesmo" (o executável não possui assinatura digital paga).
- O arquivo `.app` do Mac **não funciona no Windows** e vice-versa — gere um build separado para cada sistema.
- A consulta à API da SEFAZ AL tem timeout de 10 segundos. Se indisponível, o processamento continua normalmente sem os valores de ST externos.