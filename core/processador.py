import re
import json
import html
import pandas as pd
import xml.etree.ElementTree as ET
import requests
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Border, Side, Alignment, Font
from openpyxl.utils import get_column_letter

# ─── Padrão embalagem ────────────────────────────────────────────────────────
# Padrão 1a — ESPECÍFICO (alta prioridade): "CAIXA COM 12", "CX C/120", "PCT C/24"
# Separado do C/ genérico para evitar que "CONJUNTO C/ 03 PCS... CX C/120"
# capture o 3 em vez do 120 — re.search retorna a primeira ocorrência.
_PATTERN_EMB_ESPECIFICO = re.compile(
    r'(?:CAIXA COM|PACOTE COM|KIT COM|PCT\s*C/|CX\s*C/)\s*(\d+)',
    re.IGNORECASE
)

# Padrão 1b — GENÉRICO (baixa prioridade): "C/5", "C/ 12"
# Só é tentado se nenhum padrão mais específico encontrou match.
_PATTERN_EMB_GENERICO = re.compile(
    r'C/\s*(\d+)',
    re.IGNORECASE
)

# Padrão 2 (Oxford/fornecedores): "12 CANECAS...", "6 TIGELAS..."
# Exclui sufixos de medida para evitar falsos positivos como "2 LITROS", "500 ML", "250 G"
_PATTERN_EMB_INICIO = re.compile(
    r'^(\d+)\s+(?!(?:ML|MG|KG|GR|LT|L|CM|MM|M|UN|PC|PCS|PCT|UNID)\b)',
    re.IGNORECASE
)
# Padrão 3 (Summit/papelaria): "PT24UN", "DP 18 UN", "POTE 48 UN", "CX DP 12"
_PATTERN_EMB_FINAL = re.compile(
    r'(?:PT|DP|POTE|POLYBAG|DISPLAY)\s*(\d+)\s*(?:UN|U(?:\s|$|\]))?',
    re.IGNORECASE
)

# Padrão 4 (Display/DS): "DS NOME DO PRODUTO - 84", "DS NOME - 207 PCS"
# DS = Display; o número após o traço final é a quantidade da embalagem
_PATTERN_EMB_DS = re.compile(
    r'^DS\b.*-\s*(\d+)\s*(?:PCS)?\s*$',
    re.IGNORECASE
)

# Padrão 5 (IP-N): "IP-18 TRB", "IP-120 TRB"
# Convenção de fornecedores (ex: Affinity Trade): "Itens por Pacote"
# Presente no NOME do sistema mas ausente no XML (descrição truncada).
_PATTERN_EMB_IP = re.compile(
    r'\bIP-(\d+)\b',
    re.IGNORECASE
)

# ─── Mapeamento de colunas (1-based) ─────────────────────────────────────────
#
# FLUXO VAREJO:
#   1. C_REAL  = NF_U × MULT
#   2. FRETE   = C_REAL × FRETE%
#   3. DESP    = C_REAL × DESP%
#   4. CRED    = NF_U × CRED%  (zerado se ST/CST isento — ANT não zera)
#   5. C_ENT   = C_REAL + ST + ANT + IPI + FRETE + DESP − CRED
#   6. FEDERAL = P_VAR × FED%
#   7. CARTÃO  = P_VAR × CART%
#   8. ICMS_S  = P_VAR × ICMS%
#   9. C_SAIDA = C_ENT + FEDERAL + CARTÃO + ICMS_S
#  10. P_MIN   = C_SAIDA ÷ (1 − META%)
#
# FLUXO ATACADO:
#  11. NF_ATC      = NF_U × MULT_ATC           (parâm col 25 linha 2)
#  12. P_ATC_PED   = P_VAR × (1 − DESC_PED%)   (parâm col 26 linha 2 — pedido, 15%)
#  13. P_ATC_PDV   = P_VAR × (1 − DESC_PDV%)   (parâm col 27 linha 2 — balcão, 10%)
#  14. FED_ATC     = NF_ATC × FED%
#  15. CART_ATC    = P_ATC_PED × CART%
#  16. ICM_ATC     = NF_ATC × ICM%
#  17. C_SAIDA_ATC = C_ENT + FED_ATC + CART_ATC + ICM_ATC
#  18. MARGEM_ATC  = (P_ATC_PED − C_SAIDA_ATC) / P_ATC_PED
#
# A  B     C    D    E    F      G     H      I      J      K      L
# NF DESC  REF  SKU  QTD  NF_U   ST_U  ANT_U  IPI_U  C_REAL FRETE  DESP
#  1  2     3    4    5    6      7     8      9      10     11     12
#
# M      N    O      P      Q      R       S       T     U      V       W      X
# CRED   CST  C_ENT  FED    CART   ICMS_S  C_SAIDA META  P_MIN  P_ATUAL P_VAR  MARGEM
# 13     14   15     16     17     18      19      20    21     22      23     24
#
# Y       Z          AA         AB       AC        AD      AE             AF             AG         AH            AI         AJ         AK         AL
# NF_ATC  P_ATC_PED  P_ATC_PDV  FED_ATC  CART_ATC  ICM_ATC C_SAIDA_ATC  MARGEM_ATC_PED MARGEM_ATC_PDV P_PCT_ATC P_COMPRA_PCT AUDIT_SYS  AUDIT_EMB  AUDIT_CRED
# 25      26         27         28       29        30      31             32             33             34        35           36         37         38
#
# P_PCT_ATC    = P_ATC_PED × QTD_EMB  → preço do pacote no atacado pedido
# P_COMPRA_PCT = NF_U      × QTD_EMB  → preço de compra do pacote (custo NF por caixa)

COL = {
    'NF': 1, 'DESC': 2, 'REF': 3, 'SKU': 4, 'QTD': 5,
    'NF_U': 6, 'ST_U': 7, 'ANT_U': 8, 'IPI_U': 9,
    'C_REAL': 10, 'FRETE': 11, 'DESP': 12, 'CRED': 13,
    'CST': 14, 'C_ENT': 15,
    'FED': 16, 'CARTAO': 17, 'ICMS_S': 18, 'C_SAIDA': 19,
    'META': 20, 'P_MIN': 21,
    'P_ATUAL': 22, 'P_VAR': 23, 'MARGEM': 24,
    'NF_ATC': 25, 'P_ATC_PED': 26, 'P_ATC_PDV': 27,
    'FED_ATC': 28, 'CART_ATC': 29, 'ICM_ATC': 30,
    'C_SAIDA_ATC': 31,
    'MARGEM_ATC_PED': 32,  # margem sobre preço pedido (15% desc)
    'MARGEM_ATC_PDV': 33,  # margem sobre preço PDV balcão (10% desc)
    'P_PCT_ATC': 34,
    'P_COMPRA_PCT': 35,
    'AUDIT_SYS': 36, 'AUDIT_EMB': 37,
    'AUDIT_CRED': 38,  # taxa de crédito ICMS por produto (colorida por origem, editável)
}
TOTAL_COLS = 38

# ─── Paleta de cores por % de crédito ICMS ───────────────────────────────────
# Usada na coluna CRED do Excel e na prévia HTML
CRED_CORES = {
    0.04: 'FFB347',  # 4%  — laranja vivo   (importado)
    0.07: 'DA70D6',  # 7%  — orquídea       (SP)
    0.12: '40E0D0',  # 12% — turquesa       (PE)
    0.19: 'FF69B4',  # 19% — rosa pink      (AL)
    0.0:  'B0B0B0',  # 0%  — cinza          (ST / isento)
}
CRED_LABELS = {
    0.04: '4% — Importado',
    0.07: '7% — Nacional (SP)',
    0.12: '12% — Nacional (PE)',
    0.19: '19% — Local (AL)',
    0.0:  '0% — ST / Isento',
}

def faixa_cred(pct, tol=0.005):
    """
    Faixa de CRED_CORES mais próxima da taxa efetiva (vICMS / vProd), ou None.
    A taxa efetiva raramente é exata (vBC com IPI/frete, desconto, base reduzida),
    então a cor/legenda usa a faixa nominal mais próxima dentro de `tol`.
    """
    faixa = min(CRED_CORES, key=lambda f: abs(f - pct))
    return faixa if abs(faixa - pct) <= tol else None


def cor_cred(pct, default='FFFFFF'):
    """Hex da cor da faixa de crédito mais próxima de `pct`."""
    faixa = faixa_cred(pct)
    return CRED_CORES[faixa] if faixa is not None else default


def _chave_faixa(pct):
    """Faixa nominal se houver uma próxima, senão a própria taxa (4 casas)."""
    faixa = faixa_cred(pct)
    return faixa if faixa is not None else round(pct, 4)


def _fmt_taxa(pct):
    """'7%' para faixas nominais, '2,56%' para taxas efetivas fora das faixas."""
    if pct in CRED_CORES:
        return f'{pct*100:.0f}%'
    return f'{pct*100:.2f}%'.replace('.', ',')


def cred_fill(pct):
    """Retorna PatternFill para o percentual de crédito ICMS."""
    return PatternFill("solid", fgColor=cor_cred(pct))


# ─── Descrições dos códigos CST/CSOSN em português ───────────────────────────
CST_DESCRICOES = {
    '00':  ('Tributada integralmente',                                          '#E2EFDA'),
    '10':  ('Tributada + ST nas operações seguintes',                           '#FFF2CC'),
    '20':  ('Com redução de base de cálculo',                                   '#E2EFDA'),
    '30':  ('Isenta/não tributada + ST nas operações seguintes',                '#FFF2CC'),
    '40':  ('Isenta de ICMS',                                                   '#F2F2F2'),
    '41':  ('Não tributada',                                                    '#F2F2F2'),
    '50':  ('Suspensão — ICMS suspenso por decisão judicial ou normativa',      '#F2F2F2'),
    '51':  ('Diferimento — pagamento adiado para etapa seguinte da cadeia',     '#F2F2F2'),
    '60':  ('Cobrada anteriormente por ST — imposto já retido na entrada',      '#FCE4D6'),
    '70':  ('Com redução de base de cálculo + ST nas operações seguintes',      '#FFF2CC'),
    '90':  ('Outras (tributação mista ou não enquadrada nos anteriores)',        '#F2F2F2'),
    '101': ('Simples Nacional — tributado com permissão de crédito',            '#E2EFDA'),
    '102': ('Simples Nacional — tributado sem permissão de crédito',            '#F2F2F2'),
    '103': ('Simples Nacional — isento por faixa de receita bruta',             '#F2F2F2'),
    '201': ('Simples Nacional — com crédito + ST nas operações seguintes',      '#FFF2CC'),
    '202': ('Simples Nacional — sem crédito + ST nas operações seguintes',      '#FFF2CC'),
    '203': ('Simples Nacional — isento por faixa + ST nas operações seguintes', '#FFF2CC'),
    '300': ('Imune de ICMS',                                                    '#F2F2F2'),
    '400': ('Não tributada pelo Simples Nacional',                              '#F2F2F2'),
    '500': ('Simples — ICMS cobrado anteriormente por ST ou antecipação',       '#FCE4D6'),
    '900': ('Simples Nacional — outros',                                        '#F2F2F2'),
}


# CST/CSOSN que não geram crédito de ICMS (RULES.md §3.2).
# Simples Nacional só gera crédito com CSOSN 101/201 (e 900 quando houver pCredSN).
CST_SEM_CREDITO = frozenset({
    '40', '41', '50', '60',                        # regime normal: isento/não trib./suspensão/ST
    '102', '103', '202', '203', '300', '400', '500',  # Simples Nacional sem crédito
})


# ─── Helpers ─────────────────────────────────────────────────────────────────
def limpar_str(val):
    if pd.isna(val) or str(val).strip() == "":
        return ""
    s = str(val).replace('="', '').replace('"', '').strip()
    if s.endswith('.0'):
        s = s[:-2]
    return s


def limpar_preco(val):
    if pd.isna(val) or str(val).strip() == "":
        return 0.0
    try:
        return float(str(val).replace('R$', '').replace('.', '').replace(',', '.').strip())
    except:
        return 0.0


def extrair_credito_icms(imposto, ns):
    """
    Lê o grupo ICMS de um <imposto> e devolve (cst, v_cred).

    v_cred é o crédito de ICMS DESTACADO na nota para o item (valor total, R$):
      - Regime normal (tag CST):    vICMS         — já considera redução de base (20/70)
                                                   e diferimento parcial (51)
      - Simples Nacional (CSOSN):   vCredICMSSN   — único crédito que o Simples transfere
                                                   (LC 123/2006 art. 23); vICMS de CSOSN
                                                   900 NÃO gera crédito ao comprador
    Sem valor destacado → 0.0 (na dúvida, não credita). RULES.md §3.
    """
    cst, v_cred = "", 0.0
    icms_node = imposto.find('.//nfe:ICMS', ns) if imposto is not None else None
    if icms_node is None:
        return cst, v_cred
    for child in icms_node:
        tag_cst = child.find('nfe:CST', ns)
        if tag_cst is not None and tag_cst.text:
            cst = tag_cst.text.strip()
            campo = 'nfe:vICMS'
        else:
            tag_cst = child.find('nfe:CSOSN', ns)
            if tag_cst is None or not tag_cst.text:
                continue
            cst = tag_cst.text.strip()
            campo = 'nfe:vCredICMSSN'
        tag_v = child.find(campo, ns)
        if tag_v is not None and tag_v.text:
            try:
                v_cred = max(0.0, float(tag_v.text.strip()))
            except ValueError:
                v_cred = 0.0
    return cst, v_cred


def get_xml_text(node, xpath, ns, default=""):
    if node is None:
        return default
    child = node.find(xpath, ns)
    if child is not None and child.text is not None:
        return child.text.strip()
    return default


def _aplicar_registro_api(v_st_xml: float, registro: dict) -> tuple:
    """
    Aplica um único registro da API SEFAZ AL sobre o vICMSST do XML.

    Regra:
    - tipoImposto='ST' e v_st_xml > 0: ICMS já está em vICMSST do XML.
      Adiciona apenas valorFecoepCalculado (FECOEP não está no XML).
    - tipoImposto='ST' e v_st_xml = 0: XML não carrega ST (NF sem destaque de ST).
      Usa ICMS + FECOEP completos da API.
    - tipoImposto='ANT': XML tem zero. Adiciona ICMS + FECOEP completos da API.

    Returns: (vST_total, vANT_total)
    """
    v_icms   = float(registro.get('valorIcmsCalculado',  0) or 0)
    v_fecoep = float(registro.get('valorFecoepCalculado', 0) or 0)
    if registro.get('tipoImposto') == 'ANT':
        return v_st_xml, v_icms + v_fecoep
    if v_st_xml > 0.005:
        return v_st_xml + v_fecoep, 0.0
    return v_st_xml + v_icms + v_fecoep, 0.0


def parear_impostos_api(itens_xml: list, dados_api: list) -> dict:
    """
    Pareia itens do XML com registros da API SEFAZ AL por descrição EXATA
    (não substring) e, dentro de cada grupo de descrição idêntica, por
    POSIÇÃO — o item N-ésimo do XML com aquela descrição casa com o
    registro N-ésimo da API (ordenado por `id`) com a mesma descrição.

    Confirmado empiricamente (NF 16750, 33260649682710000138550010000167501001682880):
    a API retorna os registros na mesma ordem dos itens <det> do XML,
    inclusive preservando erros de digitação do fornecedor na descrição
    (ex: "30M X 60CM" em vez de "30CM X 60CM") — o que garante que
    `descricaoProduto` da API é o texto completo e literal do XML, não uma
    versão truncada/genérica. Isso torna a igualdade exata seguro e o
    pareamento posicional dentro do grupo verificável (valores diferentes
    de vProd dentro do mesmo grupo resultam em impostos proporcionalmente
    diferentes, na mesma ordem).

    Quando a contagem de itens do XML não bate com a contagem de registros
    da API para a mesma descrição exata (não observado até agora, mas não
    garantido pela SEFAZ), cai em rateio proporcional a vProd — e marca o
    item como 'rateado' para conferência manual, em vez de arriscar um
    pareamento posicional não verificável.

    itens_xml: lista de dicts com pelo menos 'desc_xml', 'v_st_xml', 'vProd'
    Returns: dict {indice_em_itens_xml: (vST_total, vANT_total, rateado: bool)}
    """
    from collections import defaultdict

    grupos_xml = defaultdict(list)
    for idx, item in enumerate(itens_xml):
        grupos_xml[item['desc_xml'].strip().upper()].append(idx)

    grupos_api = defaultdict(list)
    for rec in sorted(dados_api, key=lambda r: r.get('id') or 0):
        chave = str(rec.get('descricaoProduto', '')).strip().upper()
        grupos_api[chave].append(rec)

    resultados = {}
    for desc, idxs in grupos_xml.items():
        recs = grupos_api.get(desc, [])

        if len(recs) == len(idxs):
            for idx, rec in zip(idxs, recs):
                v_st_xml = itens_xml[idx]['v_st_xml']
                vst, vant = _aplicar_registro_api(v_st_xml, rec)
                resultados[idx] = (vst, vant, False)
        elif recs:
            # Contagem não bate — rateio proporcional a vProd, sinalizado.
            total_vprod = sum(max(itens_xml[i]['vProd'], 0) for i in idxs)
            pesos = (
                {i: itens_xml[i]['vProd'] / total_vprod for i in idxs}
                if total_vprod > 0 else
                {i: 1.0 / len(idxs) for i in idxs}
            )
            icms_ant = sum(float(r.get('valorIcmsCalculado', 0) or 0)   for r in recs if r.get('tipoImposto') == 'ANT')
            feco_ant = sum(float(r.get('valorFecoepCalculado', 0) or 0) for r in recs if r.get('tipoImposto') == 'ANT')
            icms_st  = sum(float(r.get('valorIcmsCalculado', 0) or 0)   for r in recs if r.get('tipoImposto') != 'ANT')
            feco_st  = sum(float(r.get('valorFecoepCalculado', 0) or 0) for r in recs if r.get('tipoImposto') != 'ANT')
            for idx in idxs:
                peso = pesos[idx]
                v_st_xml = itens_xml[idx]['v_st_xml']
                if v_st_xml > 0.005:
                    v_st = v_st_xml + peso * feco_st
                else:
                    v_st = v_st_xml + peso * (icms_st + feco_st)
                v_ant = peso * (icms_ant + feco_ant)
                resultados[idx] = (v_st, v_ant, True)
        else:
            for idx in idxs:
                resultados[idx] = (itens_xml[idx]['v_st_xml'], 0.0, False)

    return resultados


def extrair_qtd_embalagem(desc_xml, v_un_xml, p_sys, mult, desc_sys=''):
    # Padrão 0 — "IP-N" (itens por pacote): "IP-18 TRB", "IP-120 TRB"
    # Tentado primeiro em ambas as descrições (XML e sistema) pois é inequívoco.
    # Útil quando o XML está truncado e o count só aparece no NOME do sistema.
    for desc in (desc_xml, desc_sys or ''):
        m0 = _PATTERN_EMB_IP.search(desc)
        if m0:
            qtd = int(m0.group(1))
            if qtd > 1 and v_un_xml / qtd >= 0.10:
                return qtd

    # Padrão 1a — específico (alta prioridade): "CAIXA COM N", "CX C/N", "PCT C/N"
    # Tentado ANTES do genérico C/ para evitar que "CONJUNTO C/ 03... CX C/120"
    # capture o 3 em vez do 120.
    match = _PATTERN_EMB_ESPECIFICO.search(desc_xml)
    if match:
        qtd = int(match.group(1))
        padrao_explicito = True
    else:
        # Padrão 2 — número no início: "12 CANECAS", "6 TIGELAS"
        match2 = _PATTERN_EMB_INICIO.search(desc_xml)
        if match2:
            qtd = int(match2.group(1))
            padrao_explicito = False
        else:
            # Padrão 3 — sufixo explícito: "POTE 48 UN", "PT24UN", "DP 18 UN"
            match3 = _PATTERN_EMB_FINAL.search(desc_xml)
            if match3:
                qtd = int(match3.group(1))
                padrao_explicito = True
            else:
                # Padrão 4 — DS Display: "DS NOME - 84", "DS NOME - 207 PCS"
                match4 = _PATTERN_EMB_DS.search(desc_xml)
                if match4:
                    qtd = int(match4.group(1))
                    padrao_explicito = True
                else:
                    # Padrão 1b — genérico (baixa prioridade): "C/5", "C/ 12"
                    match5 = _PATTERN_EMB_GENERICO.search(desc_xml)
                    if match5:
                        qtd = int(match5.group(1))
                        padrao_explicito = True
                    else:
                        return 1

    if qtd <= 1:
        return 1

    # Padrão 2 (ambíguo) exige preço real para validar o anti-absurdo.
    # Padrões 1 e 3 são suficientemente explícitos — não precisam de preço.
    if not padrao_explicito and p_sys <= 0.02:
        return 1

    # Anti-absurdo #1: preço de sistema × embalagem > NF × mult × 3
    # Padrões explícitos (DS, CAIXA COM, etc.) são dispensados — o sistema sempre
    # tem preço por unidade; p_sys × qtd naturalmente excede o preço do display.
    if not padrao_explicito and p_sys > 0 and (p_sys * qtd) > (v_un_xml * mult * 3):
        return 1

    # Anti-absurdo #2: preço unitário resultante mínimo R$0.10
    # Evita divisões absurdas quando p_sys é placeholder (ex: 0.01)
    # e a descrição menciona pack size que já está embutido no qCom do NF.
    if v_un_xml / qtd < 0.10:
        return 1

    return qtd


# ─── Extração XML + merge CSV ─────────────────────────────────────────────────
def gerar_tabela(xml_path, csv_path, fornecedor, nota_ref, params):
    try:
        P_MULT     = float(params.get('mult_var',  2.0))
        P_DESP     = float(params.get('desp',      10))   / 100
        P_FRETE    = float(params.get('frete',     10))   / 100
        P_FED      = float(params.get('fed',       9.13)) / 100
        P_ICM      = float(params.get('icm',       21))   / 100
        P_CART     = float(params.get('cartao',    4))    / 100
        P_MULT_ATC     = float(params.get('mult_atc',      1.3))
        P_DESC_ATC     = float(params.get('desc_atc',      15))  / 100  # pedido
        P_DESC_ATC_PDV = float(params.get('desc_atc_pdv',  10))  / 100  # PDV balcão

        tree = ET.parse(xml_path)
        ns = {'nfe': 'http://www.portalfiscal.inf.br/nfe'}
        infNFe    = tree.getroot().find('.//nfe:infNFe', ns)
        num_nf    = get_xml_text(infNFe, './/nfe:ide/nfe:nNF', ns, "000")
        chave_nfe = infNFe.attrib.get('Id', '')[3:] if infNFe is not None else ""

        # Frete CIF/FOB: 0=CIF (emitente paga, frete=0%), 1=FOB (destinatário paga)
        mod_frete = get_xml_text(infNFe, './/nfe:transp/nfe:modFrete', ns, "9")
        if mod_frete == "0":  # CIF — frete por conta do emitente
            P_FRETE = 0.0
            print(f"[FRETE] modFrete=0 (CIF) → frete forçado para 0%")
        else:
            print(f"[FRETE] modFrete={mod_frete} → usando frete do formulário: {P_FRETE*100:.1f}%")

        # API SEFAZ AL
        dados_api = []
        try:
            url = (
                "https://contribuinte.sefaz.al.gov.br/cobrancadfe/"
                f"sfz-cobranca-dfe-api/api/detalhe-calculo-nfes?chaveNota.equals={chave_nfe}"
            )
            res = requests.get(
                url,
                headers={'Accept': 'application/json', 'User-Agent': 'Mozilla/5.0'},
                timeout=10
            )
            print(f"\n{'='*70}")
            print(f"[SEFAZ API] chave: {chave_nfe}")
            print(f"[SEFAZ API] status: {res.status_code}")
            if res.status_code == 200:
                dados_api = res.json()
                print(f"[SEFAZ API] {len(dados_api)} item(s) retornados:")
                for idx, it in enumerate(dados_api):
                    print(f"  [{idx}] tipo={it.get('tipoImposto')!r:6}  "
                          f"icms={it.get('valorIcmsCalculado')!r}  "
                          f"fecoep={it.get('valorFecoepCalculado')!r}  "
                          f"desc={str(it.get('descricaoProduto',''))[:60]!r}")
            else:
                print(f"[SEFAZ API] resposta inesperada: {res.text[:200]}")
            print(f"{'='*70}\n")
        except Exception as _api_err:
            print(f"[SEFAZ API] ERRO na requisição: {_api_err}")

        itens_xml = []
        for det in tree.getroot().findall('.//nfe:det', ns):
            p = det.find('nfe:prod', ns)
            i = det.find('nfe:imposto', ns)

            desc_xml = get_xml_text(p, 'nfe:xProd', ns, "SEM DESCRICAO")
            v_st_xml = float(get_xml_text(i, './/nfe:vICMSST', ns, "0"))

            cst, v_cred_xml = extrair_credito_icms(i, ns)

            ean_xml = limpar_str(get_xml_text(p, 'nfe:cEAN', ns))
            if not ean_xml or ean_xml.upper() == "SEM GTIN":
                ean_xml = "SEM GTIN"

            itens_xml.append({
                'ean_xml':  ean_xml,
                'ref_xml':  limpar_str(get_xml_text(p, 'nfe:cProd', ns)),
                'desc_xml': desc_xml,
                'qCom':     float(get_xml_text(p, 'nfe:qCom',   ns, "1")),
                'vUnCom':   float(get_xml_text(p, 'nfe:vUnCom', ns, "0")),
                'vProd':    float(get_xml_text(p, 'nfe:vProd',  ns, "0")),
                'vIPI':     float(get_xml_text(i, './/nfe:vIPI', ns, "0")),
                'v_st_xml': v_st_xml,
                'nf_base':  num_nf,
                'cst':      cst,
                'vCredICMS': v_cred_xml,
            })

        # Pareia todos os itens do XML com os registros da API de uma vez —
        # necessário para agrupar por descrição exata e casar por posição
        # dentro de cada grupo (ver docstring de parear_impostos_api).
        pareamento = parear_impostos_api(itens_xml, dados_api)

        print(f"\n{'='*70}")
        print(f"[MERGE IMPOSTOS] {len(itens_xml)} item(ns) do XML × {len(dados_api)} registro(s) da API")
        for idx, item in enumerate(itens_xml):
            vst, vant, rateado = pareamento[idx]
            item['vST']  = vst
            item['vANT'] = vant
            item['rateado'] = rateado
            flag = " [RATEADO — conferir]" if rateado else ""
            print(f"  [{idx}] {item['desc_xml'][:55]!r}  vST={vst:.4f}  vANT={vant:.4f}{flag}")
        print(f"{'='*70}\n")

        df_xml = pd.DataFrame(itens_xml)

        # CSV do sistema — FutureWarning corrigido: atribuição via .loc[]
        try:
            df_sys = pd.read_csv(csv_path, sep=';', encoding='utf-8',  on_bad_lines='skip', dtype=str)
        except:
            df_sys = pd.read_csv(csv_path, sep=';', encoding='latin1', on_bad_lines='skip', dtype=str)

        df_sys.columns = df_sys.columns.str.replace('"', '').str.strip().str.upper()
        df_sys = df_sys.copy()  # evita SettingWithCopyWarning
        df_sys.loc[:, 'ean_sys']   = df_sys['BARRA'].apply(limpar_str)      if 'BARRA'      in df_sys.columns else ''
        df_sys.loc[:, 'ref_sys']   = df_sys['REFERÊNCIA'].apply(limpar_str) if 'REFERÊNCIA' in df_sys.columns else ''
        df_sys.loc[:, 'preco_sys'] = df_sys['PREÇO'].apply(limpar_preco)    if 'PREÇO'      in df_sys.columns else 0.0
        # NOME do sistema: fallback para detectar padrões de embalagem ausentes no XML truncado
        df_sys.loc[:, 'nome_sys']  = df_sys['NOME'].apply(limpar_str)        if 'NOME'       in df_sys.columns else ''

        # Merge por EAN
        df_base = pd.merge(df_xml, df_sys[['ean_sys', 'ref_sys', 'preco_sys', 'nome_sys']],
                           left_on='ean_xml', right_on='ean_sys', how='left')

        # Fallback 1 — por REF (cProd do XML == REFERÊNCIA do CSV)
        mask = (df_base['preco_sys'].isna()) | (df_base['preco_sys'] == 0)
        if mask.any():
            uniq  = df_sys.drop_duplicates(subset=['ref_sys']).dropna(subset=['ref_sys'])
            mp    = uniq.set_index('ref_sys')['preco_sys'].to_dict()
            me    = uniq.set_index('ref_sys')['ean_sys'].to_dict()
            mn    = uniq.set_index('ref_sys')['nome_sys'].to_dict()
            df_base.loc[mask, 'preco_sys'] = df_base.loc[mask, 'ref_xml'].map(mp)
            df_base.loc[mask, 'ean_sys']   = df_base.loc[mask, 'ref_xml'].map(me)
            df_base.loc[mask, 'nome_sys']  = df_base.loc[mask, 'ref_xml'].map(mn)

        # Fallback 2 — cProd como EAN (cProd do XML == BARRA do CSV)
        # Fornecedores que colocam o GTIN em cProd e deixam cEAN="SEM GTIN" (ex: Mohnish)
        mask = (df_base['preco_sys'].isna()) | (df_base['preco_sys'] == 0)
        if mask.any():
            uniq2 = df_sys.drop_duplicates(subset=['ean_sys']).dropna(subset=['ean_sys'])
            uniq2 = uniq2[uniq2['ean_sys'] != ''].set_index('ean_sys')
            mp2   = uniq2['preco_sys'].to_dict()
            mn2   = uniq2['nome_sys'].to_dict()
            df_base.loc[mask, 'preco_sys'] = df_base.loc[mask, 'ref_xml'].map(mp2)
            df_base.loc[mask, 'ean_sys']   = df_base.loc[mask, 'ref_xml']  # ref_xml é o EAN
            df_base.loc[mask, 'nome_sys']  = df_base.loc[mask, 'ref_xml'].map(mn2)

        df_base = df_base.copy()
        df_base.loc[:, 'preco_sys'] = df_base['preco_sys'].fillna(0.0)
        df_base.loc[:, 'nome_sys']  = df_base['nome_sys'].fillna('')
        df_base.loc[:, 'sku_final'] = df_base.apply(
            lambda r: r['ean_sys'] if pd.notna(r['ean_sys']) and str(r['ean_sys']).strip() != ""
                      else r['ean_xml'], axis=1)
        df_base.loc[:, 'qtd_emb']   = df_base.apply(
            lambda r: extrair_qtd_embalagem(r['desc_xml'], r['vUnCom'], r['preco_sys'], P_MULT,
                                            str(r['nome_sys'])),
            axis=1)

        rows = []
        for _, row in df_base.iterrows():
            q       = max(float(row['qCom']), 1)
            qtd_emb = int(row['qtd_emb'])

            # Valores por unidade comercial (pré-embalagem) — usados como constantes nas fórmulas Excel
            nf_u_raw  = row['vProd'] / q
            st_u_raw  = row['vST']   / q
            ant_u_raw = row['vANT']  / q
            ipi_u_raw = row['vIPI']  / q

            # Dividir custos NF pela embalagem → tudo por UNIDADE vendida (usado no dashboard/preview)
            nf_u    = round(nf_u_raw  / qtd_emb, 2)
            st_u    = round(st_u_raw  / qtd_emb, 2)
            ant_u   = round(ant_u_raw / qtd_emb, 2)
            ipi_u   = round(ipi_u_raw / qtd_emb, 2)
            cst     = str(row['cst'])

            # Flags de cor baseadas nos valores BRUTOS (antes da divisão por embalagem)
            # Evita que qtd_emb grande faça st_u/ant_u arredondar para 0
            tem_st  = row['vST']  > 0.005
            tem_ant = row['vANT'] > 0.005 and not tem_st

            # Preço do sistema já é por unidade — NÃO multiplicar por embalagem
            p_sys_val        = float(row['preco_sys']) if pd.notna(row['preco_sys']) else 0.0
            preco_venda_base = round(p_sys_val, 2) if p_sys_val > 0.01 else 0.0

            # Crédito ICMS = valor destacado na NF ÷ valor do produto (taxa efetiva).
            # Taxa em vez de R$/un para acompanhar a divisão por embalagem e continuar
            # editável na coluna TAXA CRED do Excel.
            v_cred   = float(row['vCredICMS']) if pd.notna(row['vCredICMS']) else 0.0
            cred_pct = round(v_cred / row['vProd'], 4) if row['vProd'] > 0 else 0.0

            rows.append({
                'nf':        row['nf_base'],
                'desc':      row['desc_xml'],
                'ref':       row['ref_xml'],
                'sku':       row['sku_final'],
                'qtd':       q,
                'nf_u':      nf_u,
                'st_u':      st_u,
                'ant_u':     ant_u,
                'ipi_u':     ipi_u,
                'nf_u_raw':  nf_u_raw,
                'st_u_raw':  st_u_raw,
                'ant_u_raw': ant_u_raw,
                'ipi_u_raw': ipi_u_raw,
                'cst':       cst,
                'p_atual':   preco_venda_base,
                'p_sys_raw': p_sys_val,
                'qtd_emb':   qtd_emb,
                'cred_pct':  cred_pct,
                'tem_st':    tem_st,
                'tem_ant':   tem_ant,
                'rateado':   bool(row.get('rateado', False)),
            })

        params_out = {
            'mult': P_MULT, 'frete': P_FRETE, 'desp': P_DESP,
            'fed': P_FED, 'icm': P_ICM, 'cartao': P_CART,
            'mult_atc': P_MULT_ATC, 'desc_atc': P_DESC_ATC, 'desc_atc_pdv': P_DESC_ATC_PDV,
        }
        return True, (rows, params_out, num_nf)

    except Exception as e:
        import traceback
        return False, traceback.format_exc()


# ─── Estilos ─────────────────────────────────────────────────────────────────
def _fill(hex_color):
    return PatternFill("solid", start_color=hex_color)

def _borda():
    s = Side("thin")
    return Border(left=s, right=s, top=s, bottom=s)

F_AZUL     = _fill("BDD7EE")
F_AMAR     = _fill("FFF2CC")
F_AMAR_ANT = _fill("D1FAE5")
F_PELE     = _fill("FCE4D6")
F_CINZA    = _fill("D9D9D9")
F_VERDE    = _fill("E2EFDA")
F_PARAM    = _fill("F2F2F2")
F_RATEIO   = _fill("FF0000")  # vermelho — ST_U/ANT_U estimados por rateio (conferir manualmente)
BORDA   = _borda()
ALI_CTR = Alignment(horizontal='center', vertical='center')


def _c(ws, r, c, value=None, fill=None, font=None, fmt=None, bold=False):
    """Atalho para setar célula com valor + estilo."""
    cell = ws.cell(row=r, column=c, value=value)
    cell.border    = BORDA
    cell.alignment = ALI_CTR
    if fill: cell.fill          = fill
    if fmt:  cell.number_format = fmt
    if bold: cell.font          = Font(bold=True)
    if font: cell.font          = font
    return cell


# ─── Salvar Excel ─────────────────────────────────────────────────────────────
def salvar_excel_estilizado(dados, path):
    rows, P, num_nf = dados
    L = get_column_letter

    wb = Workbook()
    ws = wb.active
    ws.title = 'Precificação'

    # ── Linha 1: Cabeçalhos ──────────────────────────────────────────────────
    headers = [
        'NF', 'DESCRIÇÃO', 'REF', 'SKU', 'QTD',
        'NF UNIT', 'ST UNIT', 'ANT UNIT', 'IPI UNIT',
        'CUSTO REAL', 'FRETE', 'DESPESA', 'CRED ICMS',
        'CST', 'CUSTO ENTRADA',
        'FEDERAL', 'CARTÃO', 'ICMS SAÍDA', 'CUSTO SAÍDA',
        'META %', 'PREÇO MÍN VIÁVEL VRJ',
        'PREÇO ATUAL', 'PREÇO VAREJO', 'MARGEM REAL',
        'NF ATC', 'PREÇO ATC PEDIDO', 'PREÇO ATC PDV',
        'FEDERAL ATC', 'CARTÃO ATC', 'ICMS ATC',
        'CUSTO SAÍDA ATC', 'MARGEM ATC PED', 'MARGEM ATC PDV',
        'PREÇO PCT ATC', 'P. COMPRA PCT',
        'P.UNIT SISTEMA', 'QTD EMB',
        'TAXA CRED',  # col 38 — taxa de crédito ICMS por produto (colorida por origem)
    ]
    for c, h in enumerate(headers, 1):
        _c(ws, 1, c, h, fill=F_PARAM, bold=True)

    # Cabeçalhos varejo em azul
    for col_key in ('FED', 'CARTAO', 'ICMS_S', 'C_SAIDA', 'P_MIN', 'P_VAR', 'MARGEM'):
        ws.cell(1, COL[col_key]).fill = F_AZUL

    # Cabeçalhos atacado em amarelo
    for col_key in ('NF_ATC', 'P_ATC_PED', 'P_ATC_PDV', 'FED_ATC', 'CART_ATC', 'ICM_ATC',
                    'C_SAIDA_ATC', 'MARGEM_ATC_PED', 'MARGEM_ATC_PDV', 'P_PCT_ATC', 'P_COMPRA_PCT'):
        ws.cell(1, COL[col_key]).fill = F_AMAR

    # ── Linha 2: Parâmetros ──────────────────────────────────────────────────
    # col 10 = MULT (número), cols % = formatados como percentual
    param_vals = {
        COL['C_REAL']:  P['mult'],       # col 10 — multiplicador varejo
        COL['FRETE']:   P['frete'],      # col 11
        COL['DESP']:    P['desp'],       # col 12
        COL['FED']:     P['fed'],        # col 16
        COL['CARTAO']:  P['cartao'],     # col 17
        COL['ICMS_S']:  P['icm'],        # col 18
        COL['NF_ATC']:    P['mult_atc'],      # col 25 — multiplicador atacado
        COL['P_ATC_PED']: P['desc_atc'],      # col 26 — desconto pedido (15%)
        COL['P_ATC_PDV']: P['desc_atc_pdv'],  # col 27 — desconto PDV balcão (10%)
    }
    pct_cols = {COL['FRETE'], COL['DESP'], COL['FED'], COL['CARTAO'], COL['ICMS_S'],
                COL['P_ATC_PED'], COL['P_ATC_PDV']}
    for c in range(1, TOTAL_COLS + 1):
        val = param_vals.get(c)
        fmt = '0.00%' if c in pct_cols else ('0.00' if c == COL['C_REAL'] else 'General')
        _c(ws, 2, c, val, fill=F_PARAM, fmt=fmt)
    ws.cell(2, COL['META']).value = "META %"

    # ── Linhas de dados (linha 3 em diante) ──────────────────────────────────
    for idx, row in enumerate(rows):
        r       = idx + 3
        has_st  = row['tem_st']
        has_ant = row['tem_ant']
        cst     = row['cst']

        # Letra da coluna AUDIT_EMB — usada nas fórmulas de NF_U/ST_U/ANT_U/IPI_U
        AK_emb = L(COL['AUDIT_EMB'])

        # Letras das colunas para fórmulas
        F  = L(COL['NF_U'])       # NF UNIT
        G  = L(COL['ST_U'])       # ST
        H  = L(COL['ANT_U'])      # ANT
        I  = L(COL['IPI_U'])      # IPI
        N  = L(COL['C_REAL'])     # CUSTO REAL
        J  = L(COL['FRETE'])      # FRETE
        K  = L(COL['DESP'])       # DESPESA
        Lc = L(COL['CRED'])       # CRED ICMS
        Oc = L(COL['CST'])        # CST
        M  = L(COL['C_ENT'])      # CUSTO ENTRADA
        P2 = L(COL['FED'])        # FEDERAL
        Qc = L(COL['CARTAO'])     # CARTÃO
        Rc = L(COL['ICMS_S'])     # ICMS SAÍDA
        S  = L(COL['C_SAIDA'])    # CUSTO SAÍDA
        T  = L(COL['META'])       # META %
        V  = L(COL['P_ATUAL'])    # PREÇO ATUAL
        W  = L(COL['P_VAR'])      # PREÇO VAREJO
        Ya = L(COL['NF_ATC'])      # NF ATC
        Za = L(COL['P_ATC_PED'])  # PREÇO ATC PEDIDO
        Zb = L(COL['P_ATC_PDV'])  # PREÇO ATC PDV
        AA = L(COL['FED_ATC'])    # FEDERAL ATC
        AB = L(COL['CART_ATC'])   # CARTÃO ATC
        AC = L(COL['ICM_ATC'])    # ICMS ATC
        AD = L(COL['C_SAIDA_ATC'])      # CUSTO SAÍDA ATC
        AE = L(COL['MARGEM_ATC_PED'])  # MARGEM ATC PED
        AF = L(COL['MARGEM_ATC_PDV'])  # MARGEM ATC PDV
        AJc = L(COL['AUDIT_CRED'])      # taxa de crédito ICMS por produto

        # Valores fixos (vêm do XML/CSV)
        desc_final = row['desc']
        if row.get('rateado'):
            desc_final = f"⚠ CONFERIR IMPOSTO — {desc_final}"
        _c(ws, r, COL['NF'],    row['nf'])
        _c(ws, r, COL['DESC'],  desc_final)
        _c(ws, r, COL['REF'],   row['ref'])
        _c(ws, r, COL['SKU'],   row['sku'])
        _c(ws, r, COL['QTD'],   row['qtd'])
        _c(ws, r, COL['CST'],   cst)
        _c(ws, r, COL['P_ATUAL'], row['p_atual'], fmt='#,##0.00')

        # NF_U/ST_U/ANT_U/IPI_U como fórmulas referenciando AUDIT_EMB (AK)
        # Alterar AK (qtd embalagem) recalcula toda a planilha em cascata.
        # MAX(1, AK) evita divisão por zero se o usuário zerar a célula.
        def _raw_fmt(val):
            # Serializa o valor raw com casas decimais suficientes para não perder precisão
            return f"{val:.6f}"

        for col_key, raw_val in (
            ('NF_U',  row['nf_u_raw']),
            ('ST_U',  row['st_u_raw']),
            ('ANT_U', row['ant_u_raw']),
            ('IPI_U', row['ipi_u_raw']),
        ):
            cell = ws.cell(row=r, column=COL[col_key])
            cell.value         = f"=ROUND({_raw_fmt(raw_val)}/MAX(1,{AK_emb}{r}),2)"
            cell.number_format = '#,##0.00'
            cell.border        = BORDA
            cell.alignment     = ALI_CTR

        # META % — editável por produto, padrão 15%
        mc = ws.cell(row=r, column=COL['META'], value=0.15)
        mc.border = BORDA; mc.alignment = ALI_CTR
        mc.number_format = '0.00%'
        mc.fill = F_VERDE; mc.font = Font(bold=True)

        # Colunas de auditoria
        _c(ws, r, COL['AUDIT_SYS'], row['p_sys_raw'], fill=F_CINZA,
           font=Font(italic=True, color="888888"), fmt='#,##0.00')
        _c(ws, r, COL['AUDIT_EMB'], row['qtd_emb'],   fill=F_CINZA,
           font=Font(italic=True, color="888888"))

        # AUDIT_CRED — taxa de crédito ICMS por produto, colorida pela origem
        # É também o operando da fórmula CRED: alterar esta célula muda o crédito calculado
        cst_isento = cst in CST_SEM_CREDITO
        cred_rate  = 0.0 if (has_st or cst_isento) else row['cred_pct']
        cell_ac = ws.cell(r, COL['AUDIT_CRED'])
        cell_ac.value          = row['cred_pct']
        cell_ac.number_format  = '0.00%'
        cell_ac.fill           = cred_fill(cred_rate)
        cell_ac.border         = BORDA
        cell_ac.alignment      = ALI_CTR

        # ── Fórmulas ─────────────────────────────────────────────────────────

        # N - CUSTO REAL = NF_U × MULT  ← multiplicador aplicado PRIMEIRO
        ws.cell(r, COL['C_REAL']).value = f"=ROUND({F}{r}*{N}$2,2)"

        # J - FRETE = CUSTO REAL × FRETE%
        ws.cell(r, COL['FRETE']).value  = f"=ROUND({N}{r}*{J}$2,2)"

        # K - DESPESA = CUSTO REAL × DESP%
        ws.cell(r, COL['DESP']).value   = f"=ROUND({N}{r}*{K}$2,2)"

        # M (col 13) — CRED ICMS: fórmula que lê a taxa diretamente da coluna TAXA CRED
        # A COR da célula indica o percentual de origem (§16 do RULES.md)
        # A FÓRMULA zera automaticamente se ST>0.005 ou CST isento
        # O usuário pode alterar a taxa na coluna TAXA CRED e o crédito recalcula
        _cst_zero = 'OR(' + ','.join(f'{Oc}{r}="{c}"' for c in sorted(CST_SEM_CREDITO)) + ')'
        cell_cred = ws.cell(r, COL['CRED'])
        cell_cred.value = (
            f'=IF(OR({G}{r}>0.005,{_cst_zero}),0,'
            f'-ROUND({F}{r}*{AJc}{r},2))'
        )
        cell_cred.number_format = '#,##0.00'
        cell_cred.fill          = cred_fill(cred_rate)
        cell_cred.border        = BORDA
        cell_cred.alignment     = ALI_CTR

        # O - CUSTO ENTRADA = C_REAL + ST + ANT + IPI + FRETE + DESP + CRED(negativo)
        ws.cell(r, COL['C_ENT']).value  = (
            f"=ROUND({N}{r}+{G}{r}+{H}{r}+{I}{r}+{J}{r}+{K}{r}+{Lc}{r},2)"
        )

        # P - FEDERAL = P_VAR × FED%
        ws.cell(r, COL['FED']).value    = f"=ROUND({W}{r}*{P2}$2,2)"

        # Q - CARTÃO = P_VAR × CART%
        ws.cell(r, COL['CARTAO']).value = f"=ROUND({W}{r}*{Qc}$2,2)"

        # R - ICMS SAÍDA = max(0, P_VAR × ICMS% − ANT_U)
        # ST zera tudo; ANT reduz (não zera) — o valor já antecipado na entrada é abatido
        ws.cell(r, COL['ICMS_S']).value = (
            f"=IF({G}{r}>0,0,MAX(0,ROUND({W}{r}*{Rc}$2,2)-{H}{r}))"
        )

        # S - CUSTO SAÍDA = C_ENT + FED + CART + ICMS
        ws.cell(r, COL['C_SAIDA']).value = (
            f"=ROUND({M}{r}+{P2}{r}+{Qc}{r}+{Rc}{r},2)"
        )

        # U - PREÇO MÍN VIÁVEL = C_SAIDA ÷ (1 − META%)
        ws.cell(r, COL['P_MIN']).value  = (
            f"=IF({T}{r}>0,ROUND({S}{r}/(1-{T}{r}),2),0)"
        )

        # W - PREÇO VAREJO: usa PREÇO ATUAL se > 0, senão fallback analítico
        # Fallback = arredondar_x9(C_ENT / (1 - META - FED% - CART% - ICMS%))
        # Garante que P_VAR >= P_MIN sem dependência circular
        ws.cell(r, COL['P_VAR']).value  = (
            f"=IF({V}{r}>0,{V}{r},"
            f"IF({M}{r}<=0,0,"
            f"IF(1-{T}{r}-{P2}$2-{Qc}$2-IF({G}{r}>0.005,0,{Rc}$2)<=0,{M}{r},"
            f"INT({M}{r}/(1-{T}{r}-{P2}$2-{Qc}$2-IF({G}{r}>0.005,0,{Rc}$2))*10)/10+0.09)))"
        )

        # X - MARGEM REAL = (P_VAR − C_SAIDA) / P_VAR
        ws.cell(r, COL['MARGEM']).value = (
            f"=IF({W}{r}>0,ROUND(({W}{r}-{S}{r})/{W}{r},4),0)"
        )

        # ── Atacado ───────────────────────────────────────────────────────────

        # Y - NF ATC = NF_U × MULT_ATC
        ws.cell(r, COL['NF_ATC']).value = f"=ROUND({F}{r}*{Ya}$2,2)"

        # Z - PREÇO ATC PEDIDO = P_VAR × (1 − DESC_PED%)  — desconto pedido (15%)
        ws.cell(r, COL['P_ATC_PED']).value = f"=ROUND({W}{r}*(1-{Za}$2),2)"

        # AA - PREÇO ATC PDV = P_VAR × (1 − DESC_PDV%)  — desconto balcão (10%)
        ws.cell(r, COL['P_ATC_PDV']).value = f"=ROUND({W}{r}*(1-{Zb}$2),2)"
        ws.cell(r, COL['P_ATC_PDV']).number_format = '#,##0.00'

        # AB - FEDERAL ATC = NF_ATC × FED%
        ws.cell(r, COL['FED_ATC']).value = f"=ROUND({Ya}{r}*{P2}$2,2)"

        # AB - CARTÃO ATC = P_ATC × CART%
        ws.cell(r, COL['CART_ATC']).value = f"=ROUND({Za}{r}*{Qc}$2,2)"

        # AC - ICMS ATC = max(0, NF_ATC × ICM% − ANT_U)
        # ST zera tudo; ANT reduz — mesma lógica do varejo
        ws.cell(r, COL['ICM_ATC']).value = (
            f"=IF({G}{r}>0,0,MAX(0,ROUND({Ya}{r}*{Rc}$2,2)-{H}{r}))"
        )

        # AD - CUSTO SAÍDA ATC = C_ENT + FED_ATC + CART_ATC + ICM_ATC
        ws.cell(r, COL['C_SAIDA_ATC']).value = (
            f"=ROUND({M}{r}+{AA}{r}+{AB}{r}+{AC}{r},2)"
        )

        # AE - MARGEM ATC PED = (P_ATC_PED − C_SAIDA_ATC) / P_ATC_PED
        ws.cell(r, COL['MARGEM_ATC_PED']).value = (
            f"=IF({Za}{r}>0,ROUND(({Za}{r}-{AD}{r})/{Za}{r},4),0)"
        )

        # AF - MARGEM ATC PDV = (P_ATC_PDV − C_SAIDA_VRJ) / P_ATC_PDV
        # base de custo: C_SAIDA do varejo (não do atacado)
        ws.cell(r, COL['MARGEM_ATC_PDV']).value = (
            f"=IF({Zb}{r}>0,ROUND(({Zb}{r}-{S}{r})/{Zb}{r},4),0)"
        )

        # AF - PREÇO PCT ATC = P_ATC × QTD_EMB (preço de venda do pacote no atacado)
        ws.cell(r, COL['P_PCT_ATC']).value = f"=ROUND({Za}{r}*{AK_emb}{r},2)"
        ws.cell(r, COL['P_PCT_ATC']).number_format = '#,##0.00'

        # AG - P. COMPRA PCT = NF_U × QTD_EMB (preço de compra do pacote — custo NF por caixa)
        ws.cell(r, COL['P_COMPRA_PCT']).value = f"=ROUND({F}{r}*{AK_emb}{r},2)"
        ws.cell(r, COL['P_COMPRA_PCT']).number_format = '#,##0.00'

        # ── Formato numérico para colunas de fórmula ─────────────────────────
        for col_idx, fmt in [
            (COL['C_REAL'],      '#,##0.00'),
            (COL['FRETE'],       '#,##0.00'),
            (COL['DESP'],        '#,##0.00'),
            (COL['CRED'],        '#,##0.00'),
            (COL['C_ENT'],       '#,##0.00'),
            (COL['FED'],         '#,##0.00'),
            (COL['CARTAO'],      '#,##0.00'),
            (COL['ICMS_S'],      '#,##0.00'),
            (COL['C_SAIDA'],     '#,##0.00'),
            (COL['P_MIN'],       '#,##0.00'),
            (COL['P_VAR'],       '#,##0.00'),
            (COL['MARGEM'],      '0.00%'),
            (COL['NF_ATC'],      '#,##0.00'),
            (COL['P_ATC_PED'],   '#,##0.00'),
            (COL['P_ATC_PDV'],   '#,##0.00'),
            (COL['FED_ATC'],     '#,##0.00'),
            (COL['CART_ATC'],    '#,##0.00'),
            (COL['ICM_ATC'],     '#,##0.00'),
            (COL['C_SAIDA_ATC'],     '#,##0.00'),
            (COL['MARGEM_ATC_PED'], '0.00%'),
            (COL['MARGEM_ATC_PDV'], '0.00%'),
            (COL['P_PCT_ATC'],      '#,##0.00'),
            (COL['P_COMPRA_PCT'],  '#,##0.00'),
        ]:
            ws.cell(r, col_idx).number_format = fmt

        # ── Borda e alinhamento em todas as células da linha ──────────────────
        for c in range(1, TOTAL_COLS + 1):
            cell = ws.cell(r, c)
            cell.border    = BORDA
            cell.alignment = ALI_CTR

        # ── Cores ─────────────────────────────────────────────────────────────
        _skip = {COL['AUDIT_SYS'], COL['AUDIT_EMB'], COL['AUDIT_CRED'], COL['META'], COL['CRED']}
        # Colunas VAREJO destacadas em azul
        _azul = {COL['FED'], COL['CARTAO'], COL['ICMS_S'], COL['C_SAIDA'],
                 COL['P_MIN'], COL['P_VAR'], COL['MARGEM']}
        # Colunas ATACADO destacadas em amarelo
        _amar = {COL['NF_ATC'], COL['P_ATC_PED'], COL['P_ATC_PDV'], COL['FED_ATC'],
                 COL['CART_ATC'], COL['ICM_ATC'], COL['C_SAIDA_ATC'],
                 COL['MARGEM_ATC_PED'], COL['MARGEM_ATC_PDV'],
                 COL['P_PCT_ATC'], COL['P_COMPRA_PCT']}
        if has_st:
            # ST: peach em toda a linha, sem exceção
            for c in range(1, TOTAL_COLS + 1):
                if c not in _skip:
                    ws.cell(r, c).fill = F_PELE
        elif has_ant:
            # ANT: verde menta na linha inteira...
            for c in range(1, TOTAL_COLS + 1):
                if c not in _skip:
                    ws.cell(r, c).fill = F_AMAR_ANT
            # ...mas preserva as cores de varejo/atacado por cima
            for c in _azul:
                ws.cell(r, c).fill = F_AZUL
            for c in _amar:
                ws.cell(r, c).fill = F_AMAR
        else:
            for c in _azul:
                ws.cell(r, c).fill = F_AZUL
            for c in _amar:
                ws.cell(r, c).fill = F_AMAR

        # ST_U/ANT_U em vermelho quando o valor veio de rateio (contagem XML ≠
        # contagem API para a mesma descrição) — sobrepõe qualquer cor de linha,
        # pois é um alerta de conferência manual, não uma classificação fiscal.
        if row.get('rateado'):
            ws.cell(r, COL['ST_U']).fill  = F_RATEIO
            ws.cell(r, COL['ANT_U']).fill = F_RATEIO

    # ── Legenda de crédito ICMS (só se houver mais de uma faixa na NF) ────────
    # Agrupa pela faixa nominal (4/7/12/19%) quando a taxa efetiva está próxima
    taxas_usadas = sorted(set(_chave_faixa(row['cred_pct']) for row in rows))
    if taxas_usadas:  # sempre exibe legenda quando há ao menos uma faixa
        leg_row = len(rows) + 4   # 1 linha de gap após os dados
        ws.cell(leg_row, 1).value = 'LEGENDA — CRÉDITO ICMS'
        ws.cell(leg_row, 1).font  = Font(bold=True, size=10)
        ws.cell(leg_row, 1).fill  = F_PARAM
        ws.merge_cells(start_row=leg_row, start_column=1,
                       end_row=leg_row, end_column=4)
        for i, taxa in enumerate(taxas_usadas):
            lr = leg_row + 1 + i
            lbl  = CRED_LABELS.get(taxa, f'{_fmt_taxa(taxa)} — destacado na NF')
            fill = cred_fill(taxa)
            c1 = ws.cell(lr, 1)
            c1.value  = _fmt_taxa(taxa)
            c1.fill   = fill
            c1.font   = Font(bold=True)
            c1.border = BORDA
            c1.alignment = ALI_CTR
            c2 = ws.cell(lr, 2)
            c2.value  = lbl
            c2.fill   = fill
            c2.border = BORDA
            ws.merge_cells(start_row=lr, start_column=2,
                           end_row=lr, end_column=4)

    # ── Legenda de rateio (só se houver algum produto rateado) ────────────────
    if any(row.get('rateado') for row in rows):
        rat_row = len(rows) + 4 + (1 + len(taxas_usadas) + 1 if taxas_usadas else 0)
        c1 = ws.cell(rat_row, 1)
        c1.value  = 'RATEADO'
        c1.fill   = F_RATEIO
        c1.font   = Font(bold=True, color="FFFFFF")
        c1.border = BORDA
        c1.alignment = ALI_CTR
        c2 = ws.cell(rat_row, 2)
        c2.value  = ('ST/ANT estimado por rateio — a SEFAZ retornou uma quantidade de '
                     'registros diferente da quantidade de itens dessa descrição na NF. '
                     'Confira manualmente o imposto desse produto.')
        c2.border = BORDA
        ws.merge_cells(start_row=rat_row, start_column=2,
                       end_row=rat_row, end_column=8)

    # ── Largura das colunas ───────────────────────────────────────────────────
    widths = {
        1:10, 2:52, 3:14, 4:16, 5:7,
        6:10, 7:10, 8:10, 9:10,
        10:12, 11:10, 12:10, 13:12,
        14:8,  15:15,
        16:11, 17:10, 18:12, 19:12,
        20:9,  21:16,
        22:12, 23:14, 24:12,
        25:10, 26:16, 27:14,
        28:12, 29:11, 30:10,
        31:15, 32:12,
        33:14, 34:14, 35:14, 36:9,
        37:10,
    }
    for c, w in widths.items():
        ws.column_dimensions[L(c)].width = w

    ws.freeze_panes = 'A3'
    wb.calculation.calcMode      = 'auto'
    wb.calculation.fullCalcOnLoad = True
    wb.save(path)


# ─── Dashboard HTML ───────────────────────────────────────────────────────────
def gerar_dashboard_html(rows_data, lucro_total=0.0, metricas=None, num_nf=None):
    if not rows_data:
        return ""

    total     = len(rows_data)
    com_st    = sum(1 for r in rows_data if r['tem_st'])
    com_ant   = sum(1 for r in rows_data if r['tem_ant'])
    sem_ambos = total - com_st - com_ant
    sem_preco = sum(1 for r in rows_data if r['p_atual'] <= 0)
    total_nf  = sum(r['nf_u'] * r['qtd'] for r in rows_data)

    margens = []
    for r in rows_data:
        if r['p_atual'] > 0 and r['nf_u'] > 0:
            margens.append((r['p_atual'] - r['nf_u']) / r['p_atual'])
    margem_media = (sum(margens) / len(margens) * 100) if margens else 0
    cor_margem   = "#E8F5E9" if margem_media >= 15 else "#FFEBEE"
    cor_lucro    = "#E2EFDA" if lucro_total >= 0 else "#FFEBEE"

    def fmt_brl(v):
        return f"R$ {v:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')

    def card(titulo, valor, cor, sub=""):
        return f"""<div style="background:{cor};border-radius:10px;padding:16px 20px;
                    flex:1;min-width:140px;box-shadow:0 2px 6px rgba(0,0,0,.07)">
          <div style="font-size:11px;color:#666;font-weight:600;text-transform:uppercase;
                      letter-spacing:.5px;margin-bottom:4px">{titulo}</div>
          <div style="font-size:26px;font-weight:700;color:#1a1a2e">{valor}</div>
          <div style="font-size:11px;color:#888;margin-top:2px">{sub}</div>
        </div>"""

    # Embalagens detectadas
    com_emb = sum(1 for r in rows_data if r.get('qtd_emb', 1) > 1)

    # Créditos ICMS por faixa — com badges coloridos
    # Usa metricas para aplicar o mesmo filtro CST/ST do Excel (crédito já zerado lá)
    cred_faixas = {}
    for r, m in zip(rows_data, metricas or [{}] * len(rows_data)):
        if m.get('cred', 0) > 0:
            pct = _chave_faixa(r.get('cred_pct', 0))
            if pct > 0:
                cred_faixas[pct] = cred_faixas.get(pct, 0) + 1

    if cred_faixas:
        badges = []
        for pct, cnt in sorted(cred_faixas.items()):
            cor = cor_cred(pct, 'CCCCCC')
            badges.append(
                f'<span style="background:#{cor};padding:2px 8px;border-radius:12px;'
                f'font-size:11px;font-weight:700;margin:2px;display:inline-block">'
                f'{cnt}× {_fmt_taxa(pct)}</span>'
            )
        cred_resumo = ' '.join(badges)
    else:
        cred_resumo = '—'

    sem_preco_pct = (sem_preco / total * 100) if total else 0

    # Alert banner (exibido antes dos cards quando há produtos sem preço)
    alert_banner = ""
    if sem_preco > 0:
        alert_banner = (
            f'<div style="background:#FEF2F2;border:1px solid #FCA5A5;border-radius:10px;'
            f'padding:16px 20px;display:flex;align-items:center;gap:16px;margin-bottom:20px">'
            f'<div style="font-size:22px;flex-shrink:0">⚠️</div>'
            f'<div style="flex:1">'
            f'<div style="font-weight:700;color:#DC2626;font-size:13px">'
            f'Atenção: Produtos sem precificação</div>'
            f'<div style="font-size:12px;color:#555;margin-top:3px">Detectamos que '
            f'<strong>{sem_preco} iten(s) ({sem_preco_pct:.0f}%)</strong> ainda não possuem '
            f'preço de venda definido. Isso compromete a margem estimada.</div>'
            f'</div></div>'
        )

    # Cabeçalho com número da NF
    nf_ref = ""
    if num_nf:
        nf_ref = (
            f'<div style="font-size:11px;color:#888;font-weight:500;margin-bottom:4px">'
            f'Nota Fiscal: <span style="color:#3b82f6;font-weight:700">#{num_nf}</span></div>'
        )

    cards = "".join([
        card("Total Itens",     total,                      "#EBF5FB", "produtos na NF"),
        card("Com ST",          com_st,                     "#FCE4D6", "subst. tributária"),
        card("Com ANT",         com_ant,                    "#D1FAE5", "antecipação tributária"),
        card("Sem Preço Sist.", sem_preco,                  "#FFF9C4", "usarão arredondamento X.X9"),
        card("Total NF",        fmt_brl(total_nf),          "#EBF5FB", "valor de compra"),
        card("Margem Est.",     f"{margem_media:.1f}%",     cor_margem, "preço atual vs custo NF"),
        card("Embalagens",      com_emb,                    "#EBF5FB", f"de {total} detectaram embalagem"),
        card("Crédito ICMS",    cred_resumo,                "#E2EFDA", "do XML por produto"),
        card("Lucro Estimado",  fmt_brl(lucro_total),       cor_lucro, "preço varejo − custo saída"),
    ])

    pct_st    = total and (com_st    / total * 100)
    pct_ant   = total and (com_ant   / total * 100)
    pct_norm  = 100 - pct_st - pct_ant

    def _seg(w, bg, tc, label):
        if w < 1:
            return ''
        return (f'<div style="width:{w:.0f}%;background:{bg};display:flex;align-items:center;'
                f'justify-content:center;color:{tc};font-size:12px;font-weight:600">'
                f'{label} {w:.0f}%</div>')

    barra = f"""
    <div style="margin:20px 0 4px">
      <div style="font-size:11px;color:#666;font-weight:600;margin-bottom:6px;
                  text-transform:uppercase;letter-spacing:.5px">Distribuição da NF</div>
      <div style="display:flex;height:26px;border-radius:6px;overflow:hidden">
        {_seg(pct_norm, '#BDD7EE', '#1a5276', 'Normal')}
        {_seg(pct_st,   '#FCE4D6', '#784212', 'ST')}
        {_seg(pct_ant,  '#D1FAE5', '#065F46', 'ANT')}
      </div>
    </div>"""

    # ── Legenda de CST encontrados na NF ─────────────────────────────────────
    csts_nf = {}
    for r in rows_data:
        cst = str(r.get('cst', '')).strip()
        if cst:
            csts_nf[cst] = csts_nf.get(cst, 0) + 1

    cst_linhas = ""
    for cst, qtd_prod in sorted(csts_nf.items(), key=lambda x: x[0]):
        desc, cor = CST_DESCRICOES.get(cst, (f'Código {cst} — consulte a tabela ICMS', '#F2F2F2'))
        cst_linhas += (
            f'<div style="display:flex;align-items:flex-start;gap:10px;'
            f'padding:8px 12px;border-radius:6px;background:{cor};margin-bottom:6px">'
            f'<span style="font-size:13px;font-weight:800;color:#1a1a2e;min-width:36px">CST {cst}</span>'
            f'<span style="font-size:12px;color:#333;flex:1">{desc}</span>'
            f'<span style="font-size:11px;color:#666;white-space:nowrap">{qtd_prod} produto(s)</span>'
            f'</div>'
        )

    cst_html = f"""
    <div style="margin:20px 0 4px;border-top:1px solid #e8e8e8;padding-top:18px">
      <div style="font-size:11px;color:#666;font-weight:600;text-transform:uppercase;
                  letter-spacing:.5px;margin-bottom:10px">Regimes Tributários desta NF</div>
      {cst_linhas}
    </div>"""

    alertas_html = ""
    alertas = [r for r in rows_data if r['p_atual'] <= 0]
    if alertas:
        linhas = "".join(
            f"<tr><td style='padding:5px 10px'>{html.escape(str(a['ref']))}</td>"
            f"<td style='padding:5px 10px'>{html.escape(a['desc'][:60])}</td>"
            f"<td style='padding:5px 10px;color:#c0392b;font-weight:600'>Sem preço no sistema</td></tr>"
            for a in alertas
        )
        alertas_html = f"""
        <div style="margin-top:20px">
          <div style="font-weight:700;font-size:13px;color:#c0392b;margin-bottom:8px">
            ⚠️ {len(alertas)} produto(s) sem preço no sistema — preço calculado será usado
          </div>
          <table style="width:100%;border-collapse:collapse;font-size:12px;background:#fdf2f2;
                        border-radius:8px;overflow:hidden">
            <thead><tr style="background:#fdecea">
              <th style="padding:6px 10px;text-align:left">REF</th>
              <th style="padding:6px 10px;text-align:left">DESCRIÇÃO</th>
              <th style="padding:6px 10px;text-align:left">STATUS</th>
            </tr></thead>
            <tbody>{linhas}</tbody>
          </table>
        </div>"""

    rateio_alertas = [r for r in rows_data if r.get('rateado')]
    if rateio_alertas:
        linhas_rateio = "".join(
            f"<tr><td style='padding:5px 10px'>{html.escape(str(a['ref']))}</td>"
            f"<td style='padding:5px 10px'>{html.escape(a['desc'][:60])}</td>"
            f"<td style='padding:5px 10px;color:#c0392b;font-weight:600'>ST/ANT estimado por rateio</td></tr>"
            for a in rateio_alertas
        )
        alertas_html += f"""
        <div style="margin-top:20px">
          <div style="font-weight:700;font-size:13px;color:#c0392b;margin-bottom:8px">
            ⚠️ {len(rateio_alertas)} produto(s) com ST/ANT estimado por rateio — confira manualmente
          </div>
          <table style="width:100%;border-collapse:collapse;font-size:12px;background:#fdf2f2;
                        border-radius:8px;overflow:hidden">
            <thead><tr style="background:#fdecea">
              <th style="padding:6px 10px;text-align:left">REF</th>
              <th style="padding:6px 10px;text-align:left">DESCRIÇÃO</th>
              <th style="padding:6px 10px;text-align:left">STATUS</th>
            </tr></thead>
            <tbody>{linhas_rateio}</tbody>
          </table>
        </div>"""

    # ── Gráfico de lucro por produto ─────────────────────────────────────────
    chart_html = ""
    if metricas:
        produtos = sorted(
            [{'desc': r['desc'][:48], 'lucro': round(m['lucro'], 2)}
             for r, m in zip(rows_data, metricas)],
            key=lambda x: x['lucro'], reverse=True
        )
        if produtos:
            max_abs   = max(abs(p['lucro']) for p in produtos) or 1
            lucro_100 = sum(p['lucro'] for p in produtos)
            cjson     = json.dumps(produtos, ensure_ascii=False)

            bars = ""
            for i, p in enumerate(produtos):
                w   = abs(p['lucro']) / max_abs * 100
                cor = '#BDD7EE' if p['lucro'] >= 0 else '#FCE4D6'
                lc  = '#1a252f' if p['lucro'] >= 0 else '#c0392b'
                bars += (
                    f'<div style="display:flex;align-items:center;gap:10px">'
                    f'<span style="font-size:11px;color:#555;width:195px;min-width:195px;'
                    f'overflow:hidden;text-overflow:ellipsis;white-space:nowrap" '
                    f'title="{html.escape(p["desc"])}">{html.escape(p["desc"])}</span>'
                    f'<div style="flex:1;background:#f0f4f8;border-radius:4px;height:18px;overflow:hidden">'
                    f'<div id="ph-bar-{i}" style="height:100%;width:{w:.1f}%;background:{cor};'
                    f'border-radius:4px;transition:width 0.25s ease"></div></div>'
                    f'<span id="ph-lbl-{i}" style="font-size:11px;font-weight:700;color:{lc};'
                    f'min-width:100px;text-align:right">{fmt_brl(p["lucro"])}</span>'
                    f'</div>'
                )

            chart_html = f"""
      <div style="margin-top:24px;border-top:1px solid #e8e8e8;padding-top:20px">
        <div style="margin-bottom:16px">
          <div style="font-size:15px;font-weight:700;color:#1a1a2e">Projeção de Lucratividade por SKU</div>
          <div style="font-size:12px;color:#888;margin-top:3px">Lucro bruto baseado na estimativa de conversão de vendas</div>
        </div>
        <div style="display:flex;align-items:center;gap:12px;margin-bottom:14px">
          <span style="font-size:12px;color:#666;white-space:nowrap">% Vendido:</span>
          <input type="range" id="ph-slider" min="0" max="100" value="100"
                 style="flex:1;accent-color:#3b82f6;cursor:pointer"
                 oninput="phUpdate(this.value)">
          <span id="ph-pct" style="font-size:14px;font-weight:700;color:#1a1a2e;
                                    min-width:42px;text-align:right">100%</span>
        </div>
        <div style="background:#EBF5FB;border-radius:8px;padding:10px 16px;margin-bottom:14px;
                    display:flex;justify-content:space-between;align-items:center">
          <span style="font-size:11px;color:#666;font-weight:600;text-transform:uppercase;
                       letter-spacing:.5px">Total Projetado</span>
          <span id="ph-total" style="font-size:20px;font-weight:700;color:#1a252f">{fmt_brl(lucro_100)}</span>
        </div>
        <div style="max-height:380px;overflow-y:auto;padding-right:6px;
                    display:flex;flex-direction:column;gap:7px">
          {bars}
        </div>
      </div>
      <script>
      (function(){{
        var D={cjson}, MAX={max_abs};
        function brl(v){{
          var n=v<0, s=Math.abs(v).toFixed(2).replace('.',',')
            .replace(/\B(?=(\d{{3}})+(?!\d))/g,'.');
          return (n?'-':'')+'R$ '+s;
        }}
        window.phUpdate=function(pct){{
          var f=pct/100;
          document.getElementById('ph-pct').textContent=pct+'%';
          var tot=0;
          D.forEach(function(p,i){{
            var v=p.lucro*f; tot+=v;
            var b=document.getElementById('ph-bar-'+i);
            var l=document.getElementById('ph-lbl-'+i);
            if(b) b.style.width=(Math.abs(p.lucro)/MAX*100*f)+'%';
            if(l){{l.textContent=brl(v); l.style.color=v<0?'#c0392b':'#1a252f';}}
          }});
          var t=document.getElementById('ph-total');
          if(t){{t.textContent=brl(tot); t.style.color=tot<0?'#c0392b':'#1a252f';}}
        }};
      }})();
      </script>"""

    return f"""
    <div style="font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;
                margin-top:32px;padding:28px 32px;background:#fff;
                border-radius:14px;border:1px solid #e0e0e0;
                box-shadow:0 2px 12px rgba(0,0,0,.06)">
      <div style="margin-bottom:20px;padding-bottom:16px;border-bottom:2px solid #BDD7EE">
        {nf_ref}
        <h2 style="margin:0;font-size:20px;font-weight:700;color:#1a252f">
          Visão Geral da Operação
        </h2>
      </div>
      {alert_banner}
      <div style="display:flex;gap:12px;flex-wrap:wrap">{cards}</div>
      {barra}
      {cst_html}
      {chart_html}
      {alertas_html}
    </div>"""