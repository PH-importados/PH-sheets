"""
Testes de extrair_qtd_embalagem().

Esta função é a mais frágil do sistema — qualquer ajuste nos 4 padrões regex
pode introduzir regressões silenciosas. Cada padrão e cada proteção anti-absurdo
tem pelo menos um teste dedicado.

Assinatura: extrair_qtd_embalagem(desc_xml, v_un_xml, p_sys, mult)
  - desc_xml : descrição do produto na NFe
  - v_un_xml : valor unitário na NF (preço de custo)
  - p_sys    : preço atual no sistema (0 se não cadastrado)
  - mult     : multiplicador varejo (padrão 2.0)
"""
import pytest
from core.processador import extrair_qtd_embalagem


# ─── Padrão 1 — Explícito (CAIXA COM N, PCT C/N, KIT COM N) ──────────────────

class TestPadraoExplicito:
    """
    Padrão 1: r'(?:CAIXA COM|PACOTE COM|KIT COM|PCT\\s*C/|CX\\s*C/|CX/|C/)\\s*(\\d+)'
    Suficientemente explícito para dispensar o anti-absurdo de preço.
    """

    def test_caixa_com(self):
        assert extrair_qtd_embalagem("CAIXA COM 12 CANECAS", 60.0, 6.0, 2.0) == 12

    def test_pacote_com(self):
        assert extrair_qtd_embalagem("PACOTE COM 24 COPOS", 48.0, 2.5, 2.0) == 24

    def test_kit_com(self):
        assert extrair_qtd_embalagem("KIT COM 6 PECAS COLORIDAS", 18.0, 4.0, 2.0) == 6

    def test_pct_barra(self):
        assert extrair_qtd_embalagem("PCT C/24 GUARDANAPOS", 12.0, 0.80, 2.0) == 24

    def test_pct_com_espaco(self):
        assert extrair_qtd_embalagem("PCT  C/24 GUARDANAPOS", 12.0, 0.80, 2.0) == 24

    def test_cx_barra(self):
        assert extrair_qtd_embalagem("CX C/10 PRATOS BRANCO", 30.0, 4.0, 2.0) == 10

    def test_case_insensitive(self):
        assert extrair_qtd_embalagem("caixa com 6 tigelas", 18.0, 4.0, 2.0) == 6

    def test_um_nao_retorna_1(self):
        # qtd detectada = 1 → retorna 1 sem multiplicar
        assert extrair_qtd_embalagem("CAIXA COM 1 PRODUTO", 10.0, 20.0, 2.0) == 1

    # --- CX/N (KEHOME) — NÃO divide quando uCom=UN na NF ---
    # Quando a NF emite uCom=UN, a NF_U já é por unidade; CX/N é só embalagem de transporte.
    # Ex: "TACA 6PCS CX/4" com qCom=8 UN vUnCom=R$9.66 → 8 conjuntos a R$9.66 (não dividir por 4).

    def test_cx_slash_nao_detectado_ucom_un(self):
        # uCom=UN → NF_U já é por unidade → qtd_emb=1 (sem divisão)
        # Testar com NF_U=9.66 CX/4: se dividisse, daria R$2.41 (errado para uCom=UN)
        assert extrair_qtd_embalagem(
            "88002-TACA COQUETEL DIAMOND 310ML 6PCS CX/4", 9.6625, 0.01, 2.0
        ) == 1

    def test_6pcs_nao_detectado(self):
        # "6PCS" no meio da descrição = peças do produto, não embalagem
        assert extrair_qtd_embalagem("TACA 310ML 6PCS UNITARIA", 9.0, 0.01, 2.0) == 1

    def test_caneca_cx48_ucom_un_nao_divide(self):
        # uCom=UN, qCom=48 canecas: NF_U=R$1.96/caneca individual, não por caixa.
        # Anti-absurdo #2 também protegeria (1.96/48=R$0.04 < R$0.10), mas a lógica
        # principal é que CX/48 aqui é info de transporte, não divisor.
        assert extrair_qtd_embalagem(
            "88109-CANECA PERSONALIZADA VOVO 380ML CX/48", 1.959, 0.01, 2.0
        ) == 1


# ─── Padrão 2 — Número no início ("12 CANECAS") ──────────────────────────────

class TestPadraoInicio:
    """
    Padrão 2: r'^(\\d+)\\s+(?!(?:ML|MG|KG|GR|LT|L|CM|MM|M|UN|PC|PCS|PCT|UNID)\\b)'
    Ambíguo — exige p_sys > 0.02 e passa pelo anti-absurdo de preço.
    """

    def test_numero_antes_produto(self):
        # 12 canecas, p_sys=6.0, v_un_xml=30.0, mult=2.0
        # anti-absurdo: p_sys*12=72 vs 30*2*3=180 → 72 <= 180 → OK
        assert extrair_qtd_embalagem("12 CANECAS INOX 300ML", 30.0, 6.0, 2.0) == 12

    def test_numero_6_tigelas(self):
        # p_sys*6=30 vs 12*2*3=72 → OK
        assert extrair_qtd_embalagem("6 TIGELAS COLORIDAS", 12.0, 5.0, 2.0) == 6

    # --- Falsos positivos que DEVEM retornar 1 ---

    def test_falso_positivo_ml(self):
        # "500 ML" → ML está na lista de exclusão → qtd = 1
        assert extrair_qtd_embalagem("500 ML AGUA MINERAL", 2.0, 2.0, 2.0) == 1

    def test_falso_positivo_mg(self):
        assert extrair_qtd_embalagem("500 MG SUPLEMENTO", 15.0, 15.0, 2.0) == 1

    def test_falso_positivo_kg(self):
        assert extrair_qtd_embalagem("5 KG ARROZ", 10.0, 10.0, 2.0) == 1

    def test_falso_positivo_gr(self):
        assert extrair_qtd_embalagem("250 GR CAFE", 8.0, 8.0, 2.0) == 1

    def test_falso_positivo_lt(self):
        assert extrair_qtd_embalagem("2 LT OLEO", 5.0, 5.0, 2.0) == 1

    def test_falso_positivo_un(self):
        assert extrair_qtd_embalagem("1 UN PRODUTO", 10.0, 10.0, 2.0) == 1

    def test_falso_positivo_pcs(self):
        assert extrair_qtd_embalagem("10 PCS PARAFUSO", 5.0, 5.0, 2.0) == 1

    # --- Proteção: p_sys muito baixo (placeholder) ---

    def test_sem_preco_sistema_retorna_1(self):
        # p_sys <= 0.02 → ambíguo sem preço real → retorna 1
        assert extrair_qtd_embalagem("12 CANECAS INOX", 30.0, 0.0, 2.0) == 1

    def test_preco_placeholder_retorna_1(self):
        assert extrair_qtd_embalagem("6 TIGELAS", 12.0, 0.01, 2.0) == 1

    # --- Anti-absurdo: p_sys × qtd > v_un_xml × mult × 3 ---

    def test_anti_absurdo_preco(self):
        # p_sys=50.0, qtd=12 → 50*12=600 vs 30*2*3=180 → 600 > 180 → retorna 1
        # Significa: o sistema já cadastrou o preço por caixa, não por unidade
        assert extrair_qtd_embalagem("12 CANECAS", 30.0, 50.0, 2.0) == 1


# ─── Padrão 3 — Sufixo explícito (PT24UN, DP 18 UN) ─────────────────────────

class TestPadraoFinal:
    """
    Padrão 3: r'(?:PT|DP|POTE|POLYBAG|DISPLAY)\\s*(\\d+)\\s*(?:UN|U(?:\\s|$|\\]))?'
    Explícito como o Padrão 1 — dispensa anti-absurdo de preço.
    """

    def test_pt_sem_espaco(self):
        assert extrair_qtd_embalagem("CANETA ESFER PT24UN", 12.0, 0.80, 2.0) == 24

    def test_dp_com_espaco(self):
        assert extrair_qtd_embalagem("CANETA DP 18 UN AZUL", 18.0, 1.50, 2.0) == 18

    def test_pote(self):
        assert extrair_qtd_embalagem("CLIPS POTE 48 UN", 24.0, 0.60, 2.0) == 48

    def test_display(self):
        assert extrair_qtd_embalagem("DISPLAY 12 PORTA CANETAS", 60.0, 5.0, 2.0) == 12

    def test_polybag(self):
        assert extrair_qtd_embalagem("BROCHE POLYBAG 36", 18.0, 0.50, 2.0) == 36


# ─── Padrão 4 — DS Display ("DS NOME - 84") ──────────────────────────────────

class TestPadraoDS:
    """
    Padrão 4: r'^DS\\b.*-\\s*(\\d+)\\s*(?:PCS)?\\s*$'
    Fornecedores de papelaria: DS = Display com N unidades.
    Número vem SEMPRE após o último traço.
    """

    def test_ds_simples(self):
        assert extrair_qtd_embalagem("DS CANECAS OXFORD - 84", 420.0, 5.0, 2.0) == 84

    def test_ds_com_pcs(self):
        assert extrair_qtd_embalagem("DS CLIPS SORTIDOS - 207 PCS", 207.0, 1.0, 2.0) == 207

    def test_ds_case_insensitive(self):
        assert extrair_qtd_embalagem("ds porta lapis - 36", 90.0, 2.50, 2.0) == 36

    def test_ds_sem_traco_nao_detecta(self):
        # Sem traço = não é padrão DS → retorna 1
        assert extrair_qtd_embalagem("DS PRODUTO QUALQUER", 10.0, 5.0, 2.0) == 1


# ─── Anti-absurdo universal: preço unitário mínimo ───────────────────────────

class TestAntiAbsurdoUnitario:
    """
    Se v_un_xml / qtd < R$0.10 → impossível ser embalagem real → retorna 1.
    Protege contra descrições como "CAIXA COM 500" para um produto de R$2.
    """

    def test_valor_unitario_minimo(self):
        # v_un_xml=2.0, qtd=100 → unitário=0.02 < 0.10 → retorna 1
        assert extrair_qtd_embalagem("CAIXA COM 100 UNIDADES", 2.0, 0.05, 2.0) == 1

    def test_valor_unitario_ok(self):
        # v_un_xml=60.0, qtd=12 → unitário=5.0 >= 0.10 → retorna 12
        assert extrair_qtd_embalagem("CAIXA COM 12 UNIDADES", 60.0, 6.0, 2.0) == 12


# ─── Produto sem embalagem ────────────────────────────────────────────────────

class TestSemEmbalagem:
    def test_produto_simples(self):
        assert extrair_qtd_embalagem("CANETA ESFEROGRAFICA AZUL", 1.50, 3.0, 2.0) == 1

    def test_descricao_vazia(self):
        assert extrair_qtd_embalagem("", 5.0, 10.0, 2.0) == 1

    def test_descricao_numerica_no_meio(self):
        # Número no meio (não no início) não dispara Padrão 2
        assert extrair_qtd_embalagem("CANETA REF 123 AZUL", 1.50, 3.0, 2.0) == 1


# ─── Prioridade dos padrões ───────────────────────────────────────────────────

class TestPrioridadeDePatterns:
    """
    Padrão 1 tem prioridade sobre Padrão 2.
    Se ambos matcham, Padrão 1 deve vencer.
    """

    def test_padrao1_prevalece_sobre_padrao2(self):
        # "12 CANECAS CAIXA COM 6" → Padrão 1 encontra 6, Padrão 2 encontraria 12
        # Padrão 1 é testado primeiro → retorna 6
        assert extrair_qtd_embalagem("12 CANECAS CAIXA COM 6 UND", 30.0, 3.0, 2.0) == 6

    def test_cx_prevalece_sobre_c_generico(self):
        # REGRESSÃO (NF Akash): "CONJUNTO C/ 03 PCS... CX C/120"
        # C/ genérico está antes do CX C/ na descrição — o genérico matchava primeiro.
        # Após a separação em _PATTERN_EMB_ESPECIFICO vs _PATTERN_EMB_GENERICO,
        # CX C/ deve ser encontrado primeiro → retorna 120.
        desc = "CONJUNTO C/ 03 PCS DE GANCHO MULTIUSO, DE METAL COM VENTOSA CX C/120"
        assert extrair_qtd_embalagem(desc, 192.0, 2.0, 2.0) == 120

    def test_c_generico_funciona_quando_unico(self):
        # Sem padrões específicos na descrição, C/ genérico ainda deve funcionar.
        assert extrair_qtd_embalagem("PRODUTO GENERICO C/5", 25.0, 5.0, 2.0) == 5


# ─── Padrão 5 — IP-N (itens por pacote) ──────────────────────────────────────

class TestPadraoIP:
    """
    Padrão 5: r'\\bIP-(\\d+)\\b'
    Convenção de fornecedores (ex: Affinity Trade): "IP-18 TRB" = 18 itens/caixa.
    Presente no NOME do sistema, ausente no XML (descrição truncada a ~120 chars).

    REGRESSÃO (NF 62 / Affinity Trade):
      uCom=CAIXAS, vUnCom=369.85 (preço por caixa de 18 unidades).
      XML truncado não continha "IP-18"; sistema continha "- IP-18 TRB".
      Sem a detecção, nf_u = 369.85 (errado) em vez de 369.85/18 = 20.55.
    """

    def test_ip_no_desc_xml(self):
        # IP-N presente no próprio XML → detecta direto
        assert extrair_qtd_embalagem("BLOCOS DE MONTAR IP-18 TRB", 369.85, 179.90, 2.0) == 18

    def test_ip_apenas_no_desc_sys(self):
        # XML truncado, sem IP-N; NOME do sistema tem "IP-18 TRB" → desc_sys detecta
        desc_xml = "LJ-001 - MODELO REF:. MON04 - BLOCOS DE MONTAR - AZUL, VERDE, AMARELO"
        desc_sys = "LJ-001 - REF MON04 - BLOCOS DE MONTAR - AZUL VERDE AMARELO - IP-18 TRB"
        assert extrair_qtd_embalagem(desc_xml, 369.85, 179.90, 2.0, desc_sys) == 18

    @pytest.mark.parametrize("ref,vUnCom,desc_sys,esperado", [
        ("MON04", 369.85, "LJ-001 - REF MON04 - BLOCOS DE MONTAR - IP-18 TRB", 18),
        ("MON05", 383.62, "LJ-001 - REF MON05 - BLOCOS DE MONTAR - IP-36 TRB", 36),
        ("PLA007", 287.72, "LJ1-11-2 - REF PLA007 - BRINQUEDO LANCA AGUA - IP-60 TRB", 60),
        ("PLA005", 397.29, "LJ-555 - REF PLA005 - BRINQUEDO LANCA AGUA - IP-120 TRB", 120),
    ])
    def test_affinity_trade_nf62_todos_produtos(self, ref, vUnCom, desc_sys, esperado):
        """Valida os 4 produtos da NF 62 da Affinity Trade — todos uCom=CAIXAS."""
        desc_xml = f"LJ-001 - MODELO REF:. {ref} - PRODUTO BRINQUEDO"  # truncado, sem IP-N
        assert extrair_qtd_embalagem(desc_xml, vUnCom, 50.0, 2.0, desc_sys) == esperado

    def test_ip_anti_absurdo_unitario(self):
        # IP-500 numa caixa de R$2 → unitário R$0.004 < R$0.10 → retorna 1
        assert extrair_qtd_embalagem("PRODUTO IP-500", 2.0, 5.0, 2.0) == 1

    # REGRESSÃO (NF 57472 / Bambola): uCom=PC, vUnCom=6.60 já é preço por boneca.
    # "IP-12 TRB" no NOME do sistema é só a caixa de transporte — dividia por 12
    # e o marketplace precificava a boneca sobre R$0.55.
    @pytest.mark.parametrize("u_com", ['PC', 'UN', 'UND', 'pc', ' UN '])
    def test_ip_ignorado_quando_ucom_unitario(self, u_com):
        assert extrair_qtd_embalagem(
            "BONECA TITI CANTIGAS", 6.60, 49.99, 2.0,
            "BONECA TITI CANTIGAS - IP-12 TRB REF 904", u_com) == 1

    @pytest.mark.parametrize("u_com", ['CAIXAS', 'CX', ''])
    def test_ip_mantido_quando_ucom_caixa(self, u_com):
        desc_sys = "LJ-001 - REF MON04 - BLOCOS DE MONTAR - IP-18 TRB"
        assert extrair_qtd_embalagem("BLOCOS DE MONTAR", 369.85, 50.0, 2.0, desc_sys, u_com) == 18

    def test_sem_ip_sem_desc_sys(self):
        # Nenhum dos dois tem IP → cai para outros padrões / retorna 1
        assert extrair_qtd_embalagem("PRODUTO GENERICO SEM CONTAGEM", 50.0, 25.0, 2.0) == 1
