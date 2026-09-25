"""
Testes de _calcular() — motor de precificação central do PH-Sheets.

Esta função é chamada para cada produto ao gerar a prévia e calcular métricas
do dashboard. Bugs aqui afetam TODOS os produtos silenciosamente.

Regras críticas que mais causam regressões:
  1. ST > 0.005  → zera crédito ICMS, ICMS saída e ICMS atacado
  2. ANT sozinho → NÃO zera crédito ICMS (ANT não é ST)
  3. ANT sozinho → REDUZ ICMS saída pelo valor do ANT pago (max 0)
  4. CSTs/CSOSN sem crédito (CST_SEM_CREDITO) → zeram crédito ICMS
  5. p_atual > 0 → usa como preço varejo (não recalcula)
  6. p_atual = 0 → fallback analítico (p_var >= p_min)
"""
import pytest
from tests.conftest import make_row
from app import _calcular


# ─── Crédito ICMS — Regras de zeragem ────────────────────────────────────────

class TestCreditoICMS:
    """
    Regra: cred = nf_u × cred_pct  EXCETO quando há ST ou CST isento.
    ANT sozinho NÃO zera o crédito — este é o erro mais comum de regressão.
    """

    def test_produto_normal_tem_credito(self, P_zero):
        # Produto comum, sem ST, CST 00 → crédito deve ser calculado
        row = make_row(nf_u=10.0, cst='00', cred_pct=0.10, p_atual=20.0)
        m = _calcular(row, P_zero)
        assert m['cred'] == pytest.approx(1.0)  # 10.0 × 0.10

    def test_st_zera_credito(self, P_zero):
        row = make_row(nf_u=10.0, st_u=3.0, cst='00', cred_pct=0.10, p_atual=20.0)
        m = _calcular(row, P_zero)
        assert m['cred'] == 0.0

    def test_ant_sozinho_nao_zera_credito(self, P_zero):
        """
        REGRESSÃO CRÍTICA: ANT ≠ ST. Produto com antecipação mas sem ST
        ainda tem direito ao crédito de ICMS. Ver RULES.md §3.
        """
        row = make_row(nf_u=10.0, ant_u=2.0, st_u=0.0, cst='00',
                       cred_pct=0.10, p_atual=20.0)
        m = _calcular(row, P_zero)
        assert m['cred'] == pytest.approx(1.0), (
            "ANT não deve zerar crédito ICMS — apenas ST e CSTs isentos fazem isso"
        )

    @pytest.mark.parametrize("cst", ['40', '41', '50', '60', '102', '103', '202', '203', '300', '400', '500'])
    def test_cst_isento_zera_credito(self, P_zero, cst):
        row = make_row(nf_u=10.0, cst=cst, cred_pct=0.10, p_atual=20.0)
        m = _calcular(row, P_zero)
        assert m['cred'] == 0.0, f"CST {cst} deve zerar crédito ICMS"

    @pytest.mark.parametrize("cst", ['00', '10', '20', '101', '201'])
    def test_cst_tributado_mantem_credito(self, P_zero, cst):
        row = make_row(nf_u=10.0, cst=cst, cred_pct=0.10, p_atual=20.0)
        m = _calcular(row, P_zero)
        assert m['cred'] == pytest.approx(1.0), f"CST {cst} não deve zerar crédito"


# ─── ICMS Saída — Zerado por ST ───────────────────────────────────────────────

class TestICMSSaida:
    def test_produto_normal_tem_icms_saida(self, P_zero):
        P_zero['icm'] = 0.21
        row = make_row(nf_u=10.0, st_u=0.0, p_atual=20.0)
        m = _calcular(row, P_zero)
        assert m['icms_s'] == pytest.approx(4.20)  # 20.0 × 0.21

    def test_st_zera_icms_saida(self, P_zero):
        P_zero['icm'] = 0.21
        row = make_row(nf_u=10.0, st_u=3.0, p_atual=20.0)
        m = _calcular(row, P_zero)
        assert m['icms_s'] == 0.0

    def test_ant_reduz_icms_saida(self, P_zero):
        """
        ANT não zera o ICMS, mas abate o valor já antecipado.
        icms_s = max(0, p_var × icm% − ant_u) = max(0, 4.20 − 2.0) = 2.20
        """
        P_zero['icm'] = 0.21
        row = make_row(nf_u=10.0, ant_u=2.0, st_u=0.0, p_atual=20.0)
        m = _calcular(row, P_zero)
        assert m['icms_s'] == pytest.approx(2.20)

    def test_ant_nao_zera_icms_abaixo_de_zero(self, P_zero):
        """ANT maior que o ICMS calculado resulta em 0, não negativo."""
        P_zero['icm'] = 0.10
        row = make_row(nf_u=10.0, ant_u=5.0, st_u=0.0, p_atual=20.0)
        m = _calcular(row, P_zero)
        # icms normal = 20 * 0.10 = 2.0; ant = 5.0 → max(0, 2.0 - 5.0) = 0
        assert m['icms_s'] == pytest.approx(0.0)

    def test_st_zera_icm_atc(self, P_zero):
        P_zero['icm'] = 0.21
        row = make_row(nf_u=10.0, st_u=3.0, p_atual=20.0)
        m = _calcular(row, P_zero)
        assert m['icm_atc'] == 0.0


# ─── Custo de Entrada — Fórmula completa ─────────────────────────────────────

class TestCustoEntrada:
    """
    c_ent = c_real + st_u + ant_u + ipi_u + frete + desp - cred
    """

    def test_c_ent_produto_simples(self, P_zero):
        # cred_pct=0.0 para isolar: c_ent deve ser só c_real
        row = make_row(nf_u=10.0, p_atual=25.0, cred_pct=0.0)
        m = _calcular(row, P_zero)
        assert m['c_ent'] == pytest.approx(10.0)  # só c_real

    def test_c_ent_inclui_st(self, P_zero):
        # ST zera cred → c_ent = c_real + st_u
        row = make_row(nf_u=10.0, st_u=3.0, p_atual=25.0, cred_pct=0.0)
        m = _calcular(row, P_zero)
        assert m['c_ent'] == pytest.approx(13.0)  # 10 + 3

    def test_c_ent_inclui_ant(self, P_zero):
        # cred_pct=0.0 para isolar o efeito do ANT
        row = make_row(nf_u=10.0, ant_u=2.0, p_atual=25.0, cred_pct=0.0)
        m = _calcular(row, P_zero)
        assert m['c_ent'] == pytest.approx(12.0)  # 10 + 2

    def test_c_ent_inclui_ipi(self, P_zero):
        # cred_pct=0.0 para isolar o efeito do IPI
        row = make_row(nf_u=10.0, ipi_u=1.5, p_atual=25.0, cred_pct=0.0)
        m = _calcular(row, P_zero)
        assert m['c_ent'] == pytest.approx(11.5)  # 10 + 1.5

    def test_c_ent_todos_componentes(self):
        """Teste com todos os componentes ativos simultaneamente."""
        # mult=2, frete=10%, desp=10%, cred=0 (tem ST)
        P = {
            'mult': 2.0, 'frete': 0.10, 'desp': 0.10,
            'fed': 0.0, 'icm': 0.0, 'cartao': 0.0,
            'mult_atc': 1.0, 'desc_atc': 0.0,
        }
        row = make_row(nf_u=10.0, st_u=2.0, ant_u=1.0, ipi_u=0.5, p_atual=30.0)
        m = _calcular(row, P)
        # c_real = 10*2 = 20
        # frete  = 20*0.10 = 2
        # desp   = 20*0.10 = 2
        # cred   = 0 (st_u > 0.005)
        # c_ent  = 20 + 2 + 1 + 0.5 + 2 + 2 - 0 = 27.5
        assert m['c_real'] == pytest.approx(20.0)
        assert m['frete']  == pytest.approx(2.0)
        assert m['desp']   == pytest.approx(2.0)
        assert m['cred']   == 0.0
        assert m['c_ent']  == pytest.approx(27.5)

    def test_c_ent_desconta_credito(self, P_zero):
        """Crédito reduz o custo de entrada."""
        row = make_row(nf_u=10.0, cred_pct=0.10, p_atual=25.0)
        m = _calcular(row, P_zero)
        # c_real=10, cred=1.0 → c_ent = 10 - 1 = 9.0
        assert m['c_ent'] == pytest.approx(9.0)


# ─── Preço Varejo — p_atual vs fallback ──────────────────────────────────────

class TestPrecoVarejo:
    def test_usa_preco_atual_quando_disponivel(self, P_zero):
        row = make_row(nf_u=5.0, p_atual=19.99)
        m = _calcular(row, P_zero)
        assert m['p_var'] == pytest.approx(19.99)

    def test_fallback_retorna_preco_positivo(self, P_zero):
        """Sem preço no sistema, fallback deve retornar p_var > 0."""
        row = make_row(nf_u=10.0, p_atual=0.0)
        m = _calcular(row, P_zero)
        assert m['p_var'] > 0

    def test_fallback_garante_p_var_cobre_c_ent(self, P_zero):
        """Fallback analítico deve resultar em margem viável (p_var >= c_ent)."""
        row = make_row(nf_u=10.0, p_atual=0.0)
        m = _calcular(row, P_zero)
        assert m['p_var'] >= m['c_ent']

    def test_fallback_com_taxes_reais(self, P):
        """Com parâmetros reais, p_var deve cobrir c_saida."""
        row = make_row(nf_u=10.0, p_atual=0.0)
        m = _calcular(row, P)
        assert m['p_var'] >= m['c_saida']


# ─── Custo Saída e Margem ─────────────────────────────────────────────────────

class TestCustoSaidaMargem:
    def test_c_saida_formula(self, P_zero):
        """c_saida = c_ent + fed + cartao + icms_s"""
        P_zero['fed'] = 0.10
        # cred_pct=0.0 para que c_ent=10 e o cálculo seja previsível
        row = make_row(nf_u=10.0, p_atual=20.0, cred_pct=0.0)
        m = _calcular(row, P_zero)
        # c_ent=10, fed=20*0.10=2, cartao=0, icms_s=0 → c_saida=12
        assert m['c_saida'] == pytest.approx(12.0)

    def test_margem_formula(self, P_zero):
        """margem = (p_var - c_saida) / p_var"""
        # cred_pct=0.0 → c_ent=c_saida=5.0, p_var=10.0 → margem=0.5
        row = make_row(nf_u=5.0, p_atual=10.0, cred_pct=0.0)
        m = _calcular(row, P_zero)
        assert m['margem'] == pytest.approx(0.5)

    def test_margem_negativa_possivel(self, P_zero):
        """Produto vendido abaixo do custo deve retornar margem negativa."""
        row = make_row(nf_u=10.0, p_atual=8.0, cred_pct=0.0)
        m = _calcular(row, P_zero)
        assert m['margem'] < 0

    def test_p_min_formula(self, P_zero):
        """p_min = c_saida / (1 - meta) com meta=15%"""
        # cred_pct=0.0 → c_saida=10.0 → p_min=10/(1-0.15)=11.76
        row = make_row(nf_u=10.0, p_atual=20.0, cred_pct=0.0)
        m = _calcular(row, P_zero)
        assert m['p_min'] == pytest.approx(10.0 / 0.85, rel=1e-2)


# ─── Atacado ─────────────────────────────────────────────────────────────────

class TestAtacado:
    def test_nf_atc_formula(self, P_zero):
        P_zero['mult_atc'] = 1.5
        row = make_row(nf_u=10.0, p_atual=20.0)
        m = _calcular(row, P_zero)
        assert m['nf_atc'] == pytest.approx(15.0)  # 10 * 1.5

    def test_p_atc_ped_formula(self, P_zero):
        P_zero['desc_atc'] = 0.20
        row = make_row(nf_u=10.0, p_atual=20.0)
        m = _calcular(row, P_zero)
        assert m['p_atc_ped'] == pytest.approx(16.0)  # 20 * (1 - 0.20)

    def test_p_atc_pdv_formula(self, P_zero):
        """PREÇO ATC PDV usa desc_atc_pdv (padrão 10%)."""
        P_zero['desc_atc_pdv'] = 0.10
        row = make_row(nf_u=10.0, p_atual=20.0)
        m = _calcular(row, P_zero)
        assert m['p_atc_pdv'] == pytest.approx(18.0)  # 20 * (1 - 0.10)

    def test_p_atc_pdv_independente_de_ped(self, P_zero):
        """PDV e PEDIDO podem ter descontos diferentes no mesmo produto."""
        P_zero['desc_atc']     = 0.15
        P_zero['desc_atc_pdv'] = 0.05
        row = make_row(nf_u=10.0, p_atual=20.0)
        m = _calcular(row, P_zero)
        assert m['p_atc_ped'] == pytest.approx(17.0)  # 20 * 0.85
        assert m['p_atc_pdv'] == pytest.approx(19.0)  # 20 * 0.95

    def test_margem_atc_ped_formula(self, P_zero):
        P_zero['desc_atc'] = 0.0   # sem desconto → p_atc_ped = p_var
        # cred_pct=0.0 → c_ent=5.0 → c_saida_atc=5.0 → margem_atc_ped=0.5
        row = make_row(nf_u=5.0, p_atual=10.0, cred_pct=0.0)
        m = _calcular(row, P_zero)
        assert m['margem_atc_ped'] == pytest.approx(0.5)

    def test_margem_atc_pdv_formula(self, P_zero):
        P_zero['desc_atc_pdv'] = 0.0   # sem desconto → p_atc_pdv = p_var
        # cred_pct=0.0 → c_ent=5.0 → c_saida_vrj=5.0 → margem_atc_pdv=0.5
        row = make_row(nf_u=5.0, p_atual=10.0, cred_pct=0.0)
        m = _calcular(row, P_zero)
        assert m['margem_atc_pdv'] == pytest.approx(0.5)

    def test_margem_atc_pdv_usa_c_saida_varejo(self, P_zero):
        """MARGEM ATC PDV usa c_saida do varejo (não c_saida_atc).

        Com taxas de varejo não-zero, c_saida > c_saida_atc porque varejo
        aplica os impostos sobre o preço de venda (mais alto), enquanto atacado
        aplica sobre NF_ATC. A margem PDV deve refletir o custo de saída varejo.
        """
        P_zero['fed'] = 0.10
        P_zero['icm'] = 0.10
        # desc_atc_pdv=0 → p_atc_pdv = p_var = 20; sem desconto para isolar o custo
        row = make_row(nf_u=10.0, p_atual=20.0, cred_pct=0.0)
        m = _calcular(row, P_zero)
        # c_saida     = c_ent(10) + fed(2.0) + icms_s(2.0) = 14.0
        # c_saida_atc = c_ent(10) + fed_atc(1.0) + icm_atc(1.0) = 12.0
        # margem_atc_pdv = (20 - 14) / 20 = 0.3  (usa c_saida, não c_saida_atc)
        assert m['c_saida']     == pytest.approx(14.0)
        assert m['c_saida_atc'] == pytest.approx(12.0)
        assert m['margem_atc_pdv'] == pytest.approx(0.3)

    def test_margem_atc_pdv_maior_que_ped(self, P_zero):
        """PDV tem menor desconto → preço mais alto → margem maior que PED."""
        P_zero['desc_atc']     = 0.20   # 20% desc pedido
        P_zero['desc_atc_pdv'] = 0.10   # 10% desc PDV
        row = make_row(nf_u=5.0, p_atual=10.0, cred_pct=0.0)
        m = _calcular(row, P_zero)
        assert m['margem_atc_pdv'] > m['margem_atc_ped']

    def test_c_saida_atc_usa_c_ent_varejo(self, P_zero):
        """c_saida_atc usa o mesmo c_ent do varejo (custo de entrada é igual)."""
        row = make_row(nf_u=10.0, p_atual=30.0)
        m = _calcular(row, P_zero)
        # c_ent = 10.0, fed_atc=cart_atc=icm_atc=0 → c_saida_atc = 10.0
        assert m['c_saida_atc'] == pytest.approx(m['c_ent'])

    def test_st_zera_icm_atc_mas_nao_outros(self, P_zero):
        P_zero['fed'] = 0.10
        P_zero['cartao'] = 0.05
        P_zero['icm'] = 0.21
        P_zero['mult_atc'] = 1.0
        row = make_row(nf_u=10.0, st_u=2.0, p_atual=20.0)
        m = _calcular(row, P_zero)
        assert m['icm_atc'] == 0.0
        # fed_atc = nf_atc * fed = 10 * 0.10 = 1.0 → ainda calculado
        assert m['fed_atc'] == pytest.approx(1.0)
