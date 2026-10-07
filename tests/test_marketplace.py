"""
Testes do modo marketplace (core/marketplace.py) — Shopee e Mercado Livre.
"""
import os
import tempfile

import pytest
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

from tests.conftest import make_row
from core import marketplace as mkt
from core.marketplace import (
    calcular_marketplace, custo_entrada, deducoes_venda, taxas_shopee, taxas_ml,
    arredondar_x9_acima, frete_ml_item, montar_params, salvar_excel_marketplace,
    COL_MKT, ML_LIMITE_FRETE,
)


@pytest.fixture
def M():
    """Defaults do formulário do modo marketplace."""
    return montar_params({})


@pytest.fixture
def M_zero():
    """Sem Simples, despesa nem margem, multiplicador 1 — isola as taxas de marketplace."""
    return {**montar_params({}), 'mult': 1.0, 'frete': 0.0, 'desp': 0.0,
            'simples': 0.0, 'margem': 0.0, 'emb': 0.0}


class TestParams:
    def test_defaults(self, M):
        assert M['mult'] == pytest.approx(2.0)
        assert M['margem'] == pytest.approx(0.15)
        assert M['com_ml'] == pytest.approx(0.165)
        assert M['simples'] == pytest.approx(0.15)
        assert M['desp'] == pytest.approx(0.05)
        assert M['reputacao'] == 'verde'

    def test_despesa_do_varejo_nao_afeta_marketplace(self):
        # 'desp' é o campo do varejo; o marketplace tem o próprio 'desp_mkt'
        assert montar_params({'desp': '10'})['desp'] == pytest.approx(0.05)
        assert montar_params({'desp_mkt': '7'})['desp'] == pytest.approx(0.07)

    def test_frete_pct_sobrescreve_formulario(self):
        # NF CIF: gerar_tabela zera o frete e o marketplace deve respeitar
        assert montar_params({'frete': '10'}, frete_pct=0.0)['frete'] == 0.0

    def test_reputacao_invalida_vira_verde(self):
        assert montar_params({'reputacao_ml': 'azul'})['reputacao'] == 'verde'

    def test_aceita_virgula(self):
        assert montar_params({'mult_mkt': '1,8'})['mult'] == pytest.approx(1.8)


class TestCustoEntrada:
    def test_frete_sobre_custo_real(self, M):
        # c_real 20; frete 10% de 20 = 2 → 22
        e = custo_entrada(make_row(nf_u=10.0), M)
        assert e['c_real'] == pytest.approx(20.0)
        assert e['frete'] == pytest.approx(2.0)
        assert e['c_ent'] == pytest.approx(22.0)

    def test_multiplicador_variavel(self, M):
        e = custo_entrada(make_row(nf_u=10.0), {**M, 'mult': 1.5, 'frete': 0})
        assert e['c_ent'] == pytest.approx(15.0)

    def test_simples_nao_abate_credito_icms(self, M):
        # Mesmo com ICMS destacado na NF, Simples Nacional não aproveita crédito
        com = custo_entrada(make_row(nf_u=10.0, cred_pct=0.12), M)
        sem = custo_entrada(make_row(nf_u=10.0, cred_pct=0.0), M)
        assert com['c_ent'] == pytest.approx(sem['c_ent'])
        assert 'cred' not in com

    @pytest.mark.parametrize("campo", ['st_u', 'ant_u', 'ipi_u'])
    def test_st_ant_ipi_somam_no_custo(self, M, campo):
        e = custo_entrada(make_row(nf_u=10.0, **{campo: 3.0}), {**M, 'frete': 0})
        assert e['c_ent'] == pytest.approx(23.0)

    def test_despesa_nao_entra_no_custo(self, M):
        e = custo_entrada(make_row(nf_u=10.0), {**M, 'frete': 0, 'desp': 0.5})
        assert e['c_ent'] == pytest.approx(20.0)


class TestDeducoesVenda:
    def test_simples_e_despesa_sobre_o_preco(self, M):
        assert deducoes_venda(100.0, M) == (pytest.approx(15.0), pytest.approx(5.0))

    @pytest.mark.parametrize("extra", [{'st_u': 5.0}, {'ant_u': 5.0}])
    def test_st_e_ant_nao_mudam_o_simples(self, M, extra):
        # Simples é % único sobre a venda — ST/ANT não zeram nem abatem
        m = calcular_marketplace(make_row(nf_u=10.0, **extra), M)
        sh = m['shopee']
        assert sh['simples'] == pytest.approx(round(sh['preco'] * 0.15, 2))


class TestTaxas:
    @pytest.mark.parametrize("preco, esperado", [
        (50.00, 4 + 10.0),        # até 79,99: R$4 + 20%
        (79.99, 4 + 16.0),
        (80.00, 16 + 11.2),       # 80–99,99: R$16 + 14%
        (150.00, 20 + 21.0),      # 100–199,99: R$20 + 14%
        (300.00, 26 + 42.0),      # ≥ 200: R$26 + 14%
    ])
    def test_shopee_por_faixa(self, preco, esperado):
        assert taxas_shopee(preco) == pytest.approx(esperado, abs=0.01)

    @pytest.mark.parametrize("preco, esperado", [
        (10.00, 10 * (0.165 + 0.50)),     # até 12,50: comissão + 50% do item
        (20.00, 20 * 0.165 + 6.25),
        (40.00, 40 * 0.165 + 6.75),
        (60.00, 60 * 0.165 + 6.75),
        (100.00, 100 * 0.165 + 10.0),     # ≥ 79: frete 20 com 50% (verde)
    ])
    def test_ml_por_faixa(self, M, preco, esperado):
        assert taxas_ml(preco, M) == pytest.approx(esperado, abs=0.01)

    @pytest.mark.parametrize("reputacao, frete", [('verde', 10.0), ('amarela', 12.0), ('vermelha', 20.0)])
    def test_frete_ml_por_reputacao(self, M, reputacao, frete):
        assert frete_ml_item({**M, 'reputacao': reputacao}) == pytest.approx(frete)

    def test_ml_abaixo_de_79_sem_frete(self, M):
        assert calcular_marketplace(make_row(nf_u=3.0), M)['ml']['frete'] == 0.0


class TestArredondamento:
    @pytest.mark.parametrize("val, esperado", [
        (10.00, 10.09), (10.09, 10.09), (10.10, 10.19), (10.191, 10.29), (0.0, 0.0),
    ])
    def test_x9_acima(self, val, esperado):
        assert arredondar_x9_acima(val) == pytest.approx(esperado)


class TestPrecoSugerido:
    @pytest.mark.parametrize("nf_u", [0.3, 1.0, 3.0, 8.0, 15.0, 30.0, 80.0])
    @pytest.mark.parametrize("extra", [{}, {'st_u': 2.0}, {'ant_u': 1.0}])
    def test_atinge_margem_alvo(self, M, nf_u, extra):
        m = calcular_marketplace(make_row(nf_u=nf_u, **extra), M)
        # Arredondamento de centavos das taxas pode tirar até ~0,1 p.p.
        assert m['shopee']['margem'] >= M['margem'] - 0.001
        assert m['ml']['margem'] >= M['margem'] - 0.001

    def test_lucro_fecha_com_componentes(self, M):
        m = calcular_marketplace(make_row(nf_u=12.0), M)
        for canal in ('shopee', 'ml'):
            c = m[canal]
            assert c['lucro'] == pytest.approx(
                c['preco'] - m['custo_base'] - c['taxas'] - c['simples'] - c['desp'], abs=0.01)

    def test_sem_margem_nem_imposto_preco_cobre_taxas(self, M_zero):
        # custo 50 na Shopee: P = (50 + 4) / 0.8 = 67.5 → 67.59
        m = calcular_marketplace(make_row(nf_u=50.0, cred_pct=0.0), M_zero)
        assert m['shopee']['preco'] == pytest.approx(67.59)

    def test_preco_com_simples_e_despesa(self, M_zero):
        # custo 50, Simples 15%, despesa 5%: P = (50 + 4) / (1 − 0.20 − 0.15 − 0.05) = 90
        # → cai na faixa 80+ (14% + R$16): P = 66 / 0.66 = 100 → faixa 100+ (R$20): 70/0.66 = 106.06
        m = calcular_marketplace(make_row(nf_u=50.0), {**M_zero, 'simples': 0.15, 'desp': 0.05})
        assert m['shopee']['preco'] == pytest.approx(106.09)
        assert m['shopee']['lucro'] >= 0

    def test_piso_da_faixa_quando_faixa_anterior_nao_fecha(self, M_zero):
        # custo 61: faixa ≤79,99 pede 81,25 (fora); faixa 80+ pede 90,70 → fica nela
        m = calcular_marketplace(make_row(nf_u=61.0, cred_pct=0.0), M_zero)
        assert 80.0 <= m['shopee']['preco'] < 100.0
        assert m['shopee']['lucro'] >= 0

    def test_arredondamento_nao_pula_de_faixa(self, M_zero):
        # custo 59.99 → P bruto 79.99 — arredondar p/ X,X9 não pode ir a 80,09 (tarifa R$16)
        m = calcular_marketplace(make_row(nf_u=59.99, cred_pct=0.0), M_zero)
        assert m['shopee']['preco'] < 80.0

    def test_ml_acima_de_79_paga_frete(self, M):
        m = calcular_marketplace(make_row(nf_u=40.0), M)
        assert m['ml']['preco'] >= ML_LIMITE_FRETE
        assert m['ml']['frete'] == pytest.approx(10.0)

    def test_embalagem_entra_no_custo_base(self, M):
        m = calcular_marketplace(make_row(nf_u=10.0), {**M, 'emb': 2.5})
        assert m['custo_base'] == pytest.approx(m['c_ent'] + 2.5)

    def test_margem_inatingivel(self, M):
        m = calcular_marketplace(make_row(nf_u=10.0), {**M, 'margem': 0.80})
        assert m['inviavel'] is True
        assert m['shopee']['preco'] == 0.0

    def test_melhor_canal_e_o_de_maior_lucro(self, M):
        m = calcular_marketplace(make_row(nf_u=10.0), M)
        esperado = mkt.MELHOR_SHOPEE if m['shopee']['lucro'] >= m['ml']['lucro'] else mkt.MELHOR_ML
        assert m['melhor'] == esperado


class TestSerializacao:
    def test_ida_e_volta(self):
        row = make_row(nf='12', desc='X', ref='R', sku='S', rateado=False)
        assert mkt.row_de_json(mkt.row_para_json(row)) == mkt.row_para_json(row)

    def test_converte_tipos(self):
        r = mkt.row_de_json({'nf_u': '10.5', 'qtd': '3', 'cst': 60, 'tem_st': 1})
        assert r['nf_u'] == pytest.approx(10.5)
        assert r['cst'] == '60'
        assert r['tem_st'] is True



class TestDetalhamento:
    @pytest.mark.parametrize("nf_u", [1.0, 10.0, 40.0])
    def test_parcelas_fecham_o_lucro(self, M, nf_u):
        m = calcular_marketplace(make_row(nf_u=nf_u, ant_u=1.0), M)
        for canal in ('shopee', 'ml'):
            c = m[canal]
            parcelas = c['comissao'] + c['tarifa'] + c['frete'] + c['simples'] + c['desp']
            assert c['lucro'] == pytest.approx(c['preco'] - m['custo_base'] - parcelas, abs=0.01)

    def test_ml_acima_de_79_frete_no_lugar_da_tarifa(self, M):
        ml = calcular_marketplace(make_row(nf_u=40.0), M)['ml']
        assert ml['tarifa'] == 0.0
        assert ml['frete'] == pytest.approx(10.0)
        assert ml['com_pct'] == pytest.approx(0.165)

    def test_shopee_comissao_pct_da_faixa(self, M):
        sh = calcular_marketplace(make_row(nf_u=3.0), M)['shopee']
        assert sh['preco'] < 80
        assert sh['com_pct'] == pytest.approx(0.20)
        assert sh['tarifa'] == pytest.approx(4.0)

    def test_base_sem_taxas_de_marketplace(self, M):
        b = calcular_marketplace(make_row(nf_u=10.0), M)['base']
        assert b['comissao'] == 0.0 and b['tarifa'] == 0.0
        assert b['lucro'] / b['preco'] >= M['margem'] - 0.001


def _rows_teste():
    return [
        make_row(nf='1', desc='NORMAL', ref='A', sku='1', nf_u=10.0),
        make_row(nf='1', desc='COM ST', ref='B', sku='2', nf_u=10.0, st_u=3.0, tem_st=True),
        make_row(nf='1', desc='COM ANT', ref='C', sku='3', nf_u=40.0, ant_u=4.0, tem_ant=True),
    ]


def _gerar_wb(M, rows):
    ms = [calcular_marketplace(r, M) for r in rows]
    with tempfile.TemporaryDirectory() as d:
        path = os.path.join(d, 'mkt.xlsx')
        salvar_excel_marketplace(rows, ms, M, '1', path)
        return load_workbook(path), ms


class TestExcel:
    def test_gera_arquivo_com_formulas(self, M):
        wb, ms = _gerar_wb(M, _rows_teste())
        assert wb.sheetnames == ['Marketplace', 'Faixas']
        ws = wb['Marketplace']
        r = mkt.LIN_DADOS
        assert ws.cell(r, COL_MKT['DESC']).value == 'NORMAL'
        for chave in ('C_REAL', 'C_ENT', 'C_BASE', 'P_BASE', 'B_SIMPLES', 'B_DESP', 'SH_COM', 'SH_LUCRO',
                      'ML_COM', 'ML_FRETE', 'ML_LUCRO', 'MELHOR'):
            assert str(ws.cell(r, COL_MKT[chave]).value).startswith('='), chave
        # Preços sugeridos são valores (editáveis), não fórmulas
        assert ws.cell(r, COL_MKT['SH_PRECO']).value == pytest.approx(ms[0]['shopee']['preco'])
        assert ws.cell(r, COL_MKT['ML_PRECO']).value == pytest.approx(ms[0]['ml']['preco'])
        assert wb.calculation.fullCalcOnLoad is True

    @pytest.mark.parametrize("coluna, chave", [
        ('C_REAL', 'mult'), ('FRETE', 'frete'), ('EMB', 'emb'),
        ('MARGEM_ALVO', 'margem'), ('B_SIMPLES', 'simples'), ('B_DESP', 'desp'),
        ('ML_COM_PCT', 'com_ml'), ('ML_FRETE_BASE', 'frete_ml'),
    ])
    def test_variavel_no_topo_da_coluna(self, M, coluna, chave):
        wb, _ = _gerar_wb(M, _rows_teste()[:1])
        ws = wb['Marketplace']
        assert ws.cell(mkt.LIN_VAR, COL_MKT[coluna]).value == pytest.approx(M[chave])

    def test_formula_referencia_variavel_do_topo(self, M):
        wb, _ = _gerar_wb(M, _rows_teste()[:1])
        ws = wb['Marketplace']
        f = ws.cell(mkt.LIN_DADOS, COL_MKT['FRETE']).value
        letra = get_column_letter(COL_MKT['FRETE'])
        assert f'${letra}${mkt.LIN_VAR}' in f

    @pytest.mark.parametrize("base, canais", [
        ('B_SIMPLES', ('SH_SIMPLES', 'ML_SIMPLES')),
        ('B_DESP', ('SH_DESP', 'ML_DESP')),
    ])
    def test_deducoes_de_canal_espelham_variavel_base(self, M, base, canais):
        wb, _ = _gerar_wb(M, _rows_teste()[:1])
        ws = wb['Marketplace']
        letra = get_column_letter(COL_MKT[base])
        for canal in canais:
            assert ws.cell(mkt.LIN_VAR, COL_MKT[canal]).value == f'=${letra}${mkt.LIN_VAR}'

    def test_desconto_reputacao_no_topo_do_frete(self):
        M = montar_params({'reputacao_ml': 'amarela'})
        wb, _ = _gerar_wb(M, _rows_teste()[:1])
        assert wb['Marketplace'].cell(mkt.LIN_VAR, COL_MKT['ML_FRETE']).value == pytest.approx(0.40)


class TestPrevia:
    def test_mostra_variaveis_no_cabecalho(self, M):
        rows = _rows_teste()
        html_ = mkt.gerar_tabela_marketplace_html(rows, [calcular_marketplace(r, M) for r in rows], M)
        assert 'VARIÁVEIS' in html_
        assert '15,00%' in html_      # Simples
        assert '5,00%' in html_       # despesa
        assert '× 2,00' in html_      # multiplicador
        assert '(=) CUSTO ENTRADA' in html_


class TestRotaGerar:
    @pytest.fixture
    def client(self):
        from app import app
        app.config['TESTING'] = True
        return app.test_client()

    def test_gera_so_os_selecionados(self, client):
        produtos = [mkt.row_para_json(r) for r in _rows_teste()[:2]]
        res = client.post('/marketplace/gerar', json={
            'produtos': produtos, 'params': {}, 'num_nf': '99', 'fornecedor': 'TESTE', 'frete_pct': 0.1,
        }).get_json()
        assert res['sucesso'] is True
        assert res['total_itens'] == 2
        assert 'COM ANT' not in res['tabela']
        assert res['download_url'].endswith('.xlsx')

    def test_sem_produtos_da_erro(self, client):
        res = client.post('/marketplace/gerar', json={'produtos': []}).get_json()
        assert res['sucesso'] is False

    def test_pagina_marketplace(self, client):
        assert client.get('/marketplace').status_code == 200
