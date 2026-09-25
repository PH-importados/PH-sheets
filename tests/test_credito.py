"""
Testes de extrair_credito_icms() — crédito de ICMS baseado no valor DESTACADO na NF.

Regra (RULES.md §3): só há crédito se a nota destacar ICMS para o item.
  - Regime normal (CST):  crédito = vICMS
  - Simples (CSOSN):      crédito = vCredICMSSN (vICMS de CSOSN 900 não conta)
  - Nada destacado:       crédito = 0 (na dúvida, não credita)

Bug de origem: NF 875 (Plastropical, CSOSN 103) recebia 4% de crédito padrão
porque o código não estava na lista de isentos.
"""
import xml.etree.ElementTree as ET
import pytest
from core.processador import extrair_credito_icms, faixa_cred, cor_cred, CRED_CORES

NS = {'nfe': 'http://www.portalfiscal.inf.br/nfe'}


def _imposto(grupo_icms):
    xml = (
        '<imposto xmlns="http://www.portalfiscal.inf.br/nfe"><ICMS>'
        f'{grupo_icms}'
        '</ICMS></imposto>'
    )
    return ET.fromstring(xml)


class TestRegimeNormal:
    def test_cst_00_usa_vicms(self):
        imp = _imposto('<ICMS00><orig>0</orig><CST>00</CST><vBC>100.00</vBC>'
                       '<pICMS>7.00</pICMS><vICMS>7.00</vICMS></ICMS00>')
        assert extrair_credito_icms(imp, NS) == ('00', pytest.approx(7.0))

    def test_cst_20_base_reduzida_usa_vicms_nao_picms(self):
        # pICMS 12% sobre base reduzida de 100 → 58.33: crédito é 7.00, não 12.00
        imp = _imposto('<ICMS20><orig>0</orig><CST>20</CST><pRedBC>41.67</pRedBC>'
                       '<vBC>58.33</vBC><pICMS>12.00</pICMS><vICMS>7.00</vICMS></ICMS20>')
        assert extrair_credito_icms(imp, NS) == ('20', pytest.approx(7.0))

    def test_cst_51_diferimento_usa_vicms_liquido(self):
        imp = _imposto('<ICMS51><orig>0</orig><CST>51</CST><vBC>100.00</vBC>'
                       '<pICMS>12.00</pICMS><vICMSOp>12.00</vICMSOp><pDif>100</pDif>'
                       '<vICMSDif>12.00</vICMSDif><vICMS>0.00</vICMS></ICMS51>')
        assert extrair_credito_icms(imp, NS) == ('51', 0.0)

    def test_cst_51_sem_vicms_zera(self):
        imp = _imposto('<ICMS51><orig>0</orig><CST>51</CST></ICMS51>')
        assert extrair_credito_icms(imp, NS) == ('51', 0.0)

    def test_cst_10_nao_confunde_vicmsst_com_vicms(self):
        imp = _imposto('<ICMS10><orig>0</orig><CST>10</CST><vBC>100.00</vBC>'
                       '<pICMS>7.00</pICMS><vICMS>7.00</vICMS><vBCST>150.00</vBCST>'
                       '<pICMSST>19.00</pICMSST><vICMSST>21.50</vICMSST></ICMS10>')
        assert extrair_credito_icms(imp, NS) == ('10', pytest.approx(7.0))

    @pytest.mark.parametrize("grupo,cst", [
        ('<ICMS40><orig>0</orig><CST>40</CST></ICMS40>', '40'),
        ('<ICMS40><orig>0</orig><CST>41</CST></ICMS40>', '41'),
        ('<ICMS60><orig>0</orig><CST>60</CST><vBCSTRet>0.00</vBCSTRet></ICMS60>', '60'),
        ('<ICMS30><orig>0</orig><CST>30</CST><vICMSST>5.00</vICMSST></ICMS30>', '30'),
    ])
    def test_cst_sem_vicms_zera(self, grupo, cst):
        assert extrair_credito_icms(_imposto(grupo), NS) == (cst, 0.0)


class TestSimplesNacional:
    def test_csosn_101_usa_vcredicmssn(self):
        imp = _imposto('<ICMSSN101><orig>0</orig><CSOSN>101</CSOSN>'
                       '<pCredSN>2.56</pCredSN><vCredICMSSN>2.56</vCredICMSSN></ICMSSN101>')
        assert extrair_credito_icms(imp, NS) == ('101', pytest.approx(2.56))

    @pytest.mark.parametrize("csosn", ['102', '103', '300', '400'])
    def test_csosn_sem_credito_zera(self, csosn):
        # NF 875: ICMSSN102 com CSOSN 103 — nenhum valor destacado
        imp = _imposto(f'<ICMSSN102><orig>0</orig><CSOSN>{csosn}</CSOSN></ICMSSN102>')
        assert extrair_credito_icms(imp, NS) == (csosn, 0.0)

    def test_csosn_900_vicms_proprio_nao_gera_credito(self):
        imp = _imposto('<ICMSSN900><orig>0</orig><CSOSN>900</CSOSN><vBC>100.00</vBC>'
                       '<pICMS>7.00</pICMS><vICMS>7.00</vICMS></ICMSSN900>')
        assert extrair_credito_icms(imp, NS) == ('900', 0.0)

    def test_csosn_900_com_vcredicmssn_credita(self):
        imp = _imposto('<ICMSSN900><orig>0</orig><CSOSN>900</CSOSN><vICMS>7.00</vICMS>'
                       '<pCredSN>3.00</pCredSN><vCredICMSSN>3.00</vCredICMSSN></ICMSSN900>')
        assert extrair_credito_icms(imp, NS) == ('900', pytest.approx(3.0))


class TestSemGrupoICMS:
    def test_imposto_none(self):
        assert extrair_credito_icms(None, NS) == ('', 0.0)

    def test_sem_icms(self):
        imp = ET.fromstring('<imposto xmlns="http://www.portalfiscal.inf.br/nfe"/>')
        assert extrair_credito_icms(imp, NS) == ('', 0.0)


class TestFaixaCor:
    @pytest.mark.parametrize("pct,faixa", [
        (0.07, 0.07), (0.0712, 0.07), (0.1195, 0.12), (0.19, 0.19), (0.0, 0.0),
    ])
    def test_taxa_efetiva_proxima_usa_faixa_nominal(self, pct, faixa):
        assert faixa_cred(pct) == faixa

    def test_taxa_simples_fora_das_faixas(self):
        assert faixa_cred(0.0256) is None
        assert cor_cred(0.0256) == 'FFFFFF'

    def test_cor_da_faixa(self):
        assert cor_cred(0.0705) == CRED_CORES[0.07]
