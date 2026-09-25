"""
Fixtures compartilhadas entre todos os testes do PH-Sheets.
"""
import pytest


@pytest.fixture
def P():
    """Parâmetros de precificação padrão — espelha os defaults do formulário."""
    return {
        'mult':     2.0,
        'frete':    0.10,
        'desp':     0.10,
        'fed':      0.0913,
        'icm':      0.21,
        'cartao':   0.04,
        'mult_atc':     1.3,
        'desc_atc':     0.15,
        'desc_atc_pdv': 0.10,
    }


@pytest.fixture
def P_zero():
    """Parâmetros zerados — isola o comportamento sendo testado sem ruído fiscal."""
    return {
        'mult':         1.0,
        'frete':        0.0,
        'desp':         0.0,
        'fed':          0.0,
        'icm':          0.0,
        'cartao':       0.0,
        'mult_atc':     1.0,
        'desc_atc':     0.0,
        'desc_atc_pdv': 0.0,
    }


def make_row(**kwargs):
    """
    Fábrica de dicionários de produto para _calcular().
    Preenche todos os campos obrigatórios com valores neutros;
    use kwargs para sobrescrever apenas o que o teste precisa.
    """
    defaults = {
        'nf_u':    10.0,
        'st_u':    0.0,
        'ant_u':   0.0,
        'ipi_u':   0.0,
        'cst':     '00',
        'qtd':     1,
        'p_atual': 0.0,
        'cred_pct': 0.04,
        'qtd_emb': 1,
        'tem_st':  False,
        'tem_ant': False,
    }
    defaults.update(kwargs)
    return defaults
