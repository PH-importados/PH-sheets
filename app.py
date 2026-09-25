import os
import sys
import html
import tempfile
import webbrowser
from threading import Timer
from concurrent.futures import ThreadPoolExecutor, as_completed
from flask import Flask, render_template, request, jsonify, send_file

from core.processador import gerar_tabela, salvar_excel_estilizado, gerar_dashboard_html, cor_cred, CST_SEM_CREDITO


def _arredondar_x9(val):
    """Arredonda para X.X9 — mesma lógica da fórmula Excel do P_VAR."""
    if val <= 0:
        return 0.0
    return int(val * 10) / 10 + 0.09


def _calcular(row, P):
    nf_u  = row['nf_u']
    st_u  = row['st_u']
    ant_u = row['ant_u']
    ipi_u = row['ipi_u']
    cst   = str(row['cst'])
    qtd   = row['qtd']

    c_real = round(nf_u * P['mult'], 2)
    frete  = round(c_real * P['frete'], 2)
    desp   = round(c_real * P['desp'], 2)

    if st_u > 0.005 or cst in CST_SEM_CREDITO:
        cred = 0.0
    else:
        cred_pct = row.get('cred_pct', 0.0)
        cred = round(nf_u * cred_pct, 2)

    c_ent = round(c_real + st_u + ant_u + ipi_u + frete + desp - cred, 2)

    meta   = 0.15
    p_atual = row['p_atual']
    if p_atual > 0:
        p_var = p_atual
    else:
        # Fallback analítico: garante que p_var >= p_min sem dependência circular
        icms_pct_var = 0.0 if st_u > 0.005 else P['icm']
        den = 1 - meta - P['fed'] - P['cartao'] - icms_pct_var
        p_min_base = round(c_ent / den, 4) if den > 0 and c_ent > 0 else c_ent
        p_var = _arredondar_x9(p_min_base)

    fed    = round(p_var * P['fed'],    2)
    cartao = round(p_var * P['cartao'], 2)
    # ANT reduz o ICMS da saída: o valor já antecipado na entrada é abatido
    icms_s = 0.0 if st_u > 0.005 else max(0.0, round(p_var * P['icm'], 2) - ant_u)
    c_saida = round(c_ent + fed + cartao + icms_s, 2)
    p_min  = round(c_saida / (1 - meta), 2) if meta > 0 and c_saida > 0 else 0.0
    margem = round((p_var - c_saida) / p_var, 4) if p_var > 0 else 0.0
    lucro  = round((p_var - c_saida) * qtd, 2)

    # Atacado
    mult_atc     = P.get('mult_atc', 1.3)
    desc_atc     = P.get('desc_atc', 0.15)      # pedido — 15%
    desc_atc_pdv = P.get('desc_atc_pdv', 0.10)  # PDV balcão — 10%
    nf_atc      = round(nf_u * mult_atc, 2)
    p_atc_ped   = round(p_var * (1 - desc_atc), 2)
    p_atc_pdv   = round(p_var * (1 - desc_atc_pdv), 2)
    fed_atc     = round(nf_atc * P['fed'], 2)
    cart_atc    = round(p_atc_ped * P['cartao'], 2)
    icm_atc     = 0.0 if st_u > 0.005 else max(0.0, round(nf_atc * P['icm'], 2) - ant_u)
    c_saida_atc     = round(c_ent + fed_atc + cart_atc + icm_atc, 2)
    margem_atc_ped  = round((p_atc_ped - c_saida_atc) / p_atc_ped, 4) if p_atc_ped > 0 else 0.0
    margem_atc_pdv  = round((p_atc_pdv - c_saida) / p_atc_pdv, 4) if p_atc_pdv > 0 else 0.0

    return dict(c_real=c_real, frete=frete, desp=desp, cred=cred,
                c_ent=c_ent, fed=fed, cartao=cartao, icms_s=icms_s,
                c_saida=c_saida, p_min=p_min, p_var=p_var, margem=margem,
                lucro=lucro,
                nf_atc=nf_atc, p_atc_ped=p_atc_ped, p_atc_pdv=p_atc_pdv, fed_atc=fed_atc,
                cart_atc=cart_atc, icm_atc=icm_atc,
                c_saida_atc=c_saida_atc,
                margem_atc_ped=margem_atc_ped, margem_atc_pdv=margem_atc_pdv)


if getattr(sys, 'frozen', False):
    template_folder = os.path.join(sys._MEIPASS, 'templates')
    static_folder   = os.path.join(sys._MEIPASS, 'static')
    app = Flask(__name__, template_folder=template_folder, static_folder=static_folder)
else:
    app = Flask(__name__)

TEMP_DIR = tempfile.gettempdir()
app.config['MAX_CONTENT_LENGTH'] = 10 * 1024 * 1024  # 10 MB por upload


@app.route('/')
def index():
    return render_template('index.html')


_COL_HEADERS = [
    'NF', 'DESCRIÇÃO', 'REF', 'SKU', 'QTD',
    'NF UNIT', 'ST UNIT', 'ANT UNIT', 'IPI UNIT',
    'FRETE', 'DESPESA', 'CRED ICMS',
    'C. REAL', 'C. ENTRADA', 'CST',
    'FEDERAL', 'CARTÃO', 'ICMS S.', 'C. SAÍDA',
    'META %', 'PREÇO MÍN VRJ',
    'PREÇO ATUAL', 'PREÇO VAREJO', 'MARGEM',
    'NF ATC', 'PREÇO ATC PEDIDO', 'PREÇO ATC PDV',
    'FEDERAL ATC', 'CARTÃO ATC', 'ICMS ATC', 'C. SAÍDA ATC',
    'MARGEM ATC PED', 'MARGEM ATC PDV', 'PREÇO PCT ATC', 'P. COMPRA PCT',
]
_AZUL = {15, 16, 17, 18, 20, 22, 23}
_AMAR = set(range(24, 35))


def _gerar_tabela_html(rows, metricas):
    out  = '<table class="table table-sm table-bordered table-hover"><thead>'
    out += '<tr>'
    out += '<th colspan="15" style="background:#f8f9fa;text-align:center"></th>'
    out += '<th colspan="9" style="background:#dbeafe;text-align:center;font-size:0.7rem;letter-spacing:1px;color:#1d4ed8">VAREJO</th>'
    out += '<th colspan="11" style="background:#fef9c3;text-align:center;font-size:0.7rem;letter-spacing:1px;color:#92400e">ATACADO</th>'
    out += '</tr><tr>'
    out += ''.join(f'<th>{h}</th>' for h in _COL_HEADERS)
    out += '</tr></thead><tbody>'

    for row, m in zip(rows[:20], metricas[:20]):
        has_st  = row['tem_st']
        has_ant = row['tem_ant']
        out += '<tr>'
        cells = [
            row['nf'],
            html.escape(row['desc'][:55]),
            html.escape(str(row['ref'])),
            html.escape(str(row['sku'])),
            int(row['qtd']),
            f"R$ {row['nf_u']:.2f}",
            f"R$ {row['st_u']:.2f}"  if row['st_u']  > 0.001 else '-',
            f"R$ {row['ant_u']:.2f}" if row['ant_u'] > 0.001 else '-',
            f"R$ {row['ipi_u']:.2f}" if row['ipi_u'] > 0.001 else '-',
            f"R$ {m['frete']:.2f}",
            f"R$ {m['desp']:.2f}",
            m['cred'],
            f"R$ {m['c_real']:.2f}",
            f"R$ {m['c_ent']:.2f}",
            row['cst'],
            f"R$ {m['fed']:.2f}",
            f"R$ {m['cartao']:.2f}",
            f"R$ {m['icms_s']:.2f}",
            f"R$ {m['c_saida']:.2f}",
            '15%',
            f"R$ {m['p_min']:.2f}",
            f"R$ {row['p_atual']:.2f}" if row['p_atual'] > 0 else '<span style="color:#e74c3c">SEM PREÇO</span>',
            f"R$ {m['p_var']:.2f}",
            f"{m['margem']*100:.1f}%",
            f"R$ {m['nf_atc']:.2f}",
            f"R$ {m['p_atc_ped']:.2f}",
            f"R$ {m['p_atc_pdv']:.2f}",
            f"R$ {m['fed_atc']:.2f}",
            f"R$ {m['cart_atc']:.2f}",
            f"R$ {m['icm_atc']:.2f}",
            f"R$ {m['c_saida_atc']:.2f}",
            f"{m['margem_atc_ped']*100:.1f}%",
            f"{m['margem_atc_pdv']*100:.1f}%",
            f"R$ {round(m['p_atc_ped'] * row['qtd_emb'], 2):.2f}",
            f"R$ {round(row['nf_u'] * row['qtd_emb'], 2):.2f}",
        ]
        for idx, val in enumerate(cells):
            if idx == 11:
                pct_val = row.get('cred_pct', 0.0) if m['cred'] > 0 else 0.0
                cor_hex = cor_cred(pct_val)
                display = f"R$ {val:.2f}" if isinstance(val, (int, float)) else str(val)
                out += f'<td style="background:#{cor_hex};font-weight:600">{display}</td>'
                continue
            if has_st:
                style = ' style="background:#FCE4D6"'
            elif idx in _AZUL:
                style = ' style="background:#BDD7EE;font-weight:600"'
            elif idx in _AMAR:
                style = ' style="background:#FFF2CC;font-weight:600"'
            elif has_ant:
                style = ' style="background:#D1FAE5"'
            else:
                style = ''
            out += f'<td{style}>{val}</td>'
        out += '</tr>'

    out += '</tbody></table>'
    if len(rows) > 20:
        out += f'<p class="text-muted small">Mostrando 20 de {len(rows)} produtos. Baixe o Excel para ver todos.</p>'
    return out


def _processar_lote(xml_path, csv_path, params, lote_index):
    """Processa um lote independente. Seguro para rodar em thread paralela."""
    fornecedor = params.get('fornecedor', 'FORNECEDOR').strip()
    nota       = params.get('nota', '000').strip()
    label      = f"{fornecedor} — NF {nota}"
    try:
        sucesso, resultado = gerar_tabela(xml_path, csv_path, fornecedor, nota, params)
        if not sucesso:
            return lote_index, {'label': label, 'erro': resultado}

        rows, P, num_nf = resultado
        label = f"{fornecedor} — NF {num_nf}"

        nome_excel    = f"Precificacao_{fornecedor.replace(' ', '_')}_NF_{num_nf}_{lote_index}.xlsx"
        caminho_excel = os.path.join(TEMP_DIR, nome_excel)
        salvar_excel_estilizado(resultado, caminho_excel)

        metricas    = [_calcular(row, P) for row in rows]
        lucro_total = sum(m['lucro'] for m in metricas)

        return lote_index, {
            'label':        label,
            'tabela':       _gerar_tabela_html(rows, metricas),
            'dashboard':    gerar_dashboard_html(rows, lucro_total, metricas, num_nf=num_nf),
            'download_url': f'/download/{nome_excel}',
            'total_itens':  len(rows),
            'erro':         None,
        }
    except Exception:
        import traceback
        return lote_index, {'label': label, 'erro': traceback.format_exc()}


@app.route('/processar', methods=['POST'])
def processar():
    try:
        global_params = request.form.to_dict()
        lote_count    = min(int(global_params.get('lote_count', 1)), 3)

        # Salva arquivos e monta lista de lotes
        lotes = []
        for i in range(lote_count):
            xml_file = request.files.get(f'xml_{i}')
            csv_file = request.files.get(f'csv_{i}')
            if not xml_file or not csv_file:
                return jsonify({'sucesso': False, 'erro': f'Lote {i+1}: arquivos XML e CSV são obrigatórios.'})

            xml_path = os.path.join(TEMP_DIR, f'nfe_temp_{i}.xml')
            csv_path = os.path.join(TEMP_DIR, f'sys_temp_{i}.csv')
            xml_file.save(xml_path)
            csv_file.save(csv_path)

            params = {
                **global_params,
                'fornecedor': global_params.get(f'fornecedor_{i}', 'FORNECEDOR').strip(),
                'nota':       global_params.get(f'nota_{i}', '000').strip(),
            }
            lotes.append((xml_path, csv_path, params))

        # Processa em paralelo (gargalo é a chamada HTTP à API SEFAZ)
        resultados = [None] * lote_count
        with ThreadPoolExecutor(max_workers=lote_count) as ex:
            futures = {
                ex.submit(_processar_lote, xml_path, csv_path, params, i): i
                for i, (xml_path, csv_path, params) in enumerate(lotes)
            }
            for future in as_completed(futures):
                idx, resultado = future.result()
                resultados[idx] = resultado

        algum_sucesso = any(r.get('erro') is None for r in resultados)
        return jsonify({'sucesso': algum_sucesso, 'resultados': resultados})

    except Exception:
        import traceback
        return jsonify({'sucesso': False, 'erro': traceback.format_exc()})


@app.route('/download/<filename>')
def download(filename):
    # Impede path traversal: garante que o arquivo resolvido está dentro de TEMP_DIR
    caminho = os.path.realpath(os.path.join(TEMP_DIR, filename))
    if not caminho.startswith(os.path.realpath(TEMP_DIR) + os.sep):
        return jsonify({'erro': 'Arquivo inválido.'}), 400
    return send_file(caminho, as_attachment=True)


def open_browser():
    webbrowser.open_new("http://127.0.0.1:8080")


if __name__ == '__main__':
    Timer(1, open_browser).start()
    app.run(port=8080, debug=False)