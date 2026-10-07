# Deploy — PH-Sheets

## Visão geral

```
PR para main
     │
     ▼
GitHub Actions (.github/workflows/ci.yml) ── roda pytest no PR
     │
     ▼ merge
Push em main
     ├─► GitHub Actions roda pytest de novo
     └─► Render (auto-deploy) ── pip install -r requirements.txt
                                 └─ gunicorn app:app  (lê gunicorn.conf.py)
```

**Atenção:** o Render faz deploy de todo push em `main` **independente do CI**. O CI só bloqueia o merge se a branch protection estiver ativa (ver abaixo) — sem ela, um código com teste quebrado chega à produção.

---

## Produção — Render

| Item | Valor |
|------|-------|
| Serviço | `PH-sheets` (`srv-d81iaof7f7vs73dl5kc0`) — web service, plano **free**, região Ohio |
| URL | https://ph-sheets.onrender.com |
| Origem | `PH-importados/PH-sheets`, branch `main`, auto-deploy ligado |
| Build | `pip install -r requirements.txt` |
| Start | `gunicorn app:app` — porta, workers e timeout vêm do `gunicorn.conf.py` |

### Limitações do plano free

- **Hiberna após ~15 min sem acesso.** A primeira requisição depois disso leva ~40–50 s (cold start). É normal.
- **512 MB de RAM** — por isso 2 workers no `gunicorn.conf.py`.
- **Arquivos temporários:** os Excel gerados ficam no disco efêmero do container. Funcionam para download imediato; somem quando o serviço hiberna ou faz novo deploy — esperado no fluxo gera → baixa na hora.

### Render CLI

```bash
brew install render
render login
render workspace set               # escolhe "Felipe's workspace"

render services                    # lista serviços
render deploys list srv-d81iaof7f7vs73dl5kc0          # histórico (commit + status)
render logs -r srv-d81iaof7f7vs73dl5kc0 --limit 100   # logs da aplicação
render deploys create srv-d81iaof7f7vs73dl5kc0        # força um deploy do main atual
```

---

## Arquivos de deploy

| Arquivo | Função |
|---------|--------|
| `.github/workflows/ci.yml` | Roda `pytest` (Python 3.12) em todo push/PR para `main` |
| `gunicorn.conf.py` | Porta (`$PORT`), 2 workers e timeout de 120 s — lido automaticamente pelo gunicorn |
| `requirements.txt` | Usado pelo Render no build |
| `requirements-server.txt` | Usado pelo CI — versão enxuta, sem PyInstaller/Pillow/macholib |

Por que timeout de 120 s: o padrão do gunicorn é 30 s, e uma nota grande com consulta à API SEFAZ (timeout de 10 s por chamada) pode passar disso — o worker seria morto no meio do processamento.

---

## Branch protection (configurar uma vez no GitHub)

**GitHub → Settings → Branches → Add rule → `main`**

- [x] Require a pull request before merging
- [x] Require status checks to pass before merging → status check: `test`
- [x] Require branches to be up to date before merging

---

## Variáveis de ambiente

Nenhuma obrigatória. A aplicação não usa secrets, banco de dados nem API keys fixas — a API SEFAZ AL é chamada com a chave da NFe enviada pelo usuário em cada requisição. `PORT` é injetada pelo próprio Render.

Se precisar adicionar: **Render → PH-sheets → Environment**.

---

## Fluxo de trabalho

```bash
git checkout -b feat/minha-mudanca
# ... edita código ...
python -m pytest tests/ -v          # roda local antes de abrir PR
git push -u origin feat/minha-mudanca
# Abre PR → CI roda → aprova/merge → Render deploya sozinho (~2–3 min)
render deploys list srv-d81iaof7f7vs73dl5kc0   # confere se ficou "live"
```

---

## Histórico

- **Mar/2026** — distribuição por executável PyInstaller (`.app`/`.exe`), gerado e entregue manualmente. Ver README.
- **Abr/2026** — tentativa no Railway (plano trial). Funcionou em 10/04 e foi desativado ao fim do trial; não recebe mais deploys.
- **Mai/2026 →** — Render (plano free), produção atual.

---

## Repositório

`git@github.com:PH-importados/PH-sheets.git`
