#!/usr/bin/env bash
# Automação diária do findme_dashboard na VM (chamada pelo cron às 07:00 BRT).
# Atualiza SÓ o código do GitHub (preserva historico/observados/postos locais),
# sincroniza o cadastro de postos a partir do FindMe v2, roda a análise do dia
# anterior e envia o e-mail com PDF + Excel.
set -uo pipefail

R=/home/postos_findme/findme-dashboard
cd "$R" || { echo "repo nao encontrado: $R"; exit 1; }

mkdir -p logs
LOG="$R/logs/cron_$(date +%Y-%m-%d).log"
exec >> "$LOG" 2>&1

echo "===================== $(date) ====================="
export PYTHONUTF8=1 PYTHONIOENCODING=utf-8 GIT_TERMINAL_PROMPT=0
git config --global --add safe.directory "$R" 2>/dev/null

# Atualiza o código a partir do GitHub, sem tocar em historico/ observados/
# postos/ (que a VM mantém localmente). O pathspec '*.py' pega TODOS os scripts
# (raiz e subpastas), inclusive arquivos novos — não precisa manter lista fixa.
echo "--- atualizando codigo ---"
git fetch --quiet origin main 2>&1 | tail -2
if git checkout origin/main -- '*.py' requirements.txt vps 2>/dev/null; then
  echo "codigo atualizado para: $(git rev-parse --short origin/main)"
else
  echo "checkout falhou — segue com o codigo atual da VM"
fi

# Garante dependências (rápido se já instaladas).
.venv/bin/pip install -q -r requirements.txt 2>&1 | tail -1 || true

# Sincroniza o cadastro de postos (avulsas) com o FindMe v2 ANTES da análise.
# Assim, atividades novas cadastradas no sistema entram no posto certo no mesmo
# dia. Se falhar (rede/login), segue com o cadastro atual — não aborta o relatório.
echo "--- sincronizando cadastro de postos (FindMe v2) ---"
if timeout 300 .venv/bin/python sincronizar_postos.py; then
  echo "sincronizacao de postos ok"
else
  echo "sincronizador falhou (rc=$?) — segue com o cadastro atual"
fi

echo "--- rodando analise diaria ---"
.venv/bin/python analise_diaria.py
echo "fim: codigo de saida $?"
