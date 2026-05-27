#!/usr/bin/env bash
# run_diario.sh — wrapper de execução diária no VPS.
#
# Faz, em ordem:
#   1. lock (evita rodar 2x em paralelo)
#   2. git pull --rebase --autostash (puxa código novo do GitHub)
#   3. roda analise_diaria.py
#   4. se falhar → manda email de alerta via notify_failure.py
#   5. git add + commit + push (sobe snapshots/observados/postos atualizados)
#
# Tudo é logado em logs/run_diario_<data>.log no próprio repo.
#
# Variáveis de ambiente esperadas:
#   FINDME_REPO  — caminho absoluto do clone do repo no VPS
#                  (default: /home/<user>/findme-dashboard)
#   PYTHON_BIN   — binário Python (default: python3)
#
# Códigos de saída:
#   0 = tudo OK (pode ou não ter mudanças pra commitar)
#   1 = falha na análise (email de alerta foi disparado)
#   2 = falha no git pull (conflito de rebase — não tocou no resto)
#   3 = falha no git push (análise OK, snapshots locais, mas não subiram)
#
# Uso típico via systemd (ver findme-diario.service).

set -uo pipefail

REPO="${FINDME_REPO:-$HOME/findme-dashboard}"
PYTHON="${PYTHON_BIN:-python3}"

cd "$REPO" || { echo "ERRO: repo nao encontrado em $REPO" >&2; exit 1; }

# ─── Logging ──────────────────────────────────────────────────────────────────
mkdir -p logs
LOG_FILE="logs/run_diario_$(date -u +%Y-%m-%d).log"
exec > >(tee -a "$LOG_FILE") 2>&1

log() { echo "[$(date -u +'%Y-%m-%d %H:%M:%S')Z] $*"; }

# ─── Lock (impede execução paralela) ──────────────────────────────────────────
LOCK_FILE="/tmp/findme-diario.lock"
exec 200>"$LOCK_FILE"
if ! flock -n 200; then
    log "AVISO: outra execução em andamento (lock $LOCK_FILE). Saindo."
    exit 0
fi

log "════════════════════════════════════════════════════════════════"
log "Inicio  | repo=$REPO  python=$PYTHON  user=$(whoami)  host=$(hostname)"

# ─── 1) Pull do GitHub ────────────────────────────────────────────────────────
log "Git pull --rebase --autostash..."
if ! git pull --rebase --autostash origin main; then
    log "ERRO: git pull falhou (provavel conflito). Abortando rebase."
    git rebase --abort 2>/dev/null || true
    "$PYTHON" vps/notify_failure.py \
        --assunto "FALHA git pull — $(date -u +%Y-%m-%d)" \
        --log "$LOG_FILE" || true
    exit 2
fi
log "Pull OK. Commit atual: $(git rev-parse --short HEAD)"

# ─── 2) Análise diária ───────────────────────────────────────────────────────
log "Rodando analise_diaria.py..."
ANALISE_OK=0
PYTHONUTF8=1 PYTHONIOENCODING=utf-8 "$PYTHON" analise_diaria.py || ANALISE_OK=$?

if [ "$ANALISE_OK" -ne 0 ]; then
    log "ERRO: analise_diaria.py saiu com codigo $ANALISE_OK"
    "$PYTHON" vps/notify_failure.py \
        --assunto "FALHA análise diária — $(date -u +%Y-%m-%d)" \
        --log "$LOG_FILE" || true
    exit 1
fi
log "Analise OK."

# ─── 3) Commit + push de snapshots ───────────────────────────────────────────
log "Stage de mudancas..."
git add .claude/skills/findme-analyst/historico  2>/dev/null || true
git add .claude/skills/findme-analyst/observados 2>/dev/null || true
git add postos                                   2>/dev/null || true

if git diff --cached --quiet; then
    log "Nada novo pra commitar."
else
    MSG="diario: snapshot + aprendizado $(date -u +%Y-%m-%d) (vps)"
    log "Commit: $MSG"
    git commit -m "$MSG"
    log "Push..."
    if ! git push origin main; then
        log "ERRO: git push falhou. Snapshots ficam locais ate a proxima rodada."
        "$PYTHON" vps/notify_failure.py \
            --assunto "FALHA git push — $(date -u +%Y-%m-%d)" \
            --log "$LOG_FILE" || true
        exit 3
    fi
    log "Push OK: $(git rev-parse --short HEAD)"
fi

log "Fim     | tudo certo."
exit 0
