# Migração GitHub Actions → VPS

Esta pasta tem tudo o que precisa pra rodar a análise diária do FindMe num
VPS Linux (Ubuntu/Debian) em vez do GitHub Actions.

## Por que migrar

- **Sem atraso de cron** — GitHub Actions atrasa 5-30min em horário de pico;
  systemd timer dispara no segundo certo.
- **Logs persistentes e ricos** — `journalctl -u findme-diario` mostra tudo.
- **Sem limite de minutos** — GHA grátis tem cota mensal pra repos privados.
- **Mais controle** — pode instalar dependências do sistema, ajustar fuso,
  customizar tudo.
- **Alerta de falha por email** — reaproveita o SMTP já configurado.

O que **continua igual**:
- Snapshots/observados/postos continuam sendo commitados e empurrados pro
  GitHub (sincronização preservada). GitHub vira backup e fonte de verdade.

---

## Arquivos desta pasta

| Arquivo | O que é |
|---|---|
| `run_diario.sh` | Wrapper bash que orquestra: lock → pull → análise → notificação se falha → commit + push. |
| `notify_failure.py` | Helper Python que reaproveita `email_config.json` pra mandar email de alerta com as últimas 80 linhas do log. |
| `findme-diario.service` | systemd unit (`Type=oneshot`) que invoca o wrapper. |
| `findme-diario.timer` | systemd timer (`OnCalendar=07:30 UTC-3`) que dispara o service todo dia. |
| `README.md` | Este guia. |

---

## Passo-a-passo de setup

> **Convenção:** vou usar `findme` como nome do usuário Linux dedicado. Se
> usar outro, substitua nos arquivos `.service`/`.timer` (`User=`, `Group=`,
> `WorkingDirectory=`, `Environment=`).

### 1. SSH no VPS

```bash
ssh seu-usuario-admin@SEU-IP
```

### 2. Criar usuário dedicado (não rode como root)

```bash
sudo adduser --disabled-password --gecos "" findme
sudo usermod -aG sudo findme   # opcional — só se precisar sudo pra debugar
sudo -iu findme
```

### 3. Dependências do sistema

```bash
sudo apt update
sudo apt install -y python3 python3-venv python3-pip git
```

### 4. Clone do repo + venv

```bash
cd ~
git clone https://github.com/RenatoLimaRN/findme-dashboard.git
cd findme-dashboard
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
pip install openpyxl matplotlib   # extras usados pela análise
```

### 5. Criar credenciais (manual, 1x)

```bash
cp config.json.template config.json
chmod 600 config.json
nano config.json                  # preencher o campo "password"

cat > email_config.json <<'EOF'
{
  "smtp_host": "smtp.gmail.com",
  "smtp_port": 587,
  "user": "seu-email@gmail.com",
  "password": "app-password-de-16-chars",
  "from_name": "FindMe Analyst",
  "to": ["destinatario1@email.com", "destinatario2@email.com"]
}
EOF
chmod 600 email_config.json
```

> **Atenção:** esses 2 arquivos ficam no `.gitignore` — não vão pro
> repositório. São específicos do servidor.

### 6. SSH key pra git push autenticado

Sem isso, o `git push` do bot vai falhar (HTTPS exige token interativo).

```bash
ssh-keygen -t ed25519 -C "findme-bot@$(hostname)" -f ~/.ssh/id_ed25519 -N ""
cat ~/.ssh/id_ed25519.pub
```

Copia a saída (`ssh-ed25519 AAAA...`) e adiciona como **Deploy Key**
no repositório:

1. https://github.com/RenatoLimaRN/findme-dashboard/settings/keys/new
2. Title: `vps-findme` (ou nome do servidor)
3. Cola a chave pública
4. **Marca "Allow write access"** ← essencial pro push funcionar
5. Add key

Troca o remote do clone pra usar SSH:

```bash
cd ~/findme-dashboard
git remote set-url origin git@github.com:RenatoLimaRN/findme-dashboard.git

# aceitar host key do GitHub uma vez:
ssh -T git@github.com   # responde "yes" no prompt
```

### 7. Identidade do bot pros commits

```bash
git config --local user.name  "findme-vps-bot"
git config --local user.email "findme-bot@$(hostname)"
```

### 8. Testar manualmente (recomendado antes de habilitar o timer)

```bash
cd ~/findme-dashboard
bash vps/run_diario.sh
```

Deve gerar `logs/run_diario_YYYY-MM-DD.log`, mandar o email diário (se
houver dados) e fazer push de snapshots novos. Se algo der errado, você
recebe o email de alerta.

### 9. Instalar systemd timer (como root)

```bash
sudo cp ~/findme-dashboard/vps/findme-diario.service /etc/systemd/system/
sudo cp ~/findme-dashboard/vps/findme-diario.timer   /etc/systemd/system/

sudo systemctl daemon-reload
sudo systemctl enable --now findme-diario.timer
```

### 10. Conferir que o timer está ativo

```bash
systemctl list-timers findme-diario.timer
# Deve mostrar NEXT = hoje ou amanhã às 10:30 UTC
```

### 11. (Depois de validar 1-2 dias) — desativar o GitHub Actions

Edite `.github/workflows/diario.yml` e comente o bloco `schedule:`:

```yaml
on:
  # schedule:
  #   - cron: '30 10 * * *'
  workflow_dispatch: {}   # mantém execução manual disponível
```

Commit + push e pronto — o cron some, mas o botão "Run workflow" continua
funcionando como fallback.

---

## Operação no dia-a-dia

### Ver se rodou hoje

```bash
sudo journalctl -u findme-diario.service --since today
```

### Ver log completo da execução

```bash
ls ~/findme-dashboard/logs/
cat ~/findme-dashboard/logs/run_diario_$(date -u +%Y-%m-%d).log
```

### Rodar agora (sem esperar o cron)

```bash
sudo systemctl start findme-diario.service
# ou direto, pra ver a saída ao vivo:
cd ~/findme-dashboard && bash vps/run_diario.sh
```

### Ver próxima execução agendada

```bash
systemctl list-timers findme-diario.timer
```

### Parar/desparar (manutenção)

```bash
sudo systemctl stop findme-diario.timer       # pausa
sudo systemctl disable findme-diario.timer    # remove do boot
sudo systemctl enable --now findme-diario.timer  # liga de novo
```

### Atualizar o código no VPS

Não precisa! O `git pull --rebase` no início de cada execução já puxa
qualquer commit novo da `main`. Só atualize o `requirements.txt` à mão se
adicionar dependência nova (raro):

```bash
cd ~/findme-dashboard && source .venv/bin/activate && pip install -r requirements.txt
```

---

## Códigos de saída do `run_diario.sh`

| Código | Significado | Email de alerta enviado? |
|---|---|---|
| 0 | Tudo OK (pode ou não ter mudanças commitadas) | Não |
| 1 | `analise_diaria.py` falhou | **Sim** |
| 2 | `git pull` falhou (conflito de rebase) | **Sim** |
| 3 | Análise OK mas `git push` falhou | **Sim** |

---

## Troubleshooting

### `git push` falha com "Permission denied (publickey)"

A deploy key não tem **write access** ou não foi adicionada. Refaça o passo 6.

### `git pull` dá conflito de rebase

Aconteceu algo raro: alguém commitou direto no GitHub coisa que conflita
com mudanças locais não-sincronizadas (não deveria, já que o VPS só commita
após pull). Resolva manualmente:

```bash
cd ~/findme-dashboard
git status
# resolve conflito → git add → git rebase --continue
```

### Email de alerta não chega

```bash
cd ~/findme-dashboard
python3 vps/notify_failure.py --assunto "TESTE" --log logs/run_diario_$(date -u +%Y-%m-%d).log
```

Se der erro SMTP, conferir `email_config.json` (Gmail exige app password,
não a senha normal).

### O timer não está disparando

```bash
sudo systemctl status findme-diario.timer
sudo systemctl status findme-diario.service
sudo journalctl -u findme-diario.service -n 100
```

### Esqueci o que mudei nos arquivos pro meu VPS

Os arquivos `.service` e `.timer` **devem** ser ajustados pro seu setup
(usuário, caminhos). Mantenha um diff num arquivo separado se quiser:

```bash
diff /etc/systemd/system/findme-diario.service ~/findme-dashboard/vps/findme-diario.service
```

---

## Convivência VPS + GitHub Actions (período de transição)

Durante 1-2 dias, deixe os dois rodando. Vai ter **duplicação de email**
(ambos vão tentar enviar). Pra evitar isso temporariamente, mude o cron
do GHA pra rodar 30min depois do VPS — aí o `git pull` do bot do GHA
puxa o snapshot que o VPS já fez, vê que não tem nada novo pra fazer, e
sai sem mandar email.

Depois que validar o VPS, comente o cron do GHA (passo 11).
