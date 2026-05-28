# Guia de instalação no VPS — passo a passo

Este guia leva você do **zero** (acabou de logar no VPS) até a **análise
diária rodando sozinha todo dia**. Não assume conhecimento prévio de
Linux. Cada comando tem: o que faz, por que precisa, cola-e-executa, e o
que esperar de resposta.

> ⏱ **Tempo estimado:** 30-40 minutos na primeira vez.
> 💻 **Onde rodar:** dentro do terminal SSH do seu VPS.

---

## Antes de começar — coisas que você precisa ter em mãos

Pega papel e caneta (ou um bloco de notas) e anota:

| Item | Onde achar | Exemplo |
|---|---|---|
| IP ou nome do VPS | Painel do provedor (Hetzner/DO/etc) | `203.0.113.42` ou `meu-vps.exemplo.com` |
| Usuário SSH do VPS | Definido quando você criou o VPS | `root` ou `ubuntu` ou seu nome |
| Senha do FindMe | Aquela que você usa pra entrar no portal | (mantém em segredo) |
| App password do Gmail | Conta Google → Segurança → Senhas de app | 16 letras, ex.: `abcd efgh ijkl mnop` |
| Emails que recebem o relatório | Lista que você quer notificar | `voce@gmail.com, gestor@empresa.com` |

> ⚠️ **App password do Gmail ≠ senha normal do Gmail.** Se não tem, gera
> uma em: https://myaccount.google.com/apppasswords (precisa ter
> verificação em 2 etapas ativada na sua conta Google).

---

## Etapa 1 — Conectar no VPS

No seu computador (Windows: PowerShell, Mac/Linux: Terminal), digite:

```bash
ssh SEU_USUARIO@SEU_IP
```

Exemplo concreto:
```bash
ssh ubuntu@203.0.113.42
```

Vai pedir a senha (ou usar a chave SSH que você já configurou no provedor).
Depois de entrar, você vai ver algo tipo:

```
Welcome to Ubuntu 22.04 LTS
ubuntu@meu-vps:~$
```

**Esse `$` é o prompt — significa que está pronto pra receber comandos.**

> ✅ **Se entrou aqui, pode seguir.**
> ❌ **Se deu "Permission denied":** sua senha ou chave SSH não bate.
> Verifica com o suporte do provedor.

---

## Etapa 2 — Criar um usuário dedicado pro robô

**Por que:** rodar como `root` é perigoso. Se algo der errado (ou um
hacker descobrir uma senha), ele teria acesso ao servidor inteiro. Vamos
criar um usuário chamado `findme` que só pode mexer nas próprias coisas.

Cola e executa:

```bash
sudo adduser --disabled-password --gecos "" findme
```

**O que isso faz:**
- `sudo` = roda como administrador (vai pedir sua senha)
- `adduser findme` = cria o usuário chamado `findme`
- `--disabled-password` = sem senha (só vai ser usado pelo robô, não por humano)
- `--gecos ""` = não pergunta nome completo, telefone, etc.

**Resposta esperada:**
```
Adding user `findme' ...
Adding new group `findme' (1001) ...
Adding new user `findme' (1001) with group `findme' ...
Creating home directory `/home/findme' ...
```

Agora **vira** esse usuário (todos os comandos das próximas etapas rodam
como `findme`):

```bash
sudo -iu findme
```

O prompt muda pra:
```
findme@meu-vps:~$
```

> 💡 **Se em qualquer momento você precisar voltar a ser admin:** digite
> `exit`. Pra virar `findme` de novo: `sudo -iu findme`.

---

## Etapa 3 — Instalar os programas necessários

**Por que:** o VPS vem pelado. Precisamos instalar Python (linguagem que
roda a análise) e git (pra baixar o código do GitHub).

```bash
sudo apt update
sudo apt install -y python3 python3-venv python3-pip git
```

**Resposta esperada:** uma lista de pacotes sendo instalados, terminando com
`Setting up ... (1:X.Y.Z)`. Demora 1-2 minutos.

**Pra confirmar:**
```bash
python3 --version
git --version
```

Deve mostrar algo tipo `Python 3.10.12` e `git version 2.34.1`.

---

## Etapa 4 — Baixar o código do projeto

```bash
cd ~
git clone https://github.com/RenatoLimaRN/findme-dashboard.git
cd findme-dashboard
```

**O que isso faz:**
- `cd ~` = vai pra pasta home do usuário findme (`/home/findme`)
- `git clone https://...` = baixa o repositório do GitHub
- `cd findme-dashboard` = entra na pasta baixada

**Pra confirmar que deu certo:**
```bash
ls
```

Deve listar: `analise_diaria.py`, `findme_programacao.py`, `postos/`,
`vps/`, `requirements.txt`, etc.

---

## Etapa 5 — Criar ambiente Python isolado (venv)

**Por que:** Python tem milhares de bibliotecas. Se você instalar tudo no
sistema inteiro, uma versão briga com outra. Um "venv" é uma caixinha
isolada só pra esse projeto.

```bash
python3 -m venv .venv
source .venv/bin/activate
```

**O prompt agora começa com `(.venv)`:**
```
(.venv) findme@meu-vps:~/findme-dashboard$
```

Isso significa que você está dentro da caixinha. **Toda vez que abrir SSH
novo e quiser rodar coisas Python deste projeto**, precisa rodar o
`source .venv/bin/activate` de novo.

---

## Etapa 6 — Instalar as bibliotecas que o projeto usa

```bash
pip install -r requirements.txt
pip install openpyxl matplotlib
```

**O que faz:** baixa as bibliotecas (`requests`, `openpyxl`, etc) que o
código precisa pra funcionar.

**Resposta esperada:** lista de pacotes sendo instalados, terminando com
`Successfully installed ...`.

---

## Etapa 7 — Configurar as credenciais do FindMe

Esse é o passo mais delicado — você vai escrever sua senha num arquivo.
Vamos fazer com cuidado.

### 7.1 — Copia o template

```bash
cp config.json.template config.json
```

### 7.2 — Edita pra colocar a senha

Abre o arquivo no editor de texto:

```bash
nano config.json
```

Vai abrir uma tela com algo assim:
```json
{
  "email": "rnl.lima.nascimento@gmail.com",
  "password": "",
  "locations": [
    "e40c026-6bac-4422-85e3-3517984ba207",
    ...
  ]
}
```

**Acha a linha `"password": "",`** e coloca sua senha do FindMe entre as
aspas. Vai ficar tipo:
```json
  "password": "minhaSenhaDoFindMe123",
```

**Pra navegar no nano:**
- Setas do teclado pra mover o cursor
- Apaga e digita normal
- **Pra salvar:** `Ctrl + O`, depois `Enter` pra confirmar o nome
- **Pra sair:** `Ctrl + X`

### 7.3 — Protege o arquivo (impede outros usuários de lerem)

```bash
chmod 600 config.json
```

**O que faz:** só o usuário `findme` consegue ler/escrever esse arquivo.
Importante porque tem senha dentro.

---

## Etapa 8 — Configurar as credenciais do email

```bash
nano email_config.json
```

(O nano vai abrir uma tela vazia porque o arquivo não existe ainda.)

**Cola o seguinte**, substituindo pelos seus dados:

```json
{
  "smtp_host": "smtp.gmail.com",
  "smtp_port": 587,
  "user": "seu-email-remetente@gmail.com",
  "password": "abcdefghijklmnop",
  "from_name": "FindMe Analyst",
  "to": ["destinatario1@email.com", "destinatario2@email.com"]
}
```

**Atenção em 3 coisas:**
- `user` = o email **do qual** vai sair o relatório (precisa ser Gmail)
- `password` = a **app password** de 16 caracteres (sem espaços), não a senha normal
- `to` = lista de emails que vão **receber**. Pode ser 1 ou vários. Mantenha os colchetes `[ ]`.

Salva com `Ctrl + O` → `Enter` → `Ctrl + X`.

Protege:
```bash
chmod 600 email_config.json
```

---

## Etapa 9 — Criar SSH key dedicada pro robô empurrar código pro GitHub

**Por que:** o robô vai precisar fazer `git push` (subir snapshots). Pra
isso, o GitHub precisa "confiar" no robô. Vamos criar uma chave criptográfica
pra ele.

### 9.1 — Gera a chave

```bash
ssh-keygen -t ed25519 -C "findme-bot@$(hostname)" -f ~/.ssh/id_ed25519 -N ""
```

**Resposta esperada:**
```
Generating public/private ed25519 key pair.
Your identification has been saved in /home/findme/.ssh/id_ed25519
Your public key has been saved in /home/findme/.ssh/id_ed25519.pub
...
```

### 9.2 — Copia a chave pública (a parte que vai pro GitHub)

```bash
cat ~/.ssh/id_ed25519.pub
```

Vai imprimir algo tipo:
```
ssh-ed25519 AAAAC3NzaC1lZDI1NTE5AAAAILpL... findme-bot@meu-vps
```

**Seleciona toda essa linha e copia** (Ctrl+C no Windows, Cmd+C no Mac).

### 9.3 — Adiciona como Deploy Key no GitHub

Abre o navegador no seu computador e acessa:

**https://github.com/RenatoLimaRN/findme-dashboard/settings/keys/new**

- **Title:** `vps-findme` (ou qualquer nome que ajude você a lembrar)
- **Key:** cola a linha que você copiou
- **⚠️ MARCA A CAIXA "Allow write access"** ← sem isso, push falha!
- Clica **Add key**

### 9.4 — Troca a forma do projeto falar com o GitHub (de HTTPS pra SSH)

De volta no terminal do VPS:

```bash
cd ~/findme-dashboard
git remote set-url origin git@github.com:RenatoLimaRN/findme-dashboard.git
```

### 9.5 — Aceita o GitHub como host conhecido (1x só)

```bash
ssh -T git@github.com
```

Vai aparecer:
```
The authenticity of host 'github.com (...)' can't be established.
ED25519 key fingerprint is SHA256:+DiY3wvvV6TuJJhbpZisF/zLDA0zPMSvHdkr4UvCOqU.
Are you sure you want to continue connecting (yes/no/[fingerprint])?
```

**Digita** `yes` **e dá Enter.**

Resposta esperada:
```
Hi RenatoLimaRN/findme-dashboard! You've successfully authenticated, but GitHub does not provide shell access.
```

Se viu `Hi ...! You've successfully authenticated` → **deu certo!** 🎉

> ❌ Se deu `Permission denied (publickey)` → você esqueceu de marcar
> "Allow write access" na etapa 9.3, ou colou a chave errada.
> **Como corrigir:** volta no link e deleta a key, faz o passo 9.3 de novo
> com cuidado.

---

## Etapa 10 — Identificar o robô nos commits

Quando o robô subir snapshots, o GitHub precisa saber "quem fez isso":

```bash
cd ~/findme-dashboard
git config --local user.name  "findme-vps-bot"
git config --local user.email "findme-bot@$(hostname)"
```

---

## Etapa 11 — TESTE MANUAL (esse é o momento da verdade)

Rode o wrapper **uma vez à mão** pra ver se tudo funciona:

```bash
cd ~/findme-dashboard
bash vps/run_diario.sh
```

**Vai aparecer várias linhas com `[2026-05-XX HH:MM:SSZ]`** — é o log
ao vivo. Em 1-3 minutos termina.

### Se tudo deu certo:

- Você **recebeu o email diário** normal (com o .xlsx anexado)
- Última linha do terminal: `[...] Fim     | tudo certo.`
- `git log` no GitHub mostra commit novo: `diario: snapshot + aprendizado YYYY-MM-DD (vps)`

### Se algo deu errado:

- Você recebe um **email de alerta** com "[FindMe] FALHA na ..." no assunto
- O corpo do email tem as últimas 80 linhas do log explicando o que quebrou

**Erros mais comuns:**

| Erro no log | Causa | Como resolver |
|---|---|---|
| `Permission denied (publickey)` | Deploy key sem "write access" | Etapa 9.3 de novo, marca a caixa |
| `KeyError: 'password'` | Esqueceu de preencher `config.json` | Etapa 7.2 de novo |
| `ModuleNotFoundError: No module named '...'` | Faltou instalar biblioteca | `source .venv/bin/activate && pip install <nome>` |
| `Authentication failed` (SMTP) | App password do Gmail errada | Etapa 8: confere as 16 letras |
| `git pull` falhou com conflito | Estado raro | Manda o log aqui que te ajudo |

---

## Etapa 12 — Agendar pra rodar todo dia automaticamente

Agora que você sabe que funciona à mão, vamos fazer rodar sozinho às 07:30
da manhã todo dia.

### 12.1 — Copia os arquivos do systemd pro lugar certo

```bash
sudo cp ~/findme-dashboard/vps/findme-diario.service /etc/systemd/system/
sudo cp ~/findme-dashboard/vps/findme-diario.timer   /etc/systemd/system/
```

### 12.2 — Atualiza o systemd pra ler os arquivos novos

```bash
sudo systemctl daemon-reload
```

### 12.3 — Liga o timer (agendamento) e habilita pra ligar no boot

```bash
sudo systemctl enable --now findme-diario.timer
```

**Resposta esperada:**
```
Created symlink /etc/systemd/system/timers.target.wants/findme-diario.timer → /etc/systemd/system/findme-diario.timer.
```

### 12.4 — Confere que está agendado

```bash
systemctl list-timers findme-diario.timer
```

Vai mostrar uma tabela:
```
NEXT                        LEFT       LAST  PASSED  UNIT                  ACTIVATES
Tue 2026-05-26 10:30:00 UTC 12h left   -     -       findme-diario.timer   findme-diario.service
```

A coluna **NEXT** mostra a próxima execução agendada. **Se aparecer essa
linha, está tudo certo.** Pode sair do SSH e ir dormir tranquilo.

---

## Etapa 13 — No dia seguinte: confere que rodou

Loga de novo no VPS e roda:

```bash
sudo systemctl status findme-diario.service
```

Procure por:
- `Active: inactive (dead) since ...` ← é normal! O service é "oneshot",
  ele roda, termina, e fica como "inactive" até a próxima.
- `Main PID: ... (code=exited, status=0/SUCCESS)` ← `status=0/SUCCESS` é OK.

**Pra ver o log completo da última execução:**

```bash
sudo journalctl -u findme-diario.service --since today
```

Ou diretamente o arquivo de log:

```bash
cat ~/findme-dashboard/logs/run_diario_$(date -u +%Y-%m-%d).log
```

---

## Comandos do dia-a-dia (consulta rápida)

### Ver se rodou hoje
```bash
sudo systemctl status findme-diario.service
```

### Ver log da última execução
```bash
cat ~/findme-dashboard/logs/run_diario_$(date -u +%Y-%m-%d).log
```

### Ver próxima execução agendada
```bash
systemctl list-timers findme-diario.timer
```

### Rodar agora (sem esperar o cron)
```bash
sudo systemctl start findme-diario.service
```
Ou, pra ver a saída ao vivo:
```bash
cd ~/findme-dashboard && bash vps/run_diario.sh
```

### Pausar temporariamente
```bash
sudo systemctl stop findme-diario.timer
```

### Voltar a rodar
```bash
sudo systemctl start findme-diario.timer
```

### Atualizar o código (quando você fizer mudanças no GitHub)
**Não precisa fazer nada!** O `run_diario.sh` já faz `git pull --rebase`
antes de cada execução. Mas se quiser puxar agora:

```bash
sudo -iu findme
cd ~/findme-dashboard
git pull --rebase
```

### Ver logs antigos
```bash
ls ~/findme-dashboard/logs/
cat ~/findme-dashboard/logs/run_diario_2026-05-26.log
```

---

## Códigos de saída do `run_diario.sh`

Se você abrir um log e vir "exit code X" no final, isso significa:

| Código | Significado | Email de alerta? |
|---|---|---|
| 0 | Tudo OK | Não |
| 1 | `analise_diaria.py` falhou | ✅ Sim |
| 2 | `git pull` falhou (conflito) | ✅ Sim |
| 3 | Análise OK mas `git push` falhou | ✅ Sim |

---

## Troubleshooting (problemas comuns)

### Não chegou email de alerta de falha

Testa manualmente:
```bash
cd ~/findme-dashboard
source .venv/bin/activate
python3 vps/notify_failure.py --assunto "TESTE" --log /dev/null
```
Se deu erro SMTP, confere `email_config.json` (app password, não senha normal).

### O timer não está disparando

```bash
sudo systemctl status findme-diario.timer
sudo systemctl status findme-diario.service
sudo journalctl -u findme-diario.service -n 100
```

### Quero ajustar o horário

Edita o timer:
```bash
sudo nano /etc/systemd/system/findme-diario.timer
```

Acha a linha `OnCalendar=*-*-* 10:30:00 UTC` e muda. Lembrar:
**BRT = UTC - 3 horas**, então:
- `07:00 BRT` = `10:00 UTC` → `OnCalendar=*-*-* 10:00:00 UTC`
- `08:30 BRT` = `11:30 UTC` → `OnCalendar=*-*-* 11:30:00 UTC`

Depois recarrega:
```bash
sudo systemctl daemon-reload
sudo systemctl restart findme-diario.timer
```

### Senha do FindMe mudou

```bash
sudo -iu findme
nano ~/findme-dashboard/config.json
# muda o valor de "password"
# Ctrl+O, Enter, Ctrl+X
```
Pronto. A próxima execução já usa a senha nova.

### Email mudou ou quero adicionar destinatário

```bash
sudo -iu findme
nano ~/findme-dashboard/email_config.json
# edita o array "to"
```

### Acabou o espaço em disco

Limpa logs antigos:
```bash
find ~/findme-dashboard/logs -name "run_diario_*.log" -mtime +30 -delete
```
(apaga logs com mais de 30 dias)

---

## Onde os arquivos importantes ficam no VPS

```
/home/findme/
├── .ssh/
│   ├── id_ed25519        ← chave privada do bot (NUNCA compartilhe)
│   └── id_ed25519.pub    ← chave pública (foi pro GitHub Deploy Keys)
└── findme-dashboard/     ← clone do repositório
    ├── config.json          ← senha do FindMe (chmod 600)
    ├── email_config.json    ← credenciais SMTP (chmod 600)
    ├── .venv/               ← ambiente Python isolado
    ├── logs/                ← logs diários
    ├── relatorios/          ← .xlsx gerados
    └── vps/
        ├── run_diario.sh      ← o wrapper
        ├── notify_failure.py  ← envio de alertas
        ├── findme-diario.service
        └── findme-diario.timer

/etc/systemd/system/
├── findme-diario.service     ← cópia do arquivo acima
└── findme-diario.timer       ← cópia do arquivo acima
```
