# Relatório de NCs — BCMS (OGU + FEx)

Relatório automatizado de **Notas de Crédito** do Batalhão de Comunicações e Material Sigiloso (BCMS), contemplando as duas UASGs da unidade:

| Fonte | UASG | Origem dos recursos |
|-------|------|---------------------|
| 💼 **OGU** | `160329` | Orçamento Geral da União — dotação LOA do Comando do Exército (UO 52121, Órgão 52000) |
| 🏦 **FEx** | `167329` | Fundo do Exército — instituído pela **Lei nº 4.617/1965** e regulamentado pelo **Decreto nº 91.575/1985**, administrado pela SEF/DiFEx |

O relatório é gerado em dias úteis às **09h30 BRT** e entregue por e-mail em formato WhatsApp (texto + emojis), contemplando créditos recebidos e recolhidos no dia anterior e na semana corrente, com abertura por fonte e saldo líquido consolidado.

---

## Sumário

1. [Estrutura do repositório](#1-estrutura-do-repositório)
2. [Pré-requisitos](#2-pré-requisitos)
3. [Passo a passo: criar o repositório no GitHub](#3-passo-a-passo-criar-o-repositório-no-github)
4. [Cadastro dos secrets](#4-cadastro-dos-secrets)
5. [Configuração do workflow](#5-configuração-do-workflow)
6. [Execução e validação](#6-execução-e-validação)
7. [Estrutura do relatório gerado](#7-estrutura-do-relatório-gerado)
8. [Lógica técnica do script](#8-lógica-técnica-do-script)
9. [Personalização](#9-personalização)
10. [Troubleshooting](#10-troubleshooting)
11. [Fundamentação legal](#11-fundamentação-legal)

---

## 1. Estrutura do repositório

Após criado, o repositório terá a seguinte árvore:

```
relatorio-nc-bcms/
├── .github/
│   └── workflows/
│       └── relatorio_nc_bcms.yml      ← agendamento e execução
├── .gitignore
├── README.md                          ← este arquivo
├── relatorio_nc_bcms.py               ← script principal
└── requirements.txt                   ← dependências Python
```

Somente 5 arquivos. Nenhum dado sensível trafega pelo repositório — credenciais de e-mail ficam apenas nos **Secrets** do GitHub Actions.

---

## 2. Pré-requisitos

Antes de começar, você precisa ter:

1. **Conta no GitHub** — qualquer plano, inclusive Free (o limite de 2 000 min/mês de Actions é mais do que suficiente; este workflow gasta ~30 s por execução).
2. **Git instalado localmente** (opcional — pode ser feito 100 % pela interface web do GitHub).
3. **Credenciais da conta Gmail que enviará o e-mail**:
   - Endereço do remetente (ex.: `bcmssgtdecampos@gmail.com`).
   - **Senha de app de 16 caracteres** (NÃO a senha principal da conta). Para gerar: <https://myaccount.google.com/apppasswords>. Requer 2FA ativo.
4. **Lista de destinatários** — e-mails que receberão o relatório.
5. **Acesso à planilha fonte** (já compartilhada na BaApLog): <https://docs.google.com/spreadsheets/d/1Jv546wpWQSFAlep3oLRAg29hVy86iJxJ>. O ID já está hardcoded no script.

> 💡 **Por que o SPREADSHEET_ID está hardcoded e não no `secrets`?**
> Lição aprendida em sprint anterior do `relatorio_nc.py`: o secret `SPREADSHEET_ID` do GitHub ficou mal configurado e causou HTTP 404 silencioso por dias. Como a planilha é publicada por link (sem dado sensível na chave), a opção mais segura é manter o ID como literal no script. O código traz um comentário em português documentando essa decisão para evitar que uma futura refatoração reintroduza a dependência de ambiente.

---

## 3. Passo a passo: criar o repositório no GitHub

### 3.1. Criar o repositório

1. Acesse <https://github.com/new>.
2. Preencha:
   - **Repository name:** `relatorio-nc-bcms`
   - **Description:** *Relatório automatizado de NCs do BCMS (OGU 160329 + FEx 167329).*
   - **Visibility:** **Private** (recomendado — evita exposição de metadados SIAFI).
   - ✅ *Add a README file* — marque.
   - ✅ *Add .gitignore* — selecione template **Python**.
   - *License:* nenhuma (repositório institucional privado).
3. Clique em **Create repository**.

### 3.2. Adicionar os arquivos

Existem duas formas — escolha **uma**:

#### ▶ Opção A — Pela interface web (sem terminal)

Para cada arquivo do pacote:

1. Na página inicial do repositório, clique em **Add file → Create new file** (ou **Upload files** se tiver o arquivo local).
2. Crie exatamente estes caminhos/nomes (respeitar maiúsculas/minúsculas e a barra `/`):

| Caminho no repositório | Conteúdo |
|---|---|
| `relatorio_nc_bcms.py` | script Python entregue |
| `requirements.txt` | `openpyxl>=3.1.0` (uma linha) |
| `.github/workflows/relatorio_nc_bcms.yml` | workflow entregue |
| `README.md` | este documento |
| `.gitignore` | substituir o conteúdo gerado pelo template Python pelo conteúdo do arquivo entregue |

> ⚠️ Para criar a pasta `.github/workflows/`, ao digitar o nome do arquivo você escreve `.github/workflows/relatorio_nc_bcms.yml` — o GitHub cria as pastas automaticamente ao digitar a barra.

3. Ao final de cada arquivo, em **Commit new file**, deixe a mensagem padrão e confirme.

#### ▶ Opção B — Pelo terminal (Git)

```bash
# Clonar (substitua <usuario> pelo seu handle GitHub)
git clone https://github.com/<usuario>/relatorio-nc-bcms.git
cd relatorio-nc-bcms

# Copiar os arquivos entregues para a raiz
mkdir -p .github/workflows
cp ~/Downloads/relatorio_nc_bcms.py .
cp ~/Downloads/requirements.txt .
cp ~/Downloads/.gitignore .
cp ~/Downloads/workflow.yml .github/workflows/relatorio_nc_bcms.yml
# README.md — este próprio arquivo

# Commit
git add .
git commit -m "Estrutura inicial: script, workflow, requirements e docs"
git push origin main
```

---

## 4. Cadastro dos secrets

Os secrets são variáveis seguras que **nunca aparecem no código** nem nos logs do GitHub Actions. É onde ficam as credenciais.

### 4.1. Acesso à tela de secrets

No repositório, menu: **Settings → Secrets and variables → Actions → Repository secrets**.

### 4.2. Secrets obrigatórios

Clique em **New repository secret** e cadastre os três abaixo (um por vez):

| Name | Value (exemplo) | Observação |
|---|---|---|
| `EMAIL_REMETENTE` | `bcmssgtdecampos@gmail.com` | Gmail que enviará — precisa estar autorizado a usar senha de app. |
| `EMAIL_SENHA` | `` | Senha de app **sem espaços**. O GitHub rejeita espaços. |
| `EMAIL_DESTINO` | `fulano@eb.mil.br, ciclano@eb.mil.br` | Aceita múltiplos, separados por vírgula ou ponto-e-vírgula. |

> ⚠️ **Erro clássico:** se colar a senha de app como `` (com espaços), o GitHub mostra:
> *"Secret names can only contain alphanumeric characters..."*
> **Solução:** remova os espaços — o Gmail aceita a senha tanto com quanto sem espaços.

### 4.3. Secret opcional (cópia oculta)

| Name | Value | Uso |
|---|---|---|
| `EMAIL_BCC` | `cmt@eb.mil.br, s3@eb.mil.br` | Destinatários em **cópia oculta** — entregues mas não visíveis no cabeçalho do e-mail. |

Se não precisar de BCC, simplesmente não crie o secret — o script trata como vazio.

### 4.4. ⚠️ NÃO crie um secret `SPREADSHEET_ID`

O ID da planilha **está hardcoded no script** intencionalmente. Não crie este secret — e se ele já existir herdado de outro repositório, é recomendado **deletá-lo** para evitar reconfiguração errada no futuro.

---

## 5. Configuração do workflow

### 5.1. Horário da execução

O workflow está configurado para rodar em dias úteis às **09h30 BRT**. Tecnicamente, o GitHub Actions opera em UTC, e o Brasil é UTC-3 o ano todo (horário de verão foi suspenso pelo Decreto nº 9.772/2019):

```yaml
schedule:
  - cron: '30 12 * * 1-5'   # 12:30 UTC = 09:30 BRT, seg a sex
```

Para alterar o horário, edite apenas a linha do `cron`. Referência rápida:

| BRT | UTC | Cron |
|---|---|---|
| 07:30 | 10:30 | `'30 10 * * 1-5'` |
| 08:00 | 11:00 | `'0 11 * * 1-5'` |
| 09:30 | 12:30 | `'30 12 * * 1-5'` ← **padrão** |
| 11:00 | 14:00 | `'0 14 * * 1-5'` |

### 5.2. Execução manual

O workflow inclui `workflow_dispatch`, ou seja, você pode disparar on-demand pela interface do GitHub sem esperar o agendamento — útil para testes.

---

## 6. Execução e validação

### 6.1. Primeiro teste manual

1. Entre na aba **Actions** do repositório.
2. Na lista à esquerda, clique em **Relatório de NCs — BCMS**.
3. Clique em **Run workflow → Run workflow** (botão verde).
4. Aguarde 30–60 segundos e recarregue a página.

### 6.2. O que o log deve mostrar (sucesso)

```
[HH:MM] Baixando planilha...
[HH:MM] URL: https://docs.google.com/spreadsheets/d/1Jv546wpWQSFAlep3oLRAg29hVy86iJxJ/export?format=xlsx
[HH:MM] Planilha: 1569 linhas x 66 colunas
[HH:MM] NCs BCMS extraídas: <N>
[HH:MM] Conectando smtp.gmail.com:465...
[HH:MM] E-mail enviado.
  To:  ['...']
  Bcc: ['...']
[HH:MM] Concluído com sucesso.
```

### 6.3. Se houver falha

O script faz `sys.exit(1)` em qualquer erro, o que marca o job como vermelho (❌) na aba Actions. Vá na linha vermelha → clique no step **Executar relatório** → copie a última stack trace. Consulte [Troubleshooting](#10-troubleshooting) abaixo.

---

## 7. Estrutura do relatório gerado

Cada execução envia um e-mail com:

- **Assunto:** `[BCMS] Relatório de NCs — DD/MM/AAAA`
- **Corpo:** texto puro formatado no padrão WhatsApp (asteriscos para negrito, underscores para itálico, emojis)

Divisão em três módulos:

### Módulo 1 — Créditos Recebidos
- Dia anterior (último dia útil).
- Semana corrente (segunda a sexta).
- Dentro de cada janela, NCs agrupadas por **OGU (160329)** e **FEx (167329)**.
- Para cada NC: número, valor, ND, PI (código + descrição) e descrição da NC (até 95 caracteres).

### Módulo 2 — Créditos Recolhidos / Devolvidos
Mesma estrutura do Módulo 1, porém com:
- **ANULACAO DE DESCENTRALIZACAO DE CREDITO** (estornos/recolhimentos)
- **DEVOLUCAO DE DESCENTRALIZACAO DE CREDITO** (devoluções)

### Módulo 3 — Resumo Consolidado
Para o dia anterior e para a semana:
- ✅ Recebido total, com abertura OGU/FEx
- ❌ Recolhido total, com abertura OGU/FEx
- 🟢/🔴 Saldo líquido (positivo = verde; negativo = vermelho)

---

## 8. Lógica técnica do script

### 8.1. Mapeamento de colunas (Tesouro Gerencial)

Validado em **23/04/2026** contra a planilha ao vivo (1.569 linhas × 66 colunas, cabeçalho até linha 8):

| Coluna | Índice | Conteúdo |
|---|---:|---|
| UG Executora (código UASG) | 3 | `160329`, `167329`, etc. |
| UG Executora (nome) | 4 | `BCMS` |
| Número da NC | 5 | `160329000012026NC400182` |
| Ação Governo | 6 | `2000`, `212B`, etc. |
| PI — código | 7 | `E6MIPLJBIDS` |
| PI — descrição | 8 | `MATERIAL DE INTENDENCIA` |
| ND — código | 9 | `339030`, `339039`, etc. |
| ND — descrição | 10 | `MATERIAL DE CONSUMO` |
| NC — Descrição | 11 | Texto livre |
| NC — Operação (Tipo) | 12 | `DESCENTRALIZACAO...`, `ANULACAO...`, `DEVOLUCAO...` |
| NC — Dia Emissão | 13 | `05/02/2026` |
| PROVISÃO RECEBIDA (CC) | 15 | Valor numérico (pode ser negativo em anulações) |

### 8.2. Classificação das NCs

A classificação usa o campo **"NC - Operação (Tipo)"** (coluna 12), que é mais confiável que regex na descrição:

| Tipo no SIAFI | Classificação no relatório |
|---|---|
| `DESCENTRALIZACAO DE CREDITO` | **Recebida** — entra no Módulo 1 |
| `ANULACAO DE DESCENTRALIZACAO DE CREDITO` | **Recolhida** — entra no Módulo 2 |
| `DEVOLUCAO DE DESCENTRALIZACAO DE CREDITO` | **Recolhida** — entra no Módulo 2 |
| `DETALHAMENTO DE CREDITO` | *Ignorada* — é operação interna sobre crédito já recebido, não uma nova entrada/saída |

### 8.3. Filtros aplicados

- UASG deve ser **exatamente** `160329` ou `167329` (UASG de outras UGs é descartada).
- Exclui linhas com `NAO SE APLICA`, `'-9` ou `-9` nas chaves críticas (artefatos do SIAFI para linhas vazias).
- Ignora NCs com valor absoluto `< R$ 1,00` (artefatos de R$ 0,01 gerados pelo SIGA para carga de dados).
- Para data, usa `dia útil anterior` (se hoje é segunda, retorna a sexta).

### 8.4. Formatação monetária

Padrão brasileiro: `R$ 1.234.567,89` — via função `fmt_brl()`. Valores negativos recebem prefixo `-` (útil no saldo líquido quando o recolhimento supera o recebido).

### 8.5. BCC — modo SMTP correto

O script implementa BCC **no envelope SMTP**, não no cabeçalho MIME. Isso é o que efetivamente oculta os destinatários:

```python
envelope = dest_to + dest_bcc
smtp.sendmail(EMAIL_REMETENTE, envelope, msg.as_string())
# NÃO escreve msg["Bcc"] — se escrevesse, alguns MTAs manteriam o cabeçalho
# e o "oculto" deixaria de ser oculto.
```

---

## 9. Personalização

### 9.1. Alterar o horário

Já coberto em [§5.1](#51-horário-da-execução). Edite apenas a linha do cron no YAML.

### 9.2. Alterar destinatários

Sem tocar no código — edite o secret `EMAIL_DESTINO` (ou `EMAIL_BCC`) em **Settings → Secrets**. Aceita múltiplos endereços separados por vírgula ou ponto-e-vírgula.

### 9.3. Incluir outra UG no mesmo relatório

Não é recomendado — o objetivo deste repositório é um relatório enxuto só do BCMS. Se precisar, use o `relatorio_nc.py` multi-UG do repositório da BaApLog.

### 9.4. Alterar o assunto do e-mail

No `relatorio_nc_bcms.py`, função `main()`:

```python
assunto = f"[BCMS] Relatório de NCs — {hoje.strftime('%d/%m/%Y')}"
```

---

## 10. Troubleshooting

### 10.1. `HTTP Error 404: Not Found` ao baixar a planilha

- **Causa mais provável:** alguém tentou sobrescrever `SPREADSHEET_ID` via secret do GitHub, causando a mesma falha da sprint passada.
- **Solução:** confirme que o `env:` do step **não** contém `SPREADSHEET_ID`. Ele vem hardcoded do script.
- O log agora mostra a URL efetiva: `URL: https://docs.google.com/spreadsheets/d/<ID>/export?format=xlsx`. Se o ID no log não for `1Jv546wpWQSFAlep3oLRAg29hVy86iJxJ`, algum secret está sobrescrevendo.

### 10.2. `HTTP Error 503: Service Unavailable` ao baixar a planilha

- **Causa:** Google Sheets saturado em momento de pico.
- **Solução:** o script já faz 3 tentativas com backoff exponencial (2 s → 4 s). Se ainda falhar, rode manualmente alguns minutos depois.

### 10.3. `SMTPAuthenticationError: Username and Password not accepted`

- **Causa:** senha de app inválida, 2FA desligado, ou `EMAIL_REMETENTE` diferente da conta que gerou a senha de app.
- **Solução:**
  1. Confirme 2FA ativo em <https://myaccount.google.com/security>.
  2. Gere nova senha de app em <https://myaccount.google.com/apppasswords>.
  3. Atualize o secret `EMAIL_SENHA` **sem espaços**.
  4. Confirme que `EMAIL_REMETENTE` corresponde à conta que criou a senha de app.

### 10.4. `Credenciais ausentes — imprimindo localmente` no log

- **Causa:** algum secret está cadastrado mas com nome errado, ou está faltando no `env:` do step.
- **Solução:** confirme que os três secrets `EMAIL_REMETENTE`, `EMAIL_SENHA`, `EMAIL_DESTINO` existem em **Settings → Secrets → Actions** (maiúsculas exatas). E que o bloco `env:` do workflow os mapeia. Esse bloco precisa estar **no step que executa o script**, não no nível do job.

### 10.5. E-mail chega, mas sem nenhuma NC

- **Causa mais provável:** ninguém emitiu NC para o BCMS no dia anterior — é normal em dias sem movimentação, especialmente sextas e datas entre empenhos.
- **Diagnóstico rápido:** veja se o log mostra `NCs BCMS extraídas: 0` (real ausência) ou um número > 0 mas nada no e-mail (indica bug — abrir issue).
- O relatório ainda é enviado mesmo sem movimentação, para servir de controle de que a automação está viva.

### 10.6. Cron não dispara no horário esperado

- **Causa:** GitHub Actions roda cron em UTC.
- **Solução:** lembre da conversão — 09h30 BRT = 12h30 UTC. Para 07h30 BRT, seria 10h30 UTC (`'30 10 * * 1-5'`).
- ⚠️ O GitHub documenta que o cron pode atrasar em até 15 minutos em horários de pico. Isso é esperado e não é bug.

### 10.7. `UserWarning: Workbook contains no default style` nos logs

Aviso inofensivo do openpyxl. Pode ser ignorado — não afeta a leitura dos dados.

---

## 11. Fundamentação legal

Este relatório auxilia no cumprimento dos deveres do ordenador de despesas e do gestor orçamentário estabelecidos em:

- **Lei nº 4.320/1964** — normas gerais de direito financeiro, especialmente arts. 58 a 65 (empenho e descentralização de crédito).
- **Decreto nº 93.872/1986** — unificação dos recursos de caixa do Tesouro Nacional, arts. 23 e seguintes (descentralização).
- **Lei nº 4.617/1965** — instituição do **Fundo do Exército (FEx)**, base legal da UASG 167329.
- **Decreto nº 91.575/1985** — regulamentação do FEx; administrado pela **Secretaria de Economia e Finanças (SEF)** por meio da **Divisão do Fundo do Exército (DiFEx)**.
- **LOA vigente** — dotação consignada ao Comando do Exército (Órgão 52000, UO 52121), base orçamentária das UASGs 160xxx.
- **Diretriz Especial de Economia e Finanças 2025-26** (Gen Ex TOMÁS) — marcos de liquidação por ação orçamentária, referência para acompanhamento.

A distinção OGU × FEx não é meramente contábil — tem efeitos jurídico-administrativos relevantes quanto a vinculação do crédito, inscrição em Restos a Pagar e prestação de contas, conforme já documentado em sprints anteriores do projeto.

---

## Contato

- **Responsável técnico:** Sgt De Campos — BaApLog / E10 (Assessoria Financeira e Orçamentária).
- Para alterações estruturais no relatório, consolidar demanda via documento na pasta `docs/` antes de alterar o script.

---

*Documento gerado em 23/04/2026 · BaApLog — Assessoria Jurídica e Financeira*
