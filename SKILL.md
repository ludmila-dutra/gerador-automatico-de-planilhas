---
name: qa-sheet-generator
description: >
  Gera automaticamente as 3 planilhas de QA para moderação de processos judiciais no Google Sheets:
  planilha mãe (desempate) e planilhas das executoras Mariana (Mari) e Ludmila (Lud).
  Use esta skill SEMPRE que o usuário pedir para criar, gerar ou montar planilhas de QA de moderação,
  planilhas de desempate, planilhas de execução (Mari/Lud), ou qualquer planilha de Trust & Safety do Jusbrasil.
  Também ative quando o usuário mencionar "agente de saúde", "agente de família", "agente de direito de família",
  "agente criminal", "agente trabalhista", "moderação pelo agente", "QA de moderação", "processos judiciais para QA",
  ou pedir para preencher links/processos em planilhas de avaliação de agentes de IA.
  Ative também quando o usuário enviar um arquivo com links e números de processos pedindo para gerar planilhas.
---

# Gerador Automático de Planilhas de QA

## O que esta skill faz

Gera **3 planilhas XLSX** prontas para upload no Google Sheets, com todas as fórmulas preenchidas:

1. **[EXECUÇÃO MARI]** — planilha de execução da Mariana (sem fórmulas, preenchimento manual)
2. **[EXECUÇÃO LUD]** — planilha de execução da Ludmila (sem fórmulas, preenchimento manual)
3. **[PLANILHA MÃE]** — planilha de desempate com IMPORTRANGE das duas execuções + fórmulas de desempenho do agente

O script em `scripts/gerar_planilhas.py` faz todo o trabalho de geração.

---

## Coleta de Informações

Antes de gerar, confirme com o usuário via AskUserQuestion:

| # | Informação | Obrigatório? |
|---|-----------|-------------|
| 1 | **Arquivo de processos** (CSV or XLSX com colunas: link, número do processo) | ✅ Sim |
| 2 | **Tema do agente** (`familia`, `saude`, `criminal`, `eca`, `trabalhista`, `segredo`) | ✅ Sim |
| 3 | **Link da planilha Mari** no Google Sheets | ⚠️ Se disponível |
| 4 | **Link da planilha Lud** no Google Sheets | ⚠️ Se disponível |
| 5 | **Link da planilha Mãe** no Google Sheets | ⚠️ Se disponível |

Se os links das planilhas não estiverem disponíveis, use os placeholders padrão (`LINK-PLANILHA-MARI`, etc.) — o usuário substitui depois no Google Sheets.

---

## Processo de Geração

### Passo 1 — Preparar o arquivo de entrada

O arquivo de entrada deve ser CSV ou XLSX com colunas `link` e `numero` (o script aceita variações de nome).

Se o usuário enviar um arquivo em outro formato, converta antes de rodar o script:

```bash
# Verificar colunas disponíveis
python3 -c "
import openpyxl
wb = openpyxl.load_workbook('arquivo.xlsx', data_only=True)
ws = wb.active
print([c.value for c in ws[1]])
"
```

### Passo 2 — Rodar o script de geração

```bash
python3 /sessions/relaxed-pensive-hypatia/mnt/Gerador\ automático\ de\ planilhas/scripts/gerar_planilhas.py \
  <ARQUIVO_PROCESSOS> \
  <TEMA> \
  [--link-mari "URL_MARI"] \
  [--link-lud "URL_LUD"] \
  [--link-mae "URL_MAE"] \
  --output-dir /sessions/relaxed-pensive-hypatia/mnt/outputs/
```

**Temas disponíveis:** `familia` | `saude` | `criminal` | `eca` | `trabalhista` | `segredo`

**Exemplo:**
```bash
python3 /sessions/relaxed-pensive-hypatia/mnt/Gerador\ automático\ de\ planilhas/scripts/gerar_planilhas.py \
  /sessions/relaxed-pensive-hypatia/mnt/uploads/processos.csv \
  familia \
  --output-dir /sessions/relaxed-pensive-hypatia/mnt/outputs/
```

### Passo 3 — Copiar para a pasta do usuário

```bash
cp /sessions/relaxed-pensive-hypatia/mnt/outputs/*.xlsx \
   "/sessions/relaxed-pensive-hypatia/mnt/Gerador automático de planilhas/"
```

### Passo 4 — Entregar ao usuário

Forneça links `computer://` para os 3 arquivos gerados e oriente sobre os próximos passos no Google Sheets.

---

## Estrutura das Planilhas

Para detalhes completos de colunas por tema, consulte:
👉 `references/estrutura-por-tema.md`

### Resumo r�pido

**Planilhas Mari e Lud (sem fórmulas):**

| Col | Conteúdo |
|-----|----------|
| A | ID (sequencial: 1, 2, 3...) |
| B | Link do processo (Jusbrasil) |
| C | Número do processo (CNJ) |
| D | Pergunta principal (varia por tema — preenchida manualmente: SIM/NÃO) |
| E | O processo se enquadra nos critérios institucionais? (4 opções fixas) |
| F | Sendo caso de PUBLICIDADE PROCESSUAL, em qual tema? (7 opções) |
| G | Observações (texto livre) |

**Planilha Mãe — Aba Execução:**

| Col | Conteúdo |
|-----|----------|
| A–C | Dados do processo (ID, Link, Número) |
| D–K | IMPORTRANGE das planilhas Mari (cols D, E, F, G) e Lud (cols D, E, F, G) |
| L | Desempate da pergunta principal |
| M | Desempate dos critérios institucionais |
| N | Desempate do tema |

**Planilha Mãe — Aba Desempenho do Agente:**

| Col | Conteúdo |
|-----|----------|
| A–B | ID + Link |
| C | Moderação OQJ (IMPORTRANGE col L da aba Execução) |
| D | Resposta do agente (fornecida pelo usuário) |
| E | Desempenho de moderação (fórmula SE/REGEXMATCH por tema) |
| F | Classificação OQJ (IMPORTRANGE col N da aba Execução) |
| G | Classificação do agente (fornecida pelo usuário) |
| H | Desempenho de classificação (fórmula REGEXMATCH universal) |
| J | Comentários (manual) |

---

## Fórmulas

Para fórmulas completas de todas as colunas e todos os temas, consulte:
👉 `references/formulas-por-tema.md`

As principais são:

**Desempate (cols L, M, N — linha r):**
```
=SE(OU(D{r}="";E{r}="");"Pendente";SE(D{r}=E{r};D{r};"DESEMPATAR"))
```

**IMPORTRANGE (exemplo para 100 processos):**
```
=IMPORTRANGE("URL_DA_PLANILHA";"Execução!D4:D103")
```

**Desempenho de moderação (varia por tema):**
Ver `references/formulas-por-tema.md` → seção "Coluna E"

**Desempenho de classificação temática (universal):**
Ver `references/formulas-por-tema.md` → seção "Coluna H"

---

## Orientações pós-geração para o usuário

Após fazer upload das 3 planilhas no Google Sheets:

1. Copie as URLs das planilhas Mari e Lud
2. Na Planilha Mãe (aba Execução), substitua `LINK-PLANILHA-MARI` e `LINK-PLANILHA-LUD` pelas URLs reais nas células D4, E4, F4, G4, H4, I4, J4, K4
3. Clique em **"Conceder acesso"** quando o IMPORTRANGE solicitar permissão
4. Copie a URL da Planilha Mãe e substitua `LINK-PLANILHA-MAE` nas células C3 e F3 da aba Desempenho do Agente
5. Preencha as colunas D e G da aba Desempenho do Agente com as respostas do agente de IA

---

## Notas Importantes

- **IDs:** sempre sequenciais começando em 1 (ex: 1 a 100, ou 1 a 200)
- **Linha de dados:** começa na linha 4 (linhas 1–3 são cabeçalhos)
- **Fórmulas:** escritas em português do Google Sheets (`SE`, `OU`, `E`, `IMPORTRANGE`, `REGEXMATCH`, `MINÚSCULA`)
- **Planilhas Mari/Lud:** sem fórmulas — só estrutura e dados
- **Separador de argumento:** `;` (locale pt-BR). Se o Sheets estiver em inglês, use `,`
- **Amostras de 200:** ajuste o range de `D4:D103` para `D4:D203` (linha_fim = n + 3)
