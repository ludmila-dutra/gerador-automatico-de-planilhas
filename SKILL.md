---
name: ops-trust-sheet-generator
description: >
  Cria planilhas Google Sheets (.xlsx) prontas para upload, com base em instruções fornecidas.
  Use esta skill SEMPRE que o usuário pedir para criar, gerar ou montar planilhas — especialmente
  planilhas de QA de moderação, planilhas de execução (Mari/Lud), planilha mãe de desempate,
  ou qualquer planilha de Trust & Safety do Jusbrasil.
  Também ative quando o usuário mencionar "gerar planilha", "montar planilha", "criar sheets",
  "planilha de QA", "planilha de moderação", "agente criminal", "agente de família",
  "agente trabalhista", "agente de saúde", "agente ECA", "agente segredo", ou enviar um arquivo
  com processos pedindo para gerar planilhas. NUNCA gere planilhas manualmente com openpyxl —
  sempre use o script dedicado quando disponível.
---

# Gerador de Planilhas Google Sheets

## Dois modos de operação

### MODO A — Planilhas de QA de Moderação (script dedicado)

**REGRA ABSOLUTA:** Para qualquer pedido de planilha de QA de moderação, SEMPRE execute o script abaixo. NUNCA gere planilhas manualmente. O script já contém todas as cores, menus suspensos e fórmulas corretas — recriar qualquer parte dele fora do script produzirá resultados errados.

#### Informações necessárias

| # | Informação | Obrigatório? |
|---|-----------|-------------|
| 1 | **Arquivo de processos** (CSV ou XLSX com colunas: link, número do processo) | ✅ Sim |
| 2 | **Tema** (`familia`, `saude`, `criminal`, `eca`, `trabalhista`, `segredo`) | ✅ Sim |
| 3 | **Link da planilha Mari** no Google Sheets | ⚠️ Se disponível |
| 4 | **Link da planilha Lud** no Google Sheets | ⚠️ Se disponível |
| 5 | **Link da planilha Mãe** no Google Sheets | ⚠️ Se disponível |

Se os links não estiverem disponíveis, use os placeholders `LINK-PLANILHA-MARI`, `LINK-PLANILHA-LUD` e `LINK-PLANILHA-MAE` — o usuário substitui depois.

#### Fluxo de geração obrigatório

As planilhas devem ser criadas diretamente no Google Sheets, na pasta do Drive informada pelo usuário. O fluxo é:

1. Gerar o `.xlsx` localmente com toda a formatação via `bash_tool` + `openpyxl`
2. Ler o arquivo gerado como base64
3. Fazer upload para o Google Drive via `Google Drive:create_file` com:
   - `contentMimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"` (sem `disableConversionToGoogleType`) — o Drive converte automaticamente para Google Sheets
   - `parentId`: ID da pasta de destino (extraído da URL do Drive fornecida pelo usuário: `https://drive.google.com/drive/folders/<FOLDER_ID>`)
4. Capturar a URL da planilha criada (`https://docs.google.com/spreadsheets/d/<FILE_ID>`)
5. Repetir para Mari → Lud → Mãe (usando os links reais de Mari e Lud nas fórmulas da Mãe)

```python
# Exemplo: ler xlsx como base64 para upload
import base64
with open("/tmp/planilha_mari.xlsx", "rb") as f:
    b64 = base64.b64encode(f.read()).decode()
# passar b64 como base64Content no Google Drive:create_file
```

**Pasta padrão:** extrair o `FOLDER_ID` da URL fornecida pelo usuário. Ex: `https://drive.google.com/drive/folders/1ck0X9oBWBllziC3W9Nv_cesIo_V1WTcz` → `parentId = "1ck0X9oBWBllziC3W9Nv_cesIo_V1WTcz"`

**Atenção — dropdowns após conversão:** A conversão xlsx → Google Sheets preserva cores, merges e fórmulas, mas pode não preservar DataValidations (menus suspensos). Após o upload, oriente o usuário a verificar os dropdowns nas colunas D, E, F e recriá-los manualmente se necessário.

#### Ordem obrigatória de geração

1. **[EXECUÇÃO MARI]** — planilha da Mariana com menus suspensos nas colunas D, E, F
2. **[EXECUÇÃO LUD]** — planilha da Ludmila com menus suspensos nas colunas D, E, F
3. **[PLANILHA MÃE]** — gerada por último, com os links reais de Mari e Lud já embutidos nas fórmulas IMPORTRANGE das abas "Execução" e "Comparação com o agente"

Se os links não estiverem disponíveis no momento da geração, use os placeholders `LINK-PLANILHA-MARI` e `LINK-PLANILHA-LUD` — o usuário substitui depois.

#### Especificações completas (extraídas dos modelos oficiais)

**Paleta de cores (fgColor ARGB):**
- DADOS DO PROCESSO: `FF1C664F`
- EXECUÇÃO (cabeçalho): `FF6AA84F`
- Sub-cabeçalho Mari: `FF9FC5E8`
- Sub-cabeçalho Lud: `FFB4A7D6`
- DESEMPATE: `FF85200C`
- EM QUAL TEMA: `FF38761D`
- MODERAÇÃO / CLASSIFICAÇÃO / DESEMPENHO: `FF980000`
- Comentários (Comparação): `FFA2C4C9`
- Fonte cabeçalhos: `FFFFFFFF` (branco)
- ID col A (Mãe): fundo branco `FFFFFFFF`

---

##### Planilha de Execução (Mari e Lud) — 1 aba: "Execução"

**Aba "Execução":**

Larguras de coluna:
- A: 5.88 | B: 25.5 | C: 21.5 | D: 30.88 | E: 33.0 | F: 31.25 | G: 21.0

Alturas de linha: linha 1 = 27.75, linha 2 = 27.75

Merges: `B1:C1`, `D1:F1`

Estrutura de cabeçalho (3 linhas):
- Linha 1: A=vazio(`FF1C664F`), B="DADOS DO PROCESSO"(`FF1C664F`,bold,branco), D="EXECUÇÃO"(`FF6AA84F`,bold,branco), G=vazio(`FF6AA84F`)
- Linha 2: A="ID"(`FF1C664F`,bold,branco), B="Link do Processo"(`FF1C664F`,bold,branco), C="Número do processo"(`FF1C664F`,bold,branco), D=pergunta col D por tema(`FF6AA84F`,bold,branco), E="O processo se enquadra nos critérios institucionais de aprovação?"(`FF6AA84F`,bold,branco), F="Sendo caso de PUBLICIDADE PROCESSUAL ou de RESTRIÇÃO A IDENTIFICAÇÃO DA PARTE, em qual tema o processo se enquadra?"(`FF6AA84F`,bold,branco), G="Comentários"(`FF6AA84F`,bold,branco)
- Linha 3: sem conteúdo nas execuções individuais (Mari/Lud)

Dados: linhas 3–202 (200 linhas), IDs 1–200 na col A, cols B e C vazias para preenchimento

**Menus suspensos — regra obrigatória para evitar quebra de opções:**

As opções dos menus contêm vírgulas internas (ex: "SIM, o processo é criminal finalizado"). Se passadas como string inline na `formula1` do DataValidation, a vírgula é interpretada como separador e uma opção vira duas. **Solução obrigatória: usar aba oculta como fonte.**

Criar uma aba chamada `__listas__` (oculta via `ws.sheet_state = "hidden"`) com as opções em colunas separadas, e referenciar essa aba no DataValidation via `formula1="__listas__!$A$1:$A$2"` (range, não string inline).

Exemplo de estrutura da aba `__listas__` para tema `criminal`:

| Col A | Col B | Col C |
|-------|-------|-------|
| SIM, o processo é criminal finalizado | SIM, é um caso de PUBLICIDADE PROCESSUAL ou de RESTRIÇÃO A IDENTIFICAÇÃO DA PARTE | FAMÍLIA: Filiação, Adoção e Guarda / Casamento e união / Pensão alimentícia |
| NÃO é um processo criminal finalizado | SIM, é um caso de DADO PESSOAL EXPOSTO | CRIMES SEXUAIS OU DE ÓDIO: Crime Sexual / Vítima ou testemunha de crime / Crime de Ódio |
| | SIM, é um caso de ENVOLVIMENTO INCORRETO | QUESTÕES DE SAÚDE: Doença estigmatizante / Erro médico / Procedimento Médico / Doença Geral |
| | NÃO se enquadra em nenhum critério e deveria ser REPROVADO | TRABALHISTA |
| | | MENOR: Menor infrator / ECA |
| | | PENAL: Criminal Finalizado / Violência Doméstica |
| | | SEGREDO DE JUSTIÇA |

DataValidation usando range:
```python
from openpyxl.worksheet.datavalidation import DataValidation
dv_d = DataValidation(type="list", formula1="__listas__!$A$1:$A$2", showDropDown=False, sqref="D3:D202")
dv_e = DataValidation(type="list", formula1="__listas__!$B$1:$B$4", showDropDown=False, sqref="E3:E202")
dv_f = DataValidation(type="list", formula1="__listas__!$C$1:$C$7", showDropDown=False, sqref="F3:F202")
ws_exec.add_data_validation(dv_d)
ws_exec.add_data_validation(dv_e)
ws_exec.add_data_validation(dv_f)
```

**Cores nas opções (formatação condicional obrigatória):**

DataValidation não suporta cor por opção. Usar `ConditionalFormattingList` do openpyxl com regra `CellIsRule` para cada valor, aplicando `PatternFill` na célula. Aplicar nas colunas D, E e F das linhas de dados.

Paleta de cores para col D (SIM/NÃO):
- "SIM, o processo é criminal finalizado" → fundo `FF93C47D` (verde claro)
- "NÃO é um processo criminal finalizado" → fundo `FFE06666` (vermelho claro)

Paleta de cores para col E (critérios):
- "SIM, é um caso de PUBLICIDADE PROCESSUAL ou de RESTRIÇÃO A IDENTIFICAÇÃO DA PARTE" → `FFFFD966` (amarelo)
- "SIM, é um caso de DADO PESSOAL EXPOSTO" → `FFFFB347` (laranja)
- "SIM, é um caso de ENVOLVIMENTO INCORRETO" → `FFCFE2F3` (azul claro)
- "NÃO se enquadra em nenhum critério e deveria ser REPROVADO" → `FFE06666` (vermelho claro)

Paleta de cores para col F (temas):
- "FAMÍLIA: ..." → `FFEAD1DC` (rosa claro)
- "CRIMES SEXUAIS OU DE ÓDIO: ..." → `FFE06666` (vermelho claro)
- "QUESTÕES DE SAÚDE: ..." → `FFCFE2F3` (azul claro)
- "TRABALHISTA" → `FFFFF2CC` (amarelo claro)
- "MENOR: ..." → `FFD9EAD3` (verde muito claro)
- "PENAL: ..." → `FFEAD9C9` (bege)
- "SEGREDO DE JUSTIÇA" → `FFD9D9D9` (cinza claro)

Implementação:
```python
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import PatternFill

def add_color_rule(ws, col_range, value, hex_color):
    fill = PatternFill(start_color=hex_color, end_color=hex_color, fill_type="solid")
    rule = CellIsRule(operator="equal", formula=[f'"{value}"'], fill=fill)
    ws.conditional_formatting.add(col_range, rule)

# Col D — criminal
add_color_rule(ws, "D3:D202", "SIM, o processo é criminal finalizado", "FF93C47D")
add_color_rule(ws, "D3:D202", "NÃO é um processo criminal finalizado", "FFE06666")

# Col E
add_color_rule(ws, "E3:E202", "SIM, é um caso de PUBLICIDADE PROCESSUAL ou de RESTRIÇÃO A IDENTIFICAÇÃO DA PARTE", "FFFFD966")
add_color_rule(ws, "E3:E202", "SIM, é um caso de DADO PESSOAL EXPOSTO", "FFFFB347")
add_color_rule(ws, "E3:E202", "SIM, é um caso de ENVOLVIMENTO INCORRETO", "FFCFE2F3")
add_color_rule(ws, "E3:E202", "NÃO se enquadra em nenhum critério e deveria ser REPROVADO", "FFE06666")

# Col F
add_color_rule(ws, "F3:F202", "FAMÍLIA: Filiação, Adoção e Guarda / Casamento e união / Pensão alimentícia", "FFEAD1DC")
add_color_rule(ws, "F3:F202", "CRIMES SEXUAIS OU DE ÓDIO: Crime Sexual / Vítima ou testemunha de crime / Crime de Ódio", "FFE06666")
add_color_rule(ws, "F3:F202", "QUESTÕES DE SAÚDE: Doença estigmatizante / Erro médico / Procedimento Médico / Doença Geral", "FFCFE2F3")
add_color_rule(ws, "F3:F202", "TRABALHISTA", "FFFFF2CC")
add_color_rule(ws, "F3:F202", "MENOR: Menor infrator / ECA", "FFD9EAD3")
add_color_rule(ws, "F3:F202", "PENAL: Criminal Finalizado / Violência Doméstica", "FFEAD9C9")
add_color_rule(ws, "F3:F202", "SEGREDO DE JUSTIÇA", "FFD9D9D9")
```

Aplicar as mesmas regras de cores nas colunas correspondentes da Planilha Mãe (D–I nas linhas de dados e L–N no bloco DESEMPATE).

---

##### Planilha Mãe — 2 abas: "Execução" e "Comparação com o agente"

**Aba "Execução" da Mãe:**

Larguras de coluna:
- A: 5.88 | B: 25.5 | C: 23.88 | D: 30.88 | E: 24.0 | F: 33.0 | G: 31.88 | H: 31.25 | I: 31.63 | J: 21.0 | K: 22.38 | L: 32.63 | M: 32.88 | N: 51.0 | O: 10.63 | P: 31.88

Alturas de linha: linha 1 = 27.75, linha 2 = 27.75, linha 3 = 13.5

Merges: `B1:C1`, `D1:I1`, `D2:E2`, `F2:G2`, `H2:I2`, `L2:L3`, `M2:M3`, `N2:N3`

Estrutura de cabeçalho (3 linhas):
- Linha 1: A=vazio(`FF1C664F`), B="DADOS DO PROCESSO"(`FF1C664F`,bold,branco), D="EXECUÇÃO"(`FF6AA84F`,bold,branco), J=vazio(`FF6AA84F`), K=vazio(`FF6AA84F`), L="DESEMPATE"(`FF85200C`,bold,branco), O="Auditoria final LEx"(`FF1C664F`,bold,branco), P="Observações"(`FF1C664F`,bold,branco)
- Linha 2: A="ID"(`FFFFFFFF`,bold,branco), B="Link do Processo"(`FF1C664F`,bold,branco), C="Número do processo"(`FF1C664F`,bold,branco), D="O processo é criminal finalizado?"(`FF6AA84F`,bold,branco), F="O processo se enquadra nos critérios institucionais de aprovação?"(`FF6AA84F`,bold,branco), H="Sendo caso de PUBLICIDADE PROCESSUAL..."(`FF6AA84F`,bold,branco), J="Comentários"(`FF6AA84F`,bold,branco), K="Comentários"(`FF6AA84F`,bold,branco), L="O processo é criminal finalizado??"(`FF6AA84F`,bold,branco), M="O processo se enquadra nos critérios institucionais de aprovação?"(`FF6AA84F`,bold,branco), N="Sendo caso de PUBLICIDADE PROCESSUAL..."(`FF6AA84F`,bold,branco)
- Linha 3: D="Mari"(`FF9FC5E8`,bold,branco), E="Lud"(`FFB4A7D6`,bold,branco), F="Mari"(`FF9FC5E8`,bold,branco), G="Lud"(`FFB4A7D6`,bold,branco), H="Mari"(`FF9FC5E8`,bold,branco), I="Lud"(`FFB4A7D6`,bold,branco), J="Mari"(`FF9FC5E8`,bold,branco), K="Lud"(`FFB4A7D6`,bold,branco)

Dados: linhas 4–203 (200 linhas), IDs 1–200 na col A

Fórmulas IMPORTRANGE (linhas 4–203) — usar placeholder se links não fornecidos:
- D: `=IFERROR(__xludf.DUMMYFUNCTION("IMPORTRANGE(""LINK-PLANILHA-MARI"",""Execução!D3:D202"")"), "SIM, o processo é criminal finalizado")`  ← col Mari
- E: idem com LINK-PLANILHA-LUD
- F: idem col E Mari
- G: idem col E Lud
- H: idem col F Mari
- I: idem col F Lud
- J: idem col G (Comentários) Mari
- K: idem col G (Comentários) Lud

Menus suspensos (DataValidation):
- **D4:E203** — col D por tema
- **F4:G203** — col E (critérios institucionais)
- **H4:I203** — col F (temas)
- **L4:L203** — col D por tema + `",Pendente,Desempatar"`
- **M4:M203** — col E (critérios) + `",Pendente,Desempatar"`
- **N4:N203** — col F (temas) + `",Pendente,Desempatar"`

**Aba "Comparação com o agente" da Mãe:**

Larguras de coluna:
- A: 5.75 | B: 27.63 | C: 14.13 | D: 32.88 | E: 20.13 | F: 21.88 | G: 20.75 | H: 67.25 | I: 22.5 | J: 59.0 | K: 18.13

Merges: `A1:C1` (implícito pelo fill), `D1:E1`, `G1:H1`

Estrutura de cabeçalho (2 linhas):
- Linha 1: A–C=vazio(`FF1C664F`), D="O PROCESSO É CRIMINAL FINALIZADO?"(`FF1C664F`,bold,branco), F="MODERAÇÃO"(`FF980000`,bold,branco), G="EM QUAL TEMA O PROCESSO SE ENQUADRA?"(`FF38761D`,bold,branco), I="CLASSIFICAÇÃO"(`FF980000`,bold,branco), J=vazio(`FFA2C4C9`), K=vazio(`FFA2C4C9`)
- Linha 2: A–C=vazio(`FF1C664F`), D="MODERAÇÃO OQJ"(`FF1C664F`,bold,branco), E="MODERAÇÃO AGENTE"(`FF1C664F`,bold,branco), F="DESEMPENHO DO AGENTE"(`FF980000`,bold,branco), G="CLASSIFICAÇÃO OQJ"(`FF38761D`,bold,branco), H="CLASSIFICAÇÃO AGENTE (EXPANDIDO)"(`FF38761D`,bold,branco), I="DESEMPENHO DO AGENTE"(`FF980000`,bold,branco), J=vazio(`FFA2C4C9`), K="COMENTÁRIOS"(`FFA2C4C9`,bold,branco)

Dados: linhas 3–202 (200 linhas), IDs 1–200 na col A, cols B e C com número e link do processo

Fórmulas (linhas 3–202):
- D: IMPORTRANGE da Planilha Mãe aba Execução col L (LINK-PLANILHA-MAE)
- F: `=IF(OR(AND(D{n}="SIM, o processo é criminal finalizado",E{n}="aprovado"),AND(D{n}="NÃO é um processo criminal finalizado",E{n}="reprovado")),"✅ ACERTO","❌ ERRO")`
- G: IMPORTRANGE da Planilha Mãe aba Execução col N (LINK-PLANILHA-MAE)

---

##### Temas — pergunta e opções da coluna D (na aba `__listas__` col A)

| Tema | Pergunta col D | Opção 1 (A1) | Opção 2 (A2) |
|------|---------------|-------------|-------------|
| `criminal` | "O processo é criminal finalizado?" | "SIM, o processo é criminal finalizado" | "NÃO é um processo criminal finalizado" |
| `familia` | "O processo trata sobre questões de família?" | "SIM, o processo trata sobre questões de família" | "NÃO é um processo de família" |
| `saude` | "O processo trata sobre questões de saúde?" | "SIM, o processo trata sobre questões de saúde" | "NÃO é um processo de saúde" |
| `trabalhista` | "O processo trata sobre questões trabalhistas?" | "SIM, o processo trata sobre questões trabalhistas" | "NÃO é um processo trabalhista" |
| `eca` | "O processo envolve menor (ECA)?" | "SIM, o processo envolve menor (ECA)" | "NÃO é um processo ECA" |
| `segredo` | "O processo está em segredo de justiça?" | "SIM, o processo está em segredo de justiça" | "NÃO está em segredo de justiça" |

Cores col D por tema (formatação condicional — adaptar os valores de SIM/NÃO conforme o tema):
- Opção SIM → fundo `FF93C47D` (verde claro)
- Opção NÃO → fundo `FFE06666` (vermelho claro)

---

#### Regras técnicas obrigatórias

- 200 linhas de dados (IDs 1–200), cabeçalho varia por planilha (2 ou 3 linhas)
- **Menus suspensos obrigatoriamente via aba oculta `__listas__`** (`ws.sheet_state = "hidden"`) + DataValidation com `formula1` apontando para range (ex: `__listas__!$A$1:$A$2`). NUNCA usar string inline com vírgula como separador — as opções contêm vírgulas internas
- **Cores nas opções via ConditionalFormatting** (`CellIsRule`) — não via DataValidation
- Fórmulas IMPORTRANGE com `__xludf.DUMMYFUNCTION` para compatibilidade com upload no Google Sheets
- Fórmulas com separador `,` (padrão inglês/Sheets)
- Freeze panes: primeira linha de dados (ex: `A3` nas execuções, `A4` na mãe)
- Fonte padrão: Calibri ou Arial, tamanho 10–11
- Alinhamento cabeçalhos: centralizado horizontal e vertical
- wrap_text=True nos cabeçalhos com textos longos

#### Após criar as planilhas no Drive

1. Compartilhe os links das 3 planilhas com o usuário
2. Oriente a clicar em **"Conceder acesso"** quando o IMPORTRANGE solicitar permissão em cada aba
3. Verificar se os dropdowns (cols D, E, F) foram preservados — se não, recriar manualmente

---

### MODO B — Planilhas genéricas (outros casos)

Para pedidos de planilha que **não sejam** QA de moderação, gere via Python com `openpyxl`.

#### Processo

1. Entenda a estrutura desejada: colunas, tipos de dados, fórmulas, formatação
2. Se o usuário enviar um arquivo de referência, leia-o primeiro para entender o padrão
3. Escreva o script Python e execute via `bash_tool`
4. Salve em `/mnt/user-data/outputs/` e entregue com `present_files`

#### Boas práticas

- Sempre instale dependências com `--break-system-packages` se necessário
- Use `openpyxl` para .xlsx; para fórmulas em português (Google Sheets), use `;` como separador de argumentos
- Freezar a primeira linha (`ws.freeze_panes = "A2"`) melhora a usabilidade
- Larguras de coluna razoáveis: IDs ~8, links ~55, texto longo ~40
