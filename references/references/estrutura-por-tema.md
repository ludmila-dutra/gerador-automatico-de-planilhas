# Estrutura Detalhada por Tema

---

## Perguntas e Respostas por Tema (Coluna D das planilhas de execução)

| Tema | Chave | Pergunta (coluna D) | Resposta SIM | Resposta NÃO |
|------|-------|---------------------|--------------|--------------|
| Família | `familia` | O processo é de Direito de Família? | SIM, é um processo de direito de familia | NÃO é um processo de direito de familia |
| Saúde | `saude` | O processo trata sobre questões de saúde? | SIM, o processo trata sobre questões de saúde | NÃO é um processo que trata sobre questões de saúde |
| Criminal | `criminal` | O processo é de Direito Criminal Finalizado? | SIM, é um processo de direito criminal finalizado | NÃO é um processo de direito criminal finalizado |
| ECA | `eca` | O processo envolve ECA ou menor infrator? | SIM, é um processo que envolve ECA ou menor infrator | NÃO é um processo que envolve ECA ou menor infrator |
| Trabalhista | `trabalhista` | O processo é de Direito Trabalhista? | SIM, é um processo de direito trabalhista | NÃO é um processo de direito trabalhista |
| Segredo | `segredo` | O processo está em segredo de justiça? | SIM, é um processo em segredo de justiça | NÃO é um processo em segredo de justiça |

---

## Opções Fixas — Coluna E (iguais para todos os temas)

```
SIM, é um caso de PUBLICIDADE PROCESSUAL ou de RESTRIÇÃO A IDENTIFICAÇÃO DA PARTE
SIM, é um caso de DADO PESSOAL EXPOSTO
SIM, é um caso de ENVOLVIMENTO INCORRETO
NÃO se enquadra em nenhum critério e deveria ser REPROVADO
```

---

## Opções Fixas — Coluna F (iguais para todos os temas)

```
FAMÍLIA: Filiação, Adoção e Guarda / Casamento e União / Pensão Alimentícia
CRIMES SEXUAIS OU DE ÓDIO: Crime Sexual / Vítima ou testemunha de crime / Crime de Ódio
QUESTÕES DE SAÚDE: Doença Estigmatizante / Erro Médico / Procedimento Médico / Doença Geral
TRABALHISTA
MENOR: Menor Infrator / ECA
PENAL: Criminal Finalizado / Violência Doméstica
SEGREDO DE JUSTIÇA
```

---

## Layout das Planilhas de Execução (Mari / Lud)

### Cabeçalhos (linhas 1–3)

| Linha | Conteúdo |
|-------|----------|
| 1 | Grupos: `DADOS DO PROCESSO` (A–C) \| `EXECUÇÃO` (D–G) |
| 2 | Títulos das colunas (ver abaixo) |
| 3 | Sub-cabeçalho com nome da pessoa executora (nas colunas D–G) |
| 4+ | Dados dos processos |

### Colunas

| Col | Título | Conteúdo |
|-----|--------|----------|
| A | ID | Número sequencial (1, 2, 3...) |
| B | Link do Processo | URL do Jusbrasil |
| C | Número do processo | Ex: 0001234-56.2024.8.26.0000 |
| D | [Pergunta do tema] | Preenchido manualmente: SIM ou NÃO (ver respostas por tema) |
| E | O processo se enquadra nos critérios institucionais de aprovação? | Uma das 4 opções fixas |
| F | Sendo caso de PUBLICIDADE PROCESSUAL... em qual tema? | Uma das 7 opções de tema |
| G | OBSERVAÇÕES | Texto livre, preenchido manualmente |

**Sem fórmulas** — planilhas de execução são preenchidas manualmente.

---

## Layout da Planilha Mãe — Aba Execução

### Cabeçalhos (linhas 1–3)

| Linha | Conteúdo |
|-------|----------|
| 1 | Grupos: `DADOS DO PROCESSO` (A–C) \| `EXECUÇÃO` (D–K) \| `DESEMPATE` (L–N) |
| 2 | Títulos das colunas |
| 3 | Sub-cabeçalhos Mari / Lud nas colunas D–K |
| 4+ | Dados + fórmulas |

### Colunas

| Col | Título | Conteúdo |
|-----|--------|----------|
| A | ID | Sequencial |
| B | Link do Processo | URL do Jusbrasil |
| C | Número do processo | Número CNJ |
| D | [Pergunta] — Mari | IMPORTRANGE da planilha Mari (col D) |
| E | [Pergunta] — Lud | IMPORTRANGE da planilha Lud (col D) |
| F | Critérios — Mari | IMPORTRANGE da planilha Mari (col E) |
| G | Critérios — Lud | IMPORTRANGE da planilha Lud (col E) |
| H | Tema — Mari | IMPORTRANGE da planilha Mari (col F) |
| I | Tema — Lud | IMPORTRANGE da planilha Lud (col F) |
| J | Obs — Mari | IMPORTRANGE da planilha Mari (col G) |
| K | Obs — Lud | IMPORTRANGE da planilha Lud (col G) |
| L | DESEMPATE — Pergunta | Fórmula SE: acorda ou pede desempate |
| M | DESEMPATE — Critérios | Fórmula SE: acorda ou pede desempate |
| N | DESEMPATE — Tema | Fórmula SE: acorda ou pede desempate |

---

## Layout da Planilha Mãe — Aba Desempenho do Agente

### Cabeçalhos (linhas 1–2)

| Linha | Conteúdo |
|-------|----------|
| 1 | Grupos: [Pergunta do tema] (C–E) \| `EM QUAL TEMA O PROCESSO SE ENQUADRA?` (F–H) \| `COMENTÁRIOS` (J) |
| 2 | Títulos das colunas |
| 3+ | Dados + fórmulas |

### Colunas

| Col | Título | Conteúdo |
|-----|--------|----------|
| A | ID | Sequencial (1, 2, 3...) |
| B | LINK | URL do Jusbrasil |
| C | MODERAÇÃO OQJ | IMPORTRANGE da Planilha Mãe (Execução col L — desempate da pergunta) |
| D | MODERAÇÃO AGENTE (EXPANDIDO) | Resposta do agente de IA — fornecida pelo usuário |
| E | DESEMPENHO DO AGENTE | Fórmula SE/REGEXMATCH — compara C e D |
| F | CLASSIFICAÇÃO OQJ | IMPORTRANGE da Planilha Mãe (Execução col N — desempate do tema) |
| G | CLASSIFICAÇÃO AGENTE (EXPANDIDO) | Classificação do agente de IA — fornecida pelo usuário |
| H | DESEMPENHO DO AGENTE | Fórmula REGEXMATCH — compara F e G |
| J | COMENTÁRIOS | Preenchido manualmente |

---

## Formato Esperado do Arquivo de Entrada (Processos)

O arquivo CSV ou XLSX deve conter as colunas a seguir (os nomes são inferidos automaticamente):

| Coluna obrigatória | Nomes aceitos |
|---------------------|---------------|
| Link | `link`, `Link`, `LINK`, `url`, `URL` |
| Número do processo | `numero`, `Número do processo`, `numero_processo`, `NÚMERO` |

O script atribui IDs sequenciais (1, 2, 3...) na ordem das linhas do arquivo.

Exemplo mínimo de CSV:
```csv
link,numero
https://www.jusbrasil.com.br/processos/656501759/...,0001234-56.2024.8.26.0000
https://www.jusbrasil.com.br/processos/801482778/...,0005431-91.2016.8.21.0002
```
