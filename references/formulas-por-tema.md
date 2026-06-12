# Fórmulas por Tema — Referência Completa

> Todas as fórmulas são para **Google Sheets** (português).  
> Use `;` como separador. Se o Sheets estiver em inglês, troque `;` por `,` e os nomes das funções por equivalentes ingleses.

---

## IMPORTRANGE — Aba Execução da Planilha Mãe

As fórmulas de IMPORTRANGE vão na **linha 4** e o Google Sheets expande automaticamente.

```
Col D (pergunta principal - Mari):
=IMPORTRANGE("LINK-PLANILHA-MARI";"Execução!D4:D{FIM}")

Col E (pergunta principal - Lud):
=IMPORTRANGE("LINK-PLANILHA-LUD";"Execução!D4:D{FIM}")

Col F (critérios institucionais - Mari):
=IMPORTRANGE("LINK-PLANILHA-MARI";"Execução!E4:E{FIM}")

Col G (critérios institucionais - Lud):
=IMPORTRANGE("LINK-PLANILHA-LUD";"Execução!E4:E{FIM}")

Col H (tema PUBLICIDADE - Mari):
=IMPORTRANGE("LINK-PLANILHA-MARI";"Execução!F4:F{FIM}")

Col I (tema PUBLICIDADE - Lud):
=IMPORTRANGE("LINK-PLANILHA-LUD";"Execução!F4:F{FIM}")

Col J (observações - Mari):
=IMPORTRANGE("LINK-PLANILHA-MARI";"Execução!G4:G{FIM}")

Col K (observações - Lud):
=IMPORTRANGE("LINK-PLANILHA-LUD";"Execução!G4:G{FIM}")
```

> `{FIM}` = linha da última linha de dados (ex: 103 para 100 processos, 203 para 200)

---

## Desempate (Colunas L, M, N) — Aba Execução da Planilha Mãe

```
Coluna L — Desempate da pergunta principal:
=SE(OU(D{r}="";E{r}="");"Pendente";SE(D{r}=E{r};D{r};"DESEMPATAR"))

Coluna M — Desempate dos critérios institucionais:
=SE(OU(F{r}="";G{r}="");"Pendente";SE(F{r}=G{r};F{r};"DESEMPATAR"))

Coluna N — Desempate do tema:
=SE(OU(H{r}="";I{r}="");"Pendente";SE(H{r}=I{r};H{r};"DESEMPATAR"))
```

---

## IMPORTRANGE — Aba Desempenho do Agente

```
Coluna C (Moderação OQJ — desempate da pergunta):
=IMPORTRANGE("LINK-PLANILHA-MAE";"Execução!L4:L{FIM}")

Coluna F (Classificação OQJ — desempate do tema):
=IMPORTRANGE("LINK-PLANILHA-MAE";"Execução!N4:N{FIM}")
```

---

## Coluna E — Desempenho de Moderação (por tema)

### Direito de Família
```
=SE(OU(
  E(C{r}="SIM, é um processo de direito de familia";REGEXMATCH(D{r};"(?i)por isso o reporte é aprovado|aprovado"));
  E(C{r}="NÃO é um processo de direito de familia";REGEXMATCH(D{r};"(?i)reprovado"))
);"✅ ACERTO";"❌ ERRO")
```

### Saúde
```
=SE(OU(
  E(C{r}="SIM, o processo trata sobre questões de saúde";D{r}="aprovado");
  E(C{r}="NÃO é um processo que trata sobre questões de saúde";D{r}="reprovado")
);"✅ ACERTO";"❌ ERRO")
```

### Criminal / Penal
```
=SE(OU(
  E(C{r}="SIM, é um processo de direito criminal finalizado";REGEXMATCH(D{r};"(?i)aprovado|por isso o reporte é aprovado"));
  E(C{r}="NÃO é um processo de direito criminal finalizado";REGEXMATCH(D{r};"(?i)reprovado"))
);"✅ ACERTO";"❌ ERRO")
```

### ECA / Menor Infrator
```
=SE(OU(
  E(C{r}="SIM, é um processo que envolve ECA ou menor infrator";REGEXMATCH(D{r};"(?i)aprovado"));
  E(C{r}="NÃO é um processo que envolve ECA ou menor infrator";REGEXMATCH(D{r};"(?i)reprovado"))
);"✅ ACERTO";"❌ ERRO")
```

### Trabalhista
```
=SE(OU(
  E(C{r}="SIM, é um processo de direito trabalhista";REGEXMATCH(D{r};"(?i)aprovado"));
  E(C{r}="NÃO é um processo de direito trabalhista";REGEXMATCH(D{r};"(?i)reprovado"))
);"✅ ACERTO";"❌ ERRO")
```

### Segredo de Justiça
```
=SE(OU(
  E(C{r}="SIM, é um processo em segredo de justiça";REGEXMATCH(D{r};"(?i)aprovado"));
  E(C{r}="NÃO é um processo em segredo de justiça";REGEXMATCH(D{r};"(?i)reprovado"))
);"✅ ACERTO";"❌ ERRO")
```

---

## Coluna H — Desempenho de Classificação Temática (igual para todos os temas)

```
=SE(F{r}="Pendente";"Desconsiderado";SE(
  OU(
    E(CONT.SE(F{r};"*SAÚDE*");
      REGEXMATCH(MINÚSCULA(G{r});"doen[cç]a|erro m[eé]dico|interdi[cç][aã]o|curatela|estigmatizante|cid|psiqu|esquizof|hiv|aids|depress|autismo|tea|borderline|ansiedade|mental|droga|depend[êe]ncia|interna[cç][aã]o|hospitalar|m[eé]dic|cir[uú]rg|prontu|tratamento|medicamento|sa[uú]de|exame|incapacidade|aux[ií]lio-doen[cç]a|invalidez|bpc|loas|neoplas|c[aâ]ncer|tumor"));
    E(F{r}="TRABALHISTA";
      REGEXMATCH(MINÚSCULA(G{r});"trabalhista|trabalho|trt|verbas|fgts|horas extras|rescis[aã]o|clt|adicional|periculosidade|insalubridade|v[ií]nculo|empreg"));
    E(CONT.SE(F{r};"*PENAL*");
      REGEXMATCH(MINÚSCULA(G{r});"criminal|penal|furto|recepta[cç][aã]o|execu[cç][aã]o|pena|crimes|crime|inj[uú]ria|contraven[cç][aã]o|absolvi[cç][aã]o|extin[cç][aã]o|punibilidade|prescri[cç][aã]o|transitado|julgado|maria da penha|viol[êe]ncia dom[eé]stica|amea[çc]a|les[aã]o corporal|medida[s]? protetiva[s]?|inqu[eé]rito|delegacia|pris[aã]o|den[uú]ncia|sursis"));
    E(CONT.SE(F{r};"*FAMÍLIA*");
      REGEXMATCH(MINÚSCULA(G{r});"fam[ií]lia|pensa[oõ]|ado[cç][aã]o|guarda|casamento|alimentos|uni[aã]o est[aá]vel|div[oó]rcio|partilha|paternidade|filia[cç][aã]o"));
    E(CONT.SE(F{r};"*CRIMES SEXUAIS*");
      REGEXMATCH(MINÚSCULA(G{r});"sexual|[oó]dio|liberdade pessoal|honra|estupro|importuna[cç][aã]o|ass[eé]dio|vulner[aá]vel"));
    E(CONT.SE(F{r};"*MENOR*");
      REGEXMATCH(MINÚSCULA(G{r});"menor|eca|infrator|adolescente|crian[cç]a|ato infracional"));
    E(F{r}="SEGREDO DE JUSTIÇA";
      REGEXMATCH(MINÚSCULA(G{r});"segredo|sigilo|sigiloso|reprovado|n/a"))
  );
  "✅ ACERTO";
  "❌ ERRO"
))
```

---

## Notas de Implementação

- Substitua `{r}` pelo número da linha (ex: `3`, `4`, `5`...)
- Substitua `{FIM}` pela última linha de dados (ex: `103` para 100 processos; `203` para 200)
- `IMPORTRANGE` em bloco: escreva apenas na célula de início (ex: D4); o Google Sheets preenche o resto automaticamente
- Ao colar o IMPORTRANGE pela primeira vez, o Google Sheets pedirá para "Conceder acesso" — clique em permitir
- Se o Sheets reclamar de referência circular ao conectar a planilha Mãe consigo mesma, copie as URLs após fazer o upload
