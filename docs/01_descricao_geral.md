# Plano de trabalho - Validador e importador de planilhas de cursos FIC

**Aluno:** Kaio (kaiocardoso892@gmail.com) - Técnico Integrado em Informática

**Orientadores:** Prof. Bruno Gomes (@brunogomesifrn) e Entony (@franciscoentony) - NIC/IFRN

**Repositório:** https://github.com/nicifrn/mp_ific_importador_planilhas

**Período:** 16/10/2026 a 10/12/2026: 8 sprints de 1 semana (de sexta a quinta-feira). As Sprints 1 e 2 são para aprender Python; as Sprints 3 a 8 são de desenvolvimento.

**Projeto pai:** iFIC, plataforma de gestão dos Cursos de Formação Inicial e Continuada do IFRN (resumo em [`docs/projeto_ific.md`](projeto_ific.md)).

**Trabalho escrito:** TCC no modelo do IFRN, seguindo as normas de [`docs/orientacoes_escrita_tcc.md`](orientacoes_escrita_tcc.md).

> **Leia este documento inteiro antes de começar a Sprint 1.** Ele é a referência completa do seu trabalho: tudo o que você vai estudar, programar, pesquisar e escrever está aqui, sprint por sprint. As *issues* do GitHub são o resumo de cada semana e sempre apontam para a seção correspondente deste documento. Se uma *issue* e este documento parecerem dizer coisas diferentes, **vale este documento**.

---

## Sumário

- [1. Como usar este documento e o quadro de tarefas](#como-usar)
- [2. A ideia do trabalho](#ideia)
- [3. A planilha e as regras de validação](#regras)
- [4. Decisões já tomadas](#decisoes)
- [5. Cronograma](#cronograma)
- [6. Como trabalhar no repositório (fluxo do NIC)](#fluxo-git)
- [7. A estrutura de pastas que você vai construir](#estrutura)
- [8. Combinados que valem para todas as sprints](#combinados)
- [9. Como escrever o TCC](#escrita-geral)
- [Sprint 1 - Git, ambiente e Python (semana 1)](#sprint-1-dev)
- [Sprint 2 - Python (semana 2): funções, dicionários, arquivos e erros](#sprint-2-dev)
- [Sprint 3 - Leitura da planilha e conferência do cabeçalho](#sprint-3-dev)
- [Sprint 4 - As oito regras de validação](#sprint-4-dev)
- [Sprint 5 - Leitura de XLSX, relatório de erros e planilha limpa](#sprint-5-dev)
- [Sprint 6 - Massa de teste com gabarito e avaliação da detecção](#sprint-6-dev)
- [Sprint 7 - Tempo de processamento e importação para o SQLite](#sprint-7-dev)
- [Sprint 8 - README, reprodução do zero e fechamento](#sprint-8-dev)
- [10. Quadro-resumo](#quadro-resumo)
- [11. Riscos e planos B](#riscos)
- [12. Se sobrar tempo](#extras)
- [13. Referências mínimas do TCC](#referencias-minimas)

---

<a id="como-usar"></a>
## 1. Como usar este documento e o quadro de tarefas

### As *issues* e o quadro (kanban)

Eu cadastrei as suas tarefas como *issues* no repositório. Toda semana tem:

- **uma *issue* com a etiqueta `Desenvolvimento`**: o que você vai estudar ou programar;
- **uma *issue* com a etiqueta `Escrita`**: o que você vai pesquisar e escrever no TCC. Em seis semanas ela é a entrega de uma parte do texto. Em duas semanas (Sprints 1 e 6) não há entrega de capítulo, e a *issue* é uma tarefa menor: na Sprint 1, preparar o documento e ler; na Sprint 6, colocar no documento as tabelas com os números que você acabou de medir. Essas tarefas pequenas existem para o texto não ficar parado, porque **quem deixa os Resultados para o fim não entrega**.

As *issues* de uma semana estão agrupadas em um **marco** (*milestone*) com o nome e a data da sprint, por exemplo `Sprint 04 - Regras de validação (06/11 a 12/11)`. A página **Milestones** do repositório mostra quanto de cada sprint já foi concluído.

O **quadro do projeto** (aba *Projects* do repositório) tem quatro colunas:

| Coluna | Quando o card fica aqui |
|---|---|
| **A fazer** | a tarefa ainda não começou |
| **Em andamento** | você está trabalhando nela. **Mova o card assim que começar** |
| **Em revisão** | você terminou e avisou os orientadores |
| **Concluído** | os orientadores aprovaram e fecharam a *issue* |

### Como cada sprint está organizada neste documento

Cada sprint tem duas partes, **Desenvolvimento** e **Escrita**, e cada uma tem um endereço próprio (`#sprint-4-dev`, `#sprint-4-escrita`), que é para onde as *issues* apontam.

A parte de desenvolvimento tem sempre os mesmos blocos:

- **O que você vai entregar**: em uma frase, o que precisa estar funcionando no fim da semana;
- **Passo a passo**: a lista numerada do que fazer, na ordem. Siga na ordem, sem pular;
- **Como saber que deu certo**: o que você deve ver na tela ou no arquivo. Se não vir isso, ainda não terminou;
- **Entrega da sprint**: a lista de conferência que também está na *issue*;
- **Erros comuns**: os problemas que mais aparecem naquela semana e como resolver.

**Neste documento eu não te dou os programas prontos.** Eu te digo o que cada programa precisa fazer, em que ordem, quais funções usar e o que pesquisar. Quem escreve o código é você, e é isso que faz o trabalho ser seu. Os trechos de código que aparecem aqui são **peças** (uma chamada de função, um comando SQL, um esqueleto com comentários), não a solução montada.

### Como entregar

**Desenvolvimento:**

1. abra um *pull request* (PR) da sua *branch* para a `develop`, seguindo a [seção 6](#fluxo-git), com `Closes #N` na descrição (o número N de cada sprint está indicado na própria sprint);
2. comente na *issue*: `@brunogomesifrn @franciscoentony entrega da Sprint N pronta para revisão - PR #X`;
3. mova o card para **Em revisão**.

Quando um de nós aprovar e fizer o *merge* do PR, a *issue* fecha sozinha e o card vai para **Concluído**. Se pedirmos ajustes, faça os ajustes na mesma *branch* (o PR se atualiza sozinho) e comente de novo.

**Escrita:**

1. escreva no documento do Google Drive (veja a [seção 9](#escrita-geral));
2. nomeie a versão no histórico do Drive (*Arquivo → Histórico de versões → Nomear versão atual*), por exemplo `Entrega 1 - Introdução - 29/10`;
3. comente na *issue*: `@brunogomesifrn @franciscoentony Introdução pronta para revisão - versão "Entrega 1 - Introdução - 29/10"`;
4. mova o card para **Em revisão**.

Nós comentamos no próprio documento do Drive. Quando a seção estiver aprovada, fechamos a *issue*.

### Regras de ouro

- **Entregue no prazo, mesmo incompleto.** Trabalho pela metade dá para revisar; trabalho que não chega, não. Na *issue*, diga o que ficou faltando.
- **Terminou antes? Não comece a próxima sprint.** Use o tempo para entender melhor o que fez, refazer um exercício que ficou difícil e adiantar a escrita.
- **Travou mais de 40 minutos no mesmo erro? Pare e peça ajuda.** Comente na *issue* da semana, marcando nós dois, com: o que você estava tentando fazer, o comando que rodou e **a mensagem de erro inteira, copiada e colada** (não um resumo, nem um print cortado). Pedir ajuda depois de 40 minutos não é desistir, é administrar o tempo.

---

<a id="ideia"></a>
## 2. A ideia do trabalho

O IFRN oferece, em todos os seus *campi*, **Cursos de Formação Inicial e Continuada (FIC)**: cursos curtos de qualificação profissional, como Informática Básica, Eletricista Instalador Predial ou Cuidador de Idosos. O projeto **iFIC** está construindo uma plataforma para gerenciar esses cursos (leia [`docs/projeto_ific.md`](projeto_ific.md)).

Hoje, as informações desses cursos (nome, *campus*, carga horária, turno, vagas, datas) circulam em **planilhas preenchidas à mão**, por pessoas diferentes, em *campi* diferentes. Planilha preenchida à mão tem erro: campo em branco, carga horária escrita "cem", turno escrito "Noturnoo", código de curso repetido, curso que termina antes de começar, data que não existe, como 31/02. Se essas planilhas forem importadas para o iFIC sem conferência, os erros entram no sistema e aparecem para quem procura um curso.

O seu trabalho é construir e avaliar uma **ferramenta em Python que confere a planilha antes da importação**. Ela:

1. lê uma planilha de cursos FIC, em `.csv` ou `.xlsx`;
2. confere se o cabeçalho tem as colunas certas;
3. aplica **oito regras de validação** a cada linha;
4. gera um **relatório de erros**, dizendo a linha, a coluna, a regra violada e o valor encontrado;
5. gera uma **planilha limpa**, só com as linhas sem erro;
6. **importa** as linhas sem erro para um banco **SQLite**, sem duplicar nada quando a mesma planilha é importada de novo.

E, o mais importante para o TCC, você vai **provar com números** que ela funciona:

| O que é medido | Como |
|---|---|
| **Detecção de erros** | planilhas fictícias de 100, 1.000 e 10.000 linhas, geradas por um programa seu, com erros colocados de propósito e anotados em um **gabarito**. A ferramenta confere a planilha e você compara o relatório com o gabarito: quantos erros ela achou, quantos deixou passar e quantos "achou" onde não havia |
| **Tempo de processamento** | quanto tempo a ferramenta leva para conferir 100, 1.000 e 10.000 linhas, com 5 medições de cada |
| **Importação** | se o banco recebe só as linhas válidas e se importar duas vezes não duplica nada |

### As três perguntas do trabalho

Elas aparecem na Introdução, organizam os Resultados e são respondidas na Conclusão:

1. a ferramenta encontra **todos** os erros colocados nas planilhas de teste? E aponta erro onde **não** há?
2. quanto **tempo** ela leva para conferir 100, 1.000 e 10.000 linhas, e como esse tempo cresce?
3. a **importação** leva para o banco só as linhas válidas, sem duplicar registros quando a planilha é importada de novo?

### O que este trabalho **não** é

**Não é "fazer a importação dentro do iFIC".** A plataforma é o **contexto** que justifica o trabalho; o objeto do seu trabalho é a **ferramenta e a avaliação dela**, feitas fora da plataforma, no seu repositório. A integração ao iFIC é **trabalho futuro**, e isso vai escrito com todas as letras no TCC.

Guarde esta frase, porque ela vale para o código e para o texto:

> O seu trabalho **não é** "fiz um programa que confere planilha". O seu trabalho é **"construí uma ferramenta que confere planilhas de cursos FIC antes da importação e mostrei, com números, quantos erros ela encontra e em quanto tempo"**.

### Nada de dados reais

Todos os dados do trabalho são **fictícios**: você mesmo cria as planilhas de exemplo e um programa seu gera as planilhas de teste. Nenhuma planilha real de curso e nenhum dado de pessoa entram no repositório. Isso atende à Lei Geral de Proteção de Dados (LGPD, Lei nº 13.709/2018) e dispensa o Comitê de Ética em Pesquisa, porque não há participantes humanos.

---

<a id="regras"></a>
## 3. A planilha e as regras de validação

Esta seção é a **especificação** da ferramenta. Você vai voltar a ela em quase todas as sprints, e ela vai quase inteira para o capítulo 3 do TCC.

### 3.1 O formato da planilha

A planilha tem **oito colunas**, exatamente com estes nomes, na primeira linha (o cabeçalho). A ordem abaixo é a recomendada, mas a ferramenta aceita as colunas em outra ordem, porque lê cada valor pelo nome da coluna:

| Coluna | O que é | Exemplo |
|---|---|---|
| `codigo` | código do curso, único na planilha | `FIC-2027-00001` |
| `nome_curso` | nome do curso | `Informática Básica` |
| `campus` | *campus* do IFRN que oferta o curso | `Natal-Central` |
| `carga_horaria` | carga horária total, em horas | `160` |
| `turno` | turno das aulas | `noturno` |
| `vagas` | número de vagas | `30` |
| `data_inicio` | data de início, no formato `DD/MM/AAAA` | `01/03/2027` |
| `data_fim` | data de término, no formato `DD/MM/AAAA` | `30/06/2027` |

Os nomes das colunas são em minúsculas, sem acento e sem espaço, com `_` no lugar do espaço. Isso evita problema de codificação e facilita o código.

### 3.2 A conferência do cabeçalho

Antes de qualquer regra, a ferramenta confere o cabeçalho. Se faltar alguma das oito colunas, ou se houver uma coluna com nome diferente (por exemplo, `carga horaria` com espaço), **a ferramenta não valida nada**: ela mostra quais colunas faltam e quais sobram e para. Não faz sentido conferir linha por linha uma planilha com a estrutura errada.

### 3.3 As oito regras

Cada regra tem um código, que você vai usar **sempre**: no código, no relatório, nas tabelas e no texto do TCC.

| Código | Regra | Coluna(s) | Dimensão de qualidade de dados |
|---|---|---|---|
| **R01** | Nenhum campo pode estar vazio | todas | completude |
| **R02** | O `codigo` não pode se repetir na planilha | `codigo` | unicidade |
| **R03** | A `carga_horaria` tem de ser um número inteiro maior que zero | `carga_horaria` | validade |
| **R04** | As `vagas` têm de ser um número inteiro maior que zero | `vagas` | validade |
| **R05** | O `turno` tem de ser `matutino`, `vespertino`, `noturno` ou `integral` | `turno` | validade |
| **R06** | As datas têm de estar no formato `DD/MM/AAAA` e existir no calendário | `data_inicio`, `data_fim` | validade |
| **R07** | A `data_fim` tem de ser posterior à `data_inicio` | `data_fim` | consistência |
| **R08** | O `campus` tem de estar na lista de *campi* do IFRN | `campus` | validade |

A última coluna (a dimensão de qualidade de dados) é o elo entre a ferramenta e o Referencial Teórico: você vai estudar essas dimensões na Sprint 3 e explicar no TCC por que cada regra existe.

### 3.4 As decisões de validação

Algumas situações podem ser tratadas de mais de um jeito. Eu já decidi como a ferramenta se comporta em cada uma delas, para que os números da Sprint 6 façam sentido. **Siga exatamente estas decisões** e explique-as no TCC (seção 4.2):

1. **Número da linha:** o relatório usa o número que a pessoa vê quando abre a planilha no Excel. O cabeçalho é a linha 1; a primeira linha de dados é a **linha 2**.
2. **Campo vazio só gera R01.** Se a `carga_horaria` está vazia, a ferramenta registra R01 para essa coluna e **não** testa R03 nela. Sem isso, um único problema geraria dois erros.
3. **Espaços nas pontas não contam como conteúdo.** Um campo com `"   "` é vazio (R01). Um valor com espaço nas pontas, como `" 160 "`, é aceito como `160`.
4. **Turno e *campus* não diferenciam maiúsculas e minúsculas.** `Noturno`, `NOTURNO` e `noturno` são o mesmo turno. Mas `Noturnoo` e `noite` são erro (R05).
5. **Cada data inválida gera o seu próprio R06.** Se as duas datas forem inválidas, são dois erros R06, um em cada coluna.
6. **R07 só é testada quando as duas datas são válidas.** Não dá para comparar uma data com algo que não é data.
7. **Data de fim igual à de início é erro (R07).** "Posterior" quer dizer "depois", não "no mesmo dia".
8. **No código repetido (R02), a primeira ocorrência é válida.** O erro vai para a segunda ocorrência em diante, e a mensagem diz onde o código apareceu primeiro: `código já usado na linha 5`.
9. **Linha válida é a linha sem nenhum erro.** Só ela vai para a planilha limpa e para o banco.

### 3.5 Os *campi* do IFRN (regra R08)

Esta é a lista que você vai usar na regra R08:

```
Apodi, Caicó, Canguaretama, Ceará-Mirim, Currais Novos, Ipanguaçu,
João Câmara, Jucurutu, Lajes, Macau, Mossoró, Natal-Central,
Natal-Cidade Alta, Natal-Zona Leste, Natal-Zona Norte, Nova Cruz,
Parelhas, Parnamirim, Pau dos Ferros, Santa Cruz,
São Gonçalo do Amarante, São Paulo do Potengi
```

Na Sprint 3, você confere essa lista no portal do IFRN (https://portal.ifrn.edu.br, menu dos *campi*) e anota no diário a data da consulta. Se o portal mostrar algum *campus* a mais ou a menos, **ajuste a lista no código para ficar igual ao portal** e anote a diferença no diário. A lista vai para o TCC com a referência do portal e a data de acesso.

---

<a id="decisoes"></a>
## 4. Decisões já tomadas

Para você não perder tempo escolhendo ferramenta:

| Item | Decisão |
|---|---|
| Linguagem | **Python 3** |
| Onde você trabalha | **computadores do laboratório (Windows)**. Não é possível instalar programas, mas o `pip install` e o Git funcionam |
| Editor | **VS Code** |
| Terminal | **Prompt de Comando** (`cmd`), e não o PowerShell ([Sprint 1](#sprint-1-dev), passo 6) |
| Leitura de CSV | módulo `csv`, que já vem com o Python |
| Leitura de XLSX | biblioteca `openpyxl`, instalada com `pip` (é a **única** biblioteca que você instala) |
| Datas | módulo `datetime`, que já vem com o Python |
| Banco de dados | **SQLite**, pelo módulo `sqlite3`, que já vem com o Python. Não precisa instalar servidor, usuário nem senha |
| Planilhas de teste | módulo `random`, com semente fixa (`random.seed(42)`), para gerar sempre as mesmas planilhas |
| Medição de tempo | `time.perf_counter()` e o módulo `statistics` |
| Testes | um programa de verificação com o comando `assert` ([seção 6.5](#fluxo-git)) |
| Repositório | `nicifrn/mp_ific_importador_planilhas`, só seu, seguindo a política de desenvolvimento do NIC ([seção 6](#fluxo-git)) |
| Texto do TCC | **modelo de TCC do IFRN**, em um documento do Google Drive que eu compartilho com você |
| O que **não** entra | Django, MySQL, interface gráfica, página web, API, `pandas` e bibliotecas prontas de validação (Pandera, Great Expectations). Tudo isso aparece no TCC como alternativa ou como trabalho futuro |

### Por que SQLite e não MySQL?

O SQLite já vem com o Python e guarda o banco inteiro em um único arquivo. O MySQL exigiria instalar e configurar um servidor, o que não é possível nas máquinas do laboratório e não muda nada no que o trabalho quer mostrar: que só as linhas válidas chegam ao banco. Essa justificativa vai para a Metodologia.

### Por que não usar `pandas` ou uma biblioteca pronta de validação?

Por três motivos, que você também vai escrever no TCC: (1) o objetivo é que as regras de negócio dos cursos FIC fiquem **explícitas e fáceis de manter** por outros alunos do projeto; (2) a ferramenta fica com **uma única dependência externa** (`openpyxl`); (3) o `pandas` é uma biblioteca grande, que exigiria aprender outra forma de programar no meio de um trabalho de oito semanas. A comparação com essas ferramentas aparece nos Trabalhos relacionados (seção 2.4 do TCC).

### Uso de inteligência artificial

- **Pode:** usar uma ferramenta de IA para entender uma mensagem de erro ou um conceito, como faria com um colega.
- **Não pode:** pedir para a IA escrever os seus programas, os seus exercícios ou o texto do TCC. O código e o texto têm de ser seus, e você precisa saber explicar cada linha na apresentação.
- **Nunca:** colocar no TCC uma referência que você não abriu e leu. Ferramentas de IA inventam referências com naturalidade.

---

<a id="cronograma"></a>
## 5. Cronograma

Cada sprint vai de **sexta a quinta-feira**. O prazo de tudo é a **quinta-feira** que fecha a sprint.

| Sprint | Período | Desenvolvimento | Escrita |
|---|---|---|---|
| 1 | 16/10 a 22/10 | Git, ambiente e Python (semana 1) | Preparar o documento e ler para a Introdução |
| 2 | 23/10 a 29/10 | Python (semana 2): funções, dicionários, arquivos e erros | **Entrega 1:** Introdução e Objetivo geral |
| 3 | 30/10 a 05/11 | Leitura da planilha e conferência do cabeçalho | **Entrega 2:** Referencial 2.1 e 2.2 |
| 4 | 06/11 a 12/11 | As oito regras de validação | **Entrega 3:** Referencial 2.3 e 2.4 |
| 5 | 13/11 a 19/11 | Leitura de XLSX, relatório de erros e planilha limpa | **Entrega 4:** Metodologia |
| 6 | 20/11 a 26/11 | Massa de teste com gabarito e avaliação da detecção | Tarefa curta: colocar no documento as tabelas com os dados da Sprint 6 |
| 7 | 27/11 a 03/12 | Tempo de processamento e importação para o SQLite | **Entrega 5:** Resultados e discussão |
| 8 | 04/12 a 10/12 | README, reprodução do zero e fechamento | **Entrega 6:** Conclusão, Objetivos específicos, Resumo e Título |

**Tudo termina em 10/12.**

**Feriados:** 02/11 (segunda, Sprint 3) e 20/11 (sexta, primeiro dia da Sprint 6). Nenhum prazo cai em feriado, mas nessas duas semanas há **um dia útil a menos**: comece cedo.

**Três avisos sobre o cronograma:**

1. **As duas primeiras semanas são de estudo, mas têm entrega.** Os exercícios vão para o repositório, por PR, como qualquer código. É assim que você aprende o fluxo do Git antes de o projeto começar de verdade.
2. **Os Resultados vencem em 03/12, na mesma semana da medição de tempo.** Por isso a Sprint 6 já manda colocar as tabelas de detecção no documento, e a Sprint 7 manda terminar a medição de tempo **até segunda-feira, 30/11**: a terça, a quarta e a quinta ficam para escrever.
3. **A última semana é cheia.** A Conclusão precisa estar escrita até **segunda-feira, 07/12**, para sobrar a terça, a quarta e a quinta para o Resumo, o *Abstract*, o Título e a revisão final.

---

<a id="fluxo-git"></a>
## 6. Como trabalhar no repositório (fluxo do NIC)

O NIC tem uma política de desenvolvimento que vale para todos os projetos. Abaixo está o que ela significa no seu trabalho.

### 6.1 As *branches*

| *Branch* | Para que serve | Quem mexe |
|---|---|---|
| `main` | versão estável e aprovada | só os orientadores, no fim do trabalho |
| `develop` | onde as sprints aprovadas vão se juntando. É a *branch* padrão do repositório | ninguém envia direto: só entra por PR aprovado |
| `feature/...` | uma por sprint, onde você trabalha | você |

Os nomes das *branches* das sprints já estão definidos em cada sprint (por exemplo, `feature/regras-validacao`). Se precisar corrigir algo de uma sprint que já foi aprovada, use `fix/` (por exemplo, `fix/leitura-acentos`).

### 6.2 O ciclo de cada sprint

**No começo da sprint** (sexta-feira), crie a *branch* da semana a partir da `develop` atualizada:

```
git checkout develop
git pull
git checkout -b feature/nome-da-sprint
```

**Durante a semana**, envie o trabalho **ao fim de cada dia**. Código que só existe na máquina do laboratório não foi entregue, e a máquina pode ser formatada.

```
git status
git add .
git commit -m "feat: adiciona regra R03 de carga horária"
git push
```

Na primeira vez que der `git push` em uma *branch* nova, o Git pede `git push -u origin feature/nome-da-sprint`. Use exatamente o comando que ele mostrar.

**Antes do `git add .`, olhe sempre o `git status`.** Se aparecer `.venv/`, `saida/`, `dados/massas/` ou `__pycache__/`, o seu `.gitignore` está errado: corrija antes de continuar.

**No fim da sprint**, abra o PR (seção 6.4).

### 6.3 Mensagens de *commit*

Os *commits* devem ser **pequenos, objetivos e frequentes**: pelo menos 3 por semana. Use o padrão abaixo, com a mensagem em português e o verbo no presente:

```
feat: adiciona leitura de planilhas xlsx
fix: corrige número da linha no relatório de erros
docs: registra configuração da máquina
test: adiciona verificações da regra R06
refactor: separa as funções de leitura em um módulo
```

| Prefixo | Quando usar |
|---|---|
| `feat:` | algo novo funcionando (um exercício, uma regra, um programa) |
| `fix:` | correção de algo que estava errado |
| `docs:` | documentação, diário, leituras, `README.md`, prints, planilhas de resultado |
| `test:` | as verificações com `assert` |
| `refactor:` | reorganizar código sem mudar o que ele faz |

**Mensagens proibidas pela política do NIC:** `ajustes`, `mudanças`, `update`, `correções`. Se você não consegue dizer em uma frase o que o *commit* faz, ele está grande demais.

### 6.4 O *pull request* do fim da sprint

1. Confira se tudo foi enviado (`git status` deve dizer `nothing to commit`).
2. A partir da Sprint 3, rode as verificações (seção 6.5). Elas precisam terminar com a mensagem `Tudo certo`.
3. No GitHub, abra o repositório. Vai aparecer uma faixa amarela com o botão **Compare & pull request**. Se não aparecer: aba **Pull requests → New pull request**, com **base: `develop`** e **compare: a sua *branch***.
4. Preencha a descrição. O repositório já traz um modelo com os campos que a política do NIC exige:
   - **Objetivo:** o que esta sprint entrega, em uma ou duas frases;
   - **O que foi feito:** a lista do que foi implementado;
   - **Possíveis impactos:** o que muda para quem roda o projeto (um programa novo, uma regra nova, uma biblioteca nova);
   - **Evidências de teste:** a saída das verificações e dos programas (copiada do terminal) e os prints;
   - **Issue:** `Closes #N`, com o número indicado na sprint.
5. Em **Reviewers**, marque `brunogomesifrn` e `franciscoentony`.
6. Comente na *issue* e mova o card para **Em revisão** (veja a [seção 1](#como-usar)).

**Você não faz o *merge* do seu próprio PR.** A política do NIC exige que pelo menos um revisor aprove antes da integração. **Um PR por sprint**: não acumule duas sprints no mesmo PR.

### 6.5 Testes

A política do NIC pede que toda funcionalidade tenha testes compatíveis com a sua complexidade. No seu trabalho, isso significa **um programa de verificação**, `testes/verificar_validadores.py`, que começa na Sprint 3 e cresce a cada sprint. Ele confere sozinho se as funções fazem o que deveriam.

Ele usa só um comando do Python, o **`assert`**, que você aprende na Sprint 2:

```python
assert inteiro_positivo("160") == True, "160 deveria ser aceito"
```

Se a condição for **verdadeira**, nada acontece e o programa segue. Se for **falsa**, o programa para na hora e mostra `AssertionError: 160 deveria ser aceito`. Ou seja: se o programa chegar até a última linha e mostrar `Tudo certo`, todas as conferências passaram.

Rode sempre **a partir da raiz do repositório**, com o `.venv` ativo:

```
python testes\verificar_validadores.py
```

### 6.6 Segurança

Nunca envie para o repositório senhas, *tokens*, chaves ou arquivos `.env`. Neste trabalho você não precisa de nenhum deles: o SQLite não tem usuário nem senha.

---

<a id="estrutura"></a>
## 7. A estrutura de pastas que você vai construir

Você não precisa criar tudo de uma vez: cada sprint cria a sua parte. Esta é a figura final, para você saber aonde está indo. Tudo fica **na raiz do repositório** `mp_ific_importador_planilhas/`:

```
mp_ific_importador_planilhas/
├── .github/
│   └── pull_request_template.md   ← já existe (modelo do PR)
├── aprendizado/
│   ├── semana1/                   ← Sprint 1 (exercícios ex01 a ex08)
│   └── semana2/                   ← Sprint 2 (exercícios ex09 a ex17)
├── dados/
│   ├── cursos_exemplo.csv         ← Sprint 3 (10 linhas, todas válidas)
│   ├── cabecalho_errado.csv       ← Sprint 3
│   ├── cursos_com_erros.csv       ← Sprint 4
│   ├── cursos_com_erros.xlsx      ← Sprint 5
│   └── massas/                    ← Sprint 6 (gerada pelo programa; NÃO vai para o Git)
├── src/
│   ├── config.py                  ← Sprint 3 (colunas, turnos e campi)
│   ├── leitor.py                  ← Sprints 3 e 5 (leitura de CSV e XLSX)
│   ├── validadores.py             ← Sprints 3 e 4 (cabeçalho e as oito regras)
│   ├── relatorio.py               ← Sprint 5 (relatório e planilha limpa)
│   ├── converter_para_xlsx.py     ← Sprint 5
│   ├── gerador_massa.py           ← Sprint 6
│   ├── comparador.py              ← Sprint 6
│   ├── medir_tempo.py             ← Sprint 7
│   ├── banco.py                   ← Sprint 7
│   ├── ver_banco.py               ← Sprint 7
│   └── main.py                    ← Sprints 3 a 7 (o programa principal)
├── testes/
│   └── verificar_validadores.py   ← Sprints 3 a 7
├── resultados/                    ← planilhas com as medições (vão para o Git)
├── saida/                         ← relatórios, planilha limpa e banco (NÃO vão para o Git)
├── docs/
│   ├── 01_descricao_geral.md      ← este documento (já existe)
│   ├── projeto_ific.md            ← já existe
│   ├── orientacoes_escrita_tcc.md ← já existe
│   ├── leituras.md                ← fichamento das leituras (Sprint 1 em diante)
│   ├── diario.md                  ← Sprint 1 em diante
│   ├── ambiente.md                ← Sprint 1
│   └── evidencias/                ← prints de tela
├── .gitignore
├── requirements.txt
└── README.md
```

O programa principal é sempre o `src\main.py`. Os outros arquivos da pasta `src` são **módulos**: guardam funções que o `main.py` (e os outros programas) importam.

---

<a id="combinados"></a>
## 8. Combinados que valem para todas as sprints

- **Rode tudo a partir da raiz do repositório:** `python src\main.py dados\cursos_exemplo.csv`, e não de dentro da pasta `src`. Todos os caminhos (`dados\`, `saida\`, `resultados\`) partem da raiz.
- **Diário (`docs/diario.md`):** no fim de cada sprint, escreva 3 linhas: o que funcionou, o que deu errado e o que aprendi. Leva cinco minutos e vai escrever a sua Conclusão por você no fim do trabalho. Use este formato:

  ```
  ## Sprint 4 - 06/11 a 12/11
  - Funcionou: ...
  - Deu errado: ...
  - Aprendi: ...
  ```

- **Nunca apague uma medição.** A partir da Sprint 6, toda medição vai para `resultados/`. Se uma rodada saiu estranha, ela fica na planilha e você anota o motivo no diário.
- **Anote toda decisão e toda surpresa.** Um acento que virou símbolo estranho, uma data que o Excel devolveu em outro formato, um erro que a ferramenta não achou: tudo isso é conteúdo do TCC.
- **Prints (`docs/evidencias/`):** sempre que algo funcionar (ou falhar de um jeito interessante), tire um print (`Win + Shift + S`). Dê nomes que digam o que é: `s04_erros_cursos_com_erros.png`. Eles viram figuras do TCC.
- **Acompanhamento semanal:** no começo de cada sprint, você mostra aos orientadores o que fez na anterior, **rodando na tela**. A partir da Sprint 6, "está funcionando" não basta: tem que ter número.

---

<a id="escrita-geral"></a>
## 9. Como escrever o TCC

### 9.1 Onde escrever: o documento compartilhado no Google Drive

Na Sprint 1 eu compartilho com você um documento do Google Drive que **já está no modelo de TCC do IFRN** (capa, folha de rosto, resumo, sumário e as seções). Você não cria arquivo novo nem trabalha em cópia local.

- **Escreva sempre no documento compartilhado.** Nada de escrever no Word e colar depois: o histórico do Drive é o que nos permite ver o que mudou de uma semana para outra.
- **Não crie cópias** ("TCC v2", "TCC final"). O Drive já guarda tudo em **Arquivo → Histórico de versões**.
- **Marque cada entrega no histórico:** *Arquivo → Histórico de versões → Nomear versão atual*, com um rótulo como `Entrega 1 - Introdução - 29/10`.
- **Avise na *issue* quando entregar**, como explicado na [seção 1](#como-usar). Nós não ficamos olhando o documento: precisamos saber que há algo novo para ler.
- **Comentário se responde, não se apaga.** Responda a cada comentário dizendo o que você fez e só então marque como resolvido.
- **Resolva todos os comentários da entrega anterior antes de avisar sobre a próxima.**
- **Figuras e tabelas** vão dentro do documento, no lugar certo, com a identificação em cima e a fonte embaixo.
- **Não dê acesso ao documento a outras pessoas.**

### 9.2 As normas

Todas as regras de escrita e de ABNT estão em [`docs/orientacoes_escrita_tcc.md`](orientacoes_escrita_tcc.md): estrutura, citações, referências, figuras, tabelas, formatação e linguagem. **Leia esse arquivo inteiro na Sprint 1.** Este plano não repete as regras da ABNT; ele diz **o que pesquisar e o que escrever** em cada seção do seu trabalho.

> **Importante:** o que está neste plano são **orientações**, não o texto pronto. Não copie estas frases para o TCC. Elas dizem sobre o que é cada parágrafo; o texto tem de ser escrito por você, com as suas palavras e com as referências que você leu.

### 9.3 A estrutura do seu TCC

```
1 INTRODUÇÃO
1.1 Objetivo geral
1.2 Objetivos específicos
2 REFERENCIAL TEÓRICO
2.1 Cursos FIC e sistemas de informação na gestão educacional
2.2 Qualidade de dados
2.3 Validação e importação de dados
2.4 Trabalhos relacionados
3 METODOLOGIA
3.1 Classificação da pesquisa
3.2 Etapas do trabalho
3.3 Materiais e ambiente de execução
3.4 Formato da planilha e regras de validação
3.5 Arquitetura da ferramenta
3.6 Base de teste e procedimento de avaliação
4 RESULTADOS E DISCUSSÃO
4.1 A ferramenta desenvolvida
4.2 Decisões de validação
4.3 Detecção de erros
4.4 Tempo de processamento
4.5 Importação para o banco de dados
4.6 Limitações do trabalho
5 CONCLUSÃO
REFERÊNCIAS
APÊNDICE A - Exemplo de relatório de erros
```

Tamanho de referência da parte textual: **de 15 a 25 páginas**, mais ou menos assim: Introdução, 2; Referencial, 5 a 7; Metodologia, 4 a 6; Resultados, 4 a 7; Conclusão, 1 a 2.

### 9.4 Regras que valem para o texto inteiro

- **Toda afirmação técnica precisa de referência.**
- **Não empilhe citações.** Explique com as suas palavras, relacione os autores e amarre ao seu trabalho.
- **Documentação oficial** (Python, openpyxl, SQLite) serve para descrever a ferramenta, não para fundamentar conceito. Para conceito, use livro ou artigo.
- **Nada de blog, Stack Overflow ou site sem autoria** nas Referências. Você vai usar muito esse material para resolver problemas de programação (isso é normal), mas ele não entra na lista.
- **Nunca cite uma referência que você não leu.**
- **Crie a seção "Referências" no primeiro dia** e acrescente cada obra assim que ler.
- O Zotero e o Mendeley (gratuitos, funcionam no navegador) ajudam a montar as referências, mas confira cada entrada com a NBR 6023:2018: esses programas erram com frequência.
- **Meta:** de 12 a 18 referências (lista mínima na [seção 13](#referencias-minimas)).
- **Fichamento:** tudo o que você ler vai para `docs/leituras.md`, no formato abaixo. É ele que transforma leitura em texto:

  ```
  ## WANG; STRONG (1996)
  - Referência completa (ABNT): ...
  - Onde achei: (endereço)
  - Ideia principal, com as minhas palavras: ...
  - Trecho que posso citar (com página): "..." (p. X)
  - Vou usar na seção: 2.2
  ```

---

# Sprint 1: 16/10 a 22/10/2026

<a id="sprint-1-dev"></a>
## Sprint 1 · Desenvolvimento: Git, ambiente e Python (semana 1)

**Branch:** `feature/aprendizado-semana-1` · **No PR:** `Closes #1`

**O que você vai entregar:** a conta do GitHub com acesso ao repositório, o Git configurado, o repositório clonado, o ambiente virtual criado, o `.gitignore`, a descrição da máquina, o `README.md` inicial e **oito exercícios de Python** resolvidos na pasta `aprendizado/semana1/`.

**Por que começar estudando:** a ferramenta do trabalho usa variáveis, textos, condições, repetições e listas o tempo inteiro. Duas semanas de base bem feita economizam muitas semanas de travamento depois. Os exercícios já usam o assunto do trabalho (cursos, turnos, vagas), então nada do que você fizer aqui é perdido.

**Divida a semana assim:**

| Dia | O que fazer |
|---|---|
| sexta, 16/10 | passos 1 a 6 (conta, Git, repositório, ambiente e pastas) |
| segunda, 19/10 | passo 7 (estudo) e exercícios ex01 a ex03 |
| terça, 20/10 | passo 7 (estudo) e exercícios ex04 a ex06 |
| quarta, 21/10 | exercícios ex07 e ex08, passos 9 e 10 (`ambiente.md` e `README.md`) |
| quinta, 22/10 | passo 11 (diário, *commits* e PR) |

### Passo a passo

**1. Conta no GitHub.** Se você ainda não tem conta, crie uma em https://github.com/signup **com o e-mail `kaiocardoso892@gmail.com`**: foi para esse e-mail que eu mandei o convite do repositório. Escolha um nome de usuário simples (de preferência, o seu nome). Se você já tem conta com outro e-mail, adicione esse e-mail a ela em **Settings → Emails** e confirme pelo link que chega na caixa de entrada.

**2. Aceite os convites.** Na caixa de entrada do `kaiocardoso892@gmail.com` há um convite para o repositório `nicifrn/mp_ific_importador_planilhas`. Aceite-o (ou abra https://github.com/nicifrn/mp_ific_importador_planilhas/invitations). Aceite também o convite do quadro do projeto, que chega no mesmo e-mail. Depois, abra a *issue* #1 do repositório e comente `Convite aceito`: é assim que eu fico sabendo o seu nome de usuário para te colocar como responsável pelas *issues*.

**3. Se você nunca usou Git e GitHub,** assista antes às aulas iniciais do Curso de Git e GitHub do Curso em Vídeo (https://www.cursoemvideo.com/curso/curso-de-git-e-github/), até a parte em que o repositório é enviado para o GitHub. Depois leia a [seção 6](#fluxo-git) deste documento inteira.

**4. Configure o Git** na máquina do laboratório (uma vez só por máquina), com o mesmo e-mail da sua conta do GitHub. Abra o Prompt de Comando (tecle `Win`, digite `cmd`, Enter):

```
git config --global user.name "Seu Nome Completo"
git config --global user.email "kaiocardoso892@gmail.com"
```

**5. Clone o repositório e crie a *branch* da sprint:**

```
cd %USERPROFILE%
git clone https://github.com/nicifrn/mp_ific_importador_planilhas.git
cd mp_ific_importador_planilhas
git checkout develop
git checkout -b feature/aprendizado-semana-1
```

Na primeira vez que o Git falar com o GitHub, abre uma janela do navegador pedindo login. Entre com a sua conta e autorize. Depois abra a pasta `mp_ific_importador_planilhas` no VS Code (*File → Open Folder*). A partir de agora, **é sempre essa pasta que você abre**.

**6. Crie o ambiente virtual, as pastas e o `.gitignore`.** O ambiente virtual é uma "caixa" só deste projeto, onde as bibliotecas ficam instaladas sem bagunçar o Python da máquina. Na raiz do repositório, **no Prompt de Comando** (e não no PowerShell):

```
python -m venv .venv
.venv\Scripts\activate
```

Depois de ativar, o começo da linha passa a mostrar `(.venv)`. **Se não mostrar, o ambiente não está ativo.** Toda vez que abrir um terminal novo, ative de novo. Nesta semana você ainda não instala nenhuma biblioteca.

> No terminal do VS Code (`` Ctrl + ` ``), abra o seletor de terminal (setinha ao lado do `+`) e escolha **Command Prompt**. No PowerShell, o `activate` costuma dar erro de "execução de scripts desabilitada".
>
> Se o comando `python` não for reconhecido, tente `py`. Se o `py` funcionar, use `py` no lugar de `python` em todos os comandos deste documento, e anote isso no diário.

Crie as pastas `aprendizado\semana1` e `docs\evidencias`. *Pesquise:* o comando `mkdir` no Windows. O Git não guarda pastas vazias; por isso, coloque dentro de `docs\evidencias` um arquivo vazio chamado `.gitkeep`.

Crie o arquivo `.gitignore` na raiz, com este conteúdo. Ele diz ao Git o que **nunca** deve ir para o repositório (algumas dessas pastas só vão existir nas próximas sprints):

```
.venv/
__pycache__/
*.pyc
saida/
dados/massas/
*.db
```

**7. Estude.** Nesta semana, os assuntos são: `print` e `input`; variáveis e tipos (`str`, `int`, `float`, `bool`); conversão com `int()`; operadores aritméticos, inclusive `//` (divisão inteira) e `%` (resto); métodos de texto (`strip`, `lower`, `upper`, `isdigit`); f-strings; `if`, `elif` e `else`; `for` com `range`; `while` e `break`; listas (`append`, `len`, `in`, `sorted`).

Use duas fontes:

- **Curso em Vídeo, Curso de Python 3** (gratuito): https://www.cursoemvideo.com/curso/python-3-mundo-1/. Assista ao **Mundo 1 inteiro** e, no **Mundo 2**, às aulas de condições aninhadas, `for` e `while`. Os desafios do curso são opcionais: os obrigatórios são os meus, abaixo;
- **Tutorial oficial do Python, em português:** https://docs.python.org/pt-br/3/tutorial/, seções 3 (*Uma introdução informal ao Python*) e 4.1 a 4.4 (`if`, `for`, `range`, `break`). Use como consulta quando tiver dúvida de como um comando funciona.

Não tente assistir a tudo antes de começar os exercícios. Assista a um assunto, faça o exercício daquele assunto e siga.

**8. Faça os oito exercícios.** Cada exercício é um arquivo em `aprendizado\semana1\`, com o nome indicado. Rode cada um a partir da raiz do repositório, por exemplo `python aprendizado\semana1\ex01_apresentacao_curso.py`. **No começo de cada arquivo**, escreva um comentário com o enunciado em uma linha (`# ex01: lê nome, campus e carga horária e mostra uma frase`).

Para cada exercício, eu digo o que o programa faz, o que pesquisar e como testar. **Teste com todos os valores indicados** e confira se a saída é a esperada.

**ex01_apresentacao_curso.py**

- O programa pergunta o nome do curso, o *campus* e a carga horária, e mostra uma frase com os três.
- *Pesquise:* `input()`, `print()` e f-string.
- *Teste com:* `Informática Básica`, `Natal-Central`, `160`. Saída esperada: `O curso Informática Básica, do campus Natal-Central, tem 160 horas.`

**ex02_semanas.py**

- O programa pergunta a carga horária total e quantas horas de aula há por semana, e mostra quantas semanas completas o curso tem e quantas horas sobram para a última semana.
- *Pesquise:* `int()`, `//` e `%`.
- *Teste com:* `160` e `12`. Saída esperada: `13 semanas completas e 4 horas na última semana.` Teste também com `160` e `20` (8 semanas e 0 horas).

**ex03_turno.py**

- O programa pergunta o turno, limpa o texto (tira os espaços das pontas e passa para minúsculas) e diz se ele é válido. Os turnos válidos são `matutino`, `vespertino`, `noturno` e `integral`, guardados em uma lista.
- *Pesquise:* os métodos `strip()` e `lower()` e o operador `in` com listas.
- *Teste com:* `  Noturno ` (com espaços), que deve dar `Turno válido: noturno`; `NOITE`, que deve dar `Turno inválido: 'NOITE'`; e `integral`, que deve ser válido. Repare que a mensagem de erro mostra o valor **original**, como a pessoa digitou: guarde o texto original em uma variável e o texto limpo em outra.

**ex04_vagas.py**

- O programa pergunta o número de vagas, como texto, e responde:
  - se o texto não for um número inteiro: `Vagas precisa ser um número inteiro`;
  - se for zero: `Curso sem vagas`;
  - se for maior que 40: `Turma grande: confirme o número de vagas`;
  - nos outros casos: `OK: 30 vagas` (com o número digitado).
- *Pesquise:* o método `isdigit()` e a estrutura `if` / `elif` / `else`. Só converta o texto com `int()` **depois** de saber que ele é um número.
- *Teste com:* `30`, `0`, `55`, `trinta` e `-5`. Anote no diário o que aconteceu com `-5` e por quê (dica: o `-` é um dígito?). Você vai resolver esse problema de um jeito melhor na Sprint 2.

**ex05_carga_acumulada.py**

- O programa pergunta quantas horas de aula há por semana e quantas semanas o curso tem, e mostra as horas acumuladas semana a semana.
- *Pesquise:* `for` com `range(1, n + 1)`.
- *Teste com:* `4` e `5`. Saída esperada: cinco linhas, de `Semana 1: 4 horas acumuladas` até `Semana 5: 20 horas acumuladas`.

**ex06_lista_cursos.py**

- O programa pede nomes de cursos, um de cada vez, até a pessoa apertar Enter sem digitar nada. Depois mostra quantos cursos foram digitados e a lista em ordem alfabética, numerada (`1. Auxiliar Administrativo`, `2. Cuidador de Idosos`...).
- *Pesquise:* `while True` com `break`, o método `append()` e a função `sorted()`. Para numerar, use uma variável contadora dentro do `for`.
- *Teste com:* `Operador de Computador`, `Auxiliar Administrativo`, `Cuidador de Idosos` e Enter vazio. Teste também apertando Enter logo de cara: o programa deve dizer `Nenhum curso digitado` em vez de mostrar uma lista vazia.

**ex07_cargas.py**

- Escreva no próprio código a lista `cargas = [160, 40, 200, 80, 20, 160]`. Usando um `for` e **sem** usar `max()`, `min()` nem `sum()`, calcule a maior carga, a menor, a soma e a média. Depois, nas últimas linhas, confira os seus resultados com `max()`, `min()` e `sum()`.
- *Pesquise:* como guardar "o maior até agora" em uma variável dentro de um `for`.
- *Saída esperada:* maior 200, menor 20, soma 660 e média 110.0.

**ex08_conta_turnos.py**

- Escreva no código a lista:

  ```python
  turnos = ["Noturno", "noturno ", "MATUTINO", "vespertino", "Noite", "integral", " Matutino", "noturno"]
  ```

- O programa limpa cada turno (como no ex03), conta quantos há de cada turno válido e quantos são inválidos, e no fim mostra a contagem e a lista dos valores inválidos. Nesta semana, use **uma variável contadora para cada turno** (`qtd_noturno`, `qtd_matutino`...).
- *Saída esperada:* noturno 3, matutino 2, vespertino 1, integral 1 e 1 inválido (`'Noite'`).
- Este exercício é uma miniatura do seu trabalho: limpar o valor, comparar com uma lista fechada e separar o que está certo do que está errado. Na Sprint 2 você vai refazê-lo de um jeito mais curto.

**9. Preencha `docs/ambiente.md`** com a configuração da máquina do laboratório que você vai usar **sempre** (anote também o número ou a identificação dela). É nela que as medições de tempo da Sprint 7 serão feitas:

- processador e memória RAM: tecle `Win`, digite "Sobre o seu PC" e anote "Processador" e "RAM instalada";
- sistema operacional e versão: na mesma tela, "Edição" e "Versão";
- tipo de disco (SSD ou HD): abra o Gerenciador de Tarefas (`Ctrl + Shift + Esc`) → aba **Desempenho** → clique em **Disco 0**; o tipo aparece embaixo, à direita;
- versão do Python: `python --version`.

Isso vai para a Metodologia do TCC. Sem isso, os tempos medidos não significam nada.

**10. Escreva o `README.md`** do repositório (se o GitHub já tiver criado um, substitua o conteúdo). Por enquanto, quatro partes, que a política do NIC exige:

1. **o que é este projeto**, em três linhas;
2. **responsáveis:** você (aluno) e os dois orientadores, com os usuários do GitHub;
3. **instalação:** como criar e ativar o ambiente virtual;
4. **como rodar:** por enquanto, como rodar um exercício. A cada sprint você acrescenta os programas novos.

**11. Faça os *commits*, um por assunto, e abra o PR.**

1. Crie `docs/diario.md` com as 3 linhas da semana ([seção 8](#combinados)).
2. Faça os *commits* separados, por exemplo:

   ```
   git status
   git add .gitignore
   git commit -m "feat: adiciona gitignore do projeto"
   git add aprendizado
   git commit -m "feat: adiciona exercícios de Python da semana 1"
   git add docs README.md
   git commit -m "docs: adiciona ambiente, diário, leituras e README"
   git push -u origin feature/aprendizado-semana-1
   ```

   O ideal é fazer o *commit* dos exercícios **ao fim de cada dia**, e não todos na quinta.
3. Abra o PR para a `develop`, seguindo a [seção 6.4](#fluxo-git), com `Closes #1`. Em **Evidências de teste**, cole a saída do `ex07` e do `ex08`.

### Como saber que deu certo

- O terminal mostra `(.venv)` no começo da linha.
- Os oito exercícios rodam sem mensagem de erro (`Traceback`) com todos os valores de teste indicados e dão as saídas esperadas.
- O `git status` **não** mostra `.venv/`.
- O PR aparece na aba **Pull requests**, da sua *branch* para a `develop`.

### Entrega da sprint

- [ ] Convites aceitos e "Convite aceito" comentado na *issue* #1.
- [ ] Git configurado, repositório clonado e *branch* `feature/aprendizado-semana-1` criada.
- [ ] Ambiente virtual, `.gitignore` e pastas criados.
- [ ] Exercícios ex01 a ex08 em `aprendizado/semana1/`, testados com os valores indicados.
- [ ] `docs/ambiente.md` com processador, memória, disco, sistema e versão do Python.
- [ ] `README.md` com as quatro primeiras partes.
- [ ] `docs/diario.md` com a Sprint 1 (incluindo o que aconteceu com `-5` no ex04).
- [ ] PR aberto para a `develop`, com `Closes #1`.

### Erros comuns

- **`'python' não é reconhecido`:** use `py` (veja o passo 6).
- **O `activate` não funciona:** você está no PowerShell. Troque para o Command Prompt.
- **`TypeError: can only concatenate str (not "int") to str`:** você tentou juntar texto com número usando `+`. Use f-string.
- **`ValueError: invalid literal for int()`:** você chamou `int()` com um texto que não é número (por exemplo, `int("trinta")`). No ex04, teste com `isdigit()` antes de converter.
- **`IndentationError`:** o Python usa o recuo (os espaços no começo da linha) para saber o que está dentro do `if` ou do `for`. Use sempre 4 espaços; o VS Code faz isso sozinho quando você aperta Tab.

<a id="sprint-1-escrita"></a>
## Sprint 1 · Escrita: preparar o documento e ler para a Introdução

**O que você vai entregar:** o documento do TCC aberto e preparado, as normas lidas e as leituras da Introdução fichadas. Nesta semana não há entrega de capítulo.

### Passo a passo

**1. Abra o documento do TCC** que eu compartilhei com você no Google Drive (o convite chega no `kaiocardoso892@gmail.com`). Preencha a capa e a folha de rosto com o seu nome, o curso, o nome dos orientadores e o título provisório: *Validação e importação de planilhas de cursos de Formação Inicial e Continuada para a plataforma iFIC*.

**2. Crie a seção "REFERÊNCIAS"** no fim do documento, ainda vazia. A partir de hoje, toda obra que você ler entra nela no mesmo dia.

**3. Leia as normas.** Leia [`docs/orientacoes_escrita_tcc.md`](orientacoes_escrita_tcc.md) inteiro, com atenção especial a: estrutura do trabalho, Introdução, objetivo geral, como citar e como referenciar.

**4. Leia o resumo do projeto pai** em [`docs/projeto_ific.md`](projeto_ific.md). Ele é o ponto de partida da sua Introdução e já traz referências que você vai reaproveitar.

**5. Pesquise e leia, fichando em `docs/leituras.md`** (no formato da [seção 9.4](#escrita-geral); este arquivo entra no PR de desenvolvimento da semana):

- **as referências do projeto:** Dos Santos e Alliprandini (2023), Lopes, Bertini e Borelli (2023) e Laudon e Laudon (2022). As referências completas estão no fim de `docs/projeto_ific.md`. Procure os dois artigos pelo título no **Google Acadêmico** (`scholar.google.com.br`); os dois têm o PDF aberto. O livro de Laudon e Laudon tem edição em português, *Sistemas de informação gerenciais* (Pearson): procure na biblioteca do campus e use a edição que você tiver em mãos, com o ano dela. Leia o capítulo 1, sobre o que é um sistema de informação;
- **a Lei nº 11.892, de 29 de dezembro de 2008,** que criou os Institutos Federais: procure no portal da legislação federal (`planalto.gov.br`) e leia o art. 7º, que lista os objetivos dos Institutos Federais. Um deles fala exatamente dos cursos de formação inicial e continuada;
- **a Lei nº 9.394/1996 (LDB), art. 39,** e o **Decreto nº 5.154/2004, art. 3º:** são os textos que definem os cursos FIC. No `planalto.gov.br`, abra a versão **compilada** (com as alterações), porque os dois foram modificados depois. No decreto, anote o que ele diz sobre a carga horária mínima: isso vai justificar a regra R03 no TCC;
- **o Guia Pronatec de Cursos FIC,** do Ministério da Educação: procure pelo nome no portal do MEC (`gov.br/mec`) e use a edição mais recente que encontrar. Não precisa ler inteiro: leia a apresentação e veja como cada curso é descrito (nome, carga horária mínima, eixo). Você vai usar os nomes dos cursos desse guia na Sprint 6;
- **um dado concreto para a Introdução:** na **Plataforma Nilo Peçanha** (o portal oficial de estatísticas da Rede Federal; procure pelo nome), filtre pelo IFRN e pelo tipo de curso "Qualificação Profissional (FIC)" e anote o número de matrículas em cursos FIC no ano mais recente, com a data da consulta. Número dá peso imediato ao texto;
- **Raymond Panko, *What we know about spreadsheet errors*:** é o texto mais citado sobre erros em planilhas. Procure o título no Google Acadêmico; o autor mantém uma versão aberta. Fiche os números que ele apresenta sobre a frequência de erros em planilhas: eles sustentam a sua Problemática. Anote o ano da versão que você leu.

### Como saber que deu certo

- O documento do Drive tem capa, folha de rosto e a seção de Referências criadas.
- O `docs/leituras.md` tem pelo menos **8 obras fichadas**, cada uma com a referência completa no formato ABNT.
- Você tem o número de matrículas FIC do IFRN, com o ano e a data da consulta.

### Entrega

- [ ] Documento do Drive com capa, folha de rosto, título provisório e seção de Referências.
- [ ] `docs/orientacoes_escrita_tcc.md` e `docs/projeto_ific.md` lidos.
- [ ] Pelo menos 8 obras fichadas em `docs/leituras.md`, com referência ABNT.
- [ ] As obras lidas lançadas na seção de Referências do documento.

Para entregar, comente na *issue* dizendo quantas obras você fichou e mova o card para **Em revisão**.

---

# Sprint 2: 23/10 a 29/10/2026

<a id="sprint-2-dev"></a>
## Sprint 2 · Desenvolvimento: Python (semana 2): funções, dicionários, arquivos e erros

**Branch:** `feature/aprendizado-semana-2` · **No PR:** `Closes #3`

**O que você vai entregar:** nove exercícios em `aprendizado/semana2/`, que terminam em um **minivalidador**: um programa que lê um CSV de cursos e aponta as linhas com problema. É uma versão de brinquedo do seu trabalho, e na Sprint 3 você começa a versão de verdade.

**Divida a semana assim:**

| Dia | O que fazer |
|---|---|
| sexta, 23/10 | estudo de funções e `try`/`except`; ex09 e ex10 |
| segunda, 26/10 | estudo de dicionários; ex11 a ex13 |
| terça, 27/10 | estudo de arquivos, `csv` e módulos; ex14 e ex15 |
| quarta, 28/10 | ex16 (datas) e ex17 (`assert`) |
| quinta, 29/10 | diário, *commits* e PR |

### Passo a passo

**1. Atualize a `develop` e crie a *branch* da semana.** A Sprint 1 precisa ter sido integrada (*merge*) por nós. Se o seu PR da Sprint 1 ainda não foi integrado na sexta de manhã, avise na *issue* #1 e crie a *branch* a partir da `feature/aprendizado-semana-1`.

```
git checkout develop
git pull
git checkout -b feature/aprendizado-semana-2
```

Crie a pasta `aprendizado\semana2`.

**2. Estude.** Os assuntos desta semana são: funções (`def`, parâmetros, `return`); `try` / `except ValueError`; dicionários (criar, ler com `[]` e com `.get()`, alterar, percorrer com `.items()`); abrir arquivos com `open(..., encoding=...)` e `with`; o módulo `csv` e o `csv.DictReader`; `enumerate`; importar funções de outro arquivo (módulos); o módulo `datetime`; e o comando `assert`.

- **Curso em Vídeo, Mundo 3** (https://www.cursoemvideo.com/curso/python-3-mundo-3/): as aulas de listas (revisão), dicionários, funções, modularização e tratamento de erros e exceções;
- **Tutorial oficial:** seções 4.8 e 4.9 (funções), 5.5 (dicionários), 6 (módulos), 7.2 (leitura de arquivos) e 8.3 (tratamento de exceções);
- **Documentação oficial do módulo `csv`, em português:** https://docs.python.org/pt-br/3/library/csv.html. Leia a parte do `DictReader` e os exemplos.

**3. Faça os nove exercícios.** Mesmas regras da semana 1: um arquivo por exercício, comentário com o enunciado no começo, rodar a partir da raiz e testar com todos os valores indicados.

**ex09_inteiro_positivo.py**

- Escreva a função `inteiro_positivo(texto)`, que devolve `True` se o texto representa um número inteiro maior que zero e `False` em qualquer outro caso. Ela deve tirar os espaços das pontas, tentar converter com `int()` dentro de um `try` e, se der `ValueError`, devolver `False`.
- Depois da função, mostre na tela o resultado para cada um destes valores: `"160"` (True), `" 40 "` (True), `"0"` (False), `"-5"` (False), `"abc"` (False), `""` (False) e `"12.5"` (False).
- *Pesquise:* `def` e `return`; `try` / `except ValueError`.
- Compare com o ex04: agora o `-5` é tratado corretamente. Anote no diário por que o `try` resolve o que o `isdigit()` não resolvia.

**ex10_turno_valido.py**

- Escreva a função `turno_valido(texto, permitidos)`, que limpa o texto (espaços das pontas e minúsculas) e devolve `True` se ele estiver na lista `permitidos`. Teste com os mesmos valores do ex03.
- Repare que a lista entra como **parâmetro**: assim a mesma função serve para turnos, para *campi* ou para qualquer lista fechada. É exatamente o que você vai precisar na Sprint 4, para as regras R05 e R08.

**ex11_dicionario_curso.py**

- Crie um dicionário `curso` com as oito colunas da planilha ([seção 3.1](#regras)) como chaves e valores **em texto** (assim como o CSV entrega: `"160"`, e não `160`).
- Mostre cada chave e valor, um por linha, com `for chave, valor in curso.items()`.
- Mude o valor de `vagas` para `"35"` e mostre o dicionário de novo.
- Tente ler `curso["professor"]` e veja o erro que aparece. Depois troque por `curso.get("professor")` e veja o que aparece. Anote no diário a diferença: você vai usar o `.get()` na ferramenta para não quebrar quando faltar uma coluna.
- *Pesquise:* dicionários, `.items()` e `.get()`.

**ex12_conta_turnos_dict.py**

- Refaça o ex08 usando um **dicionário** para a contagem, em vez de uma variável por turno: a chave é o turno e o valor é a quantidade. Os inválidos continuam em uma lista.
- *Pesquise:* o padrão `contagem[chave] = contagem.get(chave, 0) + 1`.
- *Saída esperada:* a mesma do ex08. Repare como o código ficou menor e como ele continuaria do mesmo tamanho se houvesse 20 turnos.

**ex13_lista_de_cursos.py**

- Escreva no código uma **lista de dicionários**, um por curso, com estes quatro cursos (só as colunas `nome_curso`, `campus`, `carga_horaria` e `vagas`, em texto):

  | nome_curso | campus | carga_horaria | vagas |
  |---|---|---|---|
  | Informática Básica | Natal-Central | 160 | 30 |
  | Auxiliar Administrativo | Mossoró | 200 | 25 |
  | Cuidador de Idosos | Caicó | 160 | 40 |
  | Operador de Computador | Parnamirim | 240 | 20 |

- Mostre: os cursos com 30 vagas ou mais; o total de vagas; e o curso com a maior carga horária.
- *Saída esperada:* Informática Básica e Cuidador de Idosos; 115 vagas; Operador de Computador.
- Essa estrutura, a **lista de dicionários**, é exatamente o que a ferramenta vai usar para representar uma planilha: cada linha é um dicionário, e a planilha é a lista.

**ex14_ler_csv.py**

- Crie no VS Code o arquivo `aprendizado\semana2\cursos_mini.csv` com este conteúdo (copie exatamente):

  ```
  codigo,nome_curso,campus,carga_horaria,turno,vagas,data_inicio,data_fim
  FIC-2027-00001,Informática Básica,Natal-Central,160,noturno,30,01/03/2027,30/06/2027
  FIC-2027-00002,Auxiliar Administrativo,Mossoró,200,matutino,25,08/03/2027,15/07/2027
  FIC-2027-00003,Cuidador de Idosos,Caicó,cem,vespertino,40,01/04/2027,31/07/2027
  FIC-2027-00004,Operador de Computador,Parnamirim,240,noite,20,05/04/2027,30/08/2027
  FIC-2027-00005,Eletricista Instalador Predial,Currais Novos,200,integral,15,03/05/2027,30/09/2027
  ```

- O programa abre o arquivo, lê com `csv.DictReader` e mostra, para cada linha, `Linha 2: Informática Básica (Natal-Central)`. Antes das linhas, mostre o cabeçalho, que fica em `leitor.fieldnames`.
- Abra o arquivo assim: `open("aprendizado/semana2/cursos_mini.csv", encoding="utf-8-sig", newline="")`, dentro de um `with`. O caminho parte da raiz do repositório, que é de onde você roda o programa.
- Para numerar as linhas, use `for numero, linha in enumerate(leitor, start=2)`. **Por que começar em 2?** Porque a linha 1 do arquivo é o cabeçalho: a primeira linha de dados é a linha 2 quando a pessoa abre a planilha. A ferramenta inteira vai usar essa numeração ([seção 3.4](#regras), decisão 1).
- *Pesquise:* `with open(...)`, `csv.DictReader`, `fieldnames` e `enumerate`. *Pesquise também:* o que é o **BOM** (*byte order mark*) e por que o `utf-8-sig` é mais seguro que o `utf-8` para arquivos salvos pelo Excel. Anote no diário.

**ex15: o minivalidador (dois arquivos)**

- Crie `aprendizado\semana2\utilidades.py` com as funções `inteiro_positivo` (do ex09) e `turno_valido` (do ex10), **sem** os testes: só as funções e uma lista `TURNOS_PERMITIDOS`.
- Crie `aprendizado\semana2\ex15_mini_validador.py`, que importa as funções com `from utilidades import inteiro_positivo, turno_valido, TURNOS_PERMITIDOS`, lê o `cursos_mini.csv` (como no ex14) e, para cada linha, confere a `carga_horaria`, as `vagas` e o `turno`. Para cada problema, mostra uma linha assim:

  ```
  Linha 4: carga_horaria inválida ('cem')
  Linha 5: turno inválido ('noite')
  ```

  No fim, mostra `2 problemas encontrados em 5 linhas`.
- *Pesquise:* como importar funções de outro arquivo da mesma pasta (módulos).
- Rode com `python aprendizado\semana2\ex15_mini_validador.py`. O Python encontra o `utilidades.py` porque ele está na mesma pasta do programa que você rodou.

**ex16_datas.py**

- Escreva a função `converter_data(texto)`, que tenta transformar um texto no formato `DD/MM/AAAA` em data, usando `datetime.strptime(texto.strip(), "%d/%m/%Y").date()` dentro de um `try`. Se der certo, devolve a data; se der `ValueError`, devolve `None`.
- Mostre o resultado para: `"01/03/2027"` (uma data), `"31/02/2027"` (None), `"2027-03-01"` (None), `"1/3/2027"` e `"abc"` (None). Anote no diário o que aconteceu com `"1/3/2027"`.
- Depois, converta `"01/03/2027"` e `"30/06/2027"`, mostre se a primeira é menor que a segunda e quantos dias há entre elas. *Pesquise:* subtração de datas e o atributo `.days`. *Saída esperada:* `True` e `121 dias`.
- Repare no `"31/02/2027"`: está no formato certo, mas a data **não existe** (fevereiro não tem 31 dias). O `strptime` percebe isso sozinho. Guarde esse exemplo: ele vai aparecer no Referencial Teórico e nos Resultados.

**ex17_assert.py**

- Importe `inteiro_positivo` e `turno_valido` do `utilidades.py` e escreva **8 `assert`**: 4 para cada função, com valores válidos e inválidos (use os testes do ex09 e do ex10). No fim, `print("Tudo certo")`.

  ```python
  assert inteiro_positivo("160") == True, "160 deveria ser aceito"
  ```

- **Sabotagem proposital:** troque, de propósito, o resultado esperado de um dos `assert` (por exemplo, `== False` no `"160"`) e rode. Veja a mensagem `AssertionError`. Desfaça a sabotagem e rode de novo. Esse é o tipo de verificação que a política do NIC pede, e você vai usá-lo da Sprint 3 até o fim ([seção 6.5](#fluxo-git)).

**4. Faça os *commits* e abra o PR.** Diário da semana em `docs/diario.md`, *commits* por assunto (de preferência um por dia) e PR para a `develop` com `Closes #3`. Em **Evidências de teste**, cole a saída do `ex15` e do `ex17`.

### Como saber que deu certo

- Os nove exercícios rodam sem `Traceback` e dão as saídas esperadas.
- O `ex15` mostra exatamente os 2 problemas do `cursos_mini.csv`, com os números de linha 4 e 5.
- O `ex17` mostra `Tudo certo`, e a sabotagem gerou `AssertionError`.

### Entrega da sprint

- [ ] Exercícios ex09 a ex14, ex16 e ex17 em `aprendizado/semana2/`.
- [ ] `cursos_mini.csv`, `utilidades.py` e `ex15_mini_validador.py` funcionando juntos.
- [ ] `ex17` com 8 `assert` e a sabotagem testada.
- [ ] Diário com: `-5` no ex09, `[]` *versus* `.get()`, o BOM e o `"1/3/2027"`.
- [ ] PR aberto para a `develop`, com `Closes #3`.

### Erros comuns

- **`ModuleNotFoundError: No module named 'utilidades'`:** o `utilidades.py` não está na mesma pasta do `ex15`, ou o nome do arquivo está diferente do que está no `import`.
- **`FileNotFoundError`:** o caminho do CSV parte da pasta de onde você rodou o comando. Rode a partir da raiz do repositório.
- **Acentos estranhos (`InformÃ¡tica`):** o arquivo não está em UTF-8 ou faltou o `encoding` no `open`. No VS Code, confira a codificação no canto inferior direito.
- **`KeyError: 'nome_curso'`, mesmo com a coluna no arquivo:** o arquivo tem BOM e você abriu com `utf-8`. Use `utf-8-sig`.
- **A função não devolve nada (`None`):** faltou o `return`.

<a id="sprint-2-escrita"></a>
## Sprint 2 · Escrita: Introdução e Objetivo geral (Entrega 1)

**Prazo: 29/10/2026** · **Versão no Drive:** `Entrega 1 - Introdução - 29/10`

Escreva os **5 blocos** da Introdução, exatamente nesta ordem (é a sequência definida em `docs/orientacoes_escrita_tcc.md`), e o Objetivo geral. Cerca de 2 páginas no total. Use o seu fichamento da Sprint 1.

### Bloco 1: Contextualização (1 a 2 parágrafos)

**Escreva sobre:** o cenário do trabalho. Comece pelos cursos FIC: o que são, qual a base legal (LDB, Decreto nº 5.154/2004) e o papel dos Institutos Federais na oferta deles (Lei nº 11.892/2008). Traga o **número de matrículas FIC do IFRN** que você encontrou na Plataforma Nilo Peçanha. Em seguida, fale do papel dos sistemas de informação na gestão de instituições (Laudon e Laudon) e feche apresentando o iFIC como a plataforma que o IFRN está construindo para gerenciar esses cursos.

### Bloco 2: Problemática (1 a 2 parágrafos)

**Escreva sobre:** o problema concreto. As informações dos cursos FIC circulam em planilhas preenchidas à mão, por pessoas e *campi* diferentes. Dê exemplos dos erros típicos (campo vazio, carga horária em texto, turno escrito de formas diferentes, código repetido, data inexistente, curso que termina antes de começar). Use **Panko** para mostrar, com número, que erros em planilhas são frequentes. Diga a **consequência**: importar essas planilhas para o iFIC sem conferência leva os erros para dentro do sistema, onde são mais difíceis de corrigir, e quem procura um curso vê informação errada. Diga também por que a conferência manual não resolve: é lenta, cansativa e ela mesma sujeita a erro.

Feche o bloco com a pergunta que o trabalho responde, em uma frase, e as três perguntas menores da [seção 2](#ideia).

**Evite:** afirmar que "os erros são muito frequentes" sem fonte. Ou você cita a literatura, ou não afirma.

### Bloco 3: Caminho para a solução (1 a 2 parágrafos)

**Pesquise:** as formas de evitar ou encontrar erros em dados de planilha: a "validação de dados" do próprio Excel; a validação no formulário de entrada; a validação por programa antes da importação; e as restrições do banco de dados (`NOT NULL`, `PRIMARY KEY`). Pesquise também, pelo nome, três ferramentas prontas de validação em Python: **Pandera**, **Great Expectations** e **Frictionless Data**. Leia a página inicial da documentação de cada uma e anote em uma frase o que ela faz.

**Escreva sobre:** as alternativas, com uma vantagem e uma limitação de cada. Depois, justifique a escolha deste trabalho: uma ferramenta própria, em Python, com as regras de negócio dos cursos FIC escritas de forma explícita, com uma única dependência externa e pensada para ser levada depois ao iFIC. Use os argumentos da [seção 4](#decisoes).

**Cuidado:** não desqualifique as outras ferramentas. A justificativa é "esta se ajusta melhor ao nosso contexto", e não "as outras são ruins".

### Bloco 4: Apresentação da solução (1 parágrafo)

**Escreva sobre:** o que será desenvolvido, **no futuro** ("será desenvolvida"; na Sprint 8 você passa para o passado): a ferramenta que lê planilhas `.csv` e `.xlsx`, confere o cabeçalho, aplica oito regras, gera o relatório de erros com a localização de cada problema, gera a planilha limpa e importa as linhas válidas para um banco SQLite. Diga **como ela será avaliada**: com planilhas fictícias de 100, 1.000 e 10.000 linhas, geradas por um programa que coloca erros conhecidos e registra um gabarito; comparando o relatório com o gabarito; e medindo o tempo. Deixe explícito que **todos os dados são fictícios**, o que atende à LGPD e dispensa o Comitê de Ética em Pesquisa, já que não há participantes humanos.

### Bloco 5: Vínculo com o projeto e escopo (1 parágrafo)

**Escreva sobre:** o vínculo com o projeto iFIC, do IFRN, e a etapa atual dele (autenticação, área administrativa e gerenciamento de cursos). Explique que a ferramenta é um **protótipo independente**, feito para validar a abordagem antes de ser levado à plataforma, e que a **integração ao iFIC é trabalho futuro**, fora do escopo. Ser explícito sobre o escopo protege o trabalho: a banca não cobra o que você declarou que não faria.

Feche a Introdução com um parágrafo curto de organização do texto ("O capítulo 2 apresenta o referencial teórico; o capítulo 3 descreve a metodologia; [...]").

### 1.1 Objetivo geral

Uma **única frase**, com **uma única ação principal**, seguindo o modelo de `docs/orientacoes_escrita_tcc.md`:

> "O objetivo principal do presente trabalho consiste no desenvolvimento de [o quê], chamado de [nome], para [finalidade]."

- o verbo principal é **desenvolver**;
- o objeto nomeia as duas funções da ferramenta: **validação** e **importação** de planilhas de cursos FIC;
- a finalidade aparece: garantir a qualidade dos dados antes da carga na plataforma iFIC;
- **dê um nome curto à ferramenta.** Ele deixa o texto mais fácil de escrever e pode abrir o título no fim;
- **não** use verbos vagos ("conhecer", "estudar") nem empilhe ações ("desenvolver, avaliar e integrar"). As outras ações são objetivos **específicos**.

**Teste rápido:** tudo o que estiver no objetivo geral precisa aparecer nos Resultados. Se você não vai medir, não prometa.

Na seção **1.2 Objetivos específicos**, escreva apenas "a ser escrito na Sprint 8".

### Entrega

- [ ] Os cinco blocos, na ordem, com citações.
- [ ] A pergunta do trabalho e as três perguntas menores no fim do bloco 2.
- [ ] Objetivo geral em uma frase, com um verbo só, e a ferramenta com nome.
- [ ] Seção 1.2 com "a ser escrito na Sprint 8".
- [ ] Obras citadas lançadas nas Referências.

Nomeie a versão no Drive, comente na *issue* e mova o card para **Em revisão**.

---

# Sprint 3: 30/10 a 05/11/2026

<a id="sprint-3-dev"></a>
## Sprint 3 · Desenvolvimento: leitura da planilha e conferência do cabeçalho

**Branch:** `feature/leitura-csv` · **No PR:** `Closes #5`

**O que você vai entregar:** o começo da ferramenta de verdade: uma planilha de exemplo, o arquivo de configuração, a leitura do CSV, a conferência do cabeçalho, o programa principal mostrando um resumo e o programa de verificação com os primeiros `assert`.

**Divida a semana assim** (segunda, 02/11, é feriado):

| Dia | O que fazer |
|---|---|
| sexta, 30/10 | passos 1 a 3 (*branch*, planilha de exemplo, `config.py` e conferência dos *campi*) |
| terça, 03/11 | passos 4 e 5 (`leitor.py` e `verificar_colunas`) |
| quarta, 04/11 | passos 6 e 7 (`main.py` e a planilha com cabeçalho errado) |
| quinta, 05/11 | passos 8 e 9 (verificação, `README.md`, diário e PR) |

### Passo a passo

**1. Crie a *branch* e as pastas.**

```
git checkout develop
git pull
git checkout -b feature/leitura-csv
```

Crie as pastas `dados`, `src` e `testes`.

**2. Crie a planilha de exemplo `dados\cursos_exemplo.csv`**, com o cabeçalho e **10 linhas, todas válidas** (os erros entram na Sprint 4). Escreva direto no VS Code. As duas primeiras linhas são estas:

```
codigo,nome_curso,campus,carga_horaria,turno,vagas,data_inicio,data_fim
FIC-2027-00001,Informática Básica,Natal-Central,160,noturno,30,01/03/2027,30/06/2027
```

Para as outras nove: códigos de `FIC-2027-00002` a `FIC-2027-00010`, sem repetir; nomes de cursos que existem no Guia Pronatec (você leu na Sprint 1); *campi* diferentes, da lista da [seção 3.5](#regras); os quatro turnos aparecendo pelo menos uma vez; vagas e cargas horárias variadas; e a data de fim sempre **depois** da data de início. Confira cada linha contra as oito regras da [seção 3.3](#regras) antes de salvar.

**3. Crie `src\config.py`.** É o arquivo que guarda as listas fixas do projeto, para que nenhuma delas fique espalhada pelo código. Ele tem só três constantes (variáveis em maiúsculas, que não mudam durante a execução):

- `COLUNAS_ESPERADAS`: a lista com os oito nomes de colunas da [seção 3.1](#regras), na ordem;
- `TURNOS_PERMITIDOS`: a lista com os quatro turnos;
- `CAMPI_IFRN`: a lista com os *campi* da [seção 3.5](#regras), um por linha, entre aspas.

Antes de fechar a lista dos *campi*, **confira no portal do IFRN** (https://portal.ifrn.edu.br, menu dos *campi*). Se houver diferença, deixe a lista igual ao portal. Anote no diário a data da consulta e qualquer diferença encontrada.

*Pesquise:* "constantes em Python" e a convenção de nomes em maiúsculas (PEP 8).

**4. Crie `src\leitor.py` com a função `ler_csv(caminho)`.** Ela devolve **duas coisas**: a lista com os nomes das colunas do cabeçalho e a lista de linhas (cada linha é um dicionário). É o que você fez no ex14, transformado em função. O que ela precisa fazer, nesta ordem:

1. abrir o arquivo com `encoding="utf-8-sig"` e `newline=""`, dentro de um `with`;
2. criar o `csv.DictReader`;
3. pegar os nomes das colunas em `leitor.fieldnames`. **Atenção:** se o arquivo estiver vazio, o `fieldnames` vem como `None`; nesse caso, use uma lista vazia;
4. transformar o leitor em lista com `list(leitor)`;
5. devolver as duas coisas: `return colunas, linhas`.

Quem chama a função recebe as duas de uma vez: `colunas, linhas = ler_csv(caminho)`. *Pesquise:* "função que retorna dois valores Python" (tuplas).

> **Por que devolver as colunas separadas?** Porque uma planilha só com o cabeçalho e nenhuma linha também precisa ser conferida, e nela não existe "primeira linha" de onde tirar os nomes das colunas.

**5. Crie `src\validadores.py` com a função `verificar_colunas(colunas)`.** Ela compara as colunas do arquivo com `COLUNAS_ESPERADAS` (importada do `config.py`) e devolve duas listas:

- `faltando`: as colunas esperadas que **não** estão no arquivo;
- `sobrando`: as colunas do arquivo que **não** são esperadas.

Se as duas listas vierem vazias, o cabeçalho está certo. **A ordem das colunas não importa:** a ferramenta acessa cada valor pelo nome da coluna (`linha["vagas"]`), então uma planilha com as colunas em outra ordem é aceita.

*Pesquise:* *list comprehension*, no formato `[c for c in lista if condição]`. Cada uma das duas listas cabe em uma linha.

**6. Crie `src\main.py`, o programa principal.** Ele recebe o caminho da planilha na linha de comando:

```
python src\main.py dados\cursos_exemplo.csv
```

O que ele precisa fazer, nesta ordem:

1. importar `sys` e `os`, a função `ler_csv` do `leitor` e a função `verificar_colunas` do `validadores`;
2. se a pessoa não passou o caminho, mostrar `Uso: python src\main.py <planilha>` e encerrar;
3. se o arquivo não existe, mostrar `Arquivo não encontrado: <caminho>` e encerrar;
4. ler a planilha;
5. conferir o cabeçalho. Se houver colunas faltando ou sobrando, mostrar quais e encerrar **sem validar**;
6. se o cabeçalho estiver certo, mostrar o resumo.

*Pesquise:* `sys.argv` (a lista com o que foi digitado na linha de comando: `sys.argv[1]` é o primeiro argumento depois do nome do programa), `sys.exit(1)` e `os.path.exists`.

Com a planilha de exemplo, a saída deve ser:

```
Arquivo: dados\cursos_exemplo.csv
Colunas: codigo, nome_curso, campus, carga_horaria, turno, vagas, data_inicio, data_fim
Cabeçalho: OK
Linhas de dados: 10
```

> **Por que o `from leitor import ler_csv` funciona, se você roda o programa a partir da raiz?** Porque o Python sempre procura os módulos, primeiro, na pasta do programa que foi rodado (no caso, `src`). Por isso todos os módulos ficam juntos em `src`.

**7. Crie `dados\cabecalho_errado.csv`**, para testar o passo 5. Copie o cabeçalho e as duas primeiras linhas do `cursos_exemplo.csv` e estrague o cabeçalho de dois jeitos: troque `carga_horaria` por `carga horaria` (com espaço) e apague a coluna `vagas` (do cabeçalho **e** das duas linhas). A saída esperada é:

```
Arquivo: dados\cabecalho_errado.csv
Cabeçalho com problema. A validação não foi feita.
Colunas que faltam: carga_horaria, vagas
Colunas que não deveriam existir: carga horaria
```

Teste também o programa sem nenhum argumento e com um arquivo que não existe (`python src\main.py dados\nao_existe.csv`). Nenhum dos três casos pode mostrar `Traceback`.

**8. Comece o programa de verificação `testes\verificar_validadores.py`.** Ele cresce a cada sprint até o fim do trabalho. As primeiras linhas são sempre estas:

```python
import sys
sys.path.insert(0, "src")

from config import COLUNAS_ESPERADAS
from leitor import ler_csv
from validadores import verificar_colunas
```

O `sys.path.insert(0, "src")` diz ao Python para procurar os módulos também na pasta `src`. Sem ele, o programa da pasta `testes` não encontra o `validadores.py`. Por isso o programa também tem de ser rodado a partir da raiz.

Escreva os `assert` destes casos:

1. a lista `COLUNAS_ESPERADAS` completa dá `([], [])`;
2. as mesmas colunas em outra ordem (por exemplo, a lista invertida) também dão `([], [])`;
3. a lista sem `vagas` dá `(["vagas"], [])`;
4. a lista com `carga horaria` no lugar de `carga_horaria` dá `(["carga_horaria"], ["carga horaria"])`;
5. uma lista vazia dá as oito colunas em `faltando`;
6. o `cursos_exemplo.csv` lido com `ler_csv` tem 10 linhas e o cabeçalho certo.

Termine com `print("Tudo certo")` e rode com `python testes\verificar_validadores.py`.

**9. Atualize o `README.md`** com o comando do `main.py` e da verificação. Faça o diário, os *commits* e o PR para a `develop` com `Closes #5`. Em **Evidências de teste**, cole a saída da verificação e as saídas do `main.py` com o `cursos_exemplo.csv` e com o `cabecalho_errado.csv`.

### Como saber que deu certo

- `python src\main.py dados\cursos_exemplo.csv` mostra o resumo com `Cabeçalho: OK` e 10 linhas.
- Com o `cabecalho_errado.csv`, o programa mostra as duas colunas que faltam e a que sobra, e para.
- Sem argumento e com arquivo inexistente, o programa mostra uma mensagem clara, sem `Traceback`.
- `python testes\verificar_validadores.py` mostra `Tudo certo`.

### Entrega da sprint

- [ ] `dados/cursos_exemplo.csv` com 10 linhas válidas.
- [ ] `src/config.py` com as três listas; *campi* conferidos no portal e data anotada no diário.
- [ ] `src/leitor.py` com `ler_csv` devolvendo colunas e linhas.
- [ ] `src/validadores.py` com `verificar_colunas`.
- [ ] `src/main.py` com o resumo e as três mensagens de problema (sem argumento, arquivo inexistente, cabeçalho errado).
- [ ] `dados/cabecalho_errado.csv`.
- [ ] `testes/verificar_validadores.py` com os 6 casos, mostrando `Tudo certo`.
- [ ] `README.md` e diário atualizados.
- [ ] PR aberto para a `develop`, com `Closes #5`.

### Erros comuns

- **`ModuleNotFoundError: No module named 'config'`** na verificação: faltou o `sys.path.insert(0, "src")`, ou ele está **depois** dos `import`. Ele tem de vir antes.
- **`TypeError: cannot unpack non-iterable NoneType object`:** a função `ler_csv` não tem `return`, ou devolve só uma coisa.
- **O cabeçalho certo é recusado e aparece `﻿codigo`:** o arquivo tem BOM e foi aberto com `utf-8`. Use `utf-8-sig`.
- **O cabeçalho certo é recusado sem motivo aparente:** confira se não há espaço depois de alguma vírgula do cabeçalho (`codigo, nome_curso`). Para a ferramenta, `" nome_curso"` é outra coluna.

<a id="sprint-3-escrita"></a>
## Sprint 3 · Escrita: Referencial Teórico, seções 2.1 e 2.2 (Entrega 2)

**Prazo: 05/11/2026** · **Versão no Drive:** `Entrega 2 - Referencial 2.1 e 2.2 - 05/11`

**Assista antes de começar:** https://youtu.be/8Qztq1Q5vb0 (explicação sobre o Referencial Teórico).

O Referencial explica os conceitos que o leitor precisa para entender o resto do trabalho. Vá do mais geral para o mais específico e escreva **só sobre conceitos que aparecem de novo** na Metodologia ou nos Resultados. Comece a procurar as referências na **sexta**, não na quarta.

### 2.1 Cursos FIC e sistemas de informação na gestão educacional (cerca de 2 páginas)

**Pesquise:**

- os textos legais que você fichou na Sprint 1 (Lei nº 11.892/2008, LDB e Decreto nº 5.154/2004) e o Guia Pronatec de Cursos FIC;
- um ou dois artigos sobre a oferta de cursos FIC nos Institutos Federais. No Google Acadêmico, procure por `"formação inicial e continuada" "instituto federal"` e escolha trabalhos publicados em revista ou evento;
- o conceito de **sistema de informação** e a diferença entre **dado** e **informação**, em Laudon e Laudon (capítulo 1);
- um artigo sobre sistemas de informação na **gestão educacional** (procure por `"sistemas de informação" "gestão educacional"` no Google Acadêmico ou na SciELO).

**Escreva sobre:** o que são os cursos FIC, a carga horária mínima definida em norma, o público e o papel deles na qualificação profissional e na inclusão; a oferta no IFRN (retome o número da Plataforma Nilo Peçanha). Depois, o que é um sistema de informação, a diferença entre dado e informação e por que a gestão de uma instituição depende de dados corretos. Feche ligando os dois assuntos: a plataforma iFIC é um sistema de informação, e ela só vai produzir informação útil se os dados que entrarem nela estiverem corretos.

**Atenção:** esta seção existe também para **justificar as regras** da ferramenta. A carga horária mínima da norma explica por que a regra R03 existe. Escreva pensando nisso.

### 2.2 Qualidade de dados (cerca de 2 páginas)

**Pesquise:**

- **Wang e Strong (1996)**, *Beyond accuracy: what data quality means to data consumers*: o artigo clássico que define qualidade de dados como adequação ao uso e organiza as dimensões de qualidade;
- **Batini e Scannapieco**, *Data quality: concepts, methodologies and techniques* (Springer): livro com as definições de completude, consistência e exatidão;
- o documento da **DAMA UK**, *The six primary dimensions for data quality assessment* (2013), que define as seis dimensões mais usadas: completude, unicidade, pontualidade, validade, exatidão e consistência;
- a norma **ISO/IEC 25012**, que define um modelo de qualidade de dados (cite pela norma; não precisa ler o texto completo, que é pago: leia um artigo que a explique);
- **Panko**, que você já fichou, para os erros em planilhas.

**Escreva sobre:**

1. o que é qualidade de dados (a ideia de "adequação ao uso");
2. as **dimensões** que o seu trabalho usa: **completude**, **unicidade**, **validade** e **consistência**, cada uma com uma definição, uma referência e um exemplo tirado das planilhas de cursos FIC;
3. a ligação explícita entre as dimensões e as regras: termine a seção com um **quadro** (é quadro, porque o conteúdo é texto) que mostra cada dimensão e as regras da ferramenta que a verificam. Use a última coluna da tabela da [seção 3.3](#regras);
4. a dimensão que a ferramenta **não** verifica: a **exatidão**. Se a carga horária real de um curso é 200 horas e alguém digitou 160, o valor é válido, e nenhuma regra percebe. Explique isso agora: é uma limitação que vai voltar nos Resultados, e a banca valoriza quem a reconhece;
5. por que os problemas de qualidade aparecem tanto em dados digitados à mão em planilhas (Panko).

### Entrega

- [ ] Vídeo assistido.
- [ ] 2.1 com os cursos FIC (base legal, carga horária mínima, IFRN) e os sistemas de informação (dado e informação).
- [ ] 2.2 com a definição, as quatro dimensões com exemplos, o quadro dimensões × regras e a exatidão como limitação.
- [ ] Obras novas nas Referências e fichadas em `docs/leituras.md`.
- [ ] Comentários da Entrega 1 resolvidos.

Nomeie a versão no Drive, comente na *issue* e mova o card para **Em revisão**.

---

# Sprint 4: 06/11 a 12/11/2026

<a id="sprint-4-dev"></a>
## Sprint 4 · Desenvolvimento: as oito regras de validação

**Branch:** `feature/regras-validacao` · **No PR:** `Closes #7`

**O que você vai entregar:** as oito regras da [seção 3.3](#regras) implementadas, seguindo as decisões da [seção 3.4](#regras); o `main.py` mostrando cada erro com linha, coluna, regra, mensagem e valor; uma planilha com erros conhecidos; e a verificação conferindo que a ferramenta acha **exatamente** esses erros. Esta é a sprint mais importante do trabalho: no fim dela, a ferramenta já confere planilhas de verdade.

**Divida a semana assim:**

| Dia | O que fazer |
|---|---|
| sexta, 06/11 | passos 1 a 4 (*branch*, mensagens, `novo_erro` e as funções auxiliares) |
| segunda, 09/11 | passo 5 (`validar_linha`: R01, R03, R04, R05 e R08) |
| terça, 10/11 | passo 5 (R06 e R07) e passo 6 (R02) |
| quarta, 11/11 | passos 7 a 9 (`validar_planilha`, `main.py` e a planilha com erros) |
| quinta, 12/11 | passos 10 e 11 (verificação, diário e PR) |

### Passo a passo

**1. Crie a *branch*** `feature/regras-validacao` a partir da `develop` atualizada.

**2. Acrescente ao `config.py` as mensagens de erro**, em um dicionário `MENSAGENS`, em que a chave é o código da regra. Use exatamente estas, para que o relatório fique padronizado:

| Regra | Mensagem |
|---|---|
| R01 | `campo obrigatório vazio` |
| R02 | `código já usado na linha N` (o N é preenchido na hora; veja o passo 6) |
| R03 | `carga horária deve ser um número inteiro maior que zero` |
| R04 | `vagas deve ser um número inteiro maior que zero` |
| R05 | `turno fora da lista permitida` |
| R06 | `data inválida (use DD/MM/AAAA)` |
| R07 | `data de fim deve ser posterior à data de início` |
| R08 | `campus fora da lista de campi do IFRN` |

**3. Crie a função que representa um erro.** Cada erro é um dicionário com cinco chaves. Esta função eu te dou pronta, porque ela é só a "forma" do erro:

```python
def novo_erro(numero_linha, coluna, regra, mensagem, valor):
    return {
        "linha": numero_linha,
        "coluna": coluna,
        "regra": regra,
        "mensagem": mensagem,
        "valor": valor,
    }
```

Representar o erro como dicionário facilita tudo o que vem depois: mostrar na tela agora, gravar em CSV na Sprint 5 e comparar com o gabarito na Sprint 6.

**4. Escreva as quatro funções auxiliares** no `validadores.py`. Todas recebem o valor como veio da planilha (texto ou `None`) e **nunca** podem quebrar:

| Função | Devolve | De onde vem |
|---|---|---|
| `esta_vazio(valor)` | `True` se o valor é `None` ou só tem espaços | novo |
| `inteiro_positivo(valor)` | `True` se é inteiro maior que zero | ex09 (acrescente o tratamento do `None`) |
| `valor_em_lista(valor, lista)` | `True` se o valor limpo está na lista | ex10 (veja a atenção abaixo) |
| `converter_data(valor)` | a data, ou `None` se não for data válida | ex16 (acrescente o tratamento do `None`) |

**Atenção no `valor_em_lista`:** a lista de *campi* tem maiúsculas (`Natal-Central`). Se você comparar `"natal-central"` com a lista original, não vai achar. Passe **os dois lados** para minúsculas: o valor e cada item da lista. *Pesquise:* como montar com *list comprehension* uma lista nova com todos os itens em minúsculas.

**Atenção no `None`:** o `None` não tem `.strip()`, e chamar `None.strip()` dá `AttributeError`. Teste o `None` antes, ou acrescente `AttributeError` ao `except`. *Pesquise:* `except (ValueError, AttributeError):`.

**5. Escreva `validar_linha(linha, numero_linha)`.** Ela recebe uma linha (dicionário) e o número dela, e devolve a **lista de erros** daquela linha (lista vazia se a linha estiver certa). Siga esta ordem, que implementa as decisões da [seção 3.4](#regras):

```python
def validar_linha(linha, numero_linha):
    erros = []
    vazias = []   # colunas vazias nesta linha

    # 1. R01: para cada coluna de COLUNAS_ESPERADAS, se estiver vazia,
    #    registre o erro R01 e guarde a coluna em "vazias"

    # 2. R03 (carga_horaria) e R04 (vagas): só teste se a coluna NÃO está em "vazias"

    # 3. R05 (turno) e R08 (campus): só teste se a coluna NÃO está em "vazias"

    # 4. R06: para data_inicio e para data_fim, se não estiver vazia,
    #    tente converter; se der None, registre R06 naquela coluna.
    #    Guarde as datas convertidas em duas variáveis

    # 5. R07: só se as DUAS datas foram convertidas com sucesso,
    #    registre R07 (na coluna data_fim) se a data de fim não for maior que a de início

    return erros
```

Três cuidados:

- leia os valores com `linha.get(coluna)`, e não com `linha[coluna]` (você viu a diferença no ex11);
- no erro, o campo `valor` guarda o valor **original**, como está na planilha, sem limpar: é ele que a pessoa vai procurar no Excel;
- a mensagem vem do `MENSAGENS` do `config.py`.

**6. Escreva `validar_codigos_duplicados(linhas)`** para a regra R02. Ela é diferente das outras, porque não dá para saber se um código é repetido olhando uma linha só: é preciso olhar a planilha inteira. O caminho:

1. crie um dicionário vazio `vistos`, que vai guardar `código → número da linha onde ele apareceu primeiro`;
2. percorra as linhas com `enumerate(linhas, start=2)`;
3. limpe o código com `strip()`. Se estiver vazio, pule a linha (o R01 já cuidou dela);
4. se o código **já está** em `vistos`, registre o erro R02 com a mensagem `código já usado na linha N`, em que N é `vistos[codigo]` (use f-string);
5. se **não está**, guarde: `vistos[codigo] = numero_linha`.

*Pesquise:* por que procurar uma chave em um dicionário é muito mais rápido do que procurar um item em uma lista. Anote no diário: esse detalhe explica parte do tempo de processamento que você vai medir na Sprint 7.

**7. Escreva `validar_planilha(linhas)`**, que junta tudo: percorre as linhas com `enumerate(..., start=2)`, junta os erros de `validar_linha` de cada uma, acrescenta os erros de `validar_codigos_duplicados` e devolve a lista inteira **ordenada por linha e, dentro da linha, por regra**:

```python
erros.sort(key=lambda e: (e["linha"], e["regra"]))
```

*Pesquise:* o parâmetro `key` do `sort` e o que é uma função `lambda`. Você precisa saber explicar essa linha.

**8. Atualize o `main.py`.** Depois do cabeçalho OK, ele chama `validar_planilha`, mostra cada erro em uma linha neste formato:

```
Linha 4 | carga_horaria | R03 | carga horária deve ser um número inteiro maior que zero | valor: 'cem'
```

e termina com o resumo:

```
Linhas de dados: 15 | Linhas válidas: 3 | Linhas com erro: 12 | Total de erros: 13
```

Para contar as **linhas com erro**, não conte os erros (uma linha pode ter dois): conte os números de linha **diferentes**. *Pesquise:* `set` em Python e como montar um conjunto com os números de linha dos erros.

**9. Crie `dados\cursos_com_erros.csv`**, com o cabeçalho e **15 linhas de dados** (linhas 2 a 16), cada uma com o erro indicado abaixo. Comece copiando as linhas do `cursos_exemplo.csv` e estrague só o campo indicado. Esta tabela é o **gabarito** desta planilha: você sabe, antes de rodar, exatamente o que a ferramenta tem de encontrar.

| Linha | O que fazer | Erro esperado |
|---|---|---|
| 2 | linha válida | nenhum |
| 3 | `nome_curso` vazio | R01 |
| 4 | `carga_horaria` = `cem` | R03 |
| 5 | `vagas` = `0` | R04 |
| 6 | `turno` = `noite` | R05 |
| 7 | `data_inicio` = `31/02/2027` | R06 (e **não** R07) |
| 8 | `data_fim` dez dias **antes** da `data_inicio` | R07 |
| 9 | `campus` = `Natal` | R08 |
| 10 | `codigo` igual ao da linha 2 | R02 (a linha 2 continua válida) |
| 11 | linha válida | nenhum |
| 12 | `carga_horaria` = `-40` **e** `turno` = `Noturnoo` | R03 e R05 |
| 13 | `carga_horaria` vazia | R01 (e **não** R03) |
| 14 | `data_fim` = `2027-06-30` | R06 (e **não** R07) |
| 15 | `turno` = `  Noturno ` (com espaços) e `campus` = `natal-central` | nenhum (decisões 3 e 4) |
| 16 | `data_fim` **igual** à `data_inicio` | R07 (decisão 7) |

Total esperado: **13 erros em 12 linhas**, e **3 linhas válidas** (2, 11 e 15). Por regra: R01 = 2, R02 = 1, R03 = 2, R04 = 1, R05 = 2, R06 = 2, R07 = 2, R08 = 1.

Rode `python src\main.py dados\cursos_com_erros.csv` e confira, linha por linha, com a tabela. Se algo não bater, **o erro está no código, não na tabela**: volte às decisões da [seção 3.4](#regras).

**10. Aumente a verificação** (`testes\verificar_validadores.py`). Acrescente os `assert`:

1. de cada função auxiliar, com pelo menos um caso válido e um inválido, incluindo `None` e texto só com espaços;
2. `converter_data("31/02/2027")` dá `None`;
3. `validar_linha` de uma linha válida dá lista vazia;
4. `validar_linha` de uma linha com `carga_horaria` vazia dá **exatamente um** erro, com regra R01 (decisão 2);
5. `validar_linha` de uma linha com data de fim igual à de início dá um erro R07 (decisão 7);
6. o `cursos_com_erros.csv` dá exatamente 13 erros, e o conjunto de pares `(linha, regra)` encontrado é **igual** ao esperado pela tabela do passo 9. Escreva o conjunto esperado no próprio teste:

   ```python
   esperado = {(3, "R01"), (4, "R03"), ...}   # complete com os 13 pares
   encontrado = {(e["linha"], e["regra"]) for e in erros}
   assert encontrado == esperado, f"diferença: {encontrado ^ esperado}"
   ```

   *Pesquise:* o operador `^` entre conjuntos (diferença simétrica). Quando o teste falha, ele mostra exatamente os pares que sobraram ou faltaram.

**11. Finalize.** Acrescente ao `README.md` a tabela das oito regras. Tire um print da saída do `main.py` com o `cursos_com_erros.csv` (`docs/evidencias/s04_erros_cursos_com_erros.png`). Diário, *commits* (um por grupo de regras é um bom tamanho) e PR para a `develop` com `Closes #7`. Em **Evidências de teste**, cole a saída da verificação e a do `main.py`.

### Como saber que deu certo

- `python src\main.py dados\cursos_com_erros.csv` mostra 13 erros, um por linha, ordenados, e o resumo `Linhas de dados: 15 | Linhas válidas: 3 | Linhas com erro: 12 | Total de erros: 13`.
- A linha 13 tem só R01; as linhas 7 e 14 têm só R06; a linha 15 não aparece.
- `python src\main.py dados\cursos_exemplo.csv` mostra 0 erros.
- `python testes\verificar_validadores.py` mostra `Tudo certo`.

### Entrega da sprint

- [ ] `MENSAGENS` no `config.py` e `novo_erro` no `validadores.py`.
- [ ] As quatro funções auxiliares, sem quebrar com `None`.
- [ ] `validar_linha` com R01, R03, R04, R05, R06, R07 e R08, seguindo as decisões.
- [ ] `validar_codigos_duplicados` com R02, dizendo a linha da primeira ocorrência.
- [ ] `validar_planilha` com os erros ordenados.
- [ ] `main.py` mostrando os erros e o resumo.
- [ ] `dados/cursos_com_erros.csv` com as 15 linhas da tabela.
- [ ] Verificação com os novos casos, inclusive a comparação com os 13 pares esperados, mostrando `Tudo certo`.
- [ ] Print em `docs/evidencias/`, `README.md` com a tabela das regras e diário.
- [ ] PR aberto para a `develop`, com `Closes #7`.

### Erros comuns

- **`AttributeError: 'NoneType' object has no attribute 'strip'`:** alguma função auxiliar recebeu `None` e não tratou (passo 4).
- **`TypeError: '<' not supported between instances of 'NoneType' and 'datetime.date'`:** você comparou as datas sem conferir se as duas foram convertidas (passo 5, item 5).
- **A linha 13 aparece com R01 e R03:** a regra R03 não está conferindo a lista `vazias`.
- **A linha 15 aparece com erro:** o `valor_em_lista` não está limpando os espaços ou não está passando a lista para minúsculas.
- **A linha 2 aparece com R02:** o erro do código repetido tem de ir para a **segunda** ocorrência (decisão 8).

<a id="sprint-4-escrita"></a>
## Sprint 4 · Escrita: Referencial Teórico, seções 2.3 e 2.4 (Entrega 3)

**Prazo: 12/11/2026** · **Versão no Drive:** `Entrega 3 - Referencial 2.3 e 2.4 - 12/11`

### 2.3 Validação e importação de dados (cerca de 2 páginas)

**Pesquise:**

- **ETL** (*Extract, Transform, Load*): Kimball e Caserta, *The data warehouse ETL toolkit* (Wiley, 2004), capítulo sobre limpeza e conformidade de dados;
- **limpeza de dados:** Rahm e Do (2000), *Data cleaning: problems and current approaches* (IEEE Data Engineering Bulletin), que classifica os tipos de problema em dados;
- **o formato CSV:** a RFC 4180, de 2005, que descreve o formato (procure "RFC 4180" no buscador; é um documento público da IETF);
- **o formato XLSX:** o padrão Office Open XML (ISO/IEC 29500). Basta entender e explicar que o `.xlsx` é um pacote compactado de arquivos XML;
- **restrições de integridade em bancos de dados** (`NOT NULL`, `PRIMARY KEY`): Elmasri e Navathe, *Sistemas de banco de dados* (Pearson), capítulo sobre o modelo relacional. Procure na biblioteca do campus.

**Escreva sobre:**

1. o que é o processo de ETL e em que ponto dele acontece a validação: **antes da carga**;
2. a diferença entre validação **sintática** (o valor está no formato certo?) e **semântica** (o valor faz sentido?). O seu melhor exemplo é o `31/02/2027`: está no formato `DD/MM/AAAA` e mesmo assim é inválido. A regra R07 (fim antes do início) é outro exemplo de validação semântica;
3. as camadas em que a validação pode acontecer (na planilha, no formulário, em um programa antes da importação, no banco de dados), e por que o seu trabalho valida **antes** da carga, com o banco ainda protegendo com `PRIMARY KEY`: são duas barreiras;
4. o que são os formatos CSV e XLSX, e por que a ferramenta aceita os dois (as pessoas preenchem no Excel, que salva em `.xlsx`, mas muitos sistemas exportam `.csv`).

### 2.4 Trabalhos relacionados (cerca de 1,5 página)

**Pesquise** trabalhos (TCCs, artigos, dissertações) que tenham desenvolvido validadores de planilhas, importadores em lote ou rotinas de limpeza de dados, de preferência no contexto educacional ou de gestão pública. Onde procurar: repositório institucional do IFRN (Memoria, `memoria.ifrn.edu.br`), Biblioteca Digital Brasileira de Teses e Dissertações (BDTD, `bdtd.ibict.br`), SBC OpenLib (`sol.sbc.org.br`) e Google Acadêmico. Termos úteis: `"validação de planilhas"`, `"importação de dados" planilha`, `"qualidade de dados" educação`, `"spreadsheet validation"`, `"data validation" tool`.

**Escreva sobre:** de 3 a 5 trabalhos ou ferramentas, um parágrafo cada: o que fazem, como foram avaliados e qual a limitação deles diante do seu problema. Inclua **pelo menos dois trabalhos acadêmicos** e **pelo menos uma das ferramentas prontas** que você levantou na Introdução (Pandera, Great Expectations ou Frictionless Data).

Feche com um **quadro comparativo** entre os trabalhos encontrados e o seu, com estes critérios: formatos de entrada aceitos; regras específicas dos cursos FIC; relatório com a linha e a coluna do erro; geração de planilha limpa; importação para banco; avaliação com gabarito de erros; e dependências ou custo.

**Se uma ferramenta pronta fizer quase tudo o que a sua faz**, isso não invalida o trabalho. Escreva com honestidade o que ela já faz e mostre o que fica a seu favor: as regras dos cursos FIC escritas de forma explícita, o relatório em português pensado para quem preenche a planilha, a avaliação com gabarito e a integração prevista com o iFIC.

### Entrega

- [ ] 2.3 com ETL, validação sintática e semântica (com o 31/02), camadas de validação e os formatos CSV e XLSX.
- [ ] 2.4 com 3 a 5 trabalhos ou ferramentas, sendo pelo menos dois acadêmicos.
- [ ] Quadro comparativo com os critérios indicados.
- [ ] Obras novas nas Referências e fichadas.
- [ ] Comentários da Entrega 2 resolvidos.

Nomeie a versão no Drive, comente na *issue* e mova o card para **Em revisão**.

---

# Sprint 5: 13/11 a 19/11/2026

<a id="sprint-5-dev"></a>
## Sprint 5 · Desenvolvimento: leitura de XLSX, relatório de erros e planilha limpa

**Branch:** `feature/xlsx-relatorios` · **No PR:** `Closes #9`

**O que você vai entregar:** a ferramenta deixa de só mostrar erros na tela e passa a **entregar arquivos**. Ela lê também planilhas do Excel (`.xlsx`), grava o relatório de erros em `saida\relatorio_erros.csv` e grava a planilha limpa, padronizada, em `saida\cursos_limpos.csv`.

**Divida a semana assim:**

| Dia | O que fazer |
|---|---|
| sexta, 13/11 | passos 1 a 3 (*branch*, `openpyxl` e o conversor para XLSX) |
| segunda, 16/11 | passo 4 (`ler_xlsx`, incluindo o experimento da armadilha das datas) |
| terça, 17/11 | passos 5 e 6 (`ler_planilha` e `padronizar_linha`) |
| quarta, 18/11 | passos 7 e 8 (`relatorio.py` e o novo `main.py`) |
| quinta, 19/11 | passos 9 e 10 (verificação, diário e PR) |

### Passo a passo

**1. Crie a *branch*** `feature/xlsx-relatorios` a partir da `develop` atualizada.

**2. Instale o `openpyxl`** e registre a dependência. Com o `.venv` ativo:

```
pip install openpyxl
pip freeze > requirements.txt
```

O `requirements.txt` é a lista das bibliotecas do projeto, com as versões. É com ele que outra pessoa recria o seu ambiente, com `pip install -r requirements.txt`. Anote a versão do `openpyxl` (aparece no `requirements.txt`) no `docs/ambiente.md`.

*Pesquise:* o que é o `openpyxl` (documentação oficial: https://openpyxl.readthedocs.io), em especial o "Tutorial".

**3. Crie `src\converter_para_xlsx.py`.** Você precisa de uma versão `.xlsx` do `cursos_com_erros.csv` para testar, e ela tem de se parecer com uma planilha **digitada no Excel**. Quando alguém digita `160` no Excel, a célula guarda um **número**; quando digita `01/03/2027`, guarda uma **data**. O conversor imita isso. O programa:

1. recebe o caminho do CSV na linha de comando (`sys.argv[1]`) e lê com `ler_csv`;
2. cria uma planilha nova com `Workbook()` e pega a aba ativa;
3. acrescenta o cabeçalho com `aba.append(colunas)`;
4. para cada linha, monta a lista de valores convertendo cada um assim:
   - se o valor é um número inteiro (teste com `int()` dentro de `try`), grava o número;
   - senão, se `converter_data` consegue transformar o valor em data, grava a data;
   - senão, grava o texto como está;
5. acrescenta a lista com `aba.append(valores)`;
6. formata as células de data para aparecerem como `DD/MM/AAAA` no Excel: em cada célula que recebeu uma data, `celula.number_format = "DD/MM/YYYY"`;
7. salva com o mesmo nome do CSV, trocando a extensão para `.xlsx`. *Pesquise:* `os.path.splitext`.

Rode `python src\converter_para_xlsx.py dados\cursos_com_erros.csv` e confira se apareceu o `dados\cursos_com_erros.xlsx`. Se a máquina tiver Excel ou LibreOffice, abra e veja que as datas aparecem como datas e os números alinhados à direita. O arquivo `.xlsx` é pequeno e **vai** para o Git.

**4. Escreva `ler_xlsx(caminho)` no `leitor.py`.** Ela precisa devolver **exatamente o mesmo formato** que o `ler_csv`: a lista de colunas e a lista de linhas, com todos os valores em **texto**. Assim, nada do resto da ferramenta precisa mudar. O caminho:

1. abra com `load_workbook(caminho, data_only=True)` e pegue a aba ativa (`planilha.active`);
2. leia todas as linhas com `list(aba.iter_rows(values_only=True))`. Cada linha vem como uma tupla de valores;
3. se não vier nenhuma linha, devolva duas listas vazias;
4. a primeira linha é o cabeçalho: transforme cada valor em texto, e o `None` em `""`;
5. para cada uma das outras linhas, transforme cada valor em texto e monte o dicionário juntando cabeçalho e valores. *Pesquise:* a função `zip()`. **Ignore as colunas cujo cabeçalho está vazio** (o Excel às vezes guarda colunas vazias à direita);
6. **ignore as linhas totalmente vazias**, como o `csv.DictReader` faz com o CSV.

**Faça primeiro do jeito simples**, transformando cada valor com `str()`, e rode `python src\main.py dados\cursos_com_erros.xlsx` (no `main.py`, troque por enquanto o `ler_csv` por `ler_xlsx`). Você vai ver **um monte de erros R06 que não existem**: o `openpyxl` devolve as datas como objetos `datetime`, e o `str()` delas dá `2027-03-01 00:00:00`, que a regra R06 recusa. Também vai ver números como `160`, que vieram certos. **Tire um print dessa saída** (`docs/evidencias/s05_armadilha_datas.png`) e anote no diário: esta é a "armadilha das datas do Excel", e ela é um achado do seu trabalho que vai para os Resultados.

Agora corrija. Escreva uma função `celula_para_texto(valor)`, que transforma o valor de uma célula em texto assim:

| Valor da célula | Vira |
|---|---|
| `None` | `""` |
| data (`date` ou `datetime`) | o texto `DD/MM/AAAA`, com `.strftime("%d/%m/%Y")` |
| número decimal sem casas (`160.0`) | `"160"` |
| qualquer outro valor | `str(valor)` |

*Pesquise:* `isinstance`, o módulo `datetime` (o `datetime` é um tipo especial de `date`, então testar `isinstance(valor, date)` pega os dois) e o método `is_integer()` dos números decimais.

Use `celula_para_texto` nos passos 4 e 5 e rode de novo: os erros falsos de R06 devem desaparecer, e o resultado deve ser **igual** ao do CSV (13 erros).

**5. Escreva `ler_planilha(caminho)` no `leitor.py`.** Ela olha a extensão do arquivo (em minúsculas) e chama `ler_csv` ou `ler_xlsx`. Para qualquer outra extensão, ela lança um erro com uma mensagem clara:

```python
raise ValueError(f"formato não suportado: {extensao} (use .csv ou .xlsx)")
```

No `main.py`, troque a leitura por `ler_planilha` e trate esse erro com `try` / `except ValueError`, mostrando a mensagem e encerrando. Teste com um arquivo `.txt` ou `.xls` qualquer: a mensagem tem de aparecer, sem `Traceback`.

**6. Escreva `padronizar_linha(linha)` no `validadores.py`.** A planilha limpa (e, na Sprint 7, o banco) recebe os valores **padronizados**, e não do jeito que foram digitados. A função recebe uma linha **válida** e devolve um dicionário novo:

- só com as oito colunas de `COLUNAS_ESPERADAS`, nessa ordem (*pesquise:* *dict comprehension*);
- com os espaços das pontas retirados de todos os valores;
- com o `turno` em minúsculas;
- com o `campus` escrito **exatamente como na lista oficial**: se veio `natal-central`, vira `Natal-Central`. Percorra `CAMPI_IFRN` e pegue o item cujo `lower()` é igual ao `lower()` do valor.

Na linha 15 do `cursos_com_erros.csv`, `"  Noturno "` vira `"noturno"` e `"natal-central"` vira `"Natal-Central"`.

**7. Crie `src\relatorio.py`** com duas funções. As duas gravam CSV com `encoding="utf-8-sig"` (assim o Excel mostra os acentos certos) e `newline=""`, e criam a pasta de saída se ela não existir (*pesquise:* `os.makedirs(..., exist_ok=True)`).

- `gravar_relatorio(erros, caminho)`: grava um CSV com as colunas `linha`, `coluna`, `regra`, `mensagem` e `valor`, uma linha por erro. *Pesquise:* `csv.DictWriter`, `writeheader()` e `writerows()`. Como os erros já são dicionários com essas chaves, o `writerows(erros)` grava tudo de uma vez.
- `gravar_planilha_limpa(linhas, erros, caminho)`: monte o **conjunto** dos números de linha que têm erro; percorra as linhas com `enumerate(linhas, start=2)`; grave só as linhas cujo número **não** está no conjunto, passando cada uma por `padronizar_linha`. O cabeçalho é `COLUNAS_ESPERADAS`. A função devolve quantas linhas gravou.

A planilha limpa é **sempre** CSV, mesmo quando a entrada é `.xlsx`: CSV é o formato mais simples de importar em qualquer sistema.

**8. Atualize o `main.py`.** Depois de validar, ele:

1. mostra **no máximo 20 erros** na tela e, se houver mais, a linha `... e mais N erros (veja o relatório)`. Planilhas grandes, como as da Sprint 6, teriam milhares de linhas de erro na tela;
2. mostra o resumo da Sprint 4 e, embaixo, a **contagem por regra**, sempre com as oito regras, mesmo as que deram zero. *Pesquise:* `collections.Counter`;
3. grava o relatório e a planilha limpa e mostra onde gravou.

A saída com o `cursos_com_erros.xlsx` deve terminar assim:

```
Linhas de dados: 15 | Linhas válidas: 3 | Linhas com erro: 12 | Total de erros: 13
Erros por regra: R01=2 R02=1 R03=2 R04=1 R05=2 R06=2 R07=2 R08=1
Relatório de erros: saida\relatorio_erros.csv
Planilha limpa: saida\cursos_limpos.csv (3 linhas)
```

Abra o `saida\relatorio_erros.csv` no VS Code e confira as 13 linhas. Tire um print (`docs/evidencias/s05_relatorio_erros.png`): ele vai para o Apêndice A do TCC.

**9. Aumente a verificação.** Acrescente os `assert`:

1. `celula_para_texto` com: `None`, uma `date`, um `datetime`, `160`, `160.0`, `12.5` e `"abc"`;
2. `ler_csv` e `ler_xlsx` do `cursos_com_erros` devolvem **as mesmas linhas** (compare as duas listas com `==`) e os mesmos erros;
3. `ler_planilha` com uma extensão não suportada lança `ValueError`. Para testar que um erro **acontece**, use:

   ```python
   try:
       ler_planilha("dados/arquivo.pdf")
       assert False, "deveria ter lançado ValueError"
   except ValueError:
       pass
   ```

4. `padronizar_linha` da linha 15 dá `turno` igual a `"noturno"` e `campus` igual a `"Natal-Central"`;
5. `gravar_planilha_limpa` do `cursos_com_erros` devolve 3 (grave em `saida/teste_limpos.csv`).

**10. Finalize.** Atualize o `README.md` (o `requirements.txt`, o conversor e as saídas). Diário (com a armadilha das datas), *commits* e PR para a `develop` com `Closes #9`. Em **Evidências de teste**, cole a saída da verificação e a do `main.py` com o `.xlsx`.

### Como saber que deu certo

- O `main.py` dá o **mesmo** resultado (13 erros, mesma contagem por regra) com o `cursos_com_erros.csv` e com o `cursos_com_erros.xlsx`.
- `saida\relatorio_erros.csv` tem o cabeçalho e 13 linhas; `saida\cursos_limpos.csv` tem o cabeçalho e 3 linhas, com `noturno` e `Natal-Central` padronizados.
- Um arquivo `.pdf` ou `.xls` dá uma mensagem clara, sem `Traceback`.
- O `git status` não mostra a pasta `saida/`.
- `python testes\verificar_validadores.py` mostra `Tudo certo`.

### Entrega da sprint

- [ ] `openpyxl` instalado e `requirements.txt` no repositório; versão no `docs/ambiente.md`.
- [ ] `src/converter_para_xlsx.py` e `dados/cursos_com_erros.xlsx`.
- [ ] `ler_xlsx` com `celula_para_texto`, devolvendo o mesmo formato do `ler_csv`.
- [ ] Print da armadilha das datas e explicação no diário.
- [ ] `ler_planilha` escolhendo o leitor pela extensão.
- [ ] `padronizar_linha` (turno em minúsculas e *campus* na grafia oficial).
- [ ] `src/relatorio.py` gravando o relatório de erros e a planilha limpa.
- [ ] `main.py` com no máximo 20 erros na tela, a contagem por regra e os arquivos de saída.
- [ ] Verificação com os novos casos, mostrando `Tudo certo`.
- [ ] PR aberto para a `develop`, com `Closes #9`.

### Erros comuns

- **`ModuleNotFoundError: No module named 'openpyxl'`:** o `.venv` não está ativo, ou você instalou com o ambiente desativado. Ative e instale de novo.
- **`PermissionError` ao gravar:** o arquivo está aberto no Excel. Feche e rode de novo.
- **O relatório abre no Excel tudo em uma coluna só:** o Excel em português espera ponto e vírgula. Abra no VS Code, ou no Excel use *Dados → De Texto/CSV* e escolha a vírgula como delimitador. Anote no diário: é uma limitação conhecida, que aparece no TCC.
- **O XLSX dá mais erros que o CSV:** algum tipo de valor não está sendo tratado no `celula_para_texto`. Mostre na tela `type(valor)` dos valores da linha que deu diferença.

<a id="sprint-5-escrita"></a>
## Sprint 5 · Escrita: Metodologia (Entrega 4)

**Prazo: 19/11/2026** · **Versão no Drive:** `Entrega 4 - Metodologia - 19/11`

A Metodologia explica **como o trabalho foi feito**, com detalhe suficiente para que outra pessoa o repita sem falar com você. Escreva de forma impessoal ("foi desenvolvida", "realizou-se") e **no passado**, como é a norma, inclusive o que ainda vai acontecer nas Sprints 6 e 7: o procedimento já está decidido neste plano. Na Sprint 8, você confere se tudo aconteceu como descrito e corrige o que mudou. Cerca de 4 a 6 páginas.

**Pasta com metodologias prontas para consulta:** https://drive.google.com/drive/folders/1OwAjk47aQWuRKLf-ktesizSax9csEtXO?usp=sharing

### 3.1 Classificação da pesquisa

**Pesquise:** Gil (2017) e Prodanov e Freitas (2013); as referências completas estão em `docs/orientacoes_escrita_tcc.md`. Leia os capítulos sobre classificação de pesquisa antes de escrever.

**Escreva:** a classificação pelos quatro critérios, **justificando cada um** com uma frase que ligue o critério ao seu trabalho (o erro mais comum é só listar os rótulos):

- **natureza:** aplicada, porque busca resolver um problema prático da plataforma iFIC;
- **objetivos:** exploratória e descritiva, porque explora uma solução para um problema ainda não tratado no projeto e descreve o comportamento medido da ferramenta;
- **abordagem:** quantitativa, porque os resultados são números (erros encontrados, precisão, revocação, tempo);
- **procedimentos:** pesquisa bibliográfica (conceitos e trabalhos relacionados) e pesquisa experimental (avaliação da ferramenta com planilhas de teste controladas).

### 3.2 Etapas do trabalho

**Escreva:** as etapas em ordem, com uma **figura do fluxo** feita no draw.io (gratuito, funciona no navegador: https://app.diagrams.net). Lembre: identificação em cima, centralizada, e fonte embaixo. As etapas são:

1. estudo da linguagem Python e do fluxo de trabalho com Git;
2. levantamento das regras de validação, a partir das normas dos cursos FIC e da lista de *campi* do IFRN;
3. definição do formato da planilha de entrada;
4. implementação incremental da ferramenta, em ciclos (sprints) de uma semana, com versionamento no GitHub, *pull requests* e revisão de código pelos orientadores ao fim de cada ciclo;
5. implementação da leitura de `.xlsx`, do relatório de erros e da planilha limpa;
6. construção do gerador de planilhas de teste com gabarito;
7. avaliação da detecção de erros;
8. medição do tempo de processamento;
9. implementação da importação para o banco SQLite;
10. documentação e verificação da reprodução do trabalho.

Informe o endereço do repositório (https://github.com/nicifrn/mp_ific_importador_planilhas) e diga que o código está disponível nele.

### 3.3 Materiais e ambiente de execução

**Escreva:** para **cada** ferramenta, três coisas: **o que é** (uma frase, com referência; a documentação oficial serve), **para que foi usada** e **por que foi escolhida**. Escrever só "foi utilizado o Python" não vale nada. Feche com um **quadro-resumo** (ferramenta, versão, finalidade).

| Ferramenta | O que dizer além da versão |
|---|---|
| Python | por que Python, e a relação com o iFIC (que usa Django, um *framework* em Python) |
| Módulo `csv` | leitura e escrita de CSV sem biblioteca externa |
| `openpyxl` | leitura de `.xlsx`; por que foi preciso uma biblioteca externa |
| Módulo `datetime` | conversão e comparação de datas; recusa de datas inexistentes |
| Módulo `sqlite3` e SQLite | banco embutido, sem servidor; por que SQLite e não MySQL ([seção 4](#decisoes)) |
| Módulo `random` | geração das planilhas de teste, com semente fixa |
| `time.perf_counter` e `statistics` | medição do tempo, média e desvio padrão |
| Git, GitHub e VS Code | versionamento, revisão e desenvolvimento |

Descreva também o **ambiente de execução** das medições, com os dados do `docs/ambiente.md` (processador, memória, disco, sistema e versões). Sem isso, os tempos não significam nada.

Diga, por fim, **o que não foi usado e por quê**: `pandas` e as bibliotecas prontas de validação, MySQL e Django. Use as justificativas da [seção 4](#decisoes).

### 3.4 Formato da planilha e regras de validação

**Escreva:** o formato da planilha (um quadro com as oito colunas, o que cada uma é e um exemplo), a conferência do cabeçalho e o **quadro das oito regras**, com o código, a descrição e a dimensão de qualidade de dados (o elo com a seção 2.2). Diga de onde veio cada regra: a norma dos cursos FIC (R03), a lista oficial de *campi* do portal do IFRN, com a data de acesso (R08), e o próprio formato da planilha (as outras).

### 3.5 Arquitetura da ferramenta

**Escreva:** a divisão em módulos (`config`, `leitor`, `validadores`, `relatorio`, `banco` e `main`), o que cada um faz e **por que** separar assim: cada arquivo tem uma responsabilidade, e uma regra nova muda só o `validadores.py`. Faça uma **figura** com o caminho dos dados: planilha → leitura → conferência do cabeçalho → validação → relatório de erros e planilha limpa → banco de dados.

### 3.6 Base de teste e procedimento de avaliação

Esta é a subseção mais importante do capítulo, porque é ela que sustenta os seus resultados.

**Escreva sobre a base de teste:** as três planilhas (100, 1.000 e 10.000 linhas); como cada linha válida foi gerada (nomes de cursos do Guia Pronatec, *campi* da lista oficial, datas com o fim de 30 a 180 dias depois do início); a proporção de linhas com erro (cerca de 20%); a regra de **um erro por linha com erro**, com a regra sorteada entre as oito; os valores usados para estragar cada campo; o gabarito, com a linha e a regra de cada erro; e a **semente fixa** (`random.seed(42)`), que faz o gerador produzir sempre as mesmas planilhas. Deixe claro que **nenhum dado real foi usado**.

**Escreva sobre a avaliação da detecção:** defina, no contexto do trabalho, **verdadeiro positivo** (o gabarito tem o erro e a ferramenta o encontrou), **falso negativo** (o gabarito tem o erro e a ferramenta não o encontrou) e **falso positivo** (a ferramenta apontou um erro que não está no gabarito). Diga que a correspondência foi feita pelo **par (linha, regra)**. Apresente as fórmulas:

- **precisão** = verdadeiros positivos / (verdadeiros positivos + falsos positivos);
- **revocação** = verdadeiros positivos / (verdadeiros positivos + falsos negativos).

*Pesquise* esses dois conceitos em um livro ou artigo de recuperação da informação ou de aprendizado de máquina e cite a fonte.

**Escreva sobre a medição de tempo:** o que foi cronometrado (leitura da planilha, conferência do cabeçalho e validação; a gravação dos arquivos ficou de fora), a ferramenta (`time.perf_counter`), uma execução de aquecimento descartada, **5 medições** por tamanho, as estatísticas (média e desvio padrão) e os cuidados (mesma máquina, programas fechados).

**Escreva sobre a importação:** o que foi conferido: que o banco recebe exatamente as linhas válidas e que importar a mesma planilha duas vezes não duplica registros.

### Entrega

- [ ] 3.1 com os quatro critérios, cada um justificado.
- [ ] 3.2 com as etapas, a figura do fluxo e o endereço do repositório.
- [ ] 3.3 com o que é, para que e por que de cada ferramenta, o quadro-resumo, o ambiente e o que não foi usado.
- [ ] 3.4 com os quadros da planilha e das regras.
- [ ] 3.5 com os módulos e a figura do caminho dos dados.
- [ ] 3.6 com a base de teste, as definições, as fórmulas, a medição de tempo e a importação.
- [ ] Comentários da Entrega 3 resolvidos.

Nomeie a versão no Drive, comente na *issue* e mova o card para **Em revisão**.

---

# Sprint 6: 20/11 a 26/11/2026

<a id="sprint-6-dev"></a>
## Sprint 6 · Desenvolvimento: massa de teste com gabarito e avaliação da detecção

**Branch:** `feature/massa-gabarito` · **No PR:** `Closes #11`

**O que você vai entregar:** um gerador de planilhas fictícias que coloca erros de propósito e registra cada um em um **gabarito**; as planilhas de 100, 1.000 e 10.000 linhas; e um comparador que confere o relatório da ferramenta com o gabarito e calcula precisão e revocação. **Sem esta sprint, o TCC não tem como provar que a ferramenta funciona.**

**Divida a semana assim** (sexta, 20/11, é feriado):

| Dia | O que fazer |
|---|---|
| segunda, 23/11 | passos 1 a 3 (*branch*, estudo do `random` e `gerar_linha_valida`) |
| terça, 24/11 | passos 4 e 5 (`estragar`, o gabarito e a conferência da semente) |
| quarta, 25/11 | passo 6 (`comparador.py`) |
| quinta, 26/11 | passos 7 a 9 (rodar, conferir, verificação e PR) |

### A ideia do gabarito

Você não pode escrever no TCC que "a ferramenta funciona" só porque ela mostrou alguns erros. Você precisa saber, **antes de rodar**, exatamente quais erros existem na planilha, e depois conferir se ela achou todos e se não inventou nenhum. Você já fez isso em miniatura na Sprint 4, com as 15 linhas do `cursos_com_erros.csv` e a tabela do passo 9. Agora o programa faz o mesmo em escala: gera milhares de linhas, estraga cerca de 20% delas e anota no gabarito cada estrago.

### Passo a passo

**1. Crie a *branch*** `feature/massa-gabarito` a partir da `develop` atualizada. Crie a pasta `resultados`. A pasta `dados\massas` vai ser criada pelo próprio gerador, e ela **não vai para o Git** (já está no `.gitignore`): como a geração é reprodutível pela semente, basta enviar o programa.

**2. Estude o módulo `random`.** *Pesquise* (documentação oficial em português: https://docs.python.org/pt-br/3/library/random.html): `random.seed`, `random.random`, `random.choice` e `random.randint`. *Pesquise também:* `datetime.timedelta`, para somar e subtrair dias de uma data.

Faça um teste rápido no terminal do Python: rode `random.seed(42)` e, em seguida, `random.randint(1, 100)` três vezes. Anote os números. Repita tudo desde o `seed`: os números são **os mesmos**. É isso que a semente faz, e é isso que torna o seu experimento **reprodutível**: qualquer pessoa que rodar o seu gerador vai obter as mesmas planilhas e os mesmos números que você publicou.

**3. Crie `src\gerador_massa.py` e escreva `gerar_linha_valida(indice)`.** No começo do arquivo, importe as listas do `config.py` e defina:

- `NOMES_CURSOS`: uma lista com **15 nomes de cursos FIC** tirados do Guia Pronatec (você leu na Sprint 1). Escreva os nomes exatamente como estão no guia;
- `CARGAS`: `[160, 200, 240, 300, 400]`;
- `TAXA_ERRO = 0.20`;
- `REGRAS`: a lista `["R01", "R02", ..., "R08"]`.

A função recebe o número do curso e devolve um dicionário com as oito colunas, **todos os valores em texto**, sempre válidos:

| Coluna | Como gerar |
|---|---|
| `codigo` | `f"FIC-2027-{indice:05d}"` (*pesquise:* o `:05d` da f-string) |
| `nome_curso` | sorteado de `NOMES_CURSOS` |
| `campus` | sorteado de `CAMPI_IFRN` |
| `carga_horaria` | sorteada de `CARGAS` |
| `turno` | sorteado de `TURNOS_PERMITIDOS` |
| `vagas` | inteiro sorteado de 10 a 40 |
| `data_inicio` | 04/01/2027 mais um número sorteado de 0 a 300 dias |
| `data_fim` | a data de início mais um número sorteado de 30 a 180 dias |

As datas são montadas como `date` e transformadas em texto com `.strftime("%d/%m/%Y")`.

**4. Escreva `estragar(linha, regra, codigos_validos)`.** Ela recebe uma linha válida e o código de uma regra, altera a linha **para violar só aquela regra** e devolve o nome da coluna alterada. Use estes valores, sorteando um da lista com `random.choice`:

| Regra | O que fazer na linha | Coluna devolvida |
|---|---|---|
| R01 | sortear **uma** coluna de `COLUNAS_ESPERADAS` e deixá-la `""` | a sorteada |
| R02 | trocar o `codigo` por um sorteado de `codigos_validos` | `codigo` |
| R03 | `carga_horaria` ← `"0"`, `"-40"`, `"cem"`, `"12.5"` ou `"160h"` | `carga_horaria` |
| R04 | `vagas` ← `"0"`, `"-5"`, `"trinta"` ou `"2.5"` | `vagas` |
| R05 | `turno` ← `"noite"`, `"Noturnoo"`, `"manhã"`, `"diurno"` ou `"tarde"` | `turno` |
| R06 | sortear `data_inicio` ou `data_fim` e pôr `"31/02/2027"`, `"2027-03-01"`, `"15/13/2027"`, `"00/05/2027"` ou `"amanhã"` | a sorteada |
| R07 | `data_fim` ← a data de início **menos** 0 a 30 dias (0 dá a mesma data, que também é R07) | `data_fim` |
| R08 | `campus` ← `"Natal"`, `"Natal Central"`, `"Campus Mossoró"`, `"Recife"` ou `"Fortaleza"` | `campus` |

Para a R07, converta a `data_inicio` com o seu `converter_data` (importe do `validadores.py`), subtraia os dias com `timedelta` e grave o resultado em texto.

**5. Escreva a parte principal do gerador.** Ela fica dentro de `if __name__ == "__main__":` (*pesquise* o que isso significa; a explicação curta é: "só roda quando o arquivo é executado, e não quando é importado". Você vai precisar disso no passo 8). O que ela faz:

1. lê o número de linhas da linha de comando: `python src\gerador_massa.py 1000`;
2. chama `random.seed(42)`;
3. cria a pasta `dados\massas`, se não existir;
4. para cada `indice` de 1 a N:
   1. gera a linha válida; o número da linha na planilha é `indice + 1`;
   2. sorteia com `random.random() < TAXA_ERRO` se a linha vai ter erro;
   3. se vai: sorteia a regra. **Se a regra for R02 e a lista `codigos_validos` ainda estiver vazia** (acontece nas primeiras linhas), use R03 no lugar. Estrague a linha e registre no gabarito o número da linha, a regra e a coluna;
   4. se não vai: acrescente o código dela em `codigos_validos`;
   5. guarde a linha;
5. grava `dados\massas\massa_N.csv` (com `COLUNAS_ESPERADAS` como cabeçalho) e `dados\massas\gabarito_N.csv` (colunas `linha`, `regra` e `coluna`), os dois com `utf-8-sig`;
6. mostra um resumo: `Massa de 1000 linhas: 197 com erro (19,7%)`, e a quantidade de erros por regra.

> **Por que a R02 copia só códigos de linhas sem erro?** Se copiasse o código de uma linha que teve o código apagado (R01) ou que já é uma cópia, o resultado não seria o erro que o gabarito diz. Essa é exatamente a decisão que você vai explicar na Metodologia.

**Confira a semente:** gere a massa de 1.000 duas vezes e compare a "impressão digital" dos dois arquivos com o comando do Windows:

```
certutil -hashfile dados\massas\massa_1000.csv SHA256
```

O código que aparece tem de ser **igual** nas duas vezes. Anote-o no diário e no `README.md`: é a prova de que outra pessoa vai obter a mesma planilha.

**6. Crie `src\comparador.py`.** Para cada tamanho (100, 1.000 e 10.000), ele:

1. confere se a massa e o gabarito existem; se não existirem, mostra `Rode antes: python src\gerador_massa.py N` e passa para o próximo;
2. lê a massa com `ler_csv` e valida com `validar_planilha` (importe as funções; não chame o `main.py`);
3. monta dois **conjuntos** de pares `(linha, regra)`: o `esperado`, a partir do gabarito, e o `encontrado`, a partir dos erros. **Atenção:** o número da linha que vem do gabarito é **texto** (o CSV entrega tudo como texto); converta com `int()`, senão nenhum par vai bater;
4. calcula: verdadeiros positivos (`len(esperado & encontrado)`), falsos negativos (`len(esperado - encontrado)`), falsos positivos (`len(encontrado - esperado)`), precisão e revocação. Cuidado com a divisão por zero;
5. se houver falsos negativos ou falsos positivos, mostra até 10 exemplos de cada, para você investigar.

No fim, ele grava `resultados\deteccao.csv`, com uma linha por tamanho e as colunas `tamanho`, `linhas_com_erro`, `erros_gabarito`, `verdadeiros_positivos`, `falsos_negativos`, `falsos_positivos`, `precisao` e `revocacao`; e `resultados\deteccao_por_regra.csv`, só para a massa de 10.000, com uma linha por regra e as colunas `regra`, `no_gabarito`, `encontrados`, `falsos_negativos` e `falsos_positivos`.

*Pesquise:* as operações `&` (interseção) e `-` (diferença) entre conjuntos.

**7. Rode tudo e confira.**

```
python src\gerador_massa.py 100
python src\gerador_massa.py 1000
python src\gerador_massa.py 10000
python src\comparador.py
```

O resultado esperado é **precisão e revocação iguais a 1** (100%) nas três massas. Se não der:

- **falsos positivos:** quase sempre a culpa é do **gerador**, que criou sem querer uma linha inválida e a registrou como válida, ou de alguma decisão da [seção 3.4](#regras) que a ferramenta não está seguindo. Abra a massa na linha indicada e veja o que há nela;
- **falsos negativos:** a ferramenta deixou passar um tipo de valor. Veja qual valor da tabela do passo 4 está na linha indicada e teste a regra com ele;
- **nunca altere o gabarito à mão** para "fechar a conta". Corrija o programa que está errado, gere de novo e anote no diário o que aconteceu: um defeito encontrado e corrigido também é resultado.

Por último, rode o `main.py` com a massa de 10.000 (`python src\main.py dados\massas\massa_10000.csv`) e veja o resumo e a contagem por regra. Tire um print (`docs/evidencias/s06_main_massa_10000.png`).

**8. Aumente a verificação.** Importe `gerar_linha_valida` e `estragar` do `gerador_massa` (é por isso que a parte principal está dentro do `if __name__ == "__main__":`) e acrescente:

1. com `random.seed(1)`, 200 linhas geradas por `gerar_linha_valida` passam **todas** sem erro pelo `validar_linha`;
2. para cada regra de R01 a R08, **exceto a R02**: uma linha válida estragada com aquela regra dá **exatamente um** erro, e com aquela regra;
3. para a R02: uma planilha de duas linhas, a segunda estragada com o código da primeira, dá exatamente um erro R02, na linha 3.

**9. Finalize.** Atualize o `README.md` com os comandos do gerador e do comparador e com o código SHA-256 da massa de 1.000. Escreva no diário o que significa o seu resultado (veja o aviso abaixo). *Commits* e PR para a `develop` com `Closes #11`. Em **Evidências de teste**, cole a saída da verificação e a do `comparador.py`. Os arquivos de `resultados/` **vão** para o Git.

> **Um aviso importante sobre o resultado de 100%.** Ele é esperado, e você não deve comemorá-lo sem ressalva: os erros foram colocados pelo mesmo autor que escreveu as regras, então a base de teste só tem os tipos de erro que as regras já conhecem. Isso é uma **limitação metodológica**, e ela vai escrita nos Resultados (seção 4.6). O que o 100% prova é que a ferramenta implementa corretamente as oito regras, em qualquer tamanho de planilha. Ele **não** prova que ela acharia qualquer erro de uma planilha real.

### Como saber que deu certo

- `dados\massas\` tem seis arquivos: três massas e três gabaritos.
- O gerador, rodado duas vezes, dá o mesmo SHA-256.
- Cada massa tem perto de 20% de linhas com erro, e as oito regras aparecem no gabarito da massa de 10.000.
- O `comparador.py` mostra precisão e revocação iguais a 1 nas três massas (ou você investigou, corrigiu e anotou por que não).
- `resultados\deteccao.csv` tem 3 linhas e `resultados\deteccao_por_regra.csv` tem 8.
- `python testes\verificar_validadores.py` mostra `Tudo certo`.

### Entrega da sprint

- [ ] `src/gerador_massa.py` com `gerar_linha_valida`, `estragar` e a parte principal dentro do `if __name__ == "__main__":`.
- [ ] Semente conferida com o `certutil` e o SHA-256 anotado no diário e no `README.md`.
- [ ] Massas de 100, 1.000 e 10.000 linhas geradas (sem enviar para o Git).
- [ ] `src/comparador.py` gravando `resultados/deteccao.csv` e `resultados/deteccao_por_regra.csv`.
- [ ] Print do `main.py` com a massa de 10.000.
- [ ] Verificação com os casos do gerador, mostrando `Tudo certo`.
- [ ] Diário com a interpretação do resultado e qualquer defeito encontrado.
- [ ] PR aberto para a `develop`, com `Closes #11`.

### Erros comuns

- **O comparador dá 0 verdadeiros positivos:** o número da linha do gabarito ficou como texto. Use `int()`.
- **Ao importar o `gerador_massa` na verificação, ele gera uma massa inteira:** faltou o `if __name__ == "__main__":`.
- **O SHA-256 muda a cada execução:** o `random.seed(42)` não está antes de **todos** os sorteios, ou algum valor depende da data de hoje.
- **Falsos positivos com R07:** a data de fim válida é gerada a partir de 30 dias depois do início. Confira se você não somou dias negativos por engano.

<a id="sprint-6-escrita"></a>
## Sprint 6 · Escrita: colocar no documento as tabelas com os dados da Sprint 6

**Prazo: 26/11/2026** · **Versão no Drive:** `Tabelas dos Resultados - 26/11`

**O que você vai entregar:** o capítulo 4 montado no documento, com as tabelas, os quadros e as figuras que você já tem, para que os Resultados não fiquem para a última hora. Nesta semana não há entrega de capítulo: o texto corrido é na Sprint 7. Na quinta-feira, coloque as tabelas da detecção assim que o comparador terminar; os outros itens você faz de segunda a quarta.

### Passo a passo

**1. Crie a estrutura do capítulo 4** no documento, com os títulos das seções 4.1 a 4.6 ([seção 9.3](#escrita-geral)).

**2. Em 4.1 (A ferramenta desenvolvida),** coloque a figura com o print do `main.py` rodando sobre o `cursos_com_erros.xlsx` (Sprint 5) e a figura com a massa de 10.000 (Sprint 6).

**3. No Apêndice A,** coloque um quadro com as 13 linhas do `saida\relatorio_erros.csv` gerado com o `cursos_com_erros`. Em 4.1, escreva uma frase que remete a ele ("o relatório completo está no Apêndice A").

**4. Em 4.2 (Decisões de validação),** monte um quadro com as decisões da [seção 3.4](#regras) e, para cada uma, a linha do `cursos_com_erros.csv` que a demonstra (por exemplo: decisão 2, linha 13). Acrescente a armadilha das datas do Excel (Sprint 5), com o print `s05_armadilha_datas.png`.

**5. Em 4.3 (Detecção de erros),** coloque:

- **Tabela 1:** os dados de `resultados\deteccao.csv`, uma linha por tamanho;
- **Tabela 2:** os dados de `resultados\deteccao_por_regra.csv`, uma linha por regra.

**6. Confira o padrão ABNT** de todas as tabelas, quadros e figuras: identificação em cima (`Tabela 1 - Resultado da detecção de erros por tamanho de planilha`), centralizada; fonte embaixo (`Fonte: elaborado pelo autor (2026).`); vírgula como separador decimal; o mesmo número de casas decimais em toda a coluna (precisão e revocação com 2 casas: `1,00`).

**7. Escreva uma frase de apresentação antes de cada tabela e figura** ("A Tabela 1 apresenta..."). A discussão vem na Sprint 7.

### Entrega

- [ ] Estrutura do capítulo 4 (4.1 a 4.6).
- [ ] Figuras do `main.py` em 4.1 e o relatório no Apêndice A.
- [ ] Quadro das decisões em 4.2, com a armadilha das datas.
- [ ] Tabelas 1 e 2 em 4.3, com os números da Sprint 6.
- [ ] Identificação em cima, fonte embaixo, vírgula decimal e mesmas casas decimais.
- [ ] Uma frase de apresentação antes de cada tabela e figura.
- [ ] Comentários da Entrega 4 resolvidos.

Nomeie a versão no Drive, comente na *issue* e mova o card para **Em revisão**.

---

# Sprint 7: 27/11 a 03/12/2026

<a id="sprint-7-dev"></a>
## Sprint 7 · Desenvolvimento: tempo de processamento e importação para o SQLite

**Branch:** `feature/tempo-sqlite` · **No PR:** `Closes #13`

**O que você vai entregar:** a medição do tempo que a ferramenta leva para conferir 100, 1.000 e 10.000 linhas, com 5 medições de cada, e a importação das linhas válidas para um banco SQLite, sem duplicar registros quando a mesma planilha é importada de novo.

**A medição oficial tem de terminar até segunda-feira, 30/11.** Os Resultados vencem na quinta, e você precisa dos números na terça para escrever.

**Divida a semana assim:**

| Dia | O que fazer |
|---|---|
| sexta, 27/11 | passos 1 e 2 (*branch* e `medir_tempo.py`, com uma rodada de teste) |
| segunda, 30/11 | passo 3 (rodada oficial, logo cedo) e passo 4 (separar as linhas válidas) |
| terça, 01/12 | passo 5 (`banco.py`) |
| quarta, 02/12 | passos 6 e 7 (importação no `main.py`, `ver_banco.py` e as conferências) |
| quinta, 03/12 | passos 8 e 9 (verificação, diário e PR) |

### Passo a passo

**1. Crie a *branch*** `feature/tempo-sqlite` a partir da `develop` atualizada. Se as massas da Sprint 6 não estiverem mais na máquina, gere de novo (a semente garante que serão as mesmas).

**2. Crie `src\medir_tempo.py`.** Ele mede quanto tempo a ferramenta leva para **ler a planilha, conferir o cabeçalho e validar**. A gravação dos arquivos de saída fica de fora: o tempo de gravar em disco varia muito e não diz nada sobre as regras.

1. Escreva a função `processar(caminho)`, que faz exatamente essas três coisas (`ler_planilha`, `verificar_colunas` e `validar_planilha`) e devolve os erros.
2. A parte principal recebe o nome da rodada na linha de comando: `python src\medir_tempo.py teste` ou `python src\medir_tempo.py oficial`.
3. Para cada tamanho (100, 1.000 e 10.000):
   1. rode `processar` **uma vez sem medir** (aquecimento: a primeira execução costuma ser mais lenta, porque o sistema ainda está carregando arquivos para a memória);
   2. rode **5 vezes**, medindo cada uma assim:

      ```python
      inicio = time.perf_counter()
      processar(caminho)
      duracao = time.perf_counter() - inicio
      ```

   3. acrescente cada medição ao arquivo `resultados\tempos.csv`, com as colunas `rodada`, `data_hora`, `tamanho`, `repeticao` e `segundos`.
4. **Abra o arquivo em modo de acréscimo** (`"a"`), e não de escrita (`"w"`): nenhuma medição é apagada ([seção 8](#combinados)). Grave o cabeçalho só se o arquivo ainda não existir.
5. No fim, calcule, **só com as medições da rodada `oficial`**, a média e o desvio padrão de cada tamanho e grave `resultados\tempos_resumo.csv`, com as colunas `tamanho`, `media_s`, `desvio_s`, `ms_por_linha` (a média dividida pelo tamanho, vezes 1.000) e `fator` (a média dividida pela média do tamanho anterior; vazio na primeira linha). Mostre a tabela na tela.

*Pesquise:* `time.perf_counter` (por que ele é melhor que `time.time` para medir duração), `statistics.mean`, `statistics.stdev` e `datetime.now().strftime(...)` para a coluna `data_hora`.

Na sexta, rode `python src\medir_tempo.py teste` para ver se tudo funciona. **Não use esses números no TCC.**

**3. Faça a rodada oficial na segunda-feira, 30/11, logo cedo.**

1. Use **a mesma máquina** descrita no `docs/ambiente.md`.
2. Reinicie o computador. Depois, abra só o Prompt de Comando: nada de navegador, VS Code ou outro programa.
3. Ative o `.venv` e rode `python src\medir_tempo.py oficial`.
4. Anote no diário a data, o horário e qualquer coisa estranha (por exemplo, o antivírus começou a rodar no meio).

Olhe a coluna `fator`: se a planilha fica 10 vezes maior, o tempo fica perto de 10 vezes maior? Anote no diário o que você observou. Isso é o centro da seção 4.4 do TCC.

**4. Separe as linhas válidas.** O banco e a planilha limpa recebem as mesmas linhas, então elas devem sair de um lugar só. Crie no `validadores.py` a função `separar_validas(linhas, erros)`, que devolve a lista das linhas sem erro, já passadas por `padronizar_linha`. Depois, **reorganize** a `gravar_planilha_limpa` para receber essa lista pronta, em vez de calcular as linhas válidas ela mesma. Rode a verificação: ela tem de continuar mostrando `Tudo certo` (ajuste o teste da `gravar_planilha_limpa`). Faça um *commit* só para isso, com o prefixo `refactor:`.

**5. Crie `src\banco.py`** com quatro funções:

- `criar_banco(caminho)`: abre (ou cria) o banco com `sqlite3.connect(caminho)`, cria a tabela, se ela ainda não existir, e devolve a conexão. Use este comando:

  ```sql
  CREATE TABLE IF NOT EXISTS curso_fic (
      codigo TEXT PRIMARY KEY,
      nome_curso TEXT NOT NULL,
      campus TEXT NOT NULL,
      carga_horaria INTEGER NOT NULL,
      turno TEXT NOT NULL,
      vagas INTEGER NOT NULL,
      data_inicio TEXT NOT NULL,
      data_fim TEXT NOT NULL
  )
  ```

- `data_para_iso(texto)`: transforma `"30/06/2027"` em `"2027-06-30"`, usando o seu `converter_data` e o método `.isoformat()`. **Por que guardar a data assim?** Porque, como texto, `"30/06/2027"` vem antes de `"01/07/2026"` em ordem alfabética, e o banco ordenaria errado. No formato `AAAA-MM-DD`, a ordem do texto é a ordem das datas, e o SQLite entende esse formato nas funções de data;
- `importar(conexao, linhas_validas)`: monta, para cada linha, uma tupla com os oito valores, convertendo `carga_horaria` e `vagas` para `int` e as datas com `data_para_iso`; grava todas de uma vez e faz o `commit`. Devolve quantas linhas importou. Use:

  ```python
  conexao.executemany(
      "INSERT OR REPLACE INTO curso_fic VALUES (?, ?, ?, ?, ?, ?, ?, ?)",
      lista_de_tuplas,
  )
  ```

- `contar(conexao)`: devolve o total de cursos na tabela, com `SELECT COUNT(*) FROM curso_fic` e `.fetchone()[0]`.

*Pesquise:*

- `INSERT OR REPLACE` no SQLite: como a tabela tem `codigo` como `PRIMARY KEY`, quando chega um código que já existe, o registro antigo é **substituído**, e não duplicado. É isso que garante que importar a mesma planilha duas vezes não duplica nada. Repare que o banco é a **segunda barreira** contra códigos repetidos: a primeira é a regra R02. Essa ideia vai para os Resultados;
- por que os valores entram com `?`, e nunca montando o comando SQL com f-string ou `+`. Procure por **injeção de SQL** e anote no diário. Vale uma frase no TCC.

**6. Ligue a importação no `main.py` e crie `src\ver_banco.py`.**

- No fim do `main.py`, depois de gravar os arquivos, importe as linhas válidas para `saida\cursos_fic.db` e mostre: `Banco: saida\cursos_fic.db | importadas agora: 3 | total no banco: 3`.
- O `ver_banco.py` abre o `saida\cursos_fic.db` e mostra: o total de cursos; os 5 primeiros cursos em ordem de data de início (`ORDER BY data_inicio`); e quantos cursos há por turno (`GROUP BY turno`). *Pesquise:* `SELECT`, `ORDER BY`, `LIMIT`, `GROUP BY` e `COUNT`, e o método `fetchall()`.

**7. Faça as três conferências da importação** e registre os números em `resultados\importacao.md`:

1. apague o `saida\cursos_fic.db`; rode `python src\main.py dados\massas\massa_10000.csv`. O número de **importadas agora** tem de ser igual ao de **linhas válidas** do resumo, e o total no banco também;
2. rode **o mesmo comando de novo**. O total no banco **não pode mudar**;
3. rode `python src\ver_banco.py` e tire um print (`docs/evidencias/s07_banco.png`). Confira que os turnos aparecem só nas quatro formas padronizadas e que as datas estão em `AAAA-MM-DD`.

**8. Aumente a verificação.** Use um banco **só na memória**, que some quando o programa termina: `criar_banco(":memory:")`. Acrescente:

1. `data_para_iso("01/03/2027")` dá `"2027-03-01"`;
2. importar duas linhas válidas dá total 2;
3. importar as mesmas duas linhas de novo continua dando total 2;
4. importar uma linha com o mesmo código de uma delas, mas com outro número de vagas, continua dando total 2, e o número de vagas daquele código no banco passa a ser o novo. *Pesquise:* `SELECT vagas FROM curso_fic WHERE codigo = ?`.

**9. Finalize.** `README.md` com os comandos do `medir_tempo.py` e do `ver_banco.py`. Diário (com as condições da rodada oficial). *Commits* e PR para a `develop` com `Closes #13`. Em **Evidências de teste**, cole a saída da verificação, a tabela do `tempos_resumo.csv` e o conteúdo do `resultados\importacao.md`.

### Como saber que deu certo

- `resultados\tempos.csv` tem 15 linhas da rodada `oficial` (3 tamanhos × 5 medições), além das de teste.
- `resultados\tempos_resumo.csv` tem 3 linhas, com média, desvio, milissegundos por linha e fator.
- Importar a massa de 10.000 duas vezes deixa o mesmo total no banco, igual ao número de linhas válidas.
- O `ver_banco.py` mostra os cursos com datas em `AAAA-MM-DD` e os turnos padronizados.
- `python testes\verificar_validadores.py` mostra `Tudo certo`.

### Entrega da sprint

- [ ] `src/medir_tempo.py` com aquecimento, 5 medições, modo de acréscimo e a rodada no nome.
- [ ] Rodada oficial feita até 30/11, com as condições anotadas no diário.
- [ ] `resultados/tempos.csv` e `resultados/tempos_resumo.csv`.
- [ ] `separar_validas` e a `gravar_planilha_limpa` reorganizada, em um *commit* `refactor:`.
- [ ] `src/banco.py` com `criar_banco`, `data_para_iso`, `importar` e `contar`.
- [ ] Importação no `main.py` e `src/ver_banco.py`.
- [ ] As três conferências em `resultados/importacao.md` e o print do banco.
- [ ] Verificação com os casos do banco, mostrando `Tudo certo`.
- [ ] PR aberto para a `develop`, com `Closes #13`.

### Erros comuns

- **`sqlite3.OperationalError: table curso_fic has 9 columns but 8 values were supplied`:** a tabela foi criada antes com outra estrutura. Apague o arquivo `.db` e rode de novo.
- **O total no banco dobra na segunda importação:** a tabela foi criada sem `PRIMARY KEY` no `codigo`, ou você usou `INSERT` sem o `OR REPLACE`.
- **Nada aparece no banco:** faltou o `conexao.commit()` depois do `executemany`.
- **Os tempos da massa de 100 aparecem como `0.0`:** você arredondou demais na hora de mostrar. Mostre em milissegundos, com 2 casas decimais.

<a id="sprint-7-escrita"></a>
## Sprint 7 · Escrita: Resultados e discussão (Entrega 5)

**Prazo: 03/12/2026** · **Versão no Drive:** `Entrega 5 - Resultados - 03/12`

A seção mais importante do TCC. De 4 a 7 páginas, na ordem das seções 4.1 a 4.6, que já estão montadas desde a Sprint 6. Escreva na terça e na quarta; a quinta é para revisar.

> **Regra sem exceção:** toda figura, quadro, tabela ou trecho de código é **anunciado no texto antes** de aparecer e **comentado depois**. Nada de duas figuras seguidas sem texto entre elas.

### 4.1 A ferramenta desenvolvida

**Escreva sobre:** o que foi construído, retomando a arquitetura da seção 3.5 sem repeti-la; o comando de execução e a figura do `main.py` funcionando; os arquivos de saída (relatório, planilha limpa e banco), remetendo ao Apêndice A. Coloque **no máximo dois** trechos curtos de código, cada um com uma explicação: por exemplo, o `converter_data`, que recusa `31/02/2027`, e o `validar_codigos_duplicados`, que informa a linha da primeira ocorrência. Termine com o endereço do repositório.

### 4.2 Decisões de validação

**Escreva sobre:** cada decisão do quadro da Sprint 6, **por que** ela foi tomada e o que aconteceria sem ela (por exemplo: sem a decisão 2, um campo vazio contaria como dois erros e confundiria quem lê o relatório). Depois, a **armadilha das datas do Excel**: o que aconteceu, por que aconteceu e como foi resolvido, com a figura. São achados reais do desenvolvimento, e são exatamente o que dá substância a um trabalho técnico.

### 4.3 Detecção de erros

**Escreva sobre:** as Tabelas 1 e 2, respondendo à **primeira pergunta** do trabalho com número. Mostre que as oito regras foram exercitadas (Tabela 2). **Discuta o resultado de 100%** como explicado no aviso da Sprint 6: ele mostra que as regras foram implementadas corretamente em qualquer tamanho de planilha, mas é esperado, porque os erros foram criados pelo mesmo autor das regras. Se você encontrou e corrigiu defeitos durante a Sprint 6, conte aqui: eles mostram que o método do gabarito funciona.

### 4.4 Tempo de processamento

**Escreva sobre:** a **Tabela 3**, com os dados do `tempos_resumo.csv` (tamanho, tempo médio, desvio padrão, milissegundos por linha e fator de crescimento), e um **gráfico** de linha com o tempo médio em função do número de linhas. Faça o gráfico no Google Planilhas (*Inserir → Gráfico*), a partir do `tempos_resumo.csv`, com os eixos nomeados e com unidade, e sem título dentro da imagem (o título vai na identificação da figura).

Responda à **segunda pergunta:** o tempo cresceu na mesma proporção que o número de linhas? Diga o que isso significa na prática para uma planilha de cursos real, que costuma ter dezenas ou centenas de linhas.

Compare com a conferência manual, **deixando claro que é uma estimativa**: escolha um critério (por exemplo, 10 segundos para uma pessoa conferir as oito colunas de uma linha), diga que esse valor é uma suposição sua, faça a conta para 10.000 linhas e compare com o tempo medido. Esse contraste é um bom argumento, desde que apresentado honestamente.

### 4.5 Importação para o banco de dados

**Escreva sobre:** a **terceira pergunta**, com os números do `resultados\importacao.md` (quantas linhas foram importadas, o total depois da primeira e da segunda importação) e a figura do `ver_banco.py`. Explique as duas barreiras contra códigos repetidos (a regra R02 e a `PRIMARY KEY`), a padronização dos valores (turno, *campus*, datas em `AAAA-MM-DD`) e o uso de parâmetros (`?`) contra injeção de SQL.

### 4.6 Limitações do trabalho

Uma subseção curta e honesta:

- a avaliação usou **dados fictícios**, com erros criados pelo próprio autor das regras; erros reais podem ter padrões diferentes;
- a ferramenta não verifica a **exatidão** (um valor válido, mas errado, passa), como explicado na seção 2.2;
- a ferramenta aceita só o formato de planilha definido no trabalho, e o CSV só com vírgula, enquanto o Excel em português costuma gravar com ponto e vírgula;
- linhas totalmente vazias no meio da planilha são ignoradas na leitura e deslocam a numeração das linhas seguintes no relatório;
- não houve avaliação com as pessoas que preenchem as planilhas, o que exigiria aprovação do Comitê de Ética em Pesquisa;
- a ferramenta não está integrada ao iFIC.

Declarar limitações **fortalece** o trabalho.

### Entrega

- [ ] 4.1 a 4.6 escritas, na ordem.
- [ ] As três perguntas do trabalho respondidas com número (4.3, 4.4 e 4.5).
- [ ] Tabela 3 e o gráfico de tempo, no padrão ABNT.
- [ ] O resultado de 100% discutido com a ressalva.
- [ ] Estimativa manual apresentada como estimativa, com o critério.
- [ ] Toda figura e tabela anunciada antes e comentada depois.
- [ ] Comentários das entregas anteriores resolvidos.

Nomeie a versão no Drive, comente na *issue* e mova o card para **Em revisão**.

---

# Sprint 8: 04/12 a 10/12/2026

<a id="sprint-8-dev"></a>
## Sprint 8 · Desenvolvimento: README, reprodução do zero e fechamento

**Branch:** `feature/fechamento` · **No PR:** `Closes #15`

**O que você vai entregar:** o repositório pronto para outra pessoa usar: o `README.md` completo, o código revisado e a prova de que o trabalho roda do zero, em uma pasta limpa, com os mesmos resultados. Nesta semana o código **não ganha nada novo**: é acabamento. A maior parte do tempo vai para a escrita.

**Divida a semana assim:**

| Dia | O que fazer |
|---|---|
| sexta, 04/12 | passos 1 e 2 (*branch* e `README.md`) |
| segunda, 07/12 | passo 3 (revisão do código) |
| terça, 08/12 | passos 4 e 5 (reprodução do zero e conferência do repositório) |
| quarta e quinta, 09 e 10/12 | passo 6 (diário e PR); o resto do tempo é para a escrita |

### Passo a passo

**1. Crie a *branch*** `feature/fechamento` a partir da `develop` atualizada.

**2. Complete o `README.md`.** Ele é parte da entrega: alguém que nunca falou com você tem de conseguir usar e reproduzir o trabalho só com ele. A política do NIC (item 11) pede instruções de instalação, guia de execução, arquitetura básica e responsáveis. Organize assim:

1. **o que é:** a ferramenta, em três linhas, e o vínculo com o projeto iFIC;
2. **responsáveis:** você e os orientadores;
3. **requisitos e instalação:** Python, criação do `.venv` e `pip install -r requirements.txt`;
4. **como usar:** o comando do `main.py`, com um exemplo real e a saída resumida; o que aparece em `saida\`;
5. **as regras:** a tabela das oito regras e um link para as decisões ([seção 3.4](#regras) deste documento);
6. **arquitetura:** o que cada arquivo de `src\` faz, em uma linha cada;
7. **como reproduzir os resultados:** os comandos do gerador (com o SHA-256 da massa de 1.000), do comparador e do `medir_tempo.py`, na ordem, e o que cada um grava em `resultados\`;
8. **como rodar a verificação;**
9. **resultados principais:** precisão, revocação e os tempos médios da rodada oficial, com a data e a máquina;
10. **limitações**, em três linhas.

**3. Revise o código** como se outra pessoa fosse mantê-lo:

- no topo de cada arquivo de `src\`, um comentário dizendo o que ele faz e, se for um programa, como rodar;
- em cada função, uma *docstring* de uma linha dizendo o que ela recebe e o que devolve (*pesquise:* *docstring* em Python);
- apague código comentado, `print` de depuração e funções que não são mais usadas;
- se você precisar de mais de 10 segundos para entender uma função, melhore os nomes ou divida a função;
- nenhum caminho fixo da sua máquina (`C:\Users\...`): todos os caminhos são relativos à raiz do repositório.

Depois de cada arquivo revisado, rode a verificação. Faça os *commits* com o prefixo `refactor:` ou `docs:`.

**4. Reproduza o trabalho do zero**, em uma pasta nova, como se fosse outra pessoa. Primeiro, envie tudo (`git push`). Depois, no Prompt de Comando:

```
cd %TEMP%
git clone https://github.com/nicifrn/mp_ific_importador_planilhas.git repro
cd repro
git checkout feature/fechamento
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
```

Siga **só o que está no `README.md`** para gerar as três massas, rodar o comparador e a verificação. Confira:

- o SHA-256 da massa de 1.000 é o mesmo do `README.md`;
- o `resultados\deteccao.csv` gerado é igual ao do seu repositório. Compare com o comando `fc`:

  ```
  fc resultados\deteccao.csv %USERPROFILE%\mp_ific_importador_planilhas\resultados\deteccao.csv
  ```

  O `fc` deve dizer que não encontrou diferenças;
- a verificação mostra `Tudo certo`.

Se precisou fazer algo que **não está** no `README.md`, o `README.md` está incompleto: corrija-o. Anote no diário como foi a reprodução. Depois, apague a pasta `repro`.

**5. Confira o que está no repositório.** O comando abaixo não pode mostrar nada:

```
git ls-files | findstr /i ".db massas saida .venv"
```

Se mostrar algum arquivo, remova-o do Git (sem apagar do disco) com `git rm --cached <arquivo>` e confira o `.gitignore`.

**6. Finalize.** No diário, escreva a Sprint 8 e, no fim, um parágrafo curto com as três maiores dificuldades do trabalho inteiro e o que você aprendeu com cada uma: ele vai direto para a Conclusão. *Commits* e PR para a `develop` com `Closes #15`. Em **Evidências de teste**, cole a saída da verificação e o resultado do `fc` na reprodução do zero.

### Como saber que deu certo

- Outra pessoa, lendo só o `README.md`, conseguiria instalar, usar e reproduzir tudo.
- A reprodução em pasta limpa deu o mesmo SHA-256 e o mesmo `deteccao.csv`.
- O `git ls-files` não mostra nenhum `.db`, massa, saída ou `.venv`.
- `python testes\verificar_validadores.py` mostra `Tudo certo`.

### Entrega da sprint

- [ ] `README.md` com as dez partes.
- [ ] Comentário no topo de cada arquivo, *docstrings*, sem código morto e sem caminho fixo.
- [ ] Reprodução do zero feita e anotada no diário (SHA-256 e `fc` iguais).
- [ ] `git ls-files` limpo.
- [ ] Diário da última sprint, com as três maiores dificuldades.
- [ ] PR aberto para a `develop`, com `Closes #15`.

<a id="sprint-8-escrita"></a>
## Sprint 8 · Escrita: Conclusão, Objetivos específicos, Resumo e Título (Entrega 6)

**Prazo: 10/12/2026 (Conclusão até segunda, 07/12)** · **Versão no Drive:** `Entrega 6 - Versão final - 10/12`

### Passo a passo

**1. Passe para o passado** o bloco 4 da Introdução ("será desenvolvida" vira "foi desenvolvida") e confira se a Metodologia descreve o que realmente aconteceu nas Sprints 6 e 7. Corrija o que mudou.

**2. Confira todos os números** do texto com os arquivos de `resultados\`. Os números da Introdução, dos Resultados, da Conclusão e do Resumo têm de ser **exatamente** os mesmos.

**3. Escreva a Conclusão até segunda, 07/12** (1 a 2 páginas), nesta ordem. Nada de resultado novo, conceito novo ou citação nova aqui.

1. **Retomada do problema e do objetivo geral.** Uma ou duas frases lembrando o problema e afirmando que o objetivo foi alcançado, **com o motivo e com número**: "[...] foi alcançado, uma vez que a ferramenta encontrou todos os X erros das planilhas de teste e conferiu 10.000 linhas em Y segundos."
2. **As três perguntas do trabalho,** cada uma respondida em uma ou duas frases, com o número.
3. **Um comentário por objetivo específico,** na mesma ordem em que foram apresentados, cada um apontando o resultado que o comprova.
4. **Contribuições:** a ferramenta com o código disponível; o conjunto documentado de regras e decisões para planilhas de cursos FIC; e o método de avaliação com base fictícia e gabarito, que outros trabalhos do iFIC podem reaproveitar.
5. **Limitações e dificuldades:** retome as limitações da seção 4.6 e conte as dificuldades do desenvolvimento, usando o parágrafo final do seu diário (por exemplo: os acentos, a armadilha das datas do Excel, a construção do gabarito). Dificuldade técnica bem relatada é conteúdo, não confissão de fraqueza.
6. **Trabalhos futuros**, ligados ao que você encontrou: integrar a ferramenta ao painel administrativo do iFIC, como importação em lote (a principal); aceitar CSV com ponto e vírgula; permitir configurar as regras em um arquivo, sem mudar o código; avaliar com planilhas reais, sem dados pessoais; sugerir correções para erros de digitação comuns; e avaliar o relatório com as pessoas que preenchem as planilhas (o que exige o Comitê de Ética).

**Não escreva** que a ferramenta "resolve o problema" ou "elimina os erros". Ela encontra os erros das oito categorias previstas, e isso foi medido em planilhas fictícias. Diga exatamente isso.

**4. Escreva a seção 1.2 Objetivos específicos.** De 4 a 5 itens, com verbo no infinitivo, letra minúscula no início, ponto e vírgula no fim de cada um e ponto no último, **na mesma ordem** das etapas da Metodologia e das seções dos Resultados. Cada um tem de ser verificável. O esqueleto do que eles devem cobrir (escreva com as suas palavras):

1. a revisão da literatura sobre cursos FIC, qualidade de dados e validação de dados;
2. a especificação das regras de validação das planilhas de cursos FIC;
3. a implementação da ferramenta de validação e importação;
4. a construção de uma base de teste fictícia com gabarito de erros;
5. a avaliação da ferramenta quanto à detecção de erros, ao tempo de processamento e à importação.

Não confunda objetivo com funcionalidade: "validar a carga horária" é uma regra da ferramenta, não um objetivo do trabalho.

**5. Escreva o Resumo,** depois de todo o texto pronto. Parágrafo único, de 150 a 500 palavras, impessoal, sem citações e sem mencionar figuras ou tabelas (NBR 6028:2021). Uma ou duas frases para cada item: **contexto e problema** (planilhas de cursos FIC preenchidas à mão e o risco de importar dados com erro); **objetivo**; **método** (regras em Python, planilhas fictícias de 100, 1.000 e 10.000 linhas com gabarito); **resultados**, com os números principais (precisão, revocação, tempo e importação); e **conclusão** (o que os resultados permitem afirmar e a integração ao iFIC como próximo passo).

Escreva o **Abstract**, a versão em inglês. Se usar tradutor automático, revise frase por frase.

Escolha de 3 a 5 **palavras-chave**, separadas por ponto e vírgula, que alguém usaria para **encontrar** o trabalho (por exemplo, no campo de: qualidade de dados; validação de dados; cursos FIC; Python; sistemas de informação).

**6. Proponha o Título.** Ele diz **o que foi feito** e **em que contexto**, e combina com o objetivo geral. Cubra: a ação (validação e importação), o objeto (planilhas de cursos FIC) e o contexto (IFRN ou plataforma iFIC). Se a ferramenta tem nome, ele pode abrir o título, no formato "Nome: descrição do que faz". Evite títulos genéricos ("Um sistema em Python"). Escreva a proposta no comentário da *issue*.

**7. Faça a revisão final**, item por item. Percorra o checklist de `docs/orientacoes_escrita_tcc.md` e mais estes pontos:

- [ ] O objetivo geral, os objetivos específicos, as seções dos Resultados e os parágrafos da Conclusão estão na mesma ordem e falam das mesmas coisas?
- [ ] Os números do Resumo, dos Resultados e da Conclusão são exatamente os mesmos?
- [ ] Toda tabela, quadro e figura é anunciada antes, comentada depois e tem identificação em cima e fonte embaixo?
- [ ] Todas as citações estão nas Referências, e todas as Referências estão citadas no texto?
- [ ] Nenhuma referência entrou sem ter sido lida?
- [ ] As siglas (FIC, IFRN, LGPD, CSV, XLSX, ETL, SQL) estão por extenso na primeira vez que aparecem?
- [ ] Está dito que os dados são fictícios e que não houve participantes humanos?
- [ ] O endereço do repositório está no texto?
- [ ] Todos os comentários das versões anteriores foram resolvidos?

### Entrega

- [ ] Introdução e Metodologia no passado e conferidas.
- [ ] Números conferidos com `resultados\`.
- [ ] Conclusão até 07/12, na ordem indicada.
- [ ] 1.2 Objetivos específicos.
- [ ] Resumo, *Abstract* e palavras-chave.
- [ ] Título proposto no comentário da *issue*.
- [ ] Revisão final feita.

Nomeie a versão no Drive, comente na *issue* com a proposta de título e mova o card para **Em revisão**.

---

<a id="quadro-resumo"></a>
## 10. Quadro-resumo

| Sprint | Período | Desenvolvimento | Conceitos novos | Escrita |
|---|---|---|---|---|
| 1 | 16/10 a 22/10 | Git, ambiente e exercícios ex01 a ex08 | Git, `venv`, variáveis, `if`, `for`, `while`, listas | Preparar o documento e fichar |
| 2 | 23/10 a 29/10 | Exercícios ex09 a ex17 e o minivalidador | funções, `try`, dicionários, `csv`, módulos, `datetime`, `assert` | Entrega 1: Introdução e Objetivo geral |
| 3 | 30/10 a 05/11 | Leitura do CSV e conferência do cabeçalho | `sys.argv`, *list comprehension*, tuplas | Entrega 2: Referencial 2.1 e 2.2 |
| 4 | 06/11 a 12/11 | As oito regras | `set`, `sort` com `lambda` | Entrega 3: Referencial 2.3 e 2.4 |
| 5 | 13/11 a 19/11 | XLSX, relatório e planilha limpa | `openpyxl`, `zip`, `DictWriter`, `Counter` | Entrega 4: Metodologia |
| 6 | 20/11 a 26/11 | Gerador com gabarito e comparador | `random.seed`, `timedelta`, conjuntos, precisão e revocação | Tabelas da Sprint 6 no documento |
| 7 | 27/11 a 03/12 | Tempo e importação para o SQLite | `perf_counter`, `statistics`, `sqlite3`, `INSERT OR REPLACE` | Entrega 5: Resultados |
| 8 | 04/12 a 10/12 | README, revisão e reprodução do zero | *docstrings*, reprodutibilidade | Entrega 6: Conclusão, Objetivos específicos, Resumo e Título |

---

<a id="riscos"></a>
## 11. Riscos e planos B

| Risco | O que fazer |
|---|---|
| O comando `python` não funciona | Use `py` em todos os comandos e anote no diário. |
| O `pip install openpyxl` não funciona na rede do laboratório | Tente de novo em outro horário e em outra máquina. Se não der até quarta-feira da Sprint 5, baixe em casa, no site do PyPI (https://pypi.org), os arquivos `.whl` do `openpyxl` e do `et_xmlfile` (a dependência dele), traga em um *pen drive* e instale com `pip install <arquivo>.whl`. Se ainda assim não der, avise na *issue*: a ferramenta fica só com CSV, e a leitura de XLSX vira trabalho futuro no TCC. |
| A máquina do laboratório foi formatada | O que foi enviado ao GitHub está salvo: clone de novo e recrie o `.venv`. As massas se geram de novo pela semente. É por isso que o `git push` é diário. |
| A máquina das medições não está disponível na segunda, 30/11 | Faça a rodada oficial em outra máquina, anote a configuração dela no `docs/ambiente.md` e no diário, e use **só** essa rodada no TCC. |
| O comparador não dá 100% | Não é um problema do cronograma: é um resultado. Investigue, corrija o que for defeito, gere de novo e conte no TCC. Se não conseguir explicar até a quinta, entregue com o número que deu e a explicação do que você já descobriu. |
| Você atrasou | Corte, nesta ordem: o `GROUP BY` do `ver_banco.py`; o `deteccao_por_regra.csv`; o limite de 20 erros na tela do `main.py`. **Nunca corte:** as oito regras, o gerador com gabarito, o comparador, a medição de tempo e a verificação. Sem eles, não há Resultados. |

---

<a id="extras"></a>
## 12. Se sobrar tempo

Só comece alguma destas tarefas com a sprint da semana **entregue**. É melhor um escopo fechado bem-feito do que um escopo maior pela metade. Cada uma vira uma frase a mais nos Resultados.

- **E1. Aceitar CSV com ponto e vírgula.** *Pesquise* `csv.Sniffer`, que descobre o separador sozinho. Resolve a limitação do Excel em português.
- **E2. Medir o tempo da leitura de XLSX.** Converta as três massas com o `converter_para_xlsx.py` e meça com o `medir_tempo.py`. A diferença para o CSV costuma ser grande, e explicá-la (o XLSX é um pacote compactado de arquivos XML) é uma boa discussão.
- **E3. Sugerir a correção.** Quando o turno for `Noturnoo`, sugerir `noturno` na mensagem. *Pesquise* `difflib.get_close_matches`, que já vem com o Python.
- **E4. Relatório em XLSX,** com as células com erro pintadas de vermelho. *Pesquise* `PatternFill` no `openpyxl`.

---

<a id="referencias-minimas"></a>
## 13. Referências mínimas do TCC

Garanta que o trabalho tenha pelo menos:

- **2 obras de metodologia científica:** Gil (2017) e Prodanov e Freitas (2013), já indicadas em `docs/orientacoes_escrita_tcc.md`;
- **a legislação:** Lei nº 11.892/2008, Lei nº 9.394/1996 (LDB), Decreto nº 5.154/2004 e Lei nº 13.709/2018 (LGPD);
- **1 ou 2 documentos ou artigos sobre os cursos FIC:** o Guia Pronatec de Cursos FIC e um artigo sobre a oferta nos Institutos Federais;
- **1 obra sobre sistemas de informação:** Laudon e Laudon;
- **3 obras sobre qualidade de dados e erros em planilhas:** Wang e Strong (1996), Batini e Scannapieco, Panko e o documento da DAMA UK;
- **2 obras sobre validação, ETL e banco de dados:** Rahm e Do (2000), Kimball e Caserta (2004) ou Elmasri e Navathe;
- **3 a 5 trabalhos relacionados**, sendo pelo menos dois acadêmicos;
- **a documentação oficial das ferramentas:** Python, `openpyxl` e SQLite;
- **o portal do IFRN**, com a data de acesso, como fonte da lista de *campi*.

Anote a referência completa de **tudo** o que ler, desde o primeiro dia.
