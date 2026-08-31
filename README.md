# robo-taxas-biopsia

Automação que insere a taxa do **KIT AGULHA PARA PUNCAO PERCUTANEA DIRECIONADA** (`70016941`) nas
guias de processos **radiológicos** do sistema ISSEC (Selenium + Tkinter), distribuída no esquema
**lançador + auto-atualização**. Mesmo fluxo de navegação do `robo-oftalmo`.

## Regra de negócio

O robô percorre todas as guias do processo. Em cada guia, procura os procedimentos:

| Código | Procedimento | Taxa |
|---|---|---|
| `3213004A` | BIOP.PERC.ORIENT.POR CT,US/RX - TIREOIDE | `70016941` |
| `3213004B` | BIOP.PERC.ORIENT.POR CT,US/RX - MAMA | `70016941` |
| `3213004C` | BIOP.PERC.ORIENT.POR CT,US/RX - CORE BIOPSIA | `70016941` |
| `3213004D` | BIOP.PERC.ORIENT.POR CT,US/RX - MARCACAO | `70016941` |

**Quantidade do KIT = soma da coluna Qtd** das linhas de biópsia da guia. Uma guia com
`3213004B` (Qtd 2) + `3213004C` (Qtd 1) recebe **uma única** linha de taxa `70016941` com
**Qtd 3** (Qtd e Qtd Cobrada).

Se o KIT já estiver lançado na guia:

- com a quantidade correta → nada é alterado;
- com quantidade diferente → o robô **corrige** a quantidade;
- com quantidade ilegível, ou se a verificação falhar → **nada é alterado** (reportado na
  planilha). Na dúvida, não mexe.

Procedimentos novos entram editando `MAPEAMENTO_PROCED_TAXA`, no topo de `issec_taxa_biopsia.py`.
⚠️ Digitar os códigos em **ASCII puro** — o material de origem costuma vir com homóglifos gregos
(`Α` U+0391 no lugar do `A`), que nunca casam na comparação.

## Como funciona a distribuição

Cada usuário tem um único `Taxas_Biopsia.exe` (lançador estável). Ao abrir, ele baixa a versão
mais recente de `issec_taxa_biopsia.py` deste repositório (privado) e a executa. Atualizar a
lógica = `git push` no `issec_taxa_biopsia.py`. Só é preciso reconstruir o `.exe` se mudarem as
**dependências/imports** da automação.

## Arquivos

- `issec_taxa_biopsia.py` — a automação (baixada e executada pelo lançador em runtime).
- `lancador_Taxas_Biopsia.py` — código-fonte do lançador (espelha os imports da automação +
  auto-update). Usado para reconstruir o `.exe` com PyInstaller.
- `config.ini.exemplo` — modelo do `config.ini` que vai ao lado do `.exe`.
- `_secrets.py.exemplo` — modelo do `_secrets.py` (Supabase), opcional.

## Pasta do usuário (não versionada)

Montar com os 3 itens juntos:

- `Taxas_Biopsia.exe`
- `config.ini` (copiado do `config.ini.exemplo`, com o token do GitHub real). ⚠️ O nome tem que
  ser exatamente `config.ini`. É o único segredo visível ao usuário — use token fine-grained,
  Contents Read-only, restrito só a este repositório.
- `acesso.xlsx` (planilha com colunas `usuario` / `senha` do ISSEC).

As pastas `logs/` e `relatorios/` são criadas automaticamente no primeiro uso.

## Saídas

- `logs/biopsia_AAAA_MM.log` — log da automação (a primeira linha traz a `VERSAO` e a `origem`).
- `logs/lancador.log` — resultado de cada tentativa de download e a versão do build.
- `relatorios/Relatorio_Taxas_Biopsia_Processo_<numero>.xlsx` — **uma linha por guia**, com
  status, procedimentos encontrados, quantidade apurada e quantidade lançada.
- Supabase: tabela `execucoes_taxas_biopsia` (uma linha por execução).

## Tabela do Supabase

Rodar uma vez no SQL Editor antes do primeiro uso:

```sql
create table public.execucoes_taxas_biopsia (
  id                bigserial not null,
  numero_processo   character varying(50) not null,
  qtd_guias         integer not null default 0,
  qtd_taxas         integer not null default 0,   -- total de KITs lancados (soma das quantidades)
  qtd_guias_taxadas integer not null default 0,   -- guias que receberam insercao
  qtd_corrigidas    integer not null default 0,   -- guias em que a quantidade foi corrigida
  codigos_taxas     jsonb null default '{}'::jsonb,
  data_execucao     timestamp without time zone not null default now(),
  criado_em         timestamp without time zone not null default now(),
  constraint execucoes_taxas_biopsia_pkey primary key (id)
) TABLESPACE pg_default;

alter table public.execucoes_taxas_biopsia enable row level security;

create policy "insert anon" on public.execucoes_taxas_biopsia
  for insert to anon with check (true);
```

## Token do GitHub

Fine-grained PAT, **Contents: Read-only**, restrito a este repositório, validade longa. Vai em
`config.ini` → `[github] token = ...`. Ao expirar, troca-se só o `config.ini` (sem rebuildar).

## Reconstruir o `.exe`

Requer `truststore` instalado (`pip install truststore`). **Sem ele o `.exe` não valida o
certificado em rede corporativa que intercepta HTTPS**: o download do GitHub falha sempre e a
máquina congela na versão que estiver em cache.

```
python -m PyInstaller --onefile --name Taxas_Biopsia --collect-all selenium --hidden-import truststore lancador_Taxas_Biopsia.py --noconfirm
```

Conferir depois do build, como no resto da frota:

- `grep -a truststore dist/Taxas_Biopsia.exe` → tem que responder;
- `grep -a _create_unverified_context dist/Taxas_Biopsia.exe` → **não** pode responder. Se a
  verificação TLS falhar, o lançador cai no cache; nunca baixa código por canal inseguro.

Conferir o repositório de origem exige outro método: `OWNER`/`REPO` são literais do `.pyc` do
lançador, que vai **comprimido em zlib** dentro do PKG — `grep -a` no `.exe` não alcança e responde
`0` mesmo quando o nome está certo. Um build apontando para o repositório errado dá 404 em toda a
frota, então vale conferir descomprimindo os streams do PKG e procurando `robo-taxas-biopsia`
neles.

## Layout de `#taxas_table`

Confirmado em 28/08/2026 pelo log de produção (processo `2600031605`):

| 0 | 1 | 2 | 3 | 4, 5, 6 | 7 | 8 |
|---|---|---|---|---|---|---|
| ordem | código + descrição | **Qtd Cobrada** | **Qtd** | valores | `-` | Editar/Excluir |

As constantes `IDX_COL_QTD_COB_TAXA = 2` e `IDX_COL_QTD_TAXA = 3` refletem isso. A primeira versão
lia a coluna 6 achando que era a quantidade — é o valor total (`174,80`), que não é dígito, e por
isso toda guia caía em "quantidade ilegível" sem alterar nada.

Uma guia só é considerada correta quando **Qtd e Qtd Cobrada** batem com o esperado: uma taxa
lançada à mão com 2/1 vai para o caminho de correção.

## Ponto ainda não confirmado

O fluxo de **edição** de uma taxa (usado só quando a quantidade diverge) não foi exercido em
produção até agora. Para não depender de chute, o código descobre o que precisa em runtime:

- botão de editar: casado pelo **texto** `Editar` da própria célula;
- campos de quantidade: `_descobrir_campos_qtd_taxa()` lista os inputs visíveis do modal e escolhe
  os que terminam em `QT_SERV_HOSP` / `QT_SERV_HOSP_C` — funciona tanto se o SISWEB usar
  `novo_item_taxas_*` quanto `edit_item_taxas_*`. Os ids escolhidos vão para o log na primeira vez;
- botão de salvar: tenta o XPath do `modal_novo_item` (o mesmo que `adicionar_taxa` já usa em
  produção) e, se não houver, a classe `btn_2_novo_item`.

Se ainda assim falhar, o log traz o `outerHTML` da linha e os inputs visíveis do modal.
