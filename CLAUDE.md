# CLAUDE.md

Código VBA da planilha Excel **Expense** (antes chamada Expensive, nome que o repositório mantém): controle pessoal de movimentações bancárias, cartões de crédito, investimentos, orçamento e metas financeiras. O repositório guarda só os módulos exportados. A pasta de trabalho (`.xlsm`) fica fora do repositório, junto com as fórmulas, os nomes definidos, a formatação e os atalhos de teclado.

## Como o código chega ao Excel

- Não há build, testes automatizados nem CI. Para validar uma alteração, importe ou cole o módulo no editor VBA (Alt+F11) da pasta de trabalho, rode **Depurar > Compilar VBAProject** e teste a macro na planilha.
- Os `.bas` não têm a linha `Attribute VB_Name`, e o nome do módulo no Excel é o nome do arquivo. Não acrescente atributos sem necessidade.
- Os arquivos estão em UTF-8. O editor VBA importa em ANSI (Windows-1252), então os acentos podem precisar de conferência depois da importação.
- O editor VBA espera quebras de linha CRLF: um `.frm` ou `.cls` com LF pode ser importado de forma incorreta. O repositório usa `core.autocrlf=true`.
- `UserFormTolerancia.frm` e o binário `UserFormTolerancia.frx` andam juntos: não edite o `.frx` e não altere o cabeçalho `Begin ... End` do `.frm` à mão.

## Estrutura da pasta de trabalho

- **Planilhas mensais `Jan` a `Dez`**: todas têm o mesmo layout. A célula `E4` guarda a situação (`Aberto`/`Fechado`), e as rotinas só agem em planilha `Aberto` (`IsPlanilhaAberta`). Blocos principais:
  - movimentações em `D16:G96` (data, descrição, tipo, valor) e cartões em `J16:N96` (data, descrição, tipo, cartão, valor);
  - portfólio em `D104:N142`;
  - carteiras de Ações, FII, Tesouro Direto, Tesouro Selic, ETF, Stock, REIT, Treasury, Ouro e Cripto a partir da linha 147;
  - saldos das contas nas corretoras (XP, Avenue) em `B22:B30`.
- **Outras planilhas**: `Intro` (data da última atualização), `Alocacao` (cadastro de investimentos AdHoc/Portfólio), `Acoes365` (cotações por tipos de dados vinculados do Excel), `Orcamento`, `Retorno`, `Retrato` e planilhas de gráfico `graf*`.
- Todos os endereços de célula e nomes definidos (`orcTolerancia`, `retPlanFechada`, `intPosicao`) ficam em `Constants.bas`. Quando o layout da planilha muda, ajuste as constantes ali e não espalhe endereços pelo código.

## Módulos

| Arquivo | Responsabilidade |
|---|---|
| `Constants.bas` | Constantes globais: nomes de planilhas, ranges e erro de negócio `ERRO_TABELA_CHEIA` |
| `Comum.bas` | Utilitários compartilhados (`CongelarCalculosPlanilha`, `RetornarProximaPlanilha`, `ColetarAtivos`, `MostrarMsgErro` etc.) e UDFs |
| `EstaPasta_de_trabalho.cls` | Eventos da pasta: `Open`, `BeforeSave`, `BeforeClose` e `SheetChange` (alerta de tolerância do orçamento, crítica da descrição de investimento, aviso de imposto/DARF) |
| `FechaMes.bas` | Fecha ou reabre o mês e copia os saldos das carteiras para o mês seguinte |
| `Cartao.bas` | Move um lançamento de cartão para o mês seguinte |
| `Parcelamento.bas` | Cria parcelas `descrição (n/x)` nos meses seguintes |
| `Ordenacao.bas` | Ordena movimentos e cartões por data e as carteiras por ticket |
| `Investimento.bas` | UDFs de saldo e rendimento, crítica de investimento e lembrete de DARF no Outlook |
| `Web.bas` | Atualiza as cotações a partir de `Acoes365` (`RefreshLinkedDataType`) |
| `Relatorio.bas` | Relatório de rendimentos (visualização de impressão) e retrato do mês em PDF |
| `Avaliacao.bas` / `UserFormTolerancia.frm` | Questionário de perfil de tolerância a risco |

### Pontos de entrada que a pasta de trabalho referencia

Renomear qualquer um deles quebra a pasta de trabalho, então não renomeie sem avisar.

- **Macros com atalho** (configurado no Excel, não no código): `MoverCartao` (Ctrl+M), `PosicionarTopo` (Ctrl+T), `PuxarDataAtual` (Ctrl+D), `ReduzirTotalParaUnitario` (Ctrl+U), `AvaliarRisco` (Ctrl+A), `ProtegerPlanilha` (Ctrl+B), `OrdenarMovimentos` (Ctrl+O), `MontarParcelamento` (Ctrl+P).
- **Macros chamadas por botões ou menu**: `AtualizarDadosWeb`, `GerarRelatRend`, `GerarRelatRetrato`, `OrdenarCarteirasPlanilhaAtiva`.
- **UDFs usadas em fórmulas**: `MaxDrawdown`, `TotalReturn`, `SheetOffSet`, `CalcularSaldoAtual`, `CalcularRendAtivo`.

## Convenções de código

- `Option Explicit` em todos os módulos e indentação de **2 espaços**.
- Nomes em português, com prefixo húngaro (KB 110264): `str`, `lng`, `int`, `dbl`, `bln`, `dat`, `var`, `rg`, `ws`, `obj`, `dic`, `udt`, e `m` para variáveis de módulo. Procedimentos começam com verbo (`Retornar...`, `Calcular...`, `Copiar...`), e as funções booleanas usam `Is...` ou `Has...`.
- Constantes `MAIUSCULAS_COM_UNDERSCORE`: as globais ficam em `Constants.bas` e as locais como `Private Const` no topo do módulo. Evite números e textos soltos no meio do código.
- Cabeçalho de módulo no formato `' Módulo: X.bas / ' Descrição: ...`, e comentário curto antes de cada procedimento público.
- **Tratamento de erro** de uma macro:
  ```vba
  On Error GoTo ErroNomeProc
  ...
  FimNomeProc:
    ' restaura o ambiente (On Error Resume Next se necessário)
    Exit Sub
  ErroNomeProc:
    MostrarMsgErro "NomeProc"
    Resume FimNomeProc
  ```
- **UDFs nunca abrem `MsgBox`**: em caso de erro, devolvem `CVErr(xlErrValue)` ou `CVErr(xlErrRef)`.
- **Desempenho e estado do Excel**: envolva as operações em massa com `CongelarCalculosPlanilha True/False`, que admite aninhamento. Use uma flag `blnCongelou` para liberar apenas o que foi congelado. Leia ranges em bloco para um `Variant` em vez de célula a célula.
- **Eventos**: `Workbook_SheetChange` desliga `EnableEvents` durante o próprio processamento. As rotinas que gravam em outra planilha precisam decidir de forma consciente se os eventos ficam ligados. `MoverCartao`, por exemplo, liga os eventos para marcar a alteração.
- Valores lidos de células passam por `RetornarValorNumerico` ou por um teste `IsError`/`IsEmpty` antes do uso, porque células com erro de fórmula são comuns.
- Para avançar entre meses, use `RetornarProximaPlanilha`, que devolve `Nothing` em `Dez`, na última planilha ou numa planilha de gráfico. Quando a tabela de destino estiver cheia, gere `ERRO_TABELA_CHEIA` via `RetornarLinhaLivre...`.
- Mensagens ao usuário em português, com título de `MsgBox` em constante `TITULO_...`.
- Late binding para automação externa (`CreateObject("Outlook.Application")`, `"Scripting.Dictionary"`, `"WScript.Shell"`), sem referências adicionais no projeto.

## Git

- Branches `hotfix/<assunto>` a partir de `master`, integradas por Pull Request no GitHub (`maumss/expensive`) com merge `--no-ff`.
- Mensagens de commit curtas em português (por exemplo `fix: Ajusta cotações` ou `Corrige atualização de saldos`).
