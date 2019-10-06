' Módulo de ordenamento das tabelas de movimentos e cartões
Option Explicit

Sub OrdenarMovimentos()
  '
  ' OrdenarMovimentos Macro
  ' Ordena tabelas de Movimentos e Cartões em ordem crescente de data.
  '
  ' Atalho do teclado: Ctrl+o
  '
  On Error GoTo ErroOrdena
  If Not IsPlanilhaAberta(range(RANGE_SITUAC_PLANILHA)) Then
    Exit Sub
  End If
  Application.ScreenUpdating = False
  Application.EnableEvents = False
  range(RANGE_TAB_MOVIMENTACOES).Select
  'Selection.Sort Key1:=Range(RANGE_PRIMEIRA_DATA_MOVIMENTACOES), Order1:=xlAscending, Header:=xlGuess, _
  '      OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, _
  '      DataOption1:=xlSortNormal
  ActiveSheet.Sort.SortFields.Clear
  ActiveSheet.Sort.SortFields.Add Key:=range(RANGE_COL_DATA_MOVIMENTACOES), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
  With ActiveSheet.Sort
      .SetRange range(RANGE_TAB_MOVIMENTACOES)
      .Header = xlGuess
      .MatchCase = False
      .Orientation = xlTopToBottom
      .SortMethod = xlPinYin
      .Apply
  End With
  
  range(RANGE_TAB_CARTOES).Select
  ActiveSheet.Sort.SortFields.Clear
  ActiveSheet.Sort.SortFields.Add Key:=range(RANGE_COL_DATA_CARTOES), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
  With ActiveSheet.Sort
      .SetRange range(RANGE_TAB_CARTOES)
      .Header = xlGuess
      .MatchCase = False
      .Orientation = xlTopToBottom
      .SortMethod = xlPinYin
      .Apply
  End With
        
FimOrdena:
  Application.ScreenUpdating = True
  Application.EnableEvents = True
  ' seleciona a coluna de datas
  RetornarUltimaCelulaMovimentacoes.Select
  Exit Sub
  
ErroOrdena:
  MostrarMsgErro ("OrdenarMovimentos")
  Resume FimOrdena
End Sub



