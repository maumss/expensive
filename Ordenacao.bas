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
  If Not IsPlanilhaAberta(Range(RANGE_SITUAC_PLANILHA)) Then
    Exit Sub
  End If
  Application.ScreenUpdating = False
  Application.EnableEvents = False
  Range(RANGE_TAB_MOVIMENTACOES).Select
  'Selection.Sort Key1:=Range(RANGE_PRIMEIRA_DATA_MOVIMENTACOES), Order1:=xlAscending, Header:=xlGuess, _
  '      OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, _
  '      DataOption1:=xlSortNormal
  ActiveSheet.Sort.SortFields.Clear
  ActiveSheet.Sort.SortFields.Add Key:=Range(RANGE_COL_DATA_MOVIMENTACOES), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
  With ActiveSheet.Sort
      .SetRange Range(RANGE_TAB_MOVIMENTACOES)
      .Header = xlGuess
      .MatchCase = False
      .Orientation = xlTopToBottom
      .SortMethod = xlPinYin
      .Apply
  End With
  
  Range(RANGE_TAB_CARTOES).Select
  ActiveSheet.Sort.SortFields.Clear
  ActiveSheet.Sort.SortFields.Add Key:=Range(RANGE_COL_DATA_CARTOES), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
  With ActiveSheet.Sort
      .SetRange Range(RANGE_TAB_CARTOES)
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

Sub PuxarDataAtual()
  '
  ' dataAtual Macro
  ' Traz a data atual para a coluna de movimentos ou cartão.
  '
  ' Atalho do teclado: Ctrl+d
  ' Criado por: Mauricio SS  Em: 14/02/04
  '
  If Not IsPlanilhaAberta(Range(RANGE_SITUAC_PLANILHA)) Then
    Exit Sub
  End If
  ' variáveis
  Dim rgAlvo As Range
  Dim wsPlanilha As Worksheet
  ' principal
  On Error GoTo ErroData
  Set wsPlanilha = ActiveSheet
  Set rgAlvo = Selection
  If IsEmpty(rgAlvo) Then
    If (Not Application.Intersect(rgAlvo, Range(RANGE_COL_DATA_MOVIMENTACOES)) Is Nothing) Or _
       (Not Application.Intersect(rgAlvo, Range(RANGE_COL_DATA_CARTOES)) Is Nothing) Or _
       (Not Application.Intersect(rgAlvo, Range(RANGE_COL_DATA_ACOES)) Is Nothing) Or _
       (Not Application.Intersect(rgAlvo, Range(RANGE_COL_DATA_CART_OPCOES)) Is Nothing) Or _
       (Not Application.Intersect(rgAlvo, Range(RANGE_COL_DATA_FII)) Is Nothing) Or _
       (Not Application.Intersect(rgAlvo, Range(RANGE_COL_DATA_RF)) Is Nothing) Or _
       (Not Application.Intersect(rgAlvo, Range(RANGE_COL_DATA_SELIC)) Is Nothing) Then
      rgAlvo.Value = Date
    End If
  End If
  Exit Sub
  
ErroData:
  MostrarMsgErro ("PuxarDataAtual")
End Sub

