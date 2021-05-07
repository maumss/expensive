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
  Debug.Print "Ordenando movimentos..."
  Application.ScreenUpdating = False
  Application.EnableEvents = False
  'Range(RANGE_TAB_MOVIMENTACAO).Select
  'Selection.Sort Key1:=Range(RANGE_PRIMEIRA_DATA_MOVIMENTACAO), Order1:=xlAscending, Header:=xlGuess, _
  '      OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, _
  '      DataOption1:=xlSortNormal
  ActiveSheet.Sort.SortFields.Clear
  ActiveSheet.Sort.SortFields.Add Key:=Range(RANGE_COLUNA_DATA_MOVIMENTACAO), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
  With ActiveSheet.Sort
      .SetRange Range(RANGE_TAB_MOVIMENTACAO)
      .Header = xlGuess
      .MatchCase = False
      .Orientation = xlTopToBottom
      .SortMethod = xlPinYin
      .Apply
  End With
  
  'Range(RANGE_TAB_CARTOES).Select
  ActiveSheet.Sort.SortFields.Clear
  ActiveSheet.Sort.SortFields.Add Key:=Range(RANGE_COLUNA_DATA_CARTOES), _
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

Sub OrdenarAcoesFii()
  On Error GoTo ErroOrdenarAcoesFii
  
  Call OrdenarAcoes
  Call OrdenarFii
  Call OrdenarStock
  Call OrdenarReit
  
  Exit Sub
ErroOrdenarAcoesFii:
  MostrarMsgErro ("OrdenarAcoesFii")
End Sub


Private Sub OrdenarAcoes()
  '
  ' OrdenarAcoes Macro
  ' Ordena tabelas de Ações, Fii, Stock e Reit em ordem crescente de ticket.
  '
  On Error GoTo ErroOrdenarAcoes
  If Not IsPlanilhaAberta(Range(RANGE_SITUAC_PLANILHA)) Then
    Exit Sub
  End If
  Debug.Print "Ordenando ações..."
  Application.ScreenUpdating = False
  Application.EnableEvents = False
  Dim intPrimeiraLinha As Integer, intUltimaLinha As Integer, intColunaData As Integer
  Dim blnHasCompraOuVenda As Boolean
  
  ' Ordena book acoes
  intPrimeiraLinha = RetornarPrimeiraLinha(Range(RANGE_COLUNA_DATA_ACOES))
  intUltimaLinha = RetornarUltimaLinha(Range(RANGE_COLUNA_DATA_ACOES))
  intColunaData = RetornarPrimeiraColuna(Range(RANGE_COLUNA_DATA_ACOES))
  blnHasCompraOuVenda = HasCompraOuVenda(intPrimeiraLinha, intUltimaLinha, intColunaData)
  If blnHasCompraOuVenda Then Call InserirDataFalsa(intPrimeiraLinha, intUltimaLinha, intColunaData)
  ActiveSheet.Sort.SortFields.Clear
  ActiveSheet.Sort.SortFields.Add Key:=Range(RANGE_COLUNA_ATIVO_ACOES), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
  With ActiveSheet.Sort
      .SetRange Range(RANGE_TAB_ACOES)
      .Header = xlGuess
      .MatchCase = False
      .Orientation = xlTopToBottom
      .SortMethod = xlPinYin
      .Apply
  End With
  If blnHasCompraOuVenda Then
    ActiveSheet.Sort.SortFields.Clear
    ActiveSheet.Sort.SortFields.Add Key:=Range(RANGE_COLUNA_DATA_ACOES), _
          SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveSheet.Sort
        .SetRange Range(RANGE_TAB_ACOES)
        .Header = xlGuess
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Call RestaurarDataFalsa(intPrimeiraLinha, intUltimaLinha, intColunaData)
  End If
        
FimOrdenarAcoes:
  Application.ScreenUpdating = True
  Application.EnableEvents = True
  Exit Sub
  
ErroOrdenarAcoes:
  MostrarMsgErro ("OrdenarAcoes")
  Resume FimOrdenarAcoes
End Sub

Private Sub OrdenarFii()
  '
  ' OrdenarFii Macro
  ' Ordena tabelas de Ações, Fii, Stock e Reit em ordem crescente de ticket.
  '
  On Error GoTo ErroOrdenarFii
  If Not IsPlanilhaAberta(Range(RANGE_SITUAC_PLANILHA)) Then
    Exit Sub
  End If
  Debug.Print "Ordenando fii..."
  Application.ScreenUpdating = False
  Application.EnableEvents = False
  Dim intPrimeiraLinha As Integer, intUltimaLinha As Integer, intColunaData As Integer
  Dim blnHasCompraOuVenda As Boolean
  
  ' Ordena book fii
  intPrimeiraLinha = RetornarPrimeiraLinha(Range(RANGE_COLUNA_DATA_FII))
  intUltimaLinha = RetornarUltimaLinha(Range(RANGE_COLUNA_DATA_FII))
  intColunaData = RetornarPrimeiraColuna(Range(RANGE_COLUNA_DATA_FII))
  blnHasCompraOuVenda = HasCompraOuVenda(intPrimeiraLinha, intUltimaLinha, intColunaData)
  If blnHasCompraOuVenda Then Call InserirDataFalsa(intPrimeiraLinha, intUltimaLinha, intColunaData)
  ActiveSheet.Sort.SortFields.Clear
  ActiveSheet.Sort.SortFields.Add Key:=Range(RANGE_COLUNA_ATIVO_FII), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
  With ActiveSheet.Sort
      .SetRange Range(RANGE_TAB_FII)
      .Header = xlGuess
      .MatchCase = False
      .Orientation = xlTopToBottom
      .SortMethod = xlPinYin
      .Apply
  End With
  If blnHasCompraOuVenda Then
    ActiveSheet.Sort.SortFields.Clear
    ActiveSheet.Sort.SortFields.Add Key:=Range(RANGE_COLUNA_DATA_FII), _
          SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveSheet.Sort
        .SetRange Range(RANGE_TAB_FII)
        .Header = xlGuess
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Call RestaurarDataFalsa(intPrimeiraLinha, intUltimaLinha, intColunaData)
  End If
        
FimOrdenarFii:
  Application.ScreenUpdating = True
  Application.EnableEvents = True
  Exit Sub
  
ErroOrdenarFii:
  MostrarMsgErro ("OrdenarFii")
  Resume FimOrdenarFii
End Sub

Private Sub OrdenarStock()
  '
  ' OrdenarStock Macro
  ' Ordena tabelas de Ações, Fii, Stock e Reit em ordem crescente de ticket.
  '
  On Error GoTo ErroOrdenarStock
  If Not IsPlanilhaAberta(Range(RANGE_SITUAC_PLANILHA)) Then
    Exit Sub
  End If
  Debug.Print "Ordenando stock..."
  Application.ScreenUpdating = False
  Application.EnableEvents = False
  Dim intPrimeiraLinha As Integer, intUltimaLinha As Integer, intColunaData As Integer
  Dim blnHasCompraOuVenda As Boolean
    
  ' Ordena book stok
  intPrimeiraLinha = RetornarPrimeiraLinha(Range(RANGE_COLUNA_DATA_STOCK))
  intUltimaLinha = RetornarUltimaLinha(Range(RANGE_COLUNA_DATA_STOCK))
  intColunaData = RetornarPrimeiraColuna(Range(RANGE_COLUNA_DATA_STOCK))
  blnHasCompraOuVenda = HasCompraOuVenda(intPrimeiraLinha, intUltimaLinha, intColunaData)
  If blnHasCompraOuVenda Then Call InserirDataFalsa(intPrimeiraLinha, intUltimaLinha, intColunaData)
  ActiveSheet.Sort.SortFields.Clear
  ActiveSheet.Sort.SortFields.Add Key:=Range(RANGE_COLUNA_ATIVO_STOCK), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
  With ActiveSheet.Sort
      .SetRange Range(RANGE_TAB_STOCK)
      .Header = xlGuess
      .MatchCase = False
      .Orientation = xlTopToBottom
      .SortMethod = xlPinYin
      .Apply
  End With
  
  If blnHasCompraOuVenda Then
    ActiveSheet.Sort.SortFields.Clear
    ActiveSheet.Sort.SortFields.Add Key:=Range(RANGE_COLUNA_DATA_STOCK), _
          SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveSheet.Sort
        .SetRange Range(RANGE_TAB_STOCK)
        .Header = xlGuess
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Call RestaurarDataFalsa(intPrimeiraLinha, intUltimaLinha, intColunaData)
  End If
        
FimOrdenarStock:
  Application.ScreenUpdating = True
  Application.EnableEvents = True
  Exit Sub
  
ErroOrdenarStock:
  MostrarMsgErro ("OrdenarStock")
  Resume FimOrdenarStock
End Sub

Private Sub OrdenarReit()
  '
  ' OrdenarReit Macro
  ' Ordena tabelas de Ações, Fii, Stock e Reit em ordem crescente de ticket.
  '
  On Error GoTo ErroOrdenarReit
  If Not IsPlanilhaAberta(Range(RANGE_SITUAC_PLANILHA)) Then
    Exit Sub
  End If
  Debug.Print "Ordenando reit..."
  Application.ScreenUpdating = False
  Application.EnableEvents = False
  Dim intPrimeiraLinha As Integer, intUltimaLinha As Integer, intColunaData As Integer
  Dim blnHasCompraOuVenda As Boolean
  
  ' Ordena book reit
  intPrimeiraLinha = RetornarPrimeiraLinha(Range(RANGE_COLUNA_DATA_REIT))
  intUltimaLinha = RetornarUltimaLinha(Range(RANGE_COLUNA_DATA_REIT))
  intColunaData = RetornarPrimeiraColuna(Range(RANGE_COLUNA_DATA_REIT))
  blnHasCompraOuVenda = HasCompraOuVenda(intPrimeiraLinha, intUltimaLinha, intColunaData)
  If blnHasCompraOuVenda Then Call InserirDataFalsa(intPrimeiraLinha, intUltimaLinha, intColunaData)
  ActiveSheet.Sort.SortFields.Clear
  ActiveSheet.Sort.SortFields.Add Key:=Range(RANGE_COLUNA_ATIVO_REIT), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
  With ActiveSheet.Sort
      .SetRange Range(RANGE_TAB_REIT)
      .Header = xlGuess
      .MatchCase = False
      .Orientation = xlTopToBottom
      .SortMethod = xlPinYin
      .Apply
  End With
  
  If blnHasCompraOuVenda Then
    ActiveSheet.Sort.SortFields.Clear
    ActiveSheet.Sort.SortFields.Add Key:=Range(RANGE_COLUNA_DATA_REIT), _
          SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveSheet.Sort
        .SetRange Range(RANGE_TAB_REIT)
        .Header = xlGuess
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Call RestaurarDataFalsa(intPrimeiraLinha, intUltimaLinha, intColunaData)
  End If
        
FimOrdenarReit:
  Application.ScreenUpdating = True
  Application.EnableEvents = True
  Exit Sub
  
ErroOrdenarReit:
  MostrarMsgErro ("OrdenarReit")
  Resume FimOrdenarReit
End Sub

Private Function HasCompraOuVenda(intPrimeiraLinha As Integer, intUltimaLinha As Integer, intColunaData As Integer) As Boolean
  On Error GoTo ErroHasCompraOuVenda
  Dim intCont As Integer
  
  For intCont = intPrimeiraLinha To intUltimaLinha
    If Trim(ActiveSheet.Cells(intCont, intColunaData) & " ") <> "" Then
      HasCompraOuVenda = True
      Exit Function
    End If
  Next intCont
 HasCompraOuVenda = False
  
  Exit Function
ErroHasCompraOuVenda:
  MostrarMsgErro ("HasCompraOuVenda")
End Function

Private Sub InserirDataFalsa(intPrimeiraLinha As Integer, intUltimaLinha As Integer, intColunaData As Integer)
  On Error GoTo ErroInserirDataFalsa
  Dim intCont As Integer
  
  For intCont = intPrimeiraLinha To intUltimaLinha
    If Trim(ActiveSheet.Cells(intCont, intColunaData) & " ") = "" And Trim(ActiveSheet.Cells(intCont, intColunaData + 1) & " ") <> "" Then
      ActiveSheet.Cells(intCont, intColunaData) = #12/31/1980#
    End If
  Next intCont
  
  Exit Sub
ErroInserirDataFalsa:
  MostrarMsgErro ("InserirDataFalsa")
End Sub

Private Sub RestaurarDataFalsa(intPrimeiraLinha As Integer, intUltimaLinha As Integer, intColunaData As Integer)
  On Error GoTo ErroRestaurarDataFalsa
  Dim intCont As Integer
  For intCont = intPrimeiraLinha To intUltimaLinha
    Debug.Print ActiveSheet.Cells(intCont, intColunaData).Value
    If ActiveSheet.Cells(intCont, intColunaData).Value = #12/31/1980# Then ActiveSheet.Cells(intCont, intColunaData) = ""
  Next intCont
  
  Exit Sub
ErroRestaurarDataFalsa:
  MostrarMsgErro ("RestaurarDataFalsa")
End Sub
