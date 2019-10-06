' Módulo com rotinas comuns a todos os outros
' Convenções VB: http://support.microsoft.com/kb/110264
Option Explicit

Function IsPlanilhaAberta(rgPos As range) As Boolean
  IsPlanilhaAberta = (rgPos.Value = SITUAC_ABERTO)
End Function

Sub MostrarMsgErro(strOrigem As String)
  MsgBox strOrigem & vbNewLine & vbNewLine _
    & "Erro: " & Err.Number & vbNewLine _
    & "Descrição: " & Err.Description, vbCritical
End Sub

Function RetornarUltimaCelulaMovimentacoes() As range
  Dim intColunaDataMovimentacoes As Integer
  intColunaDataMovimentacoes = range(RANGE_HEADER_DATA_MOVIMENTACOES).Column
  Set RetornarUltimaCelulaMovimentacoes = Cells((RetornarUltimaLinhaMovimentacoes + 1), intColunaDataMovimentacoes)
End Function

Function RetornarUltimaLinhaMovimentacoes() As Long
  ' procura a última linha de Movimento preenchida
  On Error GoTo ErroUltLinhaD
  RetornarUltimaLinhaMovimentacoes = range(RANGE_HEADER_MOVIMENTACOES).End(xlDown).Row
  Exit Function
  
ErroUltLinhaD:
  MostrarMsgErro ("RetornarUltimaLinhaMovimentacoes")
End Function

Function RetornarUltimaLinhaCartao() As Long
  On Error GoTo ErroRetornarUltimaLinhaCartao
  RetornarUltimaLinhaCartao = range(RANGE_HEADER_CARTOES).End(xlDown).Row
  Exit Function
  
ErroRetornarUltimaLinhaCartao:
  MostrarMsgErro ("RetornarUltimaLinhaCartao")
End Function

Sub CongelarCalculosPlanilha(blnValor As Boolean)
  If blnValor Then
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Exit Sub
  End If
  Application.ScreenUpdating = True
  Application.Calculation = xlCalculationAutomatic
  Application.EnableEvents = True
  Application.DisplayAlerts = True
End Sub

Function RetornarPrimeiraLinha(rgRange As range) As Long
  RetornarPrimeiraLinha = rgRange.Cells(1, 1).Row
End Function

Function RetornarUltimaLinha(rgRange As range) As Long
  RetornarUltimaLinha = rgRange(rgRange.Count).Row
End Function

Function RetornarPrimeiraColuna(rgRange As range) As Long
  RetornarPrimeiraColuna = rgRange.Cells(1, 1).Column
End Function

Function RetornarUltimaColuna(rgRange As range) As Long
  RetornarUltimaColuna = rgRange(rgRange.Count).Column
End Function

Sub PosicionarTopo()
  '
  ' PosicionarTopo Macro
  ' Posiciona a planilha na célula superior esquerda.
  '
  ' Atalho do teclado: Ctrl+t
  '
  On Error GoTo erroposicionarTopo
  ActiveWindow.ScrollRow = 1
  ActiveWindow.ScrollColumn = 1
  'ActiveSheet.Cells(1, 1).Select
  Exit Sub
  
erroposicionarTopo:
  MostrarMsgErro ("PosicionarTopo")
End Sub

Function MaxDrawdown(rgArray As range) As Double
  Dim rgMyCell As range
  Dim dblCurValue As Double, dblMaxValue As Double, dblCurDd As Double, dblMaxDd As Double
  Dim blnNumeric As Boolean
  
  dblMaxValue = 0
  dblMaxDd = 0
  dblCurValue = 1000
  
  For Each rgMyCell In rgArray
    If Not IsNumeric(rgMyCell) Then
      GoTo NextInteration
    End If
    
    dblCurValue = dblCurValue * (1 + rgMyCell.Value)
    
    If dblCurValue > dblMaxValue Then
      dblMaxValue = dblCurValue
    End If
    
    If dblMaxValue = 0 Then
      GoTo NextInteration
    End If
    
    dblCurDd = 0
    If dblCurValue < dblMaxValue Then
      dblCurDd = dblCurValue / dblMaxValue - 1
    End If
    
    If dblCurDd < dblMaxDd Then
        dblMaxDd = dblCurDd
    End If
NextInteration:
  Next rgMyCell
  
  MaxDrawdown = dblMaxDd
End Function

Function TotalReturn(rgArray As range) As Double
  Dim rgMyCell As range
  Dim dblCurValue As Double
    
  dblCurValue = 1000
  
  For Each rgMyCell In rgArray
    dblCurValue = dblCurValue * (1 + rgMyCell.Value)
  Next rgMyCell
  
  TotalReturn = dblCurValue / 1000 - 1
End Function

Sub PuxarDataAtual()
  '
  ' dataAtual Macro
  ' Traz a data atual para a coluna de movimentos ou cartão.
  '
  ' Atalho do teclado: Ctrl+d
  ' Criado por: Mauricio SS  Em: 14/02/04
  '
  If Not IsPlanilhaAberta(range(RANGE_SITUAC_PLANILHA)) Then
    Exit Sub
  End If
  ' variáveis
  Dim rgAlvo As range
  Dim wsPlanilha As Worksheet
  ' principal
  On Error GoTo ErroData
  Set wsPlanilha = ActiveSheet
  Set rgAlvo = Selection
  If IsEmpty(rgAlvo) Then
    If (Not Application.Intersect(rgAlvo, range(RANGE_COL_DATA_MOVIMENTACOES)) Is Nothing) Or _
       (Not Application.Intersect(rgAlvo, range(RANGE_COL_DATA_CARTOES)) Is Nothing) Or _
       (Not Application.Intersect(rgAlvo, range(RANGE_COL_DATA_ACOES)) Is Nothing) Or _
       (Not Application.Intersect(rgAlvo, range(RANGE_COL_DATA_FII)) Is Nothing) Or _
       (Not Application.Intersect(rgAlvo, range(RANGE_COL_DATA_RF)) Is Nothing) Or _
       (Not Application.Intersect(rgAlvo, range(RANGE_COL_DATA_SELIC)) Is Nothing) Or _
       (Not Application.Intersect(rgAlvo, range(RANGE_COL_DATA_ETF)) Is Nothing) Or _
       (Not Application.Intersect(rgAlvo, range(RANGE_COL_DATA_STOCK)) Is Nothing) Or _
       (Not Application.Intersect(rgAlvo, range(RANGE_COL_DATA_REIT)) Is Nothing) Or _
       (Not Application.Intersect(rgAlvo, range(RANGE_COL_DATA_CART_BILL)) Is Nothing) Or _
       (Not Application.Intersect(rgAlvo, range(RANGE_COL_DATA_CART_COMMODITY)) Is Nothing) Or _
       (Not Application.Intersect(rgAlvo, range(RANGE_COL_DATA_CART_OPCOES)) Is Nothing) Then
      rgAlvo.Value = Date
    End If
  End If
  Exit Sub
  
ErroData:
  MostrarMsgErro ("PuxarDataAtual")
End Sub

