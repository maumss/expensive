Attribute VB_Name = "Comum"
' M�dulo com rotinas comuns a todos os outros
' Conven��es VB: http://support.microsoft.com/kb/110264
Option Explicit

Function IsPlanilhaAberta(rgPos As Range) As Boolean
  IsPlanilhaAberta = (rgPos.Value = SITUAC_ABERTO)
End Function

Sub MostrarMsgErro(strOrigem As String)
  MsgBox strOrigem & vbNewLine & vbNewLine _
    & "Erro: " & Err.Number & vbNewLine _
    & "Descri��o: " & Err.Description, vbCritical
End Sub

Function RetornarUltimaCelulaMovimentacoes() As Range
  Dim intColunaDataMovimentacoes As Integer
  intColunaDataMovimentacoes = Range(RANGE_HEADER_DATA_MOVIMENTACOES).Column
  Set RetornarUltimaCelulaMovimentacoes = Cells((RetornarUltimaLinhaMovimentacoes + 1), intColunaDataMovimentacoes)
End Function

Function RetornarUltimaLinhaMovimentacoes() As Long
  ' procura a �ltima linha de Movimento preenchida
  On Error GoTo ErroUltLinhaD
  RetornarUltimaLinhaMovimentacoes = Range(RANGE_HEADER_MOVIMENTACOES).End(xlDown).Row
  Exit Function
  
ErroUltLinhaD:
  MostrarMsgErro ("RetornarUltimaLinhaMovimentacoes")
End Function

Function RetornarUltimaLinhaCartao() As Long
  On Error GoTo ErroRetornarUltimaLinhaCartao
  RetornarUltimaLinhaCartao = Range(RANGE_HEADER_CARTOES).End(xlDown).Row
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

Function RetornarPrimeiraLinha(rgRange As Range) As Long
  RetornarPrimeiraLinha = rgRange.Cells(1, 1).Row
End Function

Function RetornarUltimaLinha(rgRange As Range) As Long
  RetornarUltimaLinha = rgRange(rgRange.count).Row
End Function

Function RetornarPrimeiraColuna(rgRange As Range) As Long
  RetornarPrimeiraColuna = rgRange.Cells(1, 1).Column
End Function

Function RetornarUltimaColuna(rgRange As Range) As Long
  RetornarUltimaColuna = rgRange(rgRange.count).Column
End Function

Sub PosicionarTopo()
  '
  ' PosicionarTopo Macro
  ' Posiciona a planilha na c�lula superior esquerda.
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

Function MaxDrawdown(rgArray As Range) As Double
  Dim rgMyCell As Range
  Dim dblCurValue As Double, dblMaxValue As Double, dblCurDd As Double, dblMaxDd As Double
  
  dblMaxValue = 0
  dblMaxDd = 0
  dblCurValue = 1000
  
  For Each rgMyCell In rgArray
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

Function TotalReturn(rgArray As Range) As Double
  Dim rgMyCell As Range
  Dim dblCurValue As Double
    
  dblCurValue = 1000
  
  For Each rgMyCell In rgArray
    dblCurValue = dblCurValue * (1 + rgMyCell.Value)
  Next rgMyCell
  
  TotalReturn = dblCurValue / 1000 - 1
End Function