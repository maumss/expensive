Attribute VB_Name = "Comum"
' Módulo com rotinas comuns a todos os outros
' Convenções VB: http://support.microsoft.com/kb/110264
Option Explicit

Function IsPlanilhaAberta(rgPos As Range) As Boolean
  IsPlanilhaAberta = (rgPos.Value = SITUAC_ABERTO)
End Function

Sub MostrarMsgErro(strOrigem As String)
  MsgBox strOrigem & vbNewLine & vbNewLine _
    & "Erro: " & Err.Number & vbNewLine _
    & "Descrição: " & Err.Description, vbCritical
End Sub

Function RetornarUltimaCelulaMovimentacoes() As Range
  Dim intColunaDataMovimentacoes As Integer
  intColunaDataMovimentacoes = Range(RANGE_HEADER_DATA_MOVIMENTACOES).Column
  Set RetornarUltimaCelulaMovimentacoes = Cells((RetornarUltimaLinhaMovimentacoes + 1), intColunaDataMovimentacoes)
End Function

Function RetornarUltimaLinhaMovimentacoes() As Long
  ' procura a última linha de Movimento preenchida
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
