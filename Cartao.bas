Attribute VB_Name = "Cartao"
'M�dulo de transfer�ncia de lan�amentos de cart�o p/ outro m�s
Option Explicit

Private Type MovimentoCartao
  datDataCartao As Date
  strDescCartao As String
  strTipoCartao As String
  strNomeCartao As String
  dblValorCartao As Double
  strFormulaValorCartao As String
End Type

Public Sub MoverCartao()
Attribute MoverCartao.VB_Description = "Move um lan�amento de cart�o para o pr�ximo m�s."
Attribute MoverCartao.VB_ProcData.VB_Invoke_Func = "m\n14"
  '
  ' MoveCartao Macro
  ' Move um lan�amento de cart�o de um m�s para o pr�ximo.
  '
  ' Atalho do teclado: Ctrl+m
  '
  On Error GoTo ErroMoverCartao
  Dim lngIndPlan As Long, lngLinhaAtual As Long
  lngIndPlan = Worksheets(ActiveSheet.Name).Index
  lngLinhaAtual = Selection.Cells(1).Row
  If IsMoverInvalido(lngIndPlan, lngLinhaAtual) Then
    Exit Sub
  End If
    
  'Pede confirma��o
  If MsgBox("Voc� deseja mover o lan�amento para o pr�ximo m�s?", _
        vbYesNo + vbQuestion, "Mover lan�amento") = vbNo Then
    Exit Sub
  End If
  
  'Come�a o processo de mover o lan�amento
  CongelarCalculosPlanilha (True)
  On Error GoTo EndMacro:
  
  ' Joga os valores da linha a mover em vari�veis
  Dim udtMovimentoCartao As MovimentoCartao
  udtMovimentoCartao = RetornarDadosMovCartao(lngLinhaAtual)
  ' Apaga linha selecionada
  Dim intColunaInicioCartao As Integer, intColunaFinalCartao As Integer
  intColunaInicioCartao = Range(RANGE_PRIMEIRA_DATA_CARTOES).Column
  intColunaFinalCartao = Range(RANGE_ULTIMO_VALOR_CARTAO).Column
  Range(Cells(lngLinhaAtual, intColunaInicioCartao), Cells(lngLinhaAtual, intColunaFinalCartao)).Select
  Selection.ClearContents
  ' Habilita os eventos p/ marcar que houve mudan�a
  Application.EnableEvents = True
  Call JogarValoresCartaoNaPlanilha(lngIndPlan + 1, udtMovimentoCartao)
  ' Volta a posi��o original
  Worksheets(lngIndPlan).Activate
  ActiveSheet.Cells(lngLinhaAtual, intColunaInicioCartao).Select
  
EndMacro:
  CongelarCalculosPlanilha (False)
  Exit Sub
  
ErroMoverCartao:
  MostrarMsgErro ("MoverCartao")
  Resume EndMacro
End Sub

Private Function IsMoverInvalido(lngIndPlan As Long, lngLinhaAtual As Long) As Boolean
  'Testa se � planilha mensal v�lida e se est� aberta
  If Not IsPlanilhaAberta(Range(RANGE_SITUAC_PLANILHA)) Then
    IsMoverInvalido = True
    Exit Function
  End If
  'Testa se j� est� na planilha de Dezembro
  If ActiveSheet.Name = NOME_PLAN_DEZ Then
    MsgBox "N�o existe planilha posterior a essa para mover valores", vbCritical
    IsMoverInvalido = True
    Exit Function
  End If
  'Testa se a planilha destino est� aberta
  If Not IsPlanilhaAberta(Worksheets(lngIndPlan + 1).Range(RANGE_SITUAC_PLANILHA)) Then
    MsgBox "A planilha destino est� Fechada para altera��es", vbCritical
    IsMoverInvalido = True
    Exit Function
  End If
  'Testa se est� dentro do lan�amento do cart�o
  Dim rgSelecao As Range
  Set rgSelecao = Selection
  If (Application.Intersect(rgSelecao, Range(RANGE_TAB_CARTOES)) Is Nothing) Then
    IsMoverInvalido = True
    Exit Function
  End If
  'Verifica se existe algo para mover
  Dim intColunaInicioCartao As Integer
  intColunaInicioCartao = Range(RANGE_PRIMEIRA_DATA_CARTOES).Column
  If IsEmpty(Cells(lngLinhaAtual, intColunaInicioCartao)) Then
    MsgBox "Voc� n�o se posicionou em uma c�lula com dados", vbCritical
    IsMoverInvalido = True
    Exit Function
  End If
  IsMoverInvalido = False
End Function

Private Function RetornarDadosMovCartao(lngLinhaAtual As Long) As MovimentoCartao
  Dim udtMovimentoCartao As MovimentoCartao
  Dim intColunaInicioCartao As Integer
  intColunaInicioCartao = Range(RANGE_PRIMEIRA_DATA_CARTOES).Column
  udtMovimentoCartao.datDataCartao = Cells(lngLinhaAtual, intColunaInicioCartao).Value
  udtMovimentoCartao.strDescCartao = Cells(lngLinhaAtual, intColunaInicioCartao + 1).Value
  udtMovimentoCartao.strTipoCartao = Cells(lngLinhaAtual, intColunaInicioCartao + 2).Value
  udtMovimentoCartao.strNomeCartao = Cells(lngLinhaAtual, intColunaInicioCartao + 3).Value
  udtMovimentoCartao.strFormulaValorCartao = ""
  If Cells(lngLinhaAtual, intColunaInicioCartao + 4).HasFormula Then
    udtMovimentoCartao.strFormulaValorCartao = Cells(lngLinhaAtual, intColunaInicioCartao + 4).Formula
  Else
    udtMovimentoCartao.dblValorCartao = Cells(lngLinhaAtual, intColunaInicioCartao + 4).Value
  End If
  RetornarDadosMovCartao = udtMovimentoCartao
End Function

Private Sub JogarValoresCartaoNaPlanilha(lngIndPlanDestino As Long, udtMovimentoCartao As MovimentoCartao)
  ' Se posiciona na pr�xima planilha
  Worksheets(lngIndPlanDestino).Activate
  ' Joga valores na pr�xima planilha
  Dim lngLinhaDest As Long
  Dim intColunaDestino As Integer
  lngLinhaDest = RetornarUltimaLinhaCartao + 1
  intColunaDestino = Range(RANGE_PRIMEIRA_DATA_CARTOES).Column
  Cells(lngLinhaDest, intColunaDestino).Value = udtMovimentoCartao.datDataCartao
  Cells(lngLinhaDest, intColunaDestino + 1).Value = udtMovimentoCartao.strDescCartao
  Cells(lngLinhaDest, intColunaDestino + 2).Value = udtMovimentoCartao.strTipoCartao
  Cells(lngLinhaDest, intColunaDestino + 3).Value = udtMovimentoCartao.strNomeCartao
  If udtMovimentoCartao.strFormulaValorCartao > "" Then
    Cells(lngLinhaDest, intColunaDestino + 4).Formula = udtMovimentoCartao.strFormulaValorCartao
  Else
    Cells(lngLinhaDest, intColunaDestino + 4).Value = udtMovimentoCartao.dblValorCartao
  End If
End Sub

