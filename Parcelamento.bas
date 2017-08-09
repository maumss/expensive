Attribute VB_Name = "Parcelamento"
Option Explicit
' M�dulo de constru��o de parcelamentos de uma conta no formato: conta (a/b)
' Cria parcelamento do valor da sele��o atual. Se a descri��o contiver o texto "(1/x)", ser�
' criada x entradas nas planilhas seguintes com o mesmo valor inicial, caso contr�rio perguntar�
' o n�mero de parcelas e dividir� o valor atual.

Private Type InfoParcelamento
  intPosAbreParent As Integer
  intPosFechaParent As Integer
  intPosBarra As Integer
  intParcelaInicial As Integer
  intTotalParcelas As Integer
  dblValorParcela As Double
End Type

Private Type Movimentacao
  datDataMov As Date
  strDescMov As String
  strTipoMov As String
  strNomeCartao As String
  dblValorMov As Double
  blnMovCartao As Boolean
End Type

'Cria entrada nas planilhas futuras com base no parcelamento da sele��o corrente
Sub MontarParcelamento()
Attribute MontarParcelamento.VB_Description = "Cria parcelamento do valor da sele��o atual. Se a descri��o contiver o texto ""(1/x)"", ser� criada x entradas nas planilhas seguintes com o mesmo valor inicial, caso contr�rio perguntar� o n�mero de parcelas e dividir� o valor atual."
Attribute MontarParcelamento.VB_ProcData.VB_Invoke_Func = "p\n14"
   On Error GoTo ErroCriarParcelamento
   Dim strTextoOriginal As String
   'Testa se � planilha mensal v�lida e se est� aberta
   If Not IsPlanilhaAberta(Range(RANGE_SITUAC_PLANILHA)) Then
     Exit Sub
   End If
   'Testa se j� est� na planilha de Dezembro
   If ActiveSheet.Name = NOME_PLAN_DEZ Then
    MsgBox "N�o existe planilha posterior a essa para criar parcelas", vbCritical
    Exit Sub
   End If
   Dim wsPlanilha As Worksheet
   Dim rgAlvo As Range
   Set wsPlanilha = ActiveSheet
   Set rgAlvo = Selection
   If IsLocalIvalido(rgAlvo) Then
     Exit Sub
   End If
   'Pede confirma��o
   If MsgBox("Voc� deseja criar parcelas com base nesta movimenta��o?", _
        vbYesNo + vbQuestion, "Mover lan�amento") = vbNo Then
     Exit Sub
   End If
  
   Dim udtMovimentacao As Movimentacao
   udtMovimentacao = RetornarMovimentacao(rgAlvo)
   If Len(udtMovimentacao.strDescMov) = 0 Then
     Exit Sub
   End If
   Dim udtInfoParcelamento As InfoParcelamento
   udtInfoParcelamento = RetornarInfoParcelamento(udtMovimentacao.strDescMov, udtMovimentacao.dblValorMov)
   If udtInfoParcelamento.intTotalParcelas = 0 Then
     ' cancelou o parcelamento
     Exit Sub
   End If
   Dim varArray As Variant
   varArray = RetornarArrayDeParcelamento(udtMovimentacao.strDescMov, udtInfoParcelamento)
   If IsArrayEmpty(varArray) Then
     Exit Sub
   End If
   Dim lngIndPlanAtual As Long, lngIndPlan As Long
   If (Not IsTextoPreParcelado(udtInfoParcelamento)) Then
     Call AlterarTextoAtual(rgAlvo, udtInfoParcelamento, udtMovimentacao)
   End If
   lngIndPlanAtual = Worksheets(ActiveSheet.Name).Index
   lngIndPlan = lngIndPlanAtual + 1
   Dim vntTexto As Variant
   Dim intParcela As Integer
   intParcela = 1
   For Each vntTexto In varArray
     Call CriarEntradaNaPlanilha(lngIndPlan, vntTexto, udtInfoParcelamento.dblValorParcela, udtMovimentacao, intParcela)
     lngIndPlan = lngIndPlan + 1
     intParcela = intParcela + 1
   Next vntTexto
   Worksheets(lngIndPlanAtual).Activate
   'rgAlvo.Select
      
   Exit Sub
ErroCriarParcelamento:
  MostrarMsgErro ("MontarParcelamento")
End Sub

Private Function IsLocalIvalido(rgAlvo As Range) As Boolean
  IsLocalIvalido = (Application.Intersect(rgAlvo, Range(RANGE_TAB_MOVIMENTACOES)) Is Nothing) And _
       (Application.Intersect(rgAlvo, Range(RANGE_TAB_CARTOES)) Is Nothing)
End Function

Private Function RetornarMovimentacao(rgAlvo As Range) As Movimentacao
  Dim intLinhaAtual As Integer, intColunaInicial As Integer, intColunaInicioCartao As Integer
  intLinhaAtual = rgAlvo.Row
  Dim udtMovimentacao As Movimentacao
  If IsMovimentoCartao(rgAlvo) Then
    intColunaInicial = Range(RANGE_PRIMEIRA_DATA_CARTOES).Column
    udtMovimentacao.strNomeCartao = Cells(intLinhaAtual, intColunaInicial + 3).Value
    udtMovimentacao.dblValorMov = Cells(intLinhaAtual, intColunaInicial + 4).Value
    udtMovimentacao.blnMovCartao = True
  Else
    intColunaInicial = Range(RANGE_PRIMEIRA_DATA_MOVIMENTACOES).Column
    udtMovimentacao.dblValorMov = Cells(intLinhaAtual, intColunaInicial + 3).Value
    udtMovimentacao.blnMovCartao = False
  End If
  udtMovimentacao.datDataMov = Cells(intLinhaAtual, intColunaInicial).Value
  udtMovimentacao.strDescMov = Cells(intLinhaAtual, intColunaInicial + 1).Value
  udtMovimentacao.strTipoMov = Cells(intLinhaAtual, intColunaInicial + 2).Value
  RetornarMovimentacao = udtMovimentacao
End Function

Private Function IsMovimentoCartao(rgAlvo As Range) As Boolean
  IsMovimentoCartao = (Not Application.Intersect(rgAlvo, Range(RANGE_TAB_CARTOES)) Is Nothing)
End Function

'Verifica o formato de parcelas dizendo a posi��o na string que os separadores ocupam
Private Function RetornarInfoParcelamento(strTextoOriginal As String, dblValorMov As Double) As InfoParcelamento
  Dim udtInfoParcelamento As InfoParcelamento
  udtInfoParcelamento.intPosAbreParent = InStr(strTextoOriginal, "(")
  udtInfoParcelamento.intPosFechaParent = InStr(strTextoOriginal, ")")
  udtInfoParcelamento.intPosBarra = InStr(strTextoOriginal, "/")
  ' verifica se o texto possui o formato de parcelas: "x----x (x/x)"
  If (udtInfoParcelamento.intPosAbreParent < 0) Or (udtInfoParcelamento.intPosFechaParent < 0) Or _
     (udtInfoParcelamento.intPosFechaParent < udtInfoParcelamento.intPosAbreParent) Or _
     (udtInfoParcelamento.intPosBarra < udtInfoParcelamento.intPosAbreParent) Or _
     (udtInfoParcelamento.intPosBarra > udtInfoParcelamento.intPosFechaParent) Then
    udtInfoParcelamento.intPosAbreParent = 0
    udtInfoParcelamento.intPosFechaParent = 0
    udtInfoParcelamento.intPosBarra = 0
  End If
  udtInfoParcelamento.intParcelaInicial = RetornarParcelaInicialMaisUm(strTextoOriginal, udtInfoParcelamento)
  udtInfoParcelamento.intTotalParcelas = RetornarTotalParcelas(strTextoOriginal, udtInfoParcelamento)
  udtInfoParcelamento.dblValorParcela = RetornarValorParcela(udtInfoParcelamento, dblValorMov)
  RetornarInfoParcelamento = udtInfoParcelamento
End Function

'Busca a parcela inicial a partir da parcela selecionada
Private Function RetornarParcelaInicialMaisUm(strTextoOriginal As String, udtInfoParcelamento As InfoParcelamento) As Integer
  Dim strParcelaInicial
  RetornarParcelaInicialMaisUm = 0
  If IsTextoPreParcelado(udtInfoParcelamento) Then
    strParcelaInicial = Mid(strTextoOriginal, udtInfoParcelamento.intPosAbreParent + 1, _
        udtInfoParcelamento.intPosBarra - udtInfoParcelamento.intPosAbreParent - 1)
    If IsNumeric(strParcelaInicial) Then
      RetornarParcelaInicialMaisUm = CInt(strParcelaInicial) + 1
    End If
    Exit Function
  End If
  'se n�o estiver no formato parcelado, sup�e que a pr�xima parcela ser� a dois
  RetornarParcelaInicialMaisUm = 2
End Function

'Busca ou define o total de parcelas que dever�o ser criadas
Private Function RetornarTotalParcelas(strTextoOriginal As String, udtInfoParcelamento As InfoParcelamento) As Integer
  Dim strTotalParcelas
  RetornarTotalParcelas = 0
  If IsTextoPreParcelado(udtInfoParcelamento) Then
    strTotalParcelas = Mid(strTextoOriginal, udtInfoParcelamento.intPosBarra + 1, _
       udtInfoParcelamento.intPosFechaParent - udtInfoParcelamento.intPosBarra - 1)
    If IsNumeric(strTotalParcelas) Then
      RetornarTotalParcelas = CInt(strTotalParcelas)
    End If
    Exit Function
  End If
  'se n�o for formato parcelado, pede que o usu�rio digite
  strTotalParcelas = InputBox("Entre o total de parcelas ou pressione Cancelar para finalizar:")
  If Len(strTotalParcelas) = 0 Then
    'pressionou cancelar
    RetornarTotalParcelas = 0
    Exit Function
  End If
  If IsNumeric(strTotalParcelas) Then
    RetornarTotalParcelas = CInt(strTotalParcelas)
  End If
End Function

'Calcula o valor da parcela
Private Function RetornarValorParcela(udtInfoParcelamento As InfoParcelamento, dblValorMov As Double) As Double
  Dim dblValorParcela As Double
  If IsTextoPreParcelado(udtInfoParcelamento) Then
    RetornarValorParcela = dblValorMov
    Exit Function
  End If
  RetornarValorParcela = dblValorMov / udtInfoParcelamento.intTotalParcelas
End Function

'Monta um array com a descri��o de cada parcela
Private Function RetornarArrayDeParcelamento(strTextoOriginal As String, udtInfoParcelamento As InfoParcelamento) As Variant
  Dim astrItems() As String
  On Error GoTo ErroRetornarArrayDeParcelamento
  If (udtInfoParcelamento.intParcelaInicial = 0 Or udtInfoParcelamento.intTotalParcelas = 0) Then
    RetornarArrayDeParcelamento = astrItems
    Exit Function
  End If
  Dim intTamArray As Integer
  intTamArray = udtInfoParcelamento.intTotalParcelas - 2
  ReDim astrItems(0 To intTamArray)

  Dim strParteFixa As String
  strParteFixa = RetornarParteFixaDoTexto(strTextoOriginal, udtInfoParcelamento)
  
  Dim intParcela As Integer, intIndex As Integer
  intIndex = 0
  For intParcela = udtInfoParcelamento.intParcelaInicial To udtInfoParcelamento.intTotalParcelas
    astrItems(intIndex) = strParteFixa & CStr(intParcela) & "/" & CStr(udtInfoParcelamento.intTotalParcelas) & ")"
    intIndex = intIndex + 1
  Next intParcela
  RetornarArrayDeParcelamento = astrItems
  
  Exit Function
ErroRetornarArrayDeParcelamento:
  MostrarMsgErro ("RetornarArrayDeParcelamento")
End Function

'Retorna o prefixo da descri��o de cada parcela
Private Function RetornarParteFixaDoTexto(strTextoOriginal As String, udtInfoParcelamento As InfoParcelamento) As String
  Dim strParteFixa As String
  RetornarParteFixaDoTexto = strTextoOriginal & " ("
  If IsTextoPreParcelado(udtInfoParcelamento) Then
    RetornarParteFixaDoTexto = Left(strTextoOriginal, udtInfoParcelamento.intPosAbreParent)
  End If
End Function

'Retorna se o texto j� pressup�e um parcelamento
Private Function IsTextoPreParcelado(udtInfoParcelamento As InfoParcelamento) As Boolean
  IsTextoPreParcelado = (udtInfoParcelamento.intPosBarra > 0)
End Function

'Verifica se um array possui algum conte�do
Private Function IsArrayEmpty(varArray As Variant) As Boolean
   ' Determina se um array cont�m algum elemento
   Dim lngUBound As Long
   On Error Resume Next
   ' Se o array estiver vazio, um erro ocorrer� quando checar os limites do array
   lngUBound = UBound(varArray)
   If Err.Number <> 0 Then
      IsArrayEmpty = True
   Else
      IsArrayEmpty = False
   End If
End Function

'Caso n�o for pre-parcelado, deve mudar o texto atual para conter a parcela 1/x
Private Sub AlterarTextoAtual(rgAlvo As Range, udtInfoParcelamento As InfoParcelamento, udtMovimentacao As Movimentacao)
  Dim lngLinhaDest As Long
  Dim intColunaDestino As Integer, intColunaValor As Integer
  lngLinhaDest = rgAlvo.Row
  If (udtMovimentacao.blnMovCartao) Then
    intColunaDestino = Range(RANGE_PRIMEIRA_DATA_CARTOES).Column
    intColunaValor = intColunaDestino + 4
  Else
    intColunaDestino = Range(RANGE_PRIMEIRA_DATA_MOVIMENTACOES).Column
    intColunaValor = intColunaDestino + 3
  End If
  Cells(lngLinhaDest, intColunaDestino + 1).Value = udtMovimentacao.strDescMov & " (1/" & _
     CStr(udtInfoParcelamento.intTotalParcelas) & ")"
  Cells(lngLinhaDest, intColunaValor).Value = udtInfoParcelamento.dblValorParcela
End Sub

'Cria tantas entradas quanto for o n�mero de parcelas
Private Sub CriarEntradaNaPlanilha(lngIndPlanDestino As Long, ByVal strTexto As String, _
     dblValorParcela As Double, udtMovimentacao As Movimentacao, intParcela As Integer)
   CongelarCalculosPlanilha (True)
   On Error GoTo EndMacro:
   ' Se posiciona na pr�xima planilha
   Worksheets(lngIndPlanDestino).Activate
   If Not IsPlanilhaAberta(Range(RANGE_SITUAC_PLANILHA)) Then
     Exit Sub
   End If
   Dim lngLinhaDest As Long
   Dim intColunaDestino As Integer
   Dim datDataParcela As Date
   datDataParcela = udtMovimentacao.datDataMov
   If (udtMovimentacao.blnMovCartao) Then
     lngLinhaDest = RetornarUltimaLinhaCartao + 1
     intColunaDestino = Range(RANGE_PRIMEIRA_DATA_CARTOES).Column
   Else
     lngLinhaDest = RetornarUltimaLinhaMovimentacoes + 1
     intColunaDestino = Range(RANGE_PRIMEIRA_DATA_MOVIMENTACOES).Column
     datDataParcela = DateAdd("m", intParcela, datDataParcela)
   End If
   Cells(lngLinhaDest, intColunaDestino).Value = datDataParcela
   Cells(lngLinhaDest, intColunaDestino + 1).Value = strTexto
   Cells(lngLinhaDest, intColunaDestino + 2).Value = udtMovimentacao.strTipoMov
   Dim intColunaValor As Integer
   intColunaValor = intColunaDestino + 3
   If (udtMovimentacao.blnMovCartao) Then
     Cells(lngLinhaDest, intColunaDestino + 3).Value = udtMovimentacao.strNomeCartao
     intColunaValor = intColunaDestino + 4
   End If
   Cells(lngLinhaDest, intColunaValor).Value = dblValorParcela
   
EndMacro:
  CongelarCalculosPlanilha (False)
  Exit Sub
  
ErroCriarEntradaNaPlanilha:
  MostrarMsgErro ("CriarEntradaNaPlanilha")
  Resume EndMacro
End Sub
