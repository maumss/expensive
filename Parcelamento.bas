Option Explicit

' --------------------------------------------------------------------------------------------------
' Módulo: Parcelamento.bas
' Descrição: Módulo de construção de parcelamentos de uma conta no formato: conta (a/b)
'   Cria parcelamento do valor da seleção atual. Se a descrição contiver o texto "(1/x)", será
'   criada x entradas nas planilhas seguintes com o mesmo valor inicial, caso contrário perguntará
'   o número de parcelas e dividirá o valor atual.
' --------------------------------------------------------------------------------------------------

Private Const TITULO_CRIAR_PARCELAS As String = "Criar parcelas"
Private Const ABRE_PARENTESES As String = "("
Private Const FECHA_PARENTESES As String = ")"
Private Const BARRA As String = "/"
Private Const PRIMEIRA_PARCELA As Integer = 1
Private Const MAXIMO_PARCELAS As Integer = 999
Private Const PARCELAS_INVALIDAS As Integer = -1

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

'
' Cria entrada nas planilhas futuras com base no parcelamento da seleção corrente
'
' Atalho do teclado: Ctrl+p
'
Public Sub MontarParcelamento()
  Dim wsPlanilhaAtual As Worksheet, wsPlanilhaDestino As Worksheet
  Dim rgAlvo As Range
  Dim udtMovimentacao As Movimentacao
  Dim udtInfoParcelamento As InfoParcelamento
  Dim varDescricoes As Variant, varDescricao As Variant
  Dim intParcela As Integer
  Dim lngParcelasCriadas As Long, lngParcelasPrevistas As Long
  Dim blnCongelou As Boolean, blnConcluiu As Boolean
  On Error GoTo ErroMontarParcelamento

  'Testa se é planilha mensal válida e se está aberta
  If Not IsPlanilhaAberta(Range(RANGE_SITUAC_PLANILHA)) Then
    Exit Sub
  End If
  'Testa se já está na planilha de Dezembro
  If ActiveSheet.Name = NOME_PLAN_DEZ Then
    MsgBox "Não existe planilha posterior a essa para criar parcelas", vbCritical, TITULO_CRIAR_PARCELAS
    Exit Sub
  End If
  If Not TypeOf Selection Is Range Then
    Exit Sub
  End If
  Set wsPlanilhaAtual = ActiveSheet
  Set rgAlvo = Selection.Cells(1)
  If IsLocalInvalido(wsPlanilhaAtual, rgAlvo) Then
    Exit Sub
  End If
  'Pede confirmação
  If MsgBox("Você deseja criar parcelas com base nesta movimentação?", _
       vbYesNo + vbQuestion, TITULO_CRIAR_PARCELAS) = vbNo Then
    Exit Sub
  End If

  udtMovimentacao = RetornarMovimentacao(wsPlanilhaAtual, rgAlvo)
  If Len(udtMovimentacao.strDescMov) = 0 Then
    Exit Sub
  End If
  udtInfoParcelamento = RetornarInfoParcelamento(udtMovimentacao.strDescMov, udtMovimentacao.dblValorMov)
  If udtInfoParcelamento.intTotalParcelas = 0 Then
    ' cancelou o parcelamento
    Exit Sub
  End If
  If Not IsParcelamentoValido(udtInfoParcelamento) Then
    MsgBox "Não há parcelas a criar: verifique o número de parcelas informado ou da descrição.", _
      vbExclamation, TITULO_CRIAR_PARCELAS
    Exit Sub
  End If
  varDescricoes = RetornarArrayDeParcelamento(udtMovimentacao.strDescMov, udtInfoParcelamento)
  lngParcelasPrevistas = UBound(varDescricoes) - LBound(varDescricoes) + 1

  CongelarCalculosPlanilha True
  blnCongelou = True
  If Not IsTextoPreParcelado(udtInfoParcelamento) Then
    AlterarTextoAtual wsPlanilhaAtual, rgAlvo, udtInfoParcelamento, udtMovimentacao
  End If
  Set wsPlanilhaDestino = wsPlanilhaAtual
  intParcela = PRIMEIRA_PARCELA
  For Each varDescricao In varDescricoes
    Set wsPlanilhaDestino = RetornarProximaPlanilha(wsPlanilhaDestino)
    If wsPlanilhaDestino Is Nothing Then
      ' chegou ao fim do ano
      Exit For
    End If
    If IsPlanilhaAberta(wsPlanilhaDestino.Range(RANGE_SITUAC_PLANILHA)) Then
      CriarEntradaNaPlanilha wsPlanilhaDestino, CStr(varDescricao), udtInfoParcelamento.dblValorParcela, _
        udtMovimentacao, intParcela
      lngParcelasCriadas = lngParcelasCriadas + 1
    End If
    intParcela = intParcela + 1
  Next varDescricao
  blnConcluiu = True

FimMontarParcelamento:
  If blnCongelou Then
    CongelarCalculosPlanilha False
  End If
  If blnConcluiu And lngParcelasCriadas < lngParcelasPrevistas Then
    MsgBox "Foram criadas " & lngParcelasCriadas & " de " & lngParcelasPrevistas & " parcelas." & vbNewLine & _
      "As demais não foram criadas porque a planilha destino não existe ou está Fechada.", _
      vbInformation, TITULO_CRIAR_PARCELAS
  End If
  Exit Sub

ErroMontarParcelamento:
  MostrarMsgErro "MontarParcelamento"
  Resume FimMontarParcelamento
End Sub

Private Function IsLocalInvalido(ByVal wsPlanilha As Worksheet, ByVal rgAlvo As Range) As Boolean
  IsLocalInvalido = (Application.Intersect(rgAlvo, wsPlanilha.Range(RANGE_TAB_MOVIMENTACAO)) Is Nothing) And _
       (Application.Intersect(rgAlvo, wsPlanilha.Range(RANGE_TAB_CARTOES)) Is Nothing)
End Function

Private Function RetornarMovimentacao(ByVal wsPlanilha As Worksheet, ByVal rgAlvo As Range) As Movimentacao
  Dim udtMovimentacao As Movimentacao
  Dim lngLinhaAtual As Long, lngColunaInicial As Long
  lngLinhaAtual = rgAlvo.Row
  udtMovimentacao.blnMovCartao = IsMovimentoCartao(wsPlanilha, rgAlvo)
  If udtMovimentacao.blnMovCartao Then
    lngColunaInicial = wsPlanilha.Range(RANGE_PRIMEIRA_DATA_CARTOES).Column
    udtMovimentacao.strNomeCartao = CStr(wsPlanilha.Cells(lngLinhaAtual, lngColunaInicial + 3).Value)
    udtMovimentacao.dblValorMov = RetornarValorNumerico(wsPlanilha.Cells(lngLinhaAtual, lngColunaInicial + 4).Value)
  Else
    lngColunaInicial = wsPlanilha.Range(RANGE_PRIMEIRA_DATA_MOVIMENTACAO).Column
    udtMovimentacao.dblValorMov = RetornarValorNumerico(wsPlanilha.Cells(lngLinhaAtual, lngColunaInicial + 3).Value)
  End If
  udtMovimentacao.datDataMov = wsPlanilha.Cells(lngLinhaAtual, lngColunaInicial).Value
  udtMovimentacao.strDescMov = CStr(wsPlanilha.Cells(lngLinhaAtual, lngColunaInicial + 1).Value)
  udtMovimentacao.strTipoMov = CStr(wsPlanilha.Cells(lngLinhaAtual, lngColunaInicial + 2).Value)
  RetornarMovimentacao = udtMovimentacao
End Function

Private Function IsMovimentoCartao(ByVal wsPlanilha As Worksheet, ByVal rgAlvo As Range) As Boolean
  IsMovimentoCartao = Not (Application.Intersect(rgAlvo, wsPlanilha.Range(RANGE_TAB_CARTOES)) Is Nothing)
End Function

'
' Verifica o formato de parcelas dizendo a posição na string que os separadores ocupam
'
Private Function RetornarInfoParcelamento(ByVal strTextoOriginal As String, ByVal dblValorMov As Double) As InfoParcelamento
  Dim udtInfoParcelamento As InfoParcelamento
  ' considera o último grupo entre parênteses, no formato: "x----x (x/x)"
  udtInfoParcelamento.intPosAbreParent = InStrRev(strTextoOriginal, ABRE_PARENTESES)
  If udtInfoParcelamento.intPosAbreParent > 0 Then
    udtInfoParcelamento.intPosBarra = InStr(udtInfoParcelamento.intPosAbreParent + 1, strTextoOriginal, BARRA)
    udtInfoParcelamento.intPosFechaParent = InStr(udtInfoParcelamento.intPosAbreParent + 1, strTextoOriginal, FECHA_PARENTESES)
  End If
  If (udtInfoParcelamento.intPosAbreParent = 0) Or (udtInfoParcelamento.intPosBarra = 0) Or _
     (udtInfoParcelamento.intPosFechaParent = 0) Or _
     (udtInfoParcelamento.intPosBarra > udtInfoParcelamento.intPosFechaParent) Then
    udtInfoParcelamento.intPosAbreParent = 0
    udtInfoParcelamento.intPosFechaParent = 0
    udtInfoParcelamento.intPosBarra = 0
  End If
  udtInfoParcelamento.intParcelaInicial = RetornarParcelaInicialMaisUm(strTextoOriginal, udtInfoParcelamento)
  udtInfoParcelamento.intTotalParcelas = RetornarTotalParcelas(strTextoOriginal, udtInfoParcelamento)
  If udtInfoParcelamento.intTotalParcelas > 0 Then
    udtInfoParcelamento.dblValorParcela = RetornarValorParcela(udtInfoParcelamento, dblValorMov)
  End If
  RetornarInfoParcelamento = udtInfoParcelamento
End Function

'
' Precisa haver ao menos uma parcela posterior à atual
'
Private Function IsParcelamentoValido(ByRef udtInfoParcelamento As InfoParcelamento) As Boolean  
  IsParcelamentoValido = (udtInfoParcelamento.intParcelaInicial > PRIMEIRA_PARCELA) And _
    (udtInfoParcelamento.intParcelaInicial <= udtInfoParcelamento.intTotalParcelas)
End Function

'
' Converte o texto em número de parcela; retorna PARCELAS_INVALIDAS se não for um inteiro válido
'
Private Function ConverterParaInteiro(ByVal strValor As String) As Integer
  Dim dblValor As Double
  ConverterParaInteiro = PARCELAS_INVALIDAS
  strValor = Trim$(strValor)
  If Len(strValor) = 0 Or Not IsNumeric(strValor) Then
    Exit Function
  End If
  dblValor = CDbl(strValor)
  If dblValor <> Fix(dblValor) Or dblValor < 0 Or dblValor > MAXIMO_PARCELAS Then
    Exit Function
  End If
  ConverterParaInteiro = CInt(dblValor)
End Function

'
' Busca a parcela inicial a partir da parcela selecionada
'
Private Function RetornarParcelaInicialMaisUm(ByVal strTextoOriginal As String, ByRef udtInfoParcelamento As InfoParcelamento) As Integer
  Dim intParcelaAtual As Integer
  If IsTextoPreParcelado(udtInfoParcelamento) Then
    intParcelaAtual = ConverterParaInteiro(Mid$(strTextoOriginal, udtInfoParcelamento.intPosAbreParent + 1, _
        udtInfoParcelamento.intPosBarra - udtInfoParcelamento.intPosAbreParent - 1))
    If intParcelaAtual > 0 Then
      RetornarParcelaInicialMaisUm = intParcelaAtual + 1
    End If
    Exit Function
  End If
  'se não estiver no formato parcelado, supõe que a próxima parcela será a dois
  RetornarParcelaInicialMaisUm = PRIMEIRA_PARCELA + 1
End Function

'
' Busca ou define o total de parcelas que deverão ser criadas (0 = cancelado)
'
Private Function RetornarTotalParcelas(ByVal strTextoOriginal As String, ByRef udtInfoParcelamento As InfoParcelamento) As Integer
  Dim strTotalParcelas As String
  If IsTextoPreParcelado(udtInfoParcelamento) Then
    strTotalParcelas = Mid$(strTextoOriginal, udtInfoParcelamento.intPosBarra + 1, _
       udtInfoParcelamento.intPosFechaParent - udtInfoParcelamento.intPosBarra - 1)
  Else
    'se não for formato parcelado, pede que o usuário digite
    strTotalParcelas = InputBox("Entre o total de parcelas ou pressione Cancelar para finalizar:", TITULO_CRIAR_PARCELAS)
    If Len(strTotalParcelas) = 0 Then
      'pressionou cancelar
      Exit Function
    End If
  End If
  RetornarTotalParcelas = ConverterParaInteiro(strTotalParcelas)
End Function

'
' Calcula o valor da parcela
'
Private Function RetornarValorParcela(ByRef udtInfoParcelamento As InfoParcelamento, ByVal dblValorMov As Double) As Double
  If IsTextoPreParcelado(udtInfoParcelamento) Then
    RetornarValorParcela = dblValorMov
  Else
    RetornarValorParcela = dblValorMov / udtInfoParcelamento.intTotalParcelas
  End If
End Function

'
' Monta um array com a descrição de cada parcela a ser criada
'
Private Function RetornarArrayDeParcelamento(ByVal strTextoOriginal As String, ByRef udtInfoParcelamento As InfoParcelamento) As Variant
  Dim astrItems() As String
  Dim strParteFixa As String
  Dim intParcela As Integer, intIndice As Integer
  ReDim astrItems(0 To udtInfoParcelamento.intTotalParcelas - udtInfoParcelamento.intParcelaInicial)
  strParteFixa = RetornarParteFixaDoTexto(strTextoOriginal, udtInfoParcelamento)
  For intParcela = udtInfoParcelamento.intParcelaInicial To udtInfoParcelamento.intTotalParcelas
    astrItems(intIndice) = strParteFixa & CStr(intParcela) & BARRA & CStr(udtInfoParcelamento.intTotalParcelas) & FECHA_PARENTESES
    intIndice = intIndice + 1
  Next intParcela
  RetornarArrayDeParcelamento = astrItems
End Function

'
' Retorna o prefixo da descrição de cada parcela
'
Private Function RetornarParteFixaDoTexto(ByVal strTextoOriginal As String, ByRef udtInfoParcelamento As InfoParcelamento) As String
  If IsTextoPreParcelado(udtInfoParcelamento) Then
    RetornarParteFixaDoTexto = Left$(strTextoOriginal, udtInfoParcelamento.intPosAbreParent)
  Else
    RetornarParteFixaDoTexto = strTextoOriginal & " " & ABRE_PARENTESES
  End If
End Function

'
' Retorna se o texto já pressupõe um parcelamento
'
Private Function IsTextoPreParcelado(ByRef udtInfoParcelamento As InfoParcelamento) As Boolean
  IsTextoPreParcelado = (udtInfoParcelamento.intPosBarra > 0)
End Function

'
' Caso não for pre-parcelado, deve mudar o texto atual para conter a parcela 1/x
'
Private Sub AlterarTextoAtual(ByVal wsPlanilha As Worksheet, ByVal rgAlvo As Range, _
    ByRef udtInfoParcelamento As InfoParcelamento, ByRef udtMovimentacao As Movimentacao)
  Dim lngLinhaDestino As Long, lngColunaDestino As Long, lngColunaValor As Long
  lngLinhaDestino = rgAlvo.Row
  If udtMovimentacao.blnMovCartao Then
    lngColunaDestino = wsPlanilha.Range(RANGE_PRIMEIRA_DATA_CARTOES).Column
    lngColunaValor = lngColunaDestino + 4
  Else
    lngColunaDestino = wsPlanilha.Range(RANGE_PRIMEIRA_DATA_MOVIMENTACAO).Column
    lngColunaValor = lngColunaDestino + 3
  End If
  wsPlanilha.Cells(lngLinhaDestino, lngColunaDestino + 1).Value = udtMovimentacao.strDescMov & " " & _
     ABRE_PARENTESES & PRIMEIRA_PARCELA & BARRA & CStr(udtInfoParcelamento.intTotalParcelas) & FECHA_PARENTESES
  wsPlanilha.Cells(lngLinhaDestino, lngColunaValor).Value = udtInfoParcelamento.dblValorParcela
End Sub

'
' Cria a entrada de uma parcela na planilha destino
'
Private Sub CriarEntradaNaPlanilha(ByVal wsPlanilhaDestino As Worksheet, ByVal strTexto As String, _
     ByVal dblValorParcela As Double, ByRef udtMovimentacao As Movimentacao, ByVal intParcela As Integer)
  Dim lngLinhaDestino As Long, lngColunaDestino As Long
  Dim datDataParcela As Date
  datDataParcela = udtMovimentacao.datDataMov
  If udtMovimentacao.blnMovCartao Then
    lngLinhaDestino = RetornarLinhaLivreCartao(wsPlanilhaDestino)
    lngColunaDestino = wsPlanilhaDestino.Range(RANGE_PRIMEIRA_DATA_CARTOES).Column
    wsPlanilhaDestino.Cells(lngLinhaDestino, lngColunaDestino).Resize(1, 5).Value = _
      Array(datDataParcela, strTexto, udtMovimentacao.strTipoMov, udtMovimentacao.strNomeCartao, dblValorParcela)
  Else
    lngLinhaDestino = RetornarLinhaLivreMovimentacoes(wsPlanilhaDestino)
    lngColunaDestino = wsPlanilhaDestino.Range(RANGE_PRIMEIRA_DATA_MOVIMENTACAO).Column
    datDataParcela = DateAdd("m", intParcela, datDataParcela)
    wsPlanilhaDestino.Cells(lngLinhaDestino, lngColunaDestino).Resize(1, 4).Value = _
      Array(datDataParcela, strTexto, udtMovimentacao.strTipoMov, dblValorParcela)
  End If
End Sub
