Option Explicit

' ---------------------------------------------------------------------
' Módulo: Comum.bas
' Descrição: Módulo com rotinas comuns a todos os outros.
'            Convenções VB de <http://support.microsoft.com/kb/110264>
' ---------------------------------------------------------------------

' Informações agregadas de um ativo de uma carteira (ações, FII, ETF etc.)
Public Type InfoAtivo
  strAtivo As String
  dblQuantidade As Double
  dblSaldo As Double
End Type

Private Const NOMES_MESES_ABREVIADOS As String = "Jan|Fev|Mar|Abr|Mai|Jun|Jul|Ago|Set|Out|Nov|Dez"
Private Const NOMES_MESES_COMPLETOS As String = "Janeiro|Fevereiro|Março|Abril|Maio|Junho|Julho|Agosto|Setembro|Outubro|Novembro|Dezembro"
Private Const SEPARADOR_MESES As String = "|"
Private Const MES_PADRAO As Integer = 12
Private Const VALOR_BASE_RETORNO As Double = 1000
Private Const FORMATO_DATA_ANO As String = "/yy"
Private Const TITULO_MSG_ERRO As String = "Expense"

' Nível de aninhamento das chamadas a CongelarCalculosPlanilha
Private mlngNivelCongelamento As Long

Public Function IsPlanilhaAberta(ByVal rgPos As Range) As Boolean
  Dim varSituacao As Variant
  varSituacao = rgPos.Value
  If IsError(varSituacao) Then
    Exit Function
  End If
  IsPlanilhaAberta = (CStr(varSituacao) = SITUAC_ABERTO)
End Function

Public Sub MostrarMsgErro(ByVal strOrigem As String)
  Dim lngNumero As Long
  Dim strDescricao As String
  lngNumero = Err.Number
  strDescricao = Err.Description
  Debug.Print "Não foi possível concluir " & strOrigem & ". Erro " & lngNumero & ": " & strDescricao
  MsgBox "Não foi possível concluir a operação." & vbNewLine & vbNewLine _
    & "Origem: " & strOrigem & vbNewLine _
    & "Erro: " & lngNumero & vbNewLine _
    & "Descrição: " & strDescricao, vbCritical, TITULO_MSG_ERRO
End Sub

Private Function RetornarPlanilhaOuAtiva(ByVal wsPlanilha As Worksheet) As Worksheet
  If wsPlanilha Is Nothing Then
    Set RetornarPlanilhaOuAtiva = ActiveSheet
  Else
    Set RetornarPlanilhaOuAtiva = wsPlanilha
  End If
End Function

Public Function RetornarUltimaCelulaMovimentacoes(Optional ByVal wsPlanilha As Worksheet) As Range
  Set wsPlanilha = RetornarPlanilhaOuAtiva(wsPlanilha)
  Set RetornarUltimaCelulaMovimentacoes = wsPlanilha.Cells(RetornarUltimaLinhaMovimentacoes(wsPlanilha) + 1, _
    wsPlanilha.Range(RANGE_HEADER_DATA_MOVIMENTACAO).Column)
End Function

'
' Procura a última linha de Movimento preenchida
'
Public Function RetornarUltimaLinhaMovimentacoes(Optional ByVal wsPlanilha As Worksheet) As Long  
  RetornarUltimaLinhaMovimentacoes = RetornarPlanilhaOuAtiva(wsPlanilha).Range(RANGE_HEADER_MOVIMENTACAO).End(xlDown).Row
End Function

'
' Procura a última linha de Cartão preenchida
'
Public Function RetornarUltimaLinhaCartao(Optional ByVal wsPlanilha As Worksheet) As Long
  RetornarUltimaLinhaCartao = RetornarPlanilhaOuAtiva(wsPlanilha).Range(RANGE_HEADER_CARTOES).End(xlDown).Row
End Function

'
' Primeira linha livre da tabela de movimentações; lança erro se a tabela estiver cheia
'
Public Function RetornarLinhaLivreMovimentacoes(ByVal wsPlanilha As Worksheet) As Long
  RetornarLinhaLivreMovimentacoes = ValidarLinhaLivre(RetornarUltimaLinhaMovimentacoes(wsPlanilha) + 1, _
    wsPlanilha.Range(RANGE_COLUNA_DATA_MOVIMENTACAO), "movimentações", wsPlanilha)
End Function

'
' Primeira linha livre da tabela de cartões; lança erro se a tabela estiver cheia
'
Public Function RetornarLinhaLivreCartao(ByVal wsPlanilha As Worksheet) As Long
  RetornarLinhaLivreCartao = ValidarLinhaLivre(RetornarUltimaLinhaCartao(wsPlanilha) + 1, _
    wsPlanilha.Range(RANGE_COLUNA_DATA_CARTOES), "cartões", wsPlanilha)
End Function

Private Function ValidarLinhaLivre(ByVal lngLinha As Long, ByVal rgTabela As Range, _
    ByVal strNomeTabela As String, ByVal wsPlanilha As Worksheet) As Long
  If lngLinha < RetornarPrimeiraLinha(rgTabela) Or lngLinha > RetornarUltimaLinha(rgTabela) Then
    Err.Raise ERRO_TABELA_CHEIA, , "Não há linhas livres na tabela de " & strNomeTabela & _
      " da planilha " & wsPlanilha.Name & "."
  End If
  ValidarLinhaLivre = lngLinha
End Function

'
' Retorna a planilha seguinte à informada, ou Nothing se ela for Dezembro, a última ou um gráfico
'
Public Function RetornarProximaPlanilha(ByVal wsPlanilha As Worksheet) As Worksheet  
  Dim objPlanilha As Object
  If wsPlanilha.Name = NOME_PLAN_DEZ Then
    Exit Function
  End If
  If wsPlanilha.Index >= wsPlanilha.Parent.Sheets.Count Then
    Exit Function
  End If
  Set objPlanilha = wsPlanilha.Parent.Sheets(wsPlanilha.Index + 1)
  If TypeOf objPlanilha Is Worksheet Then
    Set RetornarProximaPlanilha = objPlanilha
  End If
End Function

'
' Resolve um nome definido na pasta, independente da planilha ativa
'
Public Function RetornarRangeNomeado(ByVal strNome As String) As Range  
  On Error Resume Next
  Set RetornarRangeNomeado = ThisWorkbook.Names(strNome).RefersToRange
  ' Desativa o comportamento de ignorar erro acima (resume next) e volta ao comportamento padrão
  On Error GoTo 0
  If RetornarRangeNomeado Is Nothing Then
    Set RetornarRangeNomeado = ActiveSheet.Range(strNome)
  End If
End Function

'
' Converte o conteúdo de uma célula em número; textos e erros valem zero
'
Public Function RetornarValorNumerico(ByVal varValor As Variant) As Double  
  If IsError(varValor) Then
    Exit Function
  End If
  If IsNumeric(varValor) Then
    RetornarValorNumerico = CDbl(varValor)
  End If
End Function

'
' Converte o nome da planilha mensal (Jan..Dez) no número do mês; padrão: Dezembro
'
Public Function RetornarNumeroMesPlanilha(ByVal strNomePlanilha As String) As Integer  
  Dim varMeses As Variant
  Dim intIndice As Integer
  varMeses = Split(NOMES_MESES_ABREVIADOS, SEPARADOR_MESES)
  For intIndice = LBound(varMeses) To UBound(varMeses)
    If varMeses(intIndice) = strNomePlanilha Then
      RetornarNumeroMesPlanilha = intIndice + 1
      Exit Function
    End If
  Next intIndice
  RetornarNumeroMesPlanilha = MES_PADRAO
End Function

Public Function RetornarNomeMesCompleto(ByVal intMes As Integer) As String
  RetornarNomeMesCompleto = Split(NOMES_MESES_COMPLETOS, SEPARADOR_MESES)(intMes - 1)
End Function

'
' Agrupa os ativos de uma carteira pelo nome, somando as quantidades.
' Retorna o número de ativos distintos (udtAtivos fica dimensionado de 1 até esse número).
'
Public Function ColetarAtivos(ByVal wsPlanilha As Worksheet, ByVal strRangeAtivo As String, _
    ByVal strRangeQtde As String, ByVal strRangeSaldoFinal As String, ByRef udtAtivos() As InfoAtivo) As Long  
  Dim varAtivos As Variant, varQtdes As Variant, varSaldos As Variant
  Dim lngLinha As Long, lngTotal As Long, lngPosicao As Long
  Dim strAtivo As String
  ' leitura em bloco: evita acessar a planilha célula a célula
  varAtivos = wsPlanilha.Range(strRangeAtivo).Value
  varQtdes = wsPlanilha.Range(strRangeQtde).Value
  varSaldos = wsPlanilha.Range(strRangeSaldoFinal).Value
  ReDim udtAtivos(1 To UBound(varAtivos, 1))
  For lngLinha = 1 To UBound(varAtivos, 1)
    If Not IsEmpty(varAtivos(lngLinha, 1)) And Not IsError(varAtivos(lngLinha, 1)) Then
      strAtivo = CStr(varAtivos(lngLinha, 1))
      lngPosicao = RetornarPosicaoAtivo(udtAtivos, lngTotal, strAtivo)
      If lngPosicao = 0 Then
        ' ainda não existe a entrada: cria uma nova
        lngTotal = lngTotal + 1
        udtAtivos(lngTotal).strAtivo = strAtivo
        udtAtivos(lngTotal).dblQuantidade = RetornarValorNumerico(varQtdes(lngLinha, 1))
        udtAtivos(lngTotal).dblSaldo = RetornarValorNumerico(varSaldos(lngLinha, 1))
      Else
        ' já existe: atualiza a quantidade
        udtAtivos(lngPosicao).dblQuantidade = udtAtivos(lngPosicao).dblQuantidade + _
          RetornarValorNumerico(varQtdes(lngLinha, 1))
      End If
    End If
  Next lngLinha
  If lngTotal > 0 Then
    ReDim Preserve udtAtivos(1 To lngTotal)
  End If
  ColetarAtivos = lngTotal
End Function

Private Function RetornarPosicaoAtivo(ByRef udtAtivos() As InfoAtivo, ByVal lngTotal As Long, _
    ByVal strAtivo As String) As Long
  Dim lngIndice As Long
  For lngIndice = 1 To lngTotal
    If udtAtivos(lngIndice).strAtivo = strAtivo Then
      RetornarPosicaoAtivo = lngIndice
      Exit Function
    End If
  Next lngIndice
End Function

'
' Congela (True) ou libera (False) tela, cálculos e eventos.
' As chamadas podem ser aninhadas: só a última liberação restaura o Excel.
'
Public Sub CongelarCalculosPlanilha(ByVal blnValor As Boolean)  
  If blnValor Then
    mlngNivelCongelamento = mlngNivelCongelamento + 1
    If mlngNivelCongelamento > 1 Then
      Exit Sub
    End If
    Application.Cursor = xlWait
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Exit Sub
  End If
  If mlngNivelCongelamento > 1 Then
    mlngNivelCongelamento = mlngNivelCongelamento - 1
    Exit Sub
  End If
  mlngNivelCongelamento = 0
  ' garante que todas as configurações sejam restauradas, mesmo que alguma falhe
  On Error Resume Next
  Application.Cursor = xlDefault
  Application.ScreenUpdating = True
  Application.Calculation = xlCalculationAutomatic
  Application.EnableEvents = True
  Application.DisplayAlerts = True
End Sub

Public Function RetornarPrimeiraLinha(ByVal rgRange As Range) As Long
  RetornarPrimeiraLinha = rgRange.Cells(1, 1).Row
End Function

Public Function RetornarUltimaLinha(ByVal rgRange As Range) As Long
  RetornarUltimaLinha = rgRange.Cells(rgRange.Cells.Count).Row
End Function

Public Function RetornarPrimeiraColuna(ByVal rgRange As Range) As Long
  RetornarPrimeiraColuna = rgRange.Cells(1, 1).Column
End Function

Public Function RetornarUltimaColuna(ByVal rgRange As Range) As Long
  RetornarUltimaColuna = rgRange.Cells(rgRange.Cells.Count).Column
End Function

'
' PosicionarTopo Macro
' Posiciona a planilha na célula superior esquerda.
'
' Atalho do teclado: Ctrl+t
'
Public Sub PosicionarTopo()  
  On Error GoTo ErroPosicionarTopo
  If ActiveWindow Is Nothing Then
    Exit Sub
  End If
  ActiveWindow.ScrollRow = 1
  ActiveWindow.ScrollColumn = 1
  Exit Sub

ErroPosicionarTopo:
  MostrarMsgErro "PosicionarTopo"
End Sub

'
' Função de planilha: maior queda acumulada de uma série de retornos.
' Células não numéricas são ignoradas.
'
Public Function MaxDrawdown(ByVal rgArray As Range) As Variant
  Dim rgCelula As Range
  Dim varValor As Variant
  Dim dblValorAtual As Double, dblValorMaximo As Double
  Dim dblDrawdownAtual As Double, dblDrawdownMaximo As Double
  On Error GoTo ErroMaxDrawdown

  dblValorAtual = VALOR_BASE_RETORNO
  For Each rgCelula In rgArray.Cells
    varValor = rgCelula.Value
    If IsNumeric(varValor) Then
      dblValorAtual = dblValorAtual * (1 + CDbl(varValor))
      If dblValorAtual > dblValorMaximo Then
        dblValorMaximo = dblValorAtual
      End If
      If dblValorMaximo <> 0 And dblValorAtual < dblValorMaximo Then
        dblDrawdownAtual = dblValorAtual / dblValorMaximo - 1
        If dblDrawdownAtual < dblDrawdownMaximo Then
          dblDrawdownMaximo = dblDrawdownAtual
        End If
      End If
    End If
  Next rgCelula
  MaxDrawdown = dblDrawdownMaximo
  Exit Function

ErroMaxDrawdown:
  ' função de planilha não deve abrir mensagens: devolve erro para a célula
  MaxDrawdown = CVErr(xlErrValue)
End Function

'
' Função de planilha: retorno acumulado de uma série de retornos.
' Células não numéricas são ignoradas.
'
Public Function TotalReturn(ByVal rgArray As Range) As Variant  
  Dim rgCelula As Range
  Dim varValor As Variant
  Dim dblValorAtual As Double
  On Error GoTo ErroTotalReturn

  dblValorAtual = VALOR_BASE_RETORNO
  For Each rgCelula In rgArray.Cells
    varValor = rgCelula.Value
    If IsNumeric(varValor) Then
      dblValorAtual = dblValorAtual * (1 + CDbl(varValor))
    End If
  Next rgCelula
  TotalReturn = dblValorAtual / VALOR_BASE_RETORNO - 1
  Exit Function

ErroTotalReturn:
  TotalReturn = CVErr(xlErrValue)
End Function

'
' dataAtual Macro
' Traz a data atual para a coluna de movimentos ou cartão.
'
' Atalho do teclado: Ctrl+d
' Criado por: Mauricio SS  Em: 14/02/04
'
Public Sub PuxarDataAtual()  
  Dim rgAlvo As Range
  On Error GoTo ErroPuxarDataAtual
  If Not IsPlanilhaAberta(Range(RANGE_SITUAC_PLANILHA)) Then
    Exit Sub
  End If
  If Not TypeOf Selection Is Range Then
    Exit Sub
  End If
  Set rgAlvo = Selection
  If IsEmpty(rgAlvo) And InStr(rgAlvo.NumberFormat, FORMATO_DATA_ANO) > 0 Then
    rgAlvo.Value = Date
  End If
  Exit Sub

ErroPuxarDataAtual:
  MostrarMsgErro "PuxarDataAtual"
End Sub

'
' SheetOffSet
' Retorna conteúdo de células de uma referência, em planilhas offset
' Ex.: =SheetOffSet(-1;M284)
'      =PROC(2;1/(SheetOffSet(-1;$E$284:$E$313)=E284);SheetOffSet(-1;$Z$284:$Z$313))
'
' Criado por: Mauricio SS  Em: 14/05/20
' Fonte: https://answers.microsoft.com/en-us/msoffice/forum/all/the-advanced-sheetoffset-function/cdf3077e-7b35-41c9-ac74-5175163c856d
'
Public Function SheetOffSet(ByVal lngOffset As Long, Optional ByVal rgCell As Range) As Variant  
  Dim wsOrigem As Worksheet
  Dim wbPasta As Workbook
  Dim lngIndice As Long, lngIndiceOrigem As Long, lngIndiceDestino As Long
  On Error GoTo ErroSheetOffSet

  Application.Volatile 'força o recalculo sempre que houver uma mudança na pasta
  If rgCell Is Nothing Then
    Set rgCell = Application.Caller
  End If
  Set wsOrigem = Application.Caller.Parent
  Set wbPasta = wsOrigem.Parent
  For lngIndice = 1 To wbPasta.Worksheets.Count
    If wbPasta.Worksheets(lngIndice).Name = wsOrigem.Name Then
      lngIndiceOrigem = lngIndice
      Exit For
    End If
  Next lngIndice
  lngIndiceDestino = lngIndiceOrigem + lngOffset
  If lngIndiceOrigem = 0 Or lngIndiceDestino < 1 Or lngIndiceDestino > wbPasta.Worksheets.Count Then
    SheetOffSet = CVErr(xlErrRef)
    Exit Function
  End If
  Set SheetOffSet = wbPasta.Worksheets(lngIndiceDestino).Range(rgCell.Address)
  Exit Function

ErroSheetOffSet:
  SheetOffSet = CVErr(xlErrRef)
End Function

'
' ReduzirTotalParaUnitario
' Considera o valor digitado como total e divide pela quantidade para obter o unitário.
'
' Atalho do teclado: Ctrl+u
' Criado por: Mauricio SS  Em: 03/08/23
'
Public Sub ReduzirTotalParaUnitario()
  Dim wsPlanilha As Worksheet
  Dim rgAlvo As Range
  Dim dblQtde As Double
  On Error GoTo ErroReduzirTotalParaUnitario
  If Not IsPlanilhaAberta(Range(RANGE_SITUAC_PLANILHA)) Then
    Exit Sub
  End If
  If Not TypeOf Selection Is Range Then
    Exit Sub
  End If
  Set wsPlanilha = ActiveSheet
  Set rgAlvo = ActiveCell
  If RetornarValorNumerico(rgAlvo.Value) <= 0 Then
    Exit Sub
  End If
  If Application.Intersect(rgAlvo, wsPlanilha.Range(RANGE_COLUNA_SALDO_FINAL_TESOURO_DIRETO)) Is Nothing And _
     Application.Intersect(rgAlvo, wsPlanilha.Range(RANGE_COLUNA_SALDO_FINAL_TESOURO_SELIC)) Is Nothing Then
    Exit Sub
  End If
  dblQtde = RetornarValorNumerico(wsPlanilha.Cells(rgAlvo.Row, _
    RetornarPrimeiraColuna(wsPlanilha.Range(RANGE_COLUNA_QTDE_TESOURO_DIRETO))).Value)
  If dblQtde = 0 Then
    MsgBox "Não foi possível calcular o valor unitário: a quantidade não está preenchida.", vbExclamation
    Exit Sub
  End If
  rgAlvo.Value = rgAlvo.Value / dblQtde
  Exit Sub

ErroReduzirTotalParaUnitario:
  MostrarMsgErro "ReduzirTotalParaUnitario"
End Sub
