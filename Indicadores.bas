' Módulo de funções para atualização dos indicadores
Option Explicit

Sub AtualizarDadosPlanAtual()
  '
  ' AtualizarDadosPlanAtual    Data: 30/04/16
  ' Atualiza tabela com fonte de dados na web
  '
  On Error GoTo ErroAtualizarDadosPlanAtual
  
  Dim wsPlanilha As Worksheet
  Set wsPlanilha = ActiveSheet
  Call AtualizarFonteDeDadosPlanilha(wsPlanilha)
  Set wsPlanilha = Nothing
  Exit Sub
   
ErroAtualizarDadosPlanAtual:
  MostrarMsgErro ("AtualizarDadosPlanAtual")
End Sub


Sub BuscarIndicadores()
  ' BuscarIndicadores
  ' Busca os indicadores a partir de uma fonte de dados na Web
  ' De: <http://www.valor.com.br/valor-data/indices-financeiros/indicadores-de-mercado>
  '
  On Error GoTo ErroBuscarIndicadores
  
  If (Range(RANGE_SITUAC_PLANILHA).Value <> SITUAC_ABERTO) Then
    Exit Sub
  End If
  Dim wsPlanilhaAtual As Worksheet
  Set wsPlanilhaAtual = ActiveSheet
  
  If HasValorMes(wsPlanilhaAtual) Then
    'Pede confirmação
    If MsgBox("Essa planilha já possui dados nos indicadores. Deseja sobreescrever?", _
        vbYesNo + vbQuestion, "Busca indicadores") = vbNo Then
        Exit Sub
    End If
  End If
  Dim wsIndicadores As Worksheet
  Set wsIndicadores = Worksheets("Indicadores")
  Call AtualizarFonteDeDadosPlanilha(wsIndicadores)
  Call PercorrerIndicadores(wsPlanilhaAtual, wsIndicadores)
  Exit Sub
    
ErroBuscarIndicadores:
  MostrarMsgErro ("BuscarIndicadores")
End Sub

Private Function HasValorMes(wsPlanilha As Worksheet) As Boolean
  '
  ' Function HasValorMes
  ' verifica se existem valores já digitados no mês do indicador
  '
  On Error GoTo ErroHasValorMes
  
  Dim blnExisteDados As Boolean
  blnExisteDados = False
  Dim rgCell As Range
  For Each rgCell In wsPlanilha.Range(RANGE_COLUNA_MES_INDICADORES)
    If rgCell.Value > 0 Then
      blnExisteDados = True
      Exit For
    End If
  Next rgCell
  HasValorMes = blnExisteDados
  Exit Function
  
ErroHasValorMes:
  MostrarMsgErro ("HasValorMes")
End Function

Sub AtualizarFonteDeDadosPlanilha(wsPlanilha As Worksheet)
  '
  ' AtualizarFonteDeDadosPlanilha    Data: 30/04/16
  ' Atualiza funções da planilha
  '
  On Error GoTo ErroAtualizarFonteDeDadosPlanilha
  
  CongelarCalculosPlanilha (True)
  
  If wsPlanilha.ProtectContents Then
    wsPlanilha.Unprotect
  End If
    
  Dim qtDados As QueryTable
  For Each qtDados In wsPlanilha.QueryTables
    qtDados.Refresh
  Next
  
  wsPlanilha.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
  CongelarCalculosPlanilha (False)
  Exit Sub
   
ErroAtualizarFonteDeDadosPlanilha:
  CongelarCalculosPlanilha (False)
  MostrarMsgErro ("AtualizarFonteDeDadosPlanilha")
End Sub

Private Sub PercorrerIndicadores(wsPlanilha As Worksheet, wsIndicadores As Worksheet)
  '
  ' PercorrerIndicadores
  ' verifica cada descrição de indicador
  '
  On Error GoTo ErroPercorrerIndicadores
    
  Dim rgCell, rgFound As Range
  Dim strIndicador As String
  For Each rgCell In wsPlanilha.Range(RANGE_COLUNA_DESCR_INDICADORES)
    strIndicador = rgCell.Value
    'If (strIndicador = "IPCA") Then
    '   strIndicador = "IPCA (5)"
    'End If
    Set rgFound = wsIndicadores.UsedRange.Find(What:=strIndicador)
    If Not rgFound Is Nothing Then
      Dim rgCelulaAtual As Range
      Set rgCelulaAtual = rgCell
      Call TransfereDadosIndicador(rgCelulaAtual, rgFound, wsIndicadores)
    End If
    If (strIndicador = "Dólar Comercial") Then
      Dim rgDolarAtual As Range
      Set rgDolarAtual = BuscarDolarComercialAtual(wsIndicadores)
      If Not rgDolarAtual Is Nothing Then
        Call TransfereValorDolar(rgDolarAtual, wsIndicadores)
      End If
    End If
    
  Next rgCell
  Exit Sub
  
ErroPercorrerIndicadores:
  MostrarMsgErro ("PercorrerIndicadores")
End Sub

Private Sub TransfereDadosIndicador(rgIndicadorAtual As Range, rgIndicadorWeb As Range, wsIndicadores As Worksheet)
  '
  ' PercorrerIndicadores
  ' verifica cada descrição de indicador
  '
  On Error GoTo ErroTransfereDadosIndicador
  
  Dim wsPlanilha As Worksheet
  Set wsPlanilha = ActiveSheet
  Dim rgIndicadorMesAtual, rgIndicadorAnoAtual, rgIndicadorDozeMesesAtual As Range
  Set rgIndicadorMesAtual = GetRangeMesIndicador(rgIndicadorAtual)
  If IsEmpty(wsIndicadores.Cells(rgIndicadorWeb.Row, rgIndicadorWeb.Column + 1)) Then
    Exit Sub
  End If
  Set rgIndicadorAnoAtual = GetRangeAnoIndicador(rgIndicadorAtual)
  Set rgIndicadorDozeMesesAtual = GetRangeDozeMesesIndicador(rgIndicadorAtual)
  With wsPlanilha
    If (rgIndicadorAtual.Value = SP500) Then
      .Cells(rgIndicadorMesAtual.Row, rgIndicadorMesAtual.Column).Value = wsIndicadores.Cells(rgIndicadorWeb.Row, rgIndicadorWeb.Column + 4).Value
      .Cells(rgIndicadorAnoAtual.Row, rgIndicadorAnoAtual.Column).Value = wsIndicadores.Cells(rgIndicadorWeb.Row, rgIndicadorWeb.Column + 5).Value
      .Cells(rgIndicadorDozeMesesAtual.Row, rgIndicadorDozeMesesAtual.Column).Value = wsIndicadores.Cells(rgIndicadorWeb.Row, rgIndicadorWeb.Column + 6).Value
      Exit Sub
    End If
    .Cells(rgIndicadorMesAtual.Row, rgIndicadorMesAtual.Column).Value = wsIndicadores.Cells(rgIndicadorWeb.Row, rgIndicadorWeb.Column + 1).Value
    .Cells(rgIndicadorAnoAtual.Row, rgIndicadorAnoAtual.Column).Value = wsIndicadores.Cells(rgIndicadorWeb.Row, rgIndicadorWeb.Column + 7).Value
    .Cells(rgIndicadorDozeMesesAtual.Row, rgIndicadorDozeMesesAtual.Column).Value = wsIndicadores.Cells(rgIndicadorWeb.Row, rgIndicadorWeb.Column + 8).Value
  End With
  Exit Sub
  
ErroTransfereDadosIndicador:
  MostrarMsgErro ("TransfereDadosIndicador")
End Sub

Private Function GetRangeMesIndicador(rgCell As Range) As Range
  '
  ' Function GetRangeMesIndicador
  ' busca os dados do mês do indicador atual
  '
  On Error GoTo ErroGetRangeMesIndicador
  
  Dim intLinhaIndicador, intColunaIndicador As Integer
  intLinhaIndicador = rgCell.Row
  intColunaIndicador = RetornarPrimeiraColuna(Range(RANGE_COLUNA_MES_INDICADORES))
  Set GetRangeMesIndicador = Cells(intLinhaIndicador, intColunaIndicador)
  Exit Function
  
ErroGetRangeMesIndicador:
  MostrarMsgErro ("GetRangeMesIndicador")
End Function

Private Function GetRangeAnoIndicador(rgCell As Range) As Range
  '
  ' Function GetRangeAnoIndicador
  ' busca os dados do ano do indicador atual
  '
  On Error GoTo ErroGetRangeAnoIndicador
  
  Dim intLinhaIndicador, intColunaIndicador As Integer
  intLinhaIndicador = rgCell.Row
  intColunaIndicador = RetornarPrimeiraColuna(Range(RANGE_COLUNA_ANO_INDICADORES))
  Set GetRangeAnoIndicador = Cells(intLinhaIndicador, intColunaIndicador)
  Exit Function
  
ErroGetRangeAnoIndicador:
  MostrarMsgErro ("GetRangeAnoIndicador")
End Function

Private Function GetRangeDozeMesesIndicador(rgCell As Range) As Range
  '
  ' Function GetRangeDozeMesesIndicador
  ' busca os dados dos últimos doze meses do indicador atual
  '
  On Error GoTo ErroGetRangeDozeMesesIndicador
  
  Dim intLinhaIndicador, intColunaIndicador As Integer
  intLinhaIndicador = rgCell.Row
  intColunaIndicador = RetornarPrimeiraColuna(Range(RANGE_COLUNA_DOZE_MESES_INDICADORES))
  Set GetRangeDozeMesesIndicador = Cells(intLinhaIndicador, intColunaIndicador)
  Exit Function
  
ErroGetRangeDozeMesesIndicador:
  MostrarMsgErro ("GetRangeDozeMesesIndicador")
End Function

Private Function BuscarDolarComercialAtual(wsIndicadores As Worksheet) As Range
  On Error GoTo ErroBuscarDolarComercialAtual
  Dim rgTabDolar As Range
  Set rgTabDolar = wsIndicadores.UsedRange.Find(What:="Dólar & Euro")
  If Not rgTabDolar Is Nothing Then
    Set BuscarDolarComercialAtual = Cells(rgTabDolar.Row + 2, rgTabDolar.Column + 2)
    Exit Function
  End If
  Set BuscarDolarComercialAtual = Nothing
  Exit Function
  
ErroBuscarDolarComercialAtual:
  MostrarMsgErro ("BuscarDolarComercialAtual")
End Function

Private Sub TransfereValorDolar(rgIndicadorWeb As Range, wsIndicadores As Worksheet)
  '
  ' TransfereValorDolar
  ' transfere o valor do dólar atual para planilha corrente
  '
  On Error GoTo ErroTransfereValorDolar
  
  Dim wsPlanilha As Worksheet
  Set wsPlanilha = ActiveSheet
  Dim rgIndicadorValorFinalMesAtual As Range
  Set rgIndicadorValorFinalMesAtual = Range(RANGE_CELULA_DOLAR_FINAL_MES)
  If IsEmpty(wsIndicadores.Cells(rgIndicadorWeb.Row, rgIndicadorWeb.Column)) Then
    Exit Sub
  End If
  wsPlanilha.Cells(rgIndicadorValorFinalMesAtual.Row, rgIndicadorValorFinalMesAtual.Column).Value = wsIndicadores.Cells(rgIndicadorWeb.Row, rgIndicadorWeb.Column).Value
  Exit Sub
  
ErroTransfereValorDolar:
  MostrarMsgErro ("TransfereValorDolar")
End Sub
