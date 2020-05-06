' Módulo de funções para atualização dos indicadores
Option Explicit

Sub AtualizarDadosPlanAtual()
  '
  ' AtualizarDadosPlanAtual    Data: 30/04/16
  ' Atualiza tabela com fonte de dados na web
  '
  On Error GoTo ErroAtualizarDadosPlanAtual
    
  Dim wsPlanilha As Worksheet
  Dim blnOldStatusBar As Boolean
  Dim intPercentual As Integer
  blnOldStatusBar = Application.DisplayStatusBar
  intPercentual = 0
  Application.DisplayStatusBar = True
  Application.StatusBar = "Importando valores.. " & intPercentual & "% Completado."
  CongelarCalculosPlanilha (True)
  Set wsPlanilha = ActiveSheet
  Call AtualizarFonteDeDadosPlanilha(wsPlanilha, 100)
  intPercentual = 100
  Application.StatusBar = "Atualizando valores.. " & intPercentual & "% Completado."
FimAtualizarDadosPlanAtual:
  CongelarCalculosPlanilha (False)
  Application.StatusBar = False
  Application.DisplayStatusBar = blnOldStatusBar
  Set wsPlanilha = Nothing
  Exit Sub
   
ErroAtualizarDadosPlanAtual:
  MostrarMsgErro ("AtualizarDadosPlanAtual")
  Resume FimAtualizarDadosPlanAtual
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
  Dim wsPlanilhaAtual As Worksheet, wsProxPlanilha As Worksheet, wsIndicadores As Worksheet
  Dim blnOldStatusBar As Boolean
  Dim intPercentual As Integer
  Dim lngIndPlan As Long
  
  Set wsPlanilhaAtual = ActiveSheet
  If HasValorMes(wsPlanilhaAtual) Then
    'Pede confirmação
    If MsgBox("Essa planilha já possui dados nos indicadores. Deseja sobreescrever?", _
        vbYesNo + vbQuestion, "Busca indicadores") = vbNo Then
        Exit Sub
    End If
  End If
  
  blnOldStatusBar = Application.DisplayStatusBar
  intPercentual = 0
  Application.DisplayStatusBar = True
  Application.StatusBar = "Importando valores.. " & intPercentual & "% Completado."
  CongelarCalculosPlanilha (True)
  Set wsIndicadores = Worksheets("Web")
  Call AtualizarFonteDeDadosPlanilha(wsIndicadores, intPercentual)
  intPercentual = 90
  Application.StatusBar = "Atualizando valores.. " & intPercentual & "% Completado."
  
  lngIndPlan = Worksheets(ActiveSheet.Name).Index
  Set wsProxPlanilha = Worksheets(lngIndPlan + 1)
  If IsPlanilhaAberta(wsProxPlanilha.Range(RANGE_SITUAC_PLANILHA)) Then
    Application.StatusBar = "Atualizando valores.. " & intPercentual & "% Completado."
    Call TransferirDolarBacen(wsProxPlanilha, wsIndicadores)
  End If
  intPercentual = 95
  Application.StatusBar = "Atualizando valores.. " & intPercentual & "% Completado."
  
  Call PercorrerIndicadores(wsPlanilhaAtual, wsIndicadores)
  intPercentual = 100
  Application.StatusBar = "Atualizando valores.. " & intPercentual & "% Completado."
      
FimBuscarIndicadores:
  CongelarCalculosPlanilha (False)
  Application.StatusBar = False
  Application.DisplayStatusBar = blnOldStatusBar
  Set wsPlanilhaAtual = Nothing
  Set wsProxPlanilha = Nothing
  Exit Sub
    
ErroBuscarIndicadores:
  MostrarMsgErro ("BuscarIndicadores")
  Resume FimBuscarIndicadores
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

Private Sub AtualizarFonteDeDadosPlanilha(wsPlanilha As Worksheet, intTotalPercentual As Integer)
  '
  ' AtualizarFonteDeDadosPlanilha    Data: 30/04/16
  ' Atualiza funções da planilha
  '
  On Error GoTo ErroAtualizarFonteDeDadosPlanilha
  
  If wsPlanilha.ProtectContents Then
    wsPlanilha.Unprotect
  End If
    
  Dim intPercentualAtual As Integer
  intPercentualAtual = 0
  Dim lngTotalConexoes As Long
  lngTotalConexoes = ContarConexoes(wsPlanilha)
  Dim intAcrescimo As Integer
  intAcrescimo = intTotalPercentual / lngTotalConexoes
  'Atualizar queryTables
  Dim qtDados As QueryTable
  Dim tmTempo As Double
  For Each qtDados In wsPlanilha.QueryTables
    Debug.Print "Intervalo de dados [" & qtDados.Name & "]"
    Application.StatusBar = "Importando consulta " & qtDados.Name & ".. " & intPercentualAtual & "% Completado."
    tmTempo = Timer
    qtDados.Refresh
    Debug.Print "Tempo de processamento: " & Round(Timer - tmTempo, 4) & " seg."
    intPercentualAtual = intPercentualAtual + intAcrescimo
    If intPercentualAtual > intTotalPercentual Then
      intPercentualAtual = intTotalPercentual
    End If
  Next
  'Atualizar consultas
  Dim lstObjetos As ListObject
  For Each lstObjetos In wsPlanilha.ListObjects
    If lstObjetos.SourceType = xlSrcQuery Then
      Debug.Print "Consulta [" & lstObjetos.Name & "]"
      Application.StatusBar = "Importando consulta " & lstObjetos.Name & ".. " & intPercentualAtual & "% Completado."
      tmTempo = Timer
      lstObjetos.Refresh
      Debug.Print "Tempo de processamento: " & Round(Timer - tmTempo, 4) & " seg."
      intPercentualAtual = intPercentualAtual + intAcrescimo
      If intPercentualAtual > intTotalPercentual Then
        intPercentualAtual = intTotalPercentual
      End If
    End If
  Next lstObjetos
  
  'wsPlanilha.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
  Exit Sub
   
ErroAtualizarFonteDeDadosPlanilha:
  MostrarMsgErro ("AtualizarFonteDeDadosPlanilha")
End Sub

Private Function ContarConexoes(wsPlanilha As Worksheet) As Long
  On Error GoTo ErroContarConexoes
  Dim lngConexoes As Long
  Dim lstObjetos As ListObject
  lngConexoes = 0
  For Each lstObjetos In wsPlanilha.ListObjects
    Debug.Print "Conexão [" & lstObjetos.Name & "(SourceType " & lstObjetos.SourceType & ")]"
    If lstObjetos.SourceType = xlSrcQuery Then
        lngConexoes = lngConexoes + 1
    End If
  Next lstObjetos
  ContarConexoes = lngConexoes + wsPlanilha.QueryTables.Count
  Exit Function
  
ErroContarConexoes:
  MostrarMsgErro ("ContarConexoes")
End Function

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
    Set rgFound = wsIndicadores.UsedRange.Find(What:=strIndicador)
    If Not rgFound Is Nothing Then
      Dim rgCelulaAtual As Range
      Set rgCelulaAtual = rgCell
      'Call TransferirDadosIndicador(rgCelulaAtual, rgFound, wsIndicadores)
    End If
    If (strIndicador = "Dólar Comercial") Then
      Dim rgDolarAtual As Range
      Set rgDolarAtual = BuscarDolarComercialAtual(wsIndicadores)
      If Not rgDolarAtual Is Nothing Then
        Call TransferirDolarComercial(rgDolarAtual, wsIndicadores)
      End If
    End If
    
  Next rgCell
  Exit Sub
  
ErroPercorrerIndicadores:
  MostrarMsgErro ("PercorrerIndicadores")
End Sub

Private Sub TransferirDadosIndicador(rgIndicadorAtual As Range, rgIndicadorWeb As Range, wsIndicadores As Worksheet)
  '
  ' ============== Esta consulta foi removida ======================
  ' TransferirDadosIndicador
  ' verifica cada descrição de indicador
  '
  On Error GoTo ErroTransferirDadosIndicador
  
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
  
ErroTransferirDadosIndicador:
  MostrarMsgErro ("TransferirDadosIndicador")
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
  Set rgTabDolar = wsIndicadores.UsedRange.Find(What:="Moeda")
  If Not rgTabDolar Is Nothing Then
    Set BuscarDolarComercialAtual = Cells(rgTabDolar.Row + 1, rgTabDolar.Column + 1)
    Exit Function
  End If
  Set BuscarDolarComercialAtual = Nothing
  Exit Function
  
ErroBuscarDolarComercialAtual:
  MostrarMsgErro ("BuscarDolarComercialAtual")
End Function

Private Sub TransferirDolarComercial(rgIndicadorWeb As Range, wsIndicadores As Worksheet)
  '
  ' TransferirDolarComercial
  ' transfere o valor do dólar atual para planilha corrente
  '
  On Error GoTo ErroTransferirDolarComercial
  
  Dim wsPlanilha As Worksheet
  Set wsPlanilha = ActiveSheet
  Dim rgIndicadorValorFinalMesAtual As Range
  Set rgIndicadorValorFinalMesAtual = Range(RANGE_CELULA_DOLAR_FINAL_MES)
  If IsEmpty(wsIndicadores.Cells(rgIndicadorWeb.Row, rgIndicadorWeb.Column)) Then
    Exit Sub
  End If
  wsPlanilha.Cells(rgIndicadorValorFinalMesAtual.Row, rgIndicadorValorFinalMesAtual.Column).Value = wsIndicadores.Cells(rgIndicadorWeb.Row, rgIndicadorWeb.Column).Value
  Exit Sub
  
ErroTransferirDolarComercial:
  MostrarMsgErro ("ErroTransferirDolarComercial")
End Sub

Private Sub TransferirDolarBacen(wsProxPlanilha As Worksheet, wsIndicadores As Worksheet)
  On Error GoTo ErroTransferirDolarBacen

  Dim rgTabBacen As Range
  Set rgTabBacen = wsIndicadores.UsedRange.Find(What:="Mês de recebimento")
  If rgTabBacen Is Nothing Then
    Debug.Print "Nenhum valor de dólar do Bacen foi encontrado."
    Exit Sub
  End If
  wsProxPlanilha.Range(RANGE_CELULA_DOLAR_BACEN_COMPRA).Value = wsIndicadores.Cells(rgTabBacen.Row + GetLinhaMes(wsProxPlanilha), rgTabBacen.Column + 1).Value
  Debug.Print "Dólar de compra [" & wsIndicadores.Cells(rgTabBacen.Row + GetLinhaMes(wsProxPlanilha), rgTabBacen.Column + 1).Value & "]"
  wsProxPlanilha.Range(RANGE_CELULA_DOLAR_BACEN_VENDA).Value = wsIndicadores.Cells(rgTabBacen.Row + GetLinhaMes(wsProxPlanilha), rgTabBacen.Column + 2).Value
  Exit Sub

ErroTransferirDolarBacen:
  MostrarMsgErro ("TransferirDolarBacen")
End Sub

Private Function GetLinhaMes(wsProxPlanilha As Worksheet) As Integer
  On Error GoTo ErroGetLinhaMes
  Dim strNomePlanilha As String

  strNomePlanilha = wsProxPlanilha.Name
  Select Case strNomePlanilha
    Case "Jan"
      GetLinhaMes = 1
    Case "Fev"
      GetLinhaMes = 2
    Case "Mar"
      GetLinhaMes = 3
    Case "Abr"
      GetLinhaMes = 4
    Case "Mai"
      GetLinhaMes = 5
    Case "Jun"
      GetLinhaMes = 6
    Case "Jul"
      GetLinhaMes = 7
    Case "Ago"
      GetLinhaMes = 8
    Case "Set"
      GetLinhaMes = 9
    Case "Out"
      GetLinhaMes = 10
    Case "Nov"
      GetLinhaMes = 11
    Case Else
      GetLinhaMes = 12
  End Select
  Exit Function
  
ErroGetLinhaMes:
  MostrarMsgErro ("GetLinhaMes")
End Function

