Option Explicit

' -------------------------------------------------------------------------
' Módulo: EstaPasta_de_trabalho.cls
' Descrição: Funções que afetam o comportamento geral das planilhas
' -------------------------------------------------------------------------

Private Const PADRAO_PLAN_GRAFICO As String = "graf*"
Private Const FORMATO_DATA_ATUALIZACAO As String = "dddd, dd mmmm yyyy"
Private Const FORMATO_HORA_ATUALIZACAO As String = "hh:mm:ss"
Private Const TITULO_AGENDAR_DARF As String = "Agendar Darf"

' Indica se houve alteração desde o último salvamento
Private mblnAtualizouArquivo As Boolean

'
' sub Workbook_beforeSave
' atualiza a data da última alteração na Planilha Intro
'
Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)  
  Dim wsIntro As Worksheet
  Dim datAgora As Date
  If Not mblnAtualizouArquivo Then
    Exit Sub
  End If
  On Error GoTo ErroBeforeSave
  Application.EnableEvents = False
  Application.ScreenUpdating = False
  datAgora = Now
  Set wsIntro = Me.Worksheets(NOME_PLAN_INTRO)
  wsIntro.Unprotect
  wsIntro.Range(RANGE_DATA_ULTIMA_ATUALIZ).Value = _
    Format$(datAgora, FORMATO_DATA_ATUALIZACAO) & " às " & Format$(datAgora, FORMATO_HORA_ATUALIZACAO)
  RetornarRangeNomeado(RANGE_POSICAO).Value = datAgora
  mblnAtualizouArquivo = False

FimBeforeSave:
  ' garante que a Intro volte protegida e que o Excel seja restaurado
  On Error Resume Next
  If Not wsIntro Is Nothing Then
    wsIntro.Protect
  End If
  Application.ScreenUpdating = True
  Application.EnableEvents = True
  Exit Sub

ErroBeforeSave:
  MostrarMsgErro "Workbook_BeforeSave"
  Resume FimBeforeSave
End Sub

Private Sub Workbook_Open()
  On Error GoTo ErroWorkbookOpen
  mblnAtualizouArquivo = False
  Me.Worksheets(NOME_PLAN_INTRO).Activate
  PosicionarTopo
  Exit Sub

ErroWorkbookOpen:
  MostrarMsgErro "Workbook_Open"
End Sub

'
' Sub Workbook_SheetChange
' verifica se atualizou alguma planilha ou se está digitando a descrição
' de um investimento nas planilhas mensais
'
Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)  
  Dim wsPlanilha As Worksheet
  Dim blnCriticouDescInvest As Boolean, blnDesabilitouEventos As Boolean
  ' ajusta variável para avisar que houve mudanças
  mblnAtualizouArquivo = True
  If Not TypeOf Sh Is Worksheet Then
    Exit Sub
  End If
  On Error GoTo ErroSheetChange
  Set wsPlanilha = Sh
  ' testa se é uma planilha mensal aberta
  If Not IsPlanilhaAberta(wsPlanilha.Range(RANGE_SITUAC_PLANILHA)) Then
    Exit Sub
  End If
  Application.EnableEvents = False
  blnDesabilitouEventos = True

  If Target.CountLarge = 1 Then
    VerificarTolerancia wsPlanilha, Target
  End If
  blnCriticouDescInvest = CriticarDescricaoInvestimento(wsPlanilha, Target)
  If Target.CountLarge = 1 Then
    VerificarImpostoAPagar wsPlanilha, Target
  End If

FimSheetChange:
  On Error Resume Next
  ' só deve deixar a tela atualizar se passou pela crítica do investimento
  If blnCriticouDescInvest Then
    Application.ScreenUpdating = True
  End If
  If blnDesabilitouEventos Then
    Application.EnableEvents = True
  End If
  Exit Sub

ErroSheetChange:
  MostrarMsgErro "Workbook_SheetChange"
  Resume FimSheetChange
End Sub

'
' Avisa quando um lançamento faz o saldo do mês alcançar o valor de tolerância
'
Private Sub VerificarTolerancia(ByVal wsPlanilha As Worksheet, ByVal rgAlvo As Range)  
  Dim dblValor As Double
  If IsEmpty(rgAlvo.Value) Or Not IsNumeric(rgAlvo.Value) Then
    Exit Sub
  End If
  dblValor = CDbl(rgAlvo.Value)
  ' movimentações afetam o saldo da planilha atual
  If IsIntersecao(rgAlvo, wsPlanilha, RANGE_COLUNA_VALOR_MOVIMENTACAO) Then
    If IsAlcancouValorCriticoNoMes(wsPlanilha, dblValor) Then
      Alertar "Cuidado! Valor de tolerância alcançado para este mês.", vbExclamation
    End If
  End If
  ' cartões afetam o saldo da próxima planilha
  If IsIntersecao(rgAlvo, wsPlanilha, RANGE_COLUNA_VALOR_CARTOES) Then
    If IsAlcancouValorCriticoNoMes(RetornarProximaPlanilha(wsPlanilha), dblValor) Then
      Alertar "Cuidado! Valor de tolerância alcançado para o próximo mês.", vbExclamation
    End If
  End If
End Sub

'
' Critica a descrição dos investimentos digitados; retorna True se a crítica foi executada
'
Private Function CriticarDescricaoInvestimento(ByVal wsPlanilha As Worksheet, ByVal Target As Range) As Boolean  
  Dim rgValor As Range
  Dim varTipo As Variant
  Dim lngLinha As Long
  Set rgValor = Application.Intersect(Target, wsPlanilha.Range(RANGE_COLUNA_VALOR_MOVIMENTACAO))
  If rgValor Is Nothing Then
    Exit Function
  End If
  If IsEmpty(rgValor.Cells(1).Value) Then
    Exit Function
  End If
  lngLinha = rgValor.Cells(1).Row
  varTipo = wsPlanilha.Cells(lngLinha, wsPlanilha.Range(RANGE_HEADER_TIPO_MOVIMENTACAO).Column).Value
  If IsError(varTipo) Then
    Exit Function
  End If
  If CStr(varTipo) <> TIPO_LANCAMENTO_INVESTIMENTOS Then
    Exit Function
  End If
  Application.ScreenUpdating = False
  CriticarInvestimento wsPlanilha.Cells(lngLinha, wsPlanilha.Range(RANGE_HEADER_DESC_MOVIMENTACAO).Column)
  CriticarDescricaoInvestimento = True
End Function

'
' Testa se alterou os valores finais de ações, fii's ou etf's e se gerou imposto a pagar
'
Private Sub VerificarImpostoAPagar(ByVal wsPlanilha As Worksheet, ByVal rgAlvo As Range)  
  If IsEmpty(rgAlvo.Value) Then
    Exit Sub
  End If
  If Not IsIntersecao(rgAlvo, wsPlanilha, RANGE_COLUNA_SALDO_FINAL_ACOES) And _
     Not IsIntersecao(rgAlvo, wsPlanilha, RANGE_COLUNA_SALDO_FINAL_FII) And _
     Not IsIntersecao(rgAlvo, wsPlanilha, RANGE_COLUNA_SALDO_FINAL_ETF) Then
    Exit Sub
  End If
  If LerBooleano(wsPlanilha.Range(RANGE_CELULA_IGNORA_AGENDA_CRIPTO)) Then
    Exit Sub
  End If
  If Not IsGerouImpostoNoMes(wsPlanilha, rgAlvo) Then
    Exit Sub
  End If
  If Alertar("Atenção! Você gerou imposto a pagar para o próximo mês." & vbLf & _
             "Deseja criar um agendamento para este pagamento?", _
             vbYesNo + vbQuestion, TITULO_AGENDAR_DARF) = vbYes Then
    AgendarLembreteOutlook
  End If
  wsPlanilha.Range(RANGE_CELULA_IGNORA_AGENDA_CRIPTO).Value = True
End Sub

Private Function Alertar(ByVal strMensagem As String, ByVal lngBotoes As VbMsgBoxStyle, _
    Optional ByVal strTitulo As String) As VbMsgBoxResult
  Beep
  Beep
  If Len(strTitulo) = 0 Then
    Alertar = MsgBox(strMensagem, lngBotoes)
  Else
    Alertar = MsgBox(strMensagem, lngBotoes, strTitulo)
  End If
End Function

Private Function IsIntersecao(ByVal rgAlvo As Range, ByVal wsPlanilha As Worksheet, ByVal strRange As String) As Boolean
  IsIntersecao = Not (Application.Intersect(rgAlvo, wsPlanilha.Range(strRange)) Is Nothing)
End Function

'
' Lê uma célula de marcação (VERDADEIRO/FALSO); textos e erros valem Falso
'
Private Function LerBooleano(ByVal rgCelula As Range) As Boolean  
  Dim varValor As Variant
  varValor = rgCelula.Value
  If IsError(varValor) Then
    Exit Function
  End If
  If VarType(varValor) = vbBoolean Then
    LerBooleano = varValor
  ElseIf IsNumeric(varValor) Then
    LerBooleano = (RetornarValorNumerico(varValor) <> 0)
  End If
End Function

Private Function IsResultadoPositivo(ByVal wsPlanilha As Worksheet, ByVal lngLinha As Long, _
    ByVal strRangeResultado As String) As Boolean
  IsResultadoPositivo = (RetornarValorNumerico(wsPlanilha.Cells(lngLinha, _
    wsPlanilha.Range(strRangeResultado).Column).Value) > 0)
End Function

'
' Function IsAlcancouValorCriticoNoMes
' verifica se alcançou o ponto crítico neste mês
'
Private Function IsAlcancouValorCriticoNoMes(ByVal wsPlanilha As Worksheet, ByVal dblValor As Double) As Boolean  
  Dim dblValorTolerancia As Double, dblSaldoAtual As Double
  On Error GoTo ErroIsAlcancouValorCriticoNoMes
  ' Testa se a planilha existe, é mensal e está aberta
  ' este teste é necessário caso esteja se vendo a planilha do próximo mês
  If wsPlanilha Is Nothing Then
    Exit Function
  End If
  If Not IsPlanilhaAberta(wsPlanilha.Range(RANGE_SITUAC_PLANILHA)) Then
    Exit Function
  End If
  dblValorTolerancia = RetornarValorNumerico(RetornarRangeNomeado(RANGE_TOLERANCIA).Value)
  dblSaldoAtual = RetornarValorNumerico(wsPlanilha.Range(RANGE_SALDO_MES).Value)
  IsAlcancouValorCriticoNoMes = (dblSaldoAtual < dblValorTolerancia) And _
                                ((dblSaldoAtual + dblValor) >= dblValorTolerancia)
  Exit Function

ErroIsAlcancouValorCriticoNoMes:
  MostrarMsgErro "IsAlcancouValorCriticoNoMes"
End Function

'
' Function IsGerouImpostoNoMes
' verifica se gerou imposto a pagar
'
Private Function IsGerouImpostoNoMes(ByVal wsPlanilha As Worksheet, ByVal rgAlvo As Range) As Boolean  
  Dim lngLinha As Long
  On Error GoTo ErroIsGerouImpostoNoMes
  lngLinha = rgAlvo.Row

  If IsIntersecao(rgAlvo, wsPlanilha, RANGE_COLUNA_SALDO_FINAL_ACOES) Then
    If LerBooleano(wsPlanilha.Range(RANGE_CELULA_TRIBUTA_ACOES)) And _
       (IsResultadoPositivo(wsPlanilha, lngLinha, RANGE_COLUNA_RESULTADO_COMUM_ACOES) Or _
        IsResultadoPositivo(wsPlanilha, lngLinha, RANGE_COLUNA_RESULTADO_DAYTRADE_ACOES)) Then
      IsGerouImpostoNoMes = True
      Exit Function
    End If
  End If

  If IsIntersecao(rgAlvo, wsPlanilha, RANGE_COLUNA_SALDO_FINAL_FII) Then
    If IsResultadoPositivo(wsPlanilha, lngLinha, RANGE_COLUNA_RESULTADO_COMUM_FII) Or _
       IsResultadoPositivo(wsPlanilha, lngLinha, RANGE_COLUNA_RESULTADO_DAYTRADE_FII) Then
      IsGerouImpostoNoMes = True
      Exit Function
    End If
  End If

  If IsIntersecao(rgAlvo, wsPlanilha, RANGE_COLUNA_SALDO_FINAL_ETF) Then
    If IsResultadoPositivo(wsPlanilha, lngLinha, RANGE_COLUNA_RESULTADO_COMUM_ETF) Or _
       IsResultadoPositivo(wsPlanilha, lngLinha, RANGE_COLUNA_RESULTADO_DAYTRADE_ETF) Then
      IsGerouImpostoNoMes = True
      Exit Function
    End If
  End If

  ' considera a soma de Stock, Reit, Treasures e Ouro
  If IsIntersecao(rgAlvo, wsPlanilha, RANGE_COLUNA_SALDO_FINAL_STOCK) Or _
     IsIntersecao(rgAlvo, wsPlanilha, RANGE_COLUNA_SALDO_FINAL_REIT) Or _
     IsIntersecao(rgAlvo, wsPlanilha, RANGE_COLUNA_SALDO_FINAL_TREASURY) Or _
     IsIntersecao(rgAlvo, wsPlanilha, RANGE_COLUNA_SALDO_FINAL_OURO) Then
    If LerBooleano(wsPlanilha.Range(RANGE_CELULA_TRIBUTA_STOCK)) And _
       (IsResultadoPositivo(wsPlanilha, lngLinha, RANGE_COLUNA_RESULTADO_COMUM_STOCK) Or _
        IsResultadoPositivo(wsPlanilha, lngLinha, RANGE_COLUNA_RESULTADO_COMUM_REIT) Or _
        IsResultadoPositivo(wsPlanilha, lngLinha, RANGE_COLUNA_RESULTADO_COMUM_TREASURY) Or _
        IsResultadoPositivo(wsPlanilha, lngLinha, RANGE_COLUNA_RESULTADO_COMUM_OURO)) Then
      IsGerouImpostoNoMes = True
      Exit Function
    End If
  End If

  If IsIntersecao(rgAlvo, wsPlanilha, RANGE_COLUNA_SALDO_FINAL_CRIPTO) Then
    If IsResultadoPositivo(wsPlanilha, lngLinha, RANGE_COLUNA_RESULTADO_COMUM_CRIPTO) Then
      IsGerouImpostoNoMes = True
      Exit Function
    End If
  End If
  Exit Function

ErroIsGerouImpostoNoMes:
  MostrarMsgErro "IsGerouImpostoNoMes"
End Function

'
' Sub Workbook_BeforeClose
' posiciona todas as planilhas no topo
'
Private Sub Workbook_BeforeClose(Cancel As Boolean)  
  Dim wsPlanilha As Worksheet
  Dim objPlanilhaAtual As Object
  ' o fechamento da pasta não deve ser interrompido por falhas no posicionamento
  On Error Resume Next
  Application.ScreenUpdating = False
  Set objPlanilhaAtual = Me.ActiveSheet
  For Each wsPlanilha In Me.Worksheets
    If wsPlanilha.Visible = xlSheetVisible And Not LCase$(wsPlanilha.Name) Like PADRAO_PLAN_GRAFICO Then
      wsPlanilha.Activate
      ActiveWindow.ScrollRow = 1
      ActiveWindow.ScrollColumn = 1
    End If
  Next wsPlanilha
  If Not objPlanilhaAtual Is Nothing Then
    objPlanilhaAtual.Activate
  End If
  Application.ScreenUpdating = True
End Sub
