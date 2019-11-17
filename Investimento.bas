' Módulo de funções sobre o controle de Investimentos
Option Explicit

Function CalcularSaldoAtual(rgDescInvest As Range) As Double
  '
  ' Função saldoAtual   Data: 13/01/04
  ' Determina o saldo atual do investimento de acordo com
  ' o saldo inicial da Carteira da primeira planilha
  ' mensal aberta
  '
  ' Variáveis
  Dim wsPlanilha As Worksheet
  ' Principal
  On Error GoTo ErroSaldoAtual
  For Each wsPlanilha In Worksheets
     ' A primeira planilha aberta que achar é a válida
     If IsPlanilhaAberta(wsPlanilha.Range(RANGE_SITUAC_PLANILHA)) Then
       CalcularSaldoAtual = Application.WorksheetFunction.SumIf(wsPlanilha.Range(RANGE_COLUNA_ATIVO_PORTFOLIO), _
             rgDescInvest, wsPlanilha.Range(RANGE_COLUNA_SALDO_FINAL_PORTFOLIO))
       Exit Function
     End If
  Next wsPlanilha
  ' Se todas as planilhas estão fechadas, pega a de Dezembro
  CalcularSaldoAtual = Application.WorksheetFunction.SumIf(Worksheets("Dez.").Range(RANGE_COLUNA_ATIVO_PORTFOLIO), _
       rgDescInvest, Worksheets("Dez.").Range(RANGE_COLUNA_SALDO_FINAL_PORTFOLIO))
  Exit Function
  
ErroSaldoAtual:
  MostrarMsgErro ("CalcularSaldoAtual")
End Function

Function CalcularRendAtivo(rgColDescAtivo As Range, _
                        rgColSaldoInicialAtivo As Range, _
                        rgColAplicacaoAtivo As Range, _
                        rgColRetornoAtivo As Range, _
                        rgColResgateAtivo As Range, _
                        rgColSaldoFinalAtivo As Range, _
                        rgCelDescRentabilidade As Range) As Double
  '
  ' Função CalcularRendAtivo   Data: 29/04/16
  ' Retorna o valor percentual do rendimento líquedo da aplicação
  ' Nota: todas as células envolvidas devem estar nos parâmetros de
  '  entrada da função para que o Excel possa saber que deve recal-
  '  cular a função caso uma delas mude.
  ' A outra maneira é teclar Ctrl + Alt + F9
  '
  On Error GoTo ErroCalcularRendAtivo
  Dim dblSaldoInicial As Double, dblAplicacao As Double, dblRetorno As Double, dblResgate As Double, dblSaldoFinal As Double, dblCpmf As Double
  dblSaldoInicial = Application.WorksheetFunction.SumIf(rgColDescAtivo, rgCelDescRentabilidade, rgColSaldoInicialAtivo)
  dblAplicacao = Application.WorksheetFunction.SumIf(rgColDescAtivo, rgCelDescRentabilidade, rgColAplicacaoAtivo)
  dblRetorno = Application.WorksheetFunction.SumIf(rgColDescAtivo, rgCelDescRentabilidade, rgColRetornoAtivo)
  dblResgate = Application.WorksheetFunction.SumIf(rgColDescAtivo, rgCelDescRentabilidade, rgColResgateAtivo)
  dblSaldoFinal = Application.WorksheetFunction.SumIf(rgColDescAtivo, rgCelDescRentabilidade, rgColSaldoFinalAtivo)
  
  dblCpmf = dblSaldoFinal - dblSaldoInicial - dblAplicacao + dblResgate - dblRetorno
  
  
  If (dblSaldoInicial + dblAplicacao) <> 0 Then
    CalcularRendAtivo = (((dblSaldoFinal + dblResgate - dblCpmf) / (dblSaldoInicial + dblAplicacao)) - 1) * 100
  Else
    CalcularRendAtivo = 0
  End If
  Exit Function
  
ErroCalcularRendAtivo:
  MostrarMsgErro ("CalcularRendAtivo")
End Function


Sub CriticarInvestimento(ByVal rgAlvo As Range)
  '
  ' Sub CriticarInvestimento    Criado por: MSS  Em: 31.01.04
  ' critica o investimento digitado
  '
  ' variáveis
  Dim strTemp As String
  ' principal
  On Error GoTo ErroCriticarInvestimento
  If (rgAlvo.Value = BROKER Or rgAlvo.Value = RESERVA_ESTRATEGICA) Then
    Exit Sub
  End If
  If Not HasCarteira(rgAlvo.Value) Then
    strTemp = RetornarParteNomeCarteira(rgAlvo.Value)
    If strTemp > "" Then
      If MsgBox("Você se refere a " & vbLf & _
                strTemp & " ?", vbQuestion + vbYesNo, "Investimentos") = vbYes Then
        rgAlvo.Value = strTemp
        Exit Sub
      End If
    End If
    If IsReserva(rgAlvo.Value) Then
      Exit Sub
    End If
    MsgBox "Não foi encontrado um investimento com esta descrição." & vbNewLine & _
               "Procure cadastrar o mesmo em uma de suas carteiras.", vbExclamation
  End If
  Exit Sub
ErroCriticarInvestimento:
  MostrarMsgErro ("CriticarInvestimento")
End Sub

Private Function HasCarteira(strDescricao As String) As Boolean
  '
  ' Function HasCarteira
  ' procura a descrição fornecida dentro da Carteira1 e Carteira2
  '
  ' variáveis
  Dim intInicioLinhaCarteira1 As Integer, intFinalLinhaCarteira1 As Integer, _
      intInicioLinhaCarteira2 As Integer, intFinalLinhaCarteira2 As Integer, _
      intColunaCarteira As Integer
  intInicioLinhaCarteira1 = Range(RANGE_CELULA_INICIO_ADHOC).Row
  intFinalLinhaCarteira1 = Range(RANGE_CELULA_FIM_ADHOC).Row
  intInicioLinhaCarteira2 = Range(RANGE_CELULA_INICIO_PORTFOLIO).Row
  intFinalLinhaCarteira2 = Range(RANGE_CELULA_FIM_PORTFOLIO).Row
  intColunaCarteira = Range(RANGE_CELULA_INICIO_ADHOC).Column
  Dim blnAchou As Boolean
  Dim intLinha As Integer
  ' principal
  On Error GoTo ErroHasCarteira
  blnAchou = False
  For intLinha = intInicioLinhaCarteira1 To intFinalLinhaCarteira1
     If IsEmpty(Sheets("Alocacao").Cells(intLinha, intColunaCarteira)) Then
        Exit For
     End If
     If Sheets("Alocacao").Cells(intLinha, intColunaCarteira).Value = strDescricao Then
       blnAchou = True
       Exit For
     End If
  Next intLinha
  If Not blnAchou Then
    For intLinha = intInicioLinhaCarteira2 To intFinalLinhaCarteira2
      If IsEmpty(Sheets("Alocacao").Cells(intLinha, intColunaCarteira)) Then
        Exit For
      End If
      If Sheets("Alocacao").Cells(intLinha, intColunaCarteira).Value = strDescricao Then
       blnAchou = True
       Exit For
     End If
    Next intLinha
  End If
  HasCarteira = blnAchou
  Exit Function
  
ErroHasCarteira:
  MostrarMsgErro ("HasCarteira")
End Function

Private Function RetornarParteNomeCarteira(strDescricao As String) As String
  '
  ' Function achaSbstrCarteira
  ' procura por parte do nome dado na carteira1 e carteira2
  '
  ' variáveis
  Dim intInicioLinhaCarteira1 As Integer, intFinalLinhaCarteira1 As Integer, _
      intInicioLinhaCarteira2 As Integer, intFinalLinhaCarteira2 As Integer, _
      intColunaCarteira As Integer
  intInicioLinhaCarteira1 = Range(RANGE_CELULA_INICIO_ADHOC).Row
  intFinalLinhaCarteira1 = Range(RANGE_CELULA_FIM_ADHOC).Row
  intInicioLinhaCarteira2 = Range(RANGE_CELULA_INICIO_PORTFOLIO).Row
  intFinalLinhaCarteira2 = Range(RANGE_CELULA_FIM_PORTFOLIO).Row
  intColunaCarteira = Range(RANGE_CELULA_INICIO_ADHOC).Column
  Dim blnAchou As Boolean
  Dim strEncontrado As String
  Dim intLinha As Integer
  ' principal
  On Error GoTo ErroRetornarParteNomeCarteira
  blnAchou = False
  strEncontrado = ""
  For intLinha = intInicioLinhaCarteira1 To intFinalLinhaCarteira1
     If IsEmpty(Sheets("Alocacao").Cells(intLinha, intColunaCarteira)) Then
        Exit For
     End If
     If Not IsEmpty(Sheets("Alocacao").Cells(intLinha, intColunaCarteira)) And _
        InStr(1, LCase(Sheets("Alocacao").Cells(intLinha, intColunaCarteira).Value), LCase(strDescricao), 1) > 0 Then
       blnAchou = True
       strEncontrado = Sheets("Alocacao").Cells(intLinha, intColunaCarteira).Value
       Exit For
     End If
  Next intLinha
  If Not blnAchou Then
    For intLinha = intInicioLinhaCarteira2 To intFinalLinhaCarteira2
      If IsEmpty(Sheets("Alocacao").Cells(intLinha, 3)) Then
        Exit For
      End If
      If InStr(1, LCase(Sheets("Alocacao").Cells(intLinha, intColunaCarteira).Value), LCase(strDescricao), 1) > 0 Then
        blnAchou = True
        strEncontrado = Sheets("Alocacao").Cells(intLinha, intColunaCarteira).Value
        Exit For
      End If
    Next intLinha
  End If
  RetornarParteNomeCarteira = strEncontrado
  Exit Function
  
ErroRetornarParteNomeCarteira:
  MostrarMsgErro ("RetornarParteNomeCarteira")
End Function

Private Function IsReserva(strDescricao As String) As Boolean
  '
  ' Function IsReserva
  ' verifica se trata de uma Reserva Estratégica
  '
  ' variáveis
  Dim blnAchou As Boolean
  ' principal
  On Error GoTo ErroIsReserva
  blnAchou = False
  If InStr(1, strDescricao, "Reserva") = 1 Then
    blnAchou = True
  End If
  IsReserva = blnAchou
  Exit Function
  
ErroIsReserva:
  MostrarMsgErro ("IsReserva")
End Function
  
Sub AgendarLembreteOutlook()
  '
  ' Sub AgendarLembreteOutlook    Criado por: MSS  Em: 06.10.19
  ' cria um lembrete no calendário do Outlook
  '
  ' variáveis
  Dim objOutlook As Object
  Dim objAppointmentItem As Object
  Const APPOINTMENT As Integer = 1 '1 = Appointment
  Const OCUPADO As Integer = 2 '1 = Provisório, 2 = Ocupado, 3 = Ausente, 4 = Trabalhando em outro lugar, 5 = Disponível
  Const ONE_DAY As Integer = 1440
  ' principal
  On Error GoTo ErroAgendarLembreteOutlook
  Set objOutlook = CreateObject("Outlook.Application")
  Set objAppointmentItem = objOutlook.createitem(APPOINTMENT)
  With objAppointmentItem
    .Subject = "Pagar Darf"
    .Body = "Pagar imposto com código 6015, Ganhos líquidos em operações em bolsa."
    .Location = ""
    .Start = BuscarUltimoDiaUtilProxMes() + TimeValue("10:00:00") '#10/28/2019 10:00:00 AM#
    '.End = BuscarUltimoDiaUtilProxMes() + TimeValue("10:30:00") '#10/28/2019 10:30:00 AM#
    .Duration = 30 'duração em minutos
    .BusyStatus = OCUPADO
    .ReminderSet = True
    .ReminderMinutesBeforeStart = ONE_DAY
    .Save
  End With
  Debug.Print objAppointmentItem.Subject & " : " & objAppointmentItem.Start 'escreve na janela de verificação imediata
  Set objAppointmentItem = Nothing
  Set objOutlook = Nothing
  Exit Sub
  
ErroAgendarLembreteOutlook:
  MostrarMsgErro ("AgendarLembreteOutlook")
End Sub

Private Function BuscarUltimoDiaUtilProxMes() As Date
  Dim intAno, intMes, intDiaSemana As Integer
  Dim dtData As Date
  Const SABADO As Integer = 7
  Const DOMINGO As Integer = 1
  On Error GoTo ErroBuscarUltimoDiaUtilProxMesMenos3
  intAno = Year(Now)
  intMes = Month(Now)
  dtData = DateAdd("m", 1, DateSerial(intAno, intMes + 1, 0))
  intDiaSemana = Weekday(dtData)
  If (intDiaSemana = DOMINGO) Then
    dtData = dtData - 2
  ElseIf (intDiaSemana = SABADO) Then
    dtData = dtData - 1
  End If
  BuscarUltimoDiaUtilProxMes = dtData
  Exit Function
  
ErroBuscarUltimoDiaUtilProxMesMenos3:
  MostrarMsgErro ("BuscarUltimoDiaUtilProxMes")
End Function

