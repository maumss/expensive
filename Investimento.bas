Option Explicit

' ---------------------------------------------------------------------
' Módulo: Investimento.bas
' Descrição: Módulo de funções sobre o controle de Investimentos
' ---------------------------------------------------------------------

Private Const TITULO_INVESTIMENTOS As String = "Investimentos"

' Lembrete do Outlook para pagamento de DARF
Private Const OUTLOOK_APPOINTMENT_ITEM As Integer = 1 '1 = Appointment
Private Const OUTLOOK_OCUPADO As Integer = 2 '1 = Provisório, 2 = Ocupado, 3 = Ausente, 4 = Trabalhando em outro lugar, 5 = Disponível
Private Const MINUTOS_POR_DIA As Integer = 1440
Private Const DURACAO_LEMBRETE_MINUTOS As Integer = 30
Private Const HORARIO_LEMBRETE As String = "10:00:00"
Private Const ASSUNTO_LEMBRETE_DARF As String = "Pagar Darf"
Private Const TEXTO_LEMBRETE_DARF As String = "Pagar imposto com código 6015, Ganhos líquidos em operações em bolsa."

'
' Função saldoAtual   Data: 13/01/04
' Determina o saldo atual do investimento de acordo com
' o saldo inicial da Carteira da primeira planilha
' mensal aberta
'
Public Function CalcularSaldoAtual(ByVal rgDescInvest As Range) As Variant  
  Dim wsPlanilha As Worksheet
  On Error GoTo ErroCalcularSaldoAtual
  For Each wsPlanilha In ThisWorkbook.Worksheets
    ' A primeira planilha aberta que achar é a válida
    If IsPlanilhaAberta(wsPlanilha.Range(RANGE_SITUAC_PLANILHA)) Then
      CalcularSaldoAtual = SomarSaldoPortfolio(wsPlanilha, rgDescInvest)
      Exit Function
    End If
  Next wsPlanilha
  ' Se todas as planilhas estão fechadas, pega a de Dezembro
  CalcularSaldoAtual = SomarSaldoPortfolio(ThisWorkbook.Worksheets(NOME_PLAN_DEZ), rgDescInvest)
  Exit Function

ErroCalcularSaldoAtual:
  ' função de planilha não deve abrir mensagens: devolve erro para a célula
  CalcularSaldoAtual = CVErr(xlErrValue)
End Function

Private Function SomarSaldoPortfolio(ByVal wsPlanilha As Worksheet, ByVal rgDescInvest As Range) As Double
  SomarSaldoPortfolio = Application.WorksheetFunction.SumIf(wsPlanilha.Range(RANGE_COLUNA_ATIVO_PORTFOLIO), _
    rgDescInvest, wsPlanilha.Range(RANGE_COLUNA_SALDO_FINAL_PORTFOLIO))
End Function

'
' Função CalcularRendAtivo   Data: 29/04/16
' Retorna o valor percentual do rendimento líquido da aplicação
' Nota: todas as células envolvidas devem estar nos parâmetros de
'  entrada da função para que o Excel possa saber que deve recal-
'  cular a função caso uma delas mude.
' A outra maneira é teclar Ctrl + Alt + F9
'
Public Function CalcularRendAtivo(ByVal rgColDescAtivo As Range, _
                        ByVal rgColSaldoInicialAtivo As Range, _
                        ByVal rgColAplicacaoAtivo As Range, _
                        ByVal rgColRetornoAtivo As Range, _
                        ByVal rgColResgateAtivo As Range, _
                        ByVal rgColSaldoFinalAtivo As Range, _
                        ByVal rgCelDescRentabilidade As Range) As Variant  
  Dim dblSaldoInicial As Double, dblAplicacao As Double, dblRetorno As Double
  Dim dblResgate As Double, dblSaldoFinal As Double, dblCpmf As Double
  On Error GoTo ErroCalcularRendAtivo
  With Application.WorksheetFunction
    dblSaldoInicial = .SumIf(rgColDescAtivo, rgCelDescRentabilidade, rgColSaldoInicialAtivo)
    dblAplicacao = .SumIf(rgColDescAtivo, rgCelDescRentabilidade, rgColAplicacaoAtivo)
    dblRetorno = .SumIf(rgColDescAtivo, rgCelDescRentabilidade, rgColRetornoAtivo)
    dblResgate = .SumIf(rgColDescAtivo, rgCelDescRentabilidade, rgColResgateAtivo)
    dblSaldoFinal = .SumIf(rgColDescAtivo, rgCelDescRentabilidade, rgColSaldoFinalAtivo)
  End With

  dblCpmf = dblSaldoFinal - dblSaldoInicial - dblAplicacao + dblResgate - dblRetorno

  If (dblSaldoInicial + dblAplicacao) <> 0 Then
    CalcularRendAtivo = (((dblSaldoFinal + dblResgate - dblCpmf) / (dblSaldoInicial + dblAplicacao)) - 1) * 100
  Else
    CalcularRendAtivo = 0
  End If
  Exit Function

ErroCalcularRendAtivo:
  CalcularRendAtivo = CVErr(xlErrValue)
End Function

'
' Sub CriticarInvestimento    Criado por: MSS  Em: 31.01.04
' critica o investimento digitado
'
Public Sub CriticarInvestimento(ByVal rgAlvo As Range)  
  Dim strDescricao As String, strSugestao As String
  On Error GoTo ErroCriticarInvestimento
  If IsError(rgAlvo.Value) Then
    Exit Sub
  End If
  strDescricao = CStr(rgAlvo.Value)
  If Len(Trim$(strDescricao)) = 0 Then
    Exit Sub
  End If
  ' Se começar com Broker ou for Reserva estratégica
  If InStr(1, strDescricao, BROKER) = 1 Or strDescricao = RESERVA_ESTRATEGICA Then
    Exit Sub
  End If
  If Len(BuscarInvestimento(strDescricao, False)) > 0 Then
    Exit Sub
  End If
  strSugestao = BuscarInvestimento(strDescricao, True)
  If Len(strSugestao) > 0 Then
    If MsgBox("Você se refere a " & vbLf & strSugestao & " ?", _
              vbQuestion + vbYesNo, TITULO_INVESTIMENTOS) = vbYes Then
      rgAlvo.Value = strSugestao
      Exit Sub
    End If
  End If
  If IsReserva(strDescricao) Then
    Exit Sub
  End If
  MsgBox "Não foi encontrado um investimento com esta descrição." & vbNewLine & _
         "Procure cadastrar o mesmo em uma de suas carteiras.", vbExclamation, TITULO_INVESTIMENTOS
  Exit Sub

ErroCriticarInvestimento:
  MostrarMsgErro "CriticarInvestimento"
End Sub

'
' Procura a descrição nas carteiras AdHoc e Portfólio da planilha Alocacao.
' blnBuscaParcial = False: descrição idêntica; True: parte do nome, sem diferenciar maiúsculas.
' Retorna a descrição encontrada ou vazio.
'
Private Function BuscarInvestimento(ByVal strDescricao As String, ByVal blnBuscaParcial As Boolean) As String  
  Dim wsAlocacao As Worksheet
  Set wsAlocacao = ThisWorkbook.Worksheets(NOME_PLAN_ALOCACAO)
  BuscarInvestimento = BuscarInvestimentoNoIntervalo(wsAlocacao.Range(RANGE_CELULA_INICIO_ADHOC, RANGE_CELULA_FIM_ADHOC), _
    strDescricao, blnBuscaParcial)
  If Len(BuscarInvestimento) = 0 Then
    BuscarInvestimento = BuscarInvestimentoNoIntervalo(wsAlocacao.Range(RANGE_CELULA_INICIO_PORTFOLIO, RANGE_CELULA_FIM_PORTFOLIO), _
      strDescricao, blnBuscaParcial)
  End If
End Function

Private Function BuscarInvestimentoNoIntervalo(ByVal rgCarteira As Range, ByVal strDescricao As String, _
    ByVal blnBuscaParcial As Boolean) As String
  Dim varValores As Variant
  Dim lngLinha As Long
  Dim strValor As String
  varValores = rgCarteira.Value
  For lngLinha = 1 To UBound(varValores, 1)
    ' a carteira termina na primeira célula vazia
    If IsEmpty(varValores(lngLinha, 1)) Then
      Exit Function
    End If
    If Not IsError(varValores(lngLinha, 1)) Then
      strValor = CStr(varValores(lngLinha, 1))
      If blnBuscaParcial Then
        If InStr(1, strValor, strDescricao, vbTextCompare) > 0 Then
          BuscarInvestimentoNoIntervalo = strValor
          Exit Function
        End If
      ElseIf strValor = strDescricao Then
        BuscarInvestimentoNoIntervalo = strValor
        Exit Function
      End If
    End If
  Next lngLinha
End Function

'
' Verifica se trata de uma Reserva Estratégica
'
Private Function IsReserva(ByVal strDescricao As String) As Boolean  
  IsReserva = (InStr(1, strDescricao, PREFIXO_RESERVA) = 1)
End Function

'
' Sub AgendarLembreteOutlook    Criado por: MSS  Em: 06.10.19
' cria um lembrete no calendário do Outlook
'
Public Sub AgendarLembreteOutlook()  
  Dim objOutlook As Object
  Dim objAppointmentItem As Object
  On Error GoTo ErroAgendarLembreteOutlook
  Set objOutlook = CreateObject("Outlook.Application")
  Set objAppointmentItem = objOutlook.CreateItem(OUTLOOK_APPOINTMENT_ITEM)
  With objAppointmentItem
    .Subject = ASSUNTO_LEMBRETE_DARF
    .Body = TEXTO_LEMBRETE_DARF
    .Location = ""
    .Start = BuscarUltimoDiaUtilProxMes() + TimeValue(HORARIO_LEMBRETE)
    .Duration = DURACAO_LEMBRETE_MINUTOS
    .BusyStatus = OUTLOOK_OCUPADO
    .ReminderSet = True
    .ReminderMinutesBeforeStart = MINUTOS_POR_DIA
    .Save
  End With
  Debug.Print objAppointmentItem.Subject & " : " & objAppointmentItem.Start
  Exit Sub

ErroAgendarLembreteOutlook:
  MostrarMsgErro "AgendarLembreteOutlook"
End Sub

'
' Último dia útil (seg a sex) do próximo mês
'
Private Function BuscarUltimoDiaUtilProxMes() As Date  
  Dim datData As Date
  ' dia 0 do mês seguinte ao próximo = último dia do próximo mês
  datData = DateSerial(Year(Date), Month(Date) + 2, 0)
  Select Case Weekday(datData, vbSunday)
    Case vbSunday
      datData = datData - 2
    Case vbSaturday
      datData = datData - 1
  End Select
  BuscarUltimoDiaUtilProxMes = datData
End Function
