Attribute VB_Name = "Investimento"
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
       CalcularSaldoAtual = Application.WorksheetFunction.SumIf(wsPlanilha.Range(RANGE_COLUNA_ATIVO_ADHOC), _
             rgDescInvest, wsPlanilha.Range(RANGE_COLUNA_SALDO_FINAL_ADHOC)) + _
          Application.WorksheetFunction.SumIf(wsPlanilha.Range(RANGE_COLUNA_ATIVO_CONSOLIDADA), _
             rgDescInvest, wsPlanilha.Range(RANGE_COLUNA_SALDO_FINAL_CONSOLIDADA))
       Exit Function
     End If
  Next wsPlanilha
  ' Se todas as planilhas estão fechadas, pega a de Dezembro
  CalcularSaldoAtual = Application.WorksheetFunction.SumIf(Worksheets("Dez.").Range(RANGE_COLUNA_ATIVO_ADHOC), _
     rgDescInvest, Worksheets("Dez.").Range(RANGE_COLUNA_SALDO_FINAL_ADHOC)) + _
     Application.WorksheetFunction.SumIf(Worksheets("Dez.").Range(RANGE_COLUNA_ATIVO_ADHOC), _
       rgDescInvest, Worksheets("Dez.").Range(RANGE_COLUNA_SALDO_FINAL_CONSOLIDADA))
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
  If (rgAlvo.Value = "Broker" Or rgAlvo.Value = "Reserva estratégica") Then
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
  intInicioLinhaCarteira2 = Range(RANGE_CELULA_INICIO_CONSOLIDADA).Row
  intFinalLinhaCarteira2 = Range(RANGE_CELULA_FIM_CONSOLIDADA).Row
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
  intInicioLinhaCarteira2 = Range(RANGE_CELULA_INICIO_CONSOLIDADA).Row
  intFinalLinhaCarteira2 = Range(RANGE_CELULA_FIM_CONSOLIDADA).Row
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
  
