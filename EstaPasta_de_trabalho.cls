VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EstaPasta_de_trabalho"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

' Variáveis locais
Dim blnAtualizouArquivo As Boolean

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, cancel As Boolean)
  '
  ' sub Workbook_beforeSave
  ' atualiza a data da última alteração na Planilha Intro
  ' criado por: Mauricio SS  em: 23.01.04
  '
  On Error GoTo final
  If blnAtualizouArquivo Then
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Worksheets("Intro").Unprotect
    Worksheets("Intro").Range(RANGE_DATA_ULTIMA_ATUALIZ).Value = _
      Format(Now, "dddd, dd mmmm yyyy") & " às " & _
      Format(Now, "hh:mm:ss")
    Range(RANGE_POSICAO).Value = Now
    Worksheets("Intro").Protect
    blnAtualizouArquivo = False
  End If
  
final:
  Application.ScreenUpdating = True
  Application.EnableEvents = True
  Exit Sub
  
erroBeforeSave:
  MostrarMsgErro ("Workbook_BeforeSave")
  Resume final
End Sub

Private Sub Workbook_Open()
  On Error GoTo erroWorkbook_Open
  blnAtualizouArquivo = False
  ActiveWorkbook.Sheets("Intro").Activate
  Call PosicionarTopo
  Exit Sub
  
erroWorkbook_Open:
  MostrarMsgErro ("Workbook_Open")
End Sub

Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
  '
  ' Sub Workbook_SheetChange
  ' verifica se atualizou alguma planilha ou se está digitando a descrição
  ' de um investimento nas planilhas mensais
  '
  ' variaveis
  Dim blnIsCriticouDescInvest As Boolean
  ' principal
  On Error GoTo erroSheetChange
  blnIsCriticouDescInvest = False
  ' ajusta variável para avisar que houve mudanças
  blnAtualizouArquivo = True
  If Not IsPlanilhaAberta(Sh.Range(RANGE_SITUAC_PLANILHA)) Then
    Exit Sub
  End If
  
  ' testa se é uma planilha mensal
  Application.EnableEvents = False
  
  ' testa se alterou os valores de movimentações
  If Not Application.Intersect(Target, Range(RANGE_COL_VALOR_MOVIMENTACOES)) Is Nothing Then
    ' testa saldo da planilha atual
    If Target.count = 1 And Not IsEmpty(Target) Then
      If IsAlcancouValorCriticoNoMes(Worksheets(Sh.Name).Index, Target.Value) Then
        Beep
        Beep
        MsgBox "Cuidado! Valor de tolerância alcançado para este mês.", vbExclamation
      End If
    End If
  End If
  ' testa se alterou os valores de cartão
  If Not Application.Intersect(Target, Range(RANGE_COL_VALOR_CARTOES)) Is Nothing Then
    ' testa saldo da próxima planilha
    If Target.count = 1 And Not IsEmpty(Target) Then
      If IsAlcancouValorCriticoNoMes(Worksheets(Sh.Name).Index + 1, Target.Value) Then
        Beep
        Beep
        MsgBox "Cuidado! Valor de tolerância alcançado para o próximo mês.", vbExclamation
      End If
    End If
  End If
  ' critica a descrição dos investimentos digitados
  If Cells(Target.Row, Range(RANGE_HEADER_TIPO_MOVIMENTACAO).Column).Value = TIPO_LANCAMENTO_INVESTIMENTOS Then
    ' testa se alterou os valores dos créditos
    If Not Application.Intersect(Target, Range(RANGE_COL_VALOR_MOVIMENTACOES)) Is Nothing And Not IsEmpty(Target) Then
      ' critica a descrição do investimento
      blnIsCriticouDescInvest = True
      Application.ScreenUpdating = False
      Call CriticarInvestimento(Cells(Target.Row, Range(RANGE_HEADER_DESC_MOVIMENTACAO).Column))
    End If
  End If
  ' testa se alterou os valores finais de ações, opções ou fii's
  If Not Application.Intersect(Target, Range(RANGE_COLUNA_SALDO_FINAL_ACOES)) Is Nothing Or _
     Not Application.Intersect(Target, Range(RANGE_COLUNA_SALDO_FINAL_CART_OPCOES)) Is Nothing Or _
     Not Application.Intersect(Target, Range(RANGE_COLUNA_SALDO_FINAL_FII)) Is Nothing Then
    ' testa se gerou imposto a pagar
    If Target.count = 1 And Not IsEmpty(Target) Then
      If IsGerouImpostoNoMes(Target) Then
        Beep
        Beep
        MsgBox "Atenção! Você tem imposto a pagar este mês.", vbExclamation
      End If
    End If
  End If
  
fimSheetChange:
  ' só deve deixar a tela atualizar se passou pela rotina acima
  If blnIsCriticouDescInvest Then
    Application.ScreenUpdating = True
  End If
  Application.EnableEvents = True
  Exit Sub
  
erroSheetChange:
  MostrarMsgErro ("Workbook_SheetChange")
  Resume fimSheetChange
End Sub


Private Function IsAlcancouValorCriticoNoMes(intPlanilha As Integer, valor As Double) As Boolean
  '
  ' Function IsAlcancouValorCriticoNoMes
  ' verifica se alcançou o ponto crítico neste mês
  '
  ' variaveis
  Dim dblValorTolerancia, dblSaldoAtual As Double
  ' principal
  On Error GoTo erroIsAlcancouValorCriticoNoMes
  'Testa se a próxima planilha está aberta e é mensal
  'este teste é necessário caso esteja se vendo a planilha do próximo mês
  If Not IsPlanilhaAberta(Worksheets(intPlanilha).Range(RANGE_SITUAC_PLANILHA)) Then
    IsAlcancouValorCriticoNoMes = False
    Exit Function
  End If
  dblValorTolerancia = Range(RANGE_TOLERANCIA).Value
  dblSaldoAtual = Sheets(intPlanilha).Range(RANGE_SALDO_MES).Value
  If (dblSaldoAtual < dblValorTolerancia) And _
     ((dblSaldoAtual + valor) >= dblValorTolerancia) Then
    IsAlcancouValorCriticoNoMes = True
  Else
    IsAlcancouValorCriticoNoMes = False
  End If
  Exit Function
  
erroIsAlcancouValorCriticoNoMes:
  MostrarMsgErro ("IsAlcancouValorCriticoNoMes")
End Function

Private Function IsGerouImpostoNoMes(ByVal Target As Range) As Boolean
  '
  ' Function IsGerouImpostoNoMes
  ' verifica se gerou imposto a pagar
  '
  ' variaveis
  Dim dblValorTolerancia, dblSaldoAtual As Double
  Dim boolTributavel, boolImpostoComum, boolImpostoDaytrade As Boolean
  ' principal
  On Error GoTo erroIsGerouImpostoNoMes
  
  If Not Application.Intersect(Target, Range(RANGE_COLUNA_SALDO_FINAL_ACOES)) Is Nothing Then
    boolTributavel = Range(RANGE_CELULA_TRIBUTAL_ACOES).Value
    boolImpostoComum = (Cells(Target.Row, Range(RANGE_COLUNA_RESULTADO_COMUM_ACOES).Column).Value > 0)
    boolImpostoDaytrade = (Cells(Target.Row, Range(RANGE_COLUNA_RESULTADO_DAYTRADE_ACOES).Column).Value > 0)
    If boolTributavel And (boolImpostoComum Or boolImpostoDaytrade) Then
      IsGerouImpostoNoMes = True
      Exit Function
    End If
  End If
  
  If Not Application.Intersect(Target, Range(RANGE_COLUNA_SALDO_FINAL_CART_OPCOES)) Is Nothing Then
    boolImpostoComum = (Cells(Target.Row, Range(RANGE_COLUNA_RESULTADO_COMUM_OPCOES).Column).Value > 0)
    boolImpostoDaytrade = (Cells(Target.Row, Range(RANGE_COLUNA_RESULTADO_DAYTRADE_OPCOES).Column).Value > 0)
    If boolImpostoComum Or boolImpostoDaytrade Then
      IsGerouImpostoNoMes = True
      Exit Function
    End If
  End If
  
  If Not Application.Intersect(Target, Range(RANGE_COLUNA_SALDO_FINAL_FII)) Is Nothing Then
    boolImpostoComum = (Cells(Target.Row, Range(RANGE_COLUNA_RESULTADO_COMUM_FII).Column).Value > 0)
    boolImpostoDaytrade = (Cells(Target.Row, Range(RANGE_COLUNA_RESULTADO_DAYTRADE_FII).Column).Value > 0)
    If boolImpostoComum Or boolImpostoDaytrade Then
      IsGerouImpostoNoMes = True
      Exit Function
    End If
  End If
  
  IsGerouImpostoNoMes = False
  Exit Function
  
erroIsGerouImpostoNoMes:
  MostrarMsgErro ("IsGerouImpostoNoMes")
End Function

Private Sub Workbook_BeforeClose(cancel As Boolean)
  '
  ' Sub Workbook_BeforeClose
  ' posiciona todas as planilhas no topo
  '
  Dim wsPlanilha As Worksheet, wsPlanilhaAtual As Worksheet
  On Error Resume Next
  Application.ScreenUpdating = False
  Set wsPlanilhaAtual = ActiveSheet
  For Each wsPlanilha In ActiveWorkbook.Worksheets
     If wsPlanilha.Visible And Not LCase(wsPlanilha.Name) Like "graf*" Then
       wsPlanilha.Activate
       Call PosicionarTopo
     End If
  Next wsPlanilha
  wsPlanilhaAtual.Activate
  Application.ScreenUpdating = True
End Sub

