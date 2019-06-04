Option Explicit
'Módulo de fechamento do mês

Private Type infoInvest
  strAtivo As String
  intQuantidade As Double
  dblCustoAnterior As Double
  dblSaldo As Double
End Type

Sub ProtegerPlanilha()
  '
  ' Protege Macro
  ' Protege ou desprotege a planilha do mês.
  '
  ' Atalho do teclado: Ctrl+p
  '
  On Error GoTo erroProtege
  ' Verifica se é uma planilha de movimentação
  If (Range(RANGE_SITUAC_PLANILHA).Value <> SITUAC_ABERTO) And _
     (Range(RANGE_SITUAC_PLANILHA).Value <> SITUAC_FECHADO) Then
    Exit Sub
  End If
  CongelarCalculosPlanilha (True)
  If ActiveSheet.ProtectContents Then
    'MsgBox "Essa planilha será desprotegida...", vbInformation
    ActiveSheet.Unprotect
    AlterarSituacaoPlanilha (SITUAC_ABERTO)
  Else
    'MsgBox "Essa planilha será protegida...", vbInformation
    AlterarSituacaoPlanilha (SITUAC_FECHADO)
    CopiarSaldos
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
  End If
  
protege_Fim:
  PosicionarTopo
  CongelarCalculosPlanilha (False)
  Exit Sub
  
erroProtege:
  MostrarMsgErro ("ProtegerPlanilha")
  Resume protege_Fim
End Sub

Private Sub AlterarSituacaoPlanilha(strSituacao As String)
  '
  ' Muda o texto do status da planilha
  '
  On Error GoTo erroStatus
  Range(RANGE_SITUAC_PLANILHA).Select
  If strSituacao = SITUAC_FECHADO Then
    ActiveCell.FormulaR1C1 = SITUAC_FECHADO
    With ActiveCell.Characters(Start:=1, Length:=7).Font
        .Name = "Arial"
        .FontStyle = "Normal"
        .Size = 12
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = 3 'vermelho
    End With
  Else
    ActiveCell.FormulaR1C1 = SITUAC_ABERTO
    With ActiveCell.Characters(Start:=1, Length:=6).Font
       .Name = "Arial"
       .FontStyle = "Normal"
       .Size = 12
       .Strikethrough = False
       .Superscript = False
       .Subscript = False
       .OutlineFont = False
       .Shadow = False
       .Underline = xlUnderlineStyleNone
       .ColorIndex = 50 'verde
    End With
  End If
  RetornarUltimaCelulaMovimentacoes.Select
  Exit Sub
erroStatus:
  MostrarMsgErro ("AlterarSituacaoPlanilha")
End Sub

Private Sub CopiarSaldos()
  '
  ' Sub copiarSaldos        De: 17/01/04
  ' Copia os saldos dos investimentos de um mês p/ o próximo
  '
  On Error GoTo errocopiarSaldos
    
  Dim lngIndPlan As Long
  lngIndPlan = Worksheets(ActiveSheet.Name).Index
  'Testa se a planilha destino está aberta
  Dim wsProxPlanilha As Worksheet
  Set wsProxPlanilha = Worksheets(lngIndPlan + 1)
  If Not IsPlanilhaAberta(wsProxPlanilha.Range(RANGE_SITUAC_PLANILHA)) Then
    Exit Sub
  End If
  ' Verifica se existe saldo final preenchido no próximo mês do resumo de investimentos
  If HasSaldosCarteira(wsProxPlanilha) Then
    Exit Sub
  End If
  ' Verifica se existe dados a transferir
  Dim wsPlanilhaAtual As Worksheet
  Set wsPlanilhaAtual = Worksheets(lngIndPlan)
  If Not (HasDadosCarteira(wsPlanilhaAtual)) Then
    Exit Sub
  End If
  'Pede confirmação
  If MsgBox("Você deseja montar o Resumo Mensal de Carteiras do próximo mês?", _
        vbYesNo + vbQuestion, "Copiar Resumo Mensal") = vbNo Then
    Exit Sub
  End If
  
  ' Reproduz o Resumo de Investimento Atual no próximo mês
  Call CopiarSaldosCarteiraAdHoc(wsPlanilhaAtual, wsProxPlanilha)
  Call CopiarSaldosCarteiraConsolidada(wsPlanilhaAtual, wsProxPlanilha)
  Call CopiarSaldosCarteiraAcoes(wsPlanilhaAtual, wsProxPlanilha)
  Call CopiarSaldosCarteiraFii(wsPlanilhaAtual, wsProxPlanilha)
  Call CopiarSaldosCarteiraTesouroRF(wsPlanilhaAtual, wsProxPlanilha)
  Call CopiarSaldosCarteiraTesouroSelic(wsPlanilhaAtual, wsProxPlanilha)
  Call CopiarSaldosCarteiraEtf(wsPlanilhaAtual, wsProxPlanilha)
  Call CopiarSaldosCarteiraExterior(wsPlanilhaAtual, wsProxPlanilha)
  Call CopiarSaldosCarteiraOpcoes(wsPlanilhaAtual, wsProxPlanilha)
  Call CopiarSaldosContaCorretora(wsPlanilhaAtual, wsProxPlanilha)
  Exit Sub
  
errocopiarSaldos:
  MostrarMsgErro ("CopiarSaldos")
End Sub

Private Function HasSaldosCarteira(wsProxPlanilha As Worksheet) As Boolean
  '
  ' Function HasSaldosCarteira
  ' verifica se existem saldos no próximo mês
  '
  Dim blnExisteDados As Boolean
  blnExisteDados = False
  Dim rgCell As Range
  On Error GoTo ErrorHasSaldosCarteira
  For Each rgCell In wsProxPlanilha.Range(RANGE_COLUNA_SALDO_FINAL_ADHOC)
    If rgCell.Value > 0 Then
      blnExisteDados = True
      Exit For
    End If
  Next rgCell
  If Not blnExisteDados Then
    For Each rgCell In wsProxPlanilha.Range(RANGE_COLUNA_SALDO_FINAL_CONSOLIDADA)
      If rgCell.Value > 0 Then
        blnExisteDados = True
        Exit For
      End If
    Next rgCell
  End If
  HasSaldosCarteira = blnExisteDados
  Exit Function
  
ErrorHasSaldosCarteira:
  MostrarMsgErro ("HasSaldosCarteira")
End Function

Private Function HasDadosCarteira(wsPlanilhaAtual As Worksheet) As Boolean
  '
  ' Function HasDadosCarteira
  ' verifica se existem dados a transferir
  '
  Dim blnExisteDados As Boolean
  blnExisteDados = False
  Dim rgCell As Range
  On Error GoTo ErrorHasDadosCarteira
  For Each rgCell In wsPlanilhaAtual.Range(RANGE_COLUNA_ATIVO_ADHOC)
    If Not IsEmpty(rgCell) Then
      blnExisteDados = True
      Exit For
    End If
  Next rgCell
  If Not blnExisteDados Then
    For Each rgCell In wsPlanilhaAtual.Range(RANGE_COLUNA_ATIVO_CONSOLIDADA)
      If Not IsEmpty(rgCell) Then
        blnExisteDados = True
        Exit For
      End If
    Next rgCell
  End If
  If Not blnExisteDados Then
    For Each rgCell In wsPlanilhaAtual.Range(RANGE_COLUNA_ATIVO_ACOES)
      If Not IsEmpty(rgCell) Then
        blnExisteDados = True
        Exit For
      End If
    Next rgCell
  End If
  If Not blnExisteDados Then
    For Each rgCell In wsPlanilhaAtual.Range(RANGE_COLUNA_ATIVO_FII)
      If Not IsEmpty(rgCell) Then
        blnExisteDados = True
        Exit For
      End If
    Next rgCell
  End If
  If Not blnExisteDados Then
    For Each rgCell In wsPlanilhaAtual.Range(RANGE_COLUNA_ATIVO_RF)
      If Not IsEmpty(rgCell) Then
        blnExisteDados = True
        Exit For
      End If
    Next rgCell
  End If
  If Not blnExisteDados Then
    For Each rgCell In wsPlanilhaAtual.Range(RANGE_COLUNA_ATIVO_SELIC)
      If Not IsEmpty(rgCell) Then
        blnExisteDados = True
        Exit For
      End If
    Next rgCell
  End If
  HasDadosCarteira = blnExisteDados
  Exit Function
  
ErrorHasDadosCarteira:
  MostrarMsgErro ("HasDadosCarteira")
End Function

Private Sub CopiarSaldosCarteiraAdHoc(wsPlanilhaAtual As Worksheet, wsProxPlanilha As Worksheet)
  '
  ' Sub CopiarSaldosCarteiraAdHoc
  ' copiar saldos e descrições da carteira 1
  '
  Dim intPrimeiraLinhaResumo As Integer, intUltimaLinhaResumo As Integer
  Dim intColunaDescricao As Integer, intColunaSaldoInicial As Integer, intColunaSaldoFinal As Integer
  On Error GoTo ErrorCopiarSaldosCarteiraAdHoc
  intPrimeiraLinhaResumo = RetornarPrimeiraLinha(Range(RANGE_COLUNA_ATIVO_ADHOC))
  intUltimaLinhaResumo = RetornarUltimaLinha(Range(RANGE_COLUNA_ATIVO_ADHOC))
  intColunaDescricao = RetornarPrimeiraColuna(Range(RANGE_COLUNA_ATIVO_ADHOC))
  intColunaSaldoInicial = RetornarPrimeiraColuna(Range(RANGE_COLUNA_SALDO_INICIAL_ADHOC))
  intColunaSaldoFinal = RetornarPrimeiraColuna(Range(RANGE_COLUNA_SALDO_FINAL_ADHOC))
  Call CopiarSaldosCarteira(intPrimeiraLinhaResumo, intUltimaLinhaResumo, _
    intColunaDescricao, intColunaSaldoInicial, intColunaSaldoFinal, _
    wsPlanilhaAtual, wsProxPlanilha)
  Exit Sub
    
ErrorCopiarSaldosCarteiraAdHoc:
  MostrarMsgErro ("CopiarSaldosCarteiraAdHoc")
End Sub

Private Sub CopiarSaldosCarteiraConsolidada(wsPlanilhaAtual As Worksheet, wsProxPlanilha As Worksheet)
  '
  ' Sub CopiarSaldosCarteiraConsolidada
  ' copiar saldos e descrições da carteira 2
  '
  Dim intPrimeiraLinhaResumo As Integer, intUltimaLinhaResumo As Integer
  Dim intColunaDescricao As Integer, intColunaSaldoInicial As Integer, intColunaSaldoFinal As Integer
  On Error GoTo ErrorCopiarSaldosCarteiraConsolidada
  intPrimeiraLinhaResumo = RetornarPrimeiraLinha(Range(RANGE_COLUNA_ATIVO_CONSOLIDADA))
  intUltimaLinhaResumo = RetornarUltimaLinha(Range(RANGE_COLUNA_ATIVO_CONSOLIDADA))
  intColunaDescricao = RetornarPrimeiraColuna(Range(RANGE_COLUNA_ATIVO_CONSOLIDADA))
  intColunaSaldoInicial = RetornarPrimeiraColuna(Range(RANGE_COLUNA_SALDO_INICIAL_CONSOLIDADA))
  intColunaSaldoFinal = RetornarPrimeiraColuna(Range(RANGE_COLUNA_SALDO_FINAL_CONSOLIDADA))
  Call CopiarSaldosCarteira(intPrimeiraLinhaResumo, intUltimaLinhaResumo, _
    intColunaDescricao, intColunaSaldoInicial, intColunaSaldoFinal, _
    wsPlanilhaAtual, wsProxPlanilha)
  Exit Sub
    
ErrorCopiarSaldosCarteiraConsolidada:
  MostrarMsgErro ("CopiarSaldosCarteiraConsolidada")
End Sub

Private Sub CopiarSaldosCarteira(intPrimeiraLinhaResumo As Integer, intUltimaLinhaResumo As Integer, _
   intColunaDescricao As Integer, intColunaSaldoInicial As Integer, intColunaSaldoFinal As Integer, _
   wsPlanilhaAtual As Worksheet, wsProxPlanilha As Worksheet)
  '
  ' Sub CopiarSaldosCarteira
  ' copiar saldos e descrições da carteira
  '
  ' Retirar colunas intColunaQtdeInicial e intColunaQtdeFinal
  On Error GoTo ErrorCopiarSaldosCarteira
  
  Dim intCont As Integer, intLinhaDestino As Integer
  Dim strFormula As String
  For intCont = intPrimeiraLinhaResumo To intUltimaLinhaResumo
     If (Not IsEmpty(wsPlanilhaAtual.Cells(intCont, intColunaDescricao))) And _
        (wsPlanilhaAtual.Cells(intCont, intColunaSaldoFinal).Value <> 0) Then
       'se existir tipo de investimento na planilha atual
       intLinhaDestino = intCont
       With wsProxPlanilha
          If (intLinhaDestino > intPrimeiraLinhaResumo) And (IsEmpty(.Cells(intLinhaDestino - 1, intColunaDescricao))) Then
            intLinhaDestino = GetPrimeiraLinhaLivreDaCarteira(intPrimeiraLinhaResumo, intUltimaLinhaResumo, intColunaDescricao, wsProxPlanilha)
          End If
          .Cells(intLinhaDestino, intColunaDescricao).Value = wsPlanilhaAtual.Cells(intCont, intColunaDescricao).Value
          If wsPlanilhaAtual.Cells(intCont, intColunaSaldoFinal).HasFormula And _
             wsPlanilhaAtual.Cells(intCont, intColunaSaldoInicial).HasFormula And _
             InStr(wsPlanilhaAtual.Cells(intCont, intColunaSaldoInicial).Formula, "D") > 0 Then
            If (intLinhaDestino = intCont) Then
              .Cells(intLinhaDestino, intColunaSaldoInicial).Formula = wsPlanilhaAtual.Cells(intCont, intColunaSaldoInicial).Formula
            Else
              strFormula = wsPlanilhaAtual.Cells(intCont, intColunaSaldoInicial).Formula
              strFormula = Replace(strFormula, "D" & intCont, "D" & intLinhaDestino)
              .Cells(intLinhaDestino, intColunaSaldoInicial).Formula = strFormula
            End If
          Else
            .Cells(intLinhaDestino, intColunaSaldoInicial).Value = wsPlanilhaAtual.Cells(intCont, intColunaSaldoFinal).Value
          End If
          If wsPlanilhaAtual.Cells(intCont, intColunaSaldoFinal).HasFormula Then
            If (intLinhaDestino = intCont) Then
              .Cells(intLinhaDestino, intColunaSaldoFinal).Formula = wsPlanilhaAtual.Cells(intCont, intColunaSaldoFinal).Formula
            Else
              strFormula = wsPlanilhaAtual.Cells(intCont, intColunaSaldoFinal).Formula
              strFormula = Replace(strFormula, "D" & intCont, "D" & intLinhaDestino)
              .Cells(intLinhaDestino, intColunaSaldoFinal).Formula = strFormula
            End If
          Else
            .Cells(intLinhaDestino, intColunaSaldoFinal).Value = wsPlanilhaAtual.Cells(intCont, intColunaSaldoFinal).Value
          End If
       End With
     End If
  Next intCont
  Exit Sub
  
ErrorCopiarSaldosCarteira:
  MostrarMsgErro ("CopiarSaldosCarteira")
End Sub

Private Sub CopiarSaldosCarteiraAcoes(wsPlanilhaAtual As Worksheet, wsProxPlanilha As Worksheet)
  '
  ' Sub CopiarSaldosCarteiraAcoes
  ' copiar saldos, quantidades e descrições da carteira 3
  '
  Dim intPrimeiraLinha As Integer, intUltimaLinha As Integer
  Dim intColunaAtivo As Integer, intColunaQtde As Integer, intColunaCustoAnterior As Integer
  Dim intColunaSaldoInicial As Integer, intColunaSaldoFinal As Integer
  Dim intColunaOperacao As Integer, intColunaCustoMedio As Integer
  Dim infoInvests() As infoInvest
  On Error GoTo ErrorCopiarSaldosCarteiraAcoes
  intPrimeiraLinha = RetornarPrimeiraLinha(Range(RANGE_COLUNA_ATIVO_ACOES))
  intUltimaLinha = RetornarUltimaLinha(Range(RANGE_COLUNA_ATIVO_ACOES))
  intColunaAtivo = RetornarPrimeiraColuna(Range(RANGE_COLUNA_ATIVO_ACOES))
  intColunaQtde = RetornarPrimeiraColuna(Range(RANGE_COLUNA_QTDE_ACOES))
  intColunaSaldoInicial = RetornarPrimeiraColuna(Range(RANGE_COLUNA_SALDO_INICIAL_ACOES))
  intColunaSaldoFinal = RetornarPrimeiraColuna(Range(RANGE_COLUNA_SALDO_FINAL_ACOES))
  intColunaCustoMedio = RetornarPrimeiraColuna(Range(RANGE_COLUNA_CUSTO_MEDIO_ACOES))
  
  Call ColetarInformacoes(intPrimeiraLinha, intUltimaLinha, _
   intColunaAtivo, intColunaSaldoFinal, _
   intColunaQtde, _
   intColunaCustoMedio, _
   wsPlanilhaAtual, infoInvests)
  If (IsArrayEmpty(infoInvests) = True) Then
    Exit Sub
  End If
  Call CopiarRendaVariavel(intPrimeiraLinha, intUltimaLinha, _
    intColunaAtivo, intColunaSaldoInicial, intColunaSaldoFinal, _
    intColunaQtde, _
    wsPlanilhaAtual, wsProxPlanilha, infoInvests)
  Exit Sub
    
ErrorCopiarSaldosCarteiraAcoes:
  MostrarMsgErro ("CopiarSaldosCarteiraAcoes")
End Sub

Private Sub CopiarSaldosCarteiraEtf(wsPlanilhaAtual As Worksheet, wsProxPlanilha As Worksheet)
  '
  ' Sub CopiarSaldosCarteiraEtf
  ' copiar saldos, quantidades e descrições da carteira etf
  '
  Dim intPrimeiraLinha As Integer, intUltimaLinha As Integer
  Dim intColunaAtivo As Integer, intColunaQtde As Integer, intColunaCustoAnterior As Integer
  Dim intColunaSaldoInicial As Integer, intColunaSaldoFinal As Integer
  Dim intColunaOperacao As Integer, intColunaCustoMedio As Integer
  Dim infoInvests() As infoInvest
  On Error GoTo ErrorCopiarSaldosCarteiraEtf
  intPrimeiraLinha = RetornarPrimeiraLinha(Range(RANGE_COLUNA_ATIVO_ETF))
  intUltimaLinha = RetornarUltimaLinha(Range(RANGE_COLUNA_ATIVO_ETF))
  intColunaAtivo = RetornarPrimeiraColuna(Range(RANGE_COLUNA_ATIVO_ETF))
  intColunaQtde = RetornarPrimeiraColuna(Range(RANGE_COLUNA_QTDE_ETF))
  intColunaSaldoInicial = RetornarPrimeiraColuna(Range(RANGE_COLUNA_SALDO_INICIAL_ETF))
  intColunaSaldoFinal = RetornarPrimeiraColuna(Range(RANGE_COLUNA_SALDO_FINAL_ETF))
  intColunaCustoMedio = RetornarPrimeiraColuna(Range(RANGE_COLUNA_CUSTO_MEDIO_ETF))
  
  Call ColetarInformacoes(intPrimeiraLinha, intUltimaLinha, _
   intColunaAtivo, intColunaSaldoFinal, _
   intColunaQtde, _
   intColunaCustoMedio, _
   wsPlanilhaAtual, infoInvests)
  If (IsArrayEmpty(infoInvests) = True) Then
    Exit Sub
  End If
  Call CopiarRendaVariavel(intPrimeiraLinha, intUltimaLinha, _
    intColunaAtivo, intColunaSaldoInicial, intColunaSaldoFinal, _
    intColunaQtde, _
    wsPlanilhaAtual, wsProxPlanilha, infoInvests)
  Exit Sub
    
ErrorCopiarSaldosCarteiraEtf:
  MostrarMsgErro ("CopiarSaldosCarteiraEtf")
End Sub

Private Sub CopiarSaldosCarteiraExterior(wsPlanilhaAtual As Worksheet, wsProxPlanilha As Worksheet)
  '
  ' Sub CopiarSaldosCarteiraExterior
  ' copiar saldos, quantidades e descrições da carteira etf
  '
  Dim intPrimeiraLinha As Integer, intUltimaLinha As Integer
  Dim intColunaAtivo As Integer, intColunaQtde As Integer, intColunaCustoAnterior As Integer
  Dim intColunaSaldoInicial As Integer, intColunaSaldoFinal As Integer
  Dim intColunaOperacao As Integer, intColunaCustoMedio As Integer
  Dim infoInvests() As infoInvest
  On Error GoTo ErrorCopiarSaldosCarteiraExterior
  intPrimeiraLinha = RetornarPrimeiraLinha(Range(RANGE_COLUNA_ATIVO_EXTERIOR))
  intUltimaLinha = RetornarUltimaLinha(Range(RANGE_COLUNA_ATIVO_EXTERIOR))
  intColunaAtivo = RetornarPrimeiraColuna(Range(RANGE_COLUNA_ATIVO_EXTERIOR))
  intColunaQtde = RetornarPrimeiraColuna(Range(RANGE_COLUNA_QTDE_EXTERIOR))
  intColunaSaldoInicial = RetornarPrimeiraColuna(Range(RANGE_COLUNA_SALDO_INICIAL_EXTERIOR))
  intColunaSaldoFinal = RetornarPrimeiraColuna(Range(RANGE_COLUNA_SALDO_FINAL_EXTERIOR))
  intColunaCustoMedio = RetornarPrimeiraColuna(Range(RANGE_COLUNA_CUSTO_MEDIO_EXTERIOR))
  
  Call ColetarInformacoes(intPrimeiraLinha, intUltimaLinha, _
   intColunaAtivo, intColunaSaldoFinal, _
   intColunaQtde, _
   intColunaCustoMedio, _
   wsPlanilhaAtual, infoInvests)
  If (IsArrayEmpty(infoInvests) = True) Then
    Exit Sub
  End If
  Call CopiarRendaVariavel(intPrimeiraLinha, intUltimaLinha, _
    intColunaAtivo, intColunaSaldoInicial, intColunaSaldoFinal, _
    intColunaQtde, _
    wsPlanilhaAtual, wsProxPlanilha, infoInvests)
  Exit Sub
    
ErrorCopiarSaldosCarteiraExterior:
  MostrarMsgErro ("CopiarSaldosCarteiraExterior")
End Sub

Private Sub CopiarSaldosCarteiraOpcoes(wsPlanilhaAtual As Worksheet, wsProxPlanilha As Worksheet)
  '
  ' Sub CopiarSaldosCarteiraOpcoes
  ' copiar saldos, quantidades e descrições da carteira opcoes
  '
  Dim intPrimeiraLinha As Integer, intUltimaLinha As Integer
  Dim intColunaAtivo As Integer, intColunaQtde As Integer, intColunaCustoAnterior As Integer
  Dim intColunaSaldoInicial As Integer, intColunaSaldoFinal As Integer
  Dim intColunaOperacao As Integer, intColunaCustoMedio As Integer
  Dim infoInvests() As infoInvest
  On Error GoTo ErrorCopiarSaldosCarteiraOpcoes
  intPrimeiraLinha = RetornarPrimeiraLinha(Range(RANGE_COLUNA_ATIVO_CART_OPCOES))
  intUltimaLinha = RetornarUltimaLinha(Range(RANGE_COLUNA_ATIVO_CART_OPCOES))
  intColunaAtivo = RetornarPrimeiraColuna(Range(RANGE_COLUNA_ATIVO_CART_OPCOES))
  intColunaQtde = RetornarPrimeiraColuna(Range(RANGE_COLUNA_QTDE_CART_OPCOES))
  intColunaSaldoInicial = RetornarPrimeiraColuna(Range(RANGE_COLUNA_SALDO_INICIAL_CART_OPCOES))
  intColunaSaldoFinal = RetornarPrimeiraColuna(Range(RANGE_COLUNA_SALDO_FINAL_CART_OPCOES))
  intColunaCustoMedio = RetornarPrimeiraColuna(Range(RANGE_COLUNA_CUSTO_MEDIO_CART_OPCOES))
  
  Call ColetarInformacoes(intPrimeiraLinha, intUltimaLinha, _
   intColunaAtivo, intColunaSaldoFinal, _
   intColunaQtde, _
   intColunaCustoMedio, _
   wsPlanilhaAtual, infoInvests)
  If (IsArrayEmpty(infoInvests) = True) Then
    Exit Sub
  End If
  Call CopiarRendaVariavel(intPrimeiraLinha, intUltimaLinha, _
    intColunaAtivo, intColunaSaldoInicial, intColunaSaldoFinal, _
    intColunaQtde, _
    wsPlanilhaAtual, wsProxPlanilha, infoInvests)
  Exit Sub
    
ErrorCopiarSaldosCarteiraOpcoes:
  MostrarMsgErro ("CopiarSaldosCarteiraOpcoes")
End Sub

Private Sub CopiarSaldosCarteiraFii(wsPlanilhaAtual As Worksheet, wsProxPlanilha As Worksheet)
  '
  ' Sub CopiarSaldosCarteiraFii
  ' copiar saldos, quantidades e descrições da carteira 4
  '
  Dim intPrimeiraLinha As Integer, intUltimaLinha As Integer
  Dim intColunaAtivo As Integer, intColunaQtde As Integer, intColunaCustoAnterior As Integer
  Dim intColunaSaldoInicial As Integer, intColunaSaldoFinal As Integer
  Dim intColunaOperacao As Integer, intColunaCustoMedio As Integer
  Dim infoInvests() As infoInvest
  On Error GoTo ErrorCopiarSaldosCarteiraFii
  intPrimeiraLinha = RetornarPrimeiraLinha(Range(RANGE_COLUNA_ATIVO_FII))
  intUltimaLinha = RetornarUltimaLinha(Range(RANGE_COLUNA_ATIVO_FII))
  intColunaAtivo = RetornarPrimeiraColuna(Range(RANGE_COLUNA_ATIVO_FII))
  intColunaQtde = RetornarPrimeiraColuna(Range(RANGE_COLUNA_QTDE_FII))
  intColunaSaldoInicial = RetornarPrimeiraColuna(Range(RANGE_COLUNA_SALDO_INICIAL_FII))
  intColunaSaldoFinal = RetornarPrimeiraColuna(Range(RANGE_COLUNA_SALDO_FINAL_FII))
  intColunaCustoMedio = RetornarPrimeiraColuna(Range(RANGE_COLUNA_CUSTO_MEDIO_FII))
  
  Call ColetarInformacoes(intPrimeiraLinha, intUltimaLinha, _
   intColunaAtivo, intColunaSaldoFinal, _
   intColunaQtde, _
   intColunaCustoMedio, _
   wsPlanilhaAtual, infoInvests)
  If (IsArrayEmpty(infoInvests) = True) Then
    Exit Sub
  End If
  Call CopiarRendaVariavel(intPrimeiraLinha, intUltimaLinha, _
    intColunaAtivo, intColunaSaldoInicial, intColunaSaldoFinal, _
    intColunaQtde, _
    wsPlanilhaAtual, wsProxPlanilha, infoInvests)
  Exit Sub
    
ErrorCopiarSaldosCarteiraFii:
  MostrarMsgErro ("CopiarSaldosCarteiraFii")
End Sub

Private Sub CopiarSaldosCarteiraTesouroRF(wsPlanilhaAtual As Worksheet, wsProxPlanilha As Worksheet)
  '
  ' Sub CopiarSaldosCarteiraTesouroRF
  ' copiar saldos, quantidades e descrições da carteira 5
  '
  Dim intPrimeiraLinha As Integer, intUltimaLinha As Integer
  Dim intColunaAtivo As Integer, intColunaQtde As Integer
  Dim intColunaSaldoInicial As Integer, intColunaSaldoFinal As Integer
  Dim intColunaOperacao As Integer
  Dim infoInvests() As infoInvest
  On Error GoTo ErrorCopiarSaldosCarteiraTesouroRF
  intPrimeiraLinha = RetornarPrimeiraLinha(Range(RANGE_COLUNA_ATIVO_RF))
  intUltimaLinha = RetornarUltimaLinha(Range(RANGE_COLUNA_ATIVO_RF))
  intColunaAtivo = RetornarPrimeiraColuna(Range(RANGE_COLUNA_ATIVO_RF))
  intColunaQtde = RetornarPrimeiraColuna(Range(RANGE_COLUNA_QTDE_RF))
  intColunaSaldoInicial = RetornarPrimeiraColuna(Range(RANGE_COLUNA_SALDO_INICIAL_RF))
  intColunaSaldoFinal = RetornarPrimeiraColuna(Range(RANGE_COLUNA_SALDO_FINAL_RF))
  
  Call ColetarInformacoes(intPrimeiraLinha, intUltimaLinha, _
   intColunaAtivo, intColunaSaldoFinal, _
   intColunaQtde, _
   0, _
   wsPlanilhaAtual, infoInvests)
  If (IsArrayEmpty(infoInvests) = True) Then
    Exit Sub
  End If
  Call CopiarRendaVariavel(intPrimeiraLinha, intUltimaLinha, _
    intColunaAtivo, intColunaSaldoInicial, intColunaSaldoFinal, _
    intColunaQtde, _
    wsPlanilhaAtual, wsProxPlanilha, infoInvests)
  Exit Sub
    
ErrorCopiarSaldosCarteiraTesouroRF:
  MostrarMsgErro ("CopiarSaldosCarteiraTesouroRF")
End Sub

Private Sub CopiarSaldosCarteiraTesouroSelic(wsPlanilhaAtual As Worksheet, wsProxPlanilha As Worksheet)
  '
  ' Sub CopiarSaldosCarteiraTesouroSelic
  ' copiar saldos, quantidades e descrições da carteira 6
  '
  Dim intPrimeiraLinha As Integer, intUltimaLinha As Integer
  Dim intColunaAtivo As Integer, intColunaQtde As Integer
  Dim intColunaSaldoInicial As Integer, intColunaSaldoFinal As Integer
  Dim intColunaOperacao As Integer
  Dim infoInvests() As infoInvest
  On Error GoTo ErrorCopiarSaldosCarteiraTesouroSelic
  intPrimeiraLinha = RetornarPrimeiraLinha(Range(RANGE_COLUNA_ATIVO_SELIC))
  intUltimaLinha = RetornarUltimaLinha(Range(RANGE_COLUNA_ATIVO_SELIC))
  intColunaAtivo = RetornarPrimeiraColuna(Range(RANGE_COLUNA_ATIVO_SELIC))
  intColunaQtde = RetornarPrimeiraColuna(Range(RANGE_COLUNA_QTDE_SELIC))
  intColunaSaldoInicial = RetornarPrimeiraColuna(Range(RANGE_COLUNA_SALDO_INICIAL_SELIC))
  intColunaSaldoFinal = RetornarPrimeiraColuna(Range(RANGE_COLUNA_SALDO_FINAL_SELIC))
  
  Call ColetarInformacoes(intPrimeiraLinha, intUltimaLinha, _
   intColunaAtivo, intColunaSaldoFinal, _
   intColunaQtde, _
   0, _
   wsPlanilhaAtual, infoInvests)
  If (IsArrayEmpty(infoInvests) = True) Then
    Exit Sub
  End If
  Call CopiarRendaVariavel(intPrimeiraLinha, intUltimaLinha, _
    intColunaAtivo, intColunaSaldoInicial, intColunaSaldoFinal, _
    intColunaQtde, _
    wsPlanilhaAtual, wsProxPlanilha, infoInvests)
  Exit Sub
    
ErrorCopiarSaldosCarteiraTesouroSelic:
  MostrarMsgErro ("CopiarSaldosCarteiraTesouroSelic")
End Sub

Private Sub ColetarInformacoes(intPrimeiraLinha As Integer, intUltimaLinha As Integer, _
   intColunaAtivo As Integer, intColunaSaldoFinal As Integer, _
   intColunaQtde As Integer, _
   intColunaCustoMedio As Integer, _
   wsPlanilhaAtual As Worksheet, ByRef infoInvests() As infoInvest)
   On Error GoTo ErrorColetarInformacoes
   Dim intPosArray As Integer, intCont As Integer
   For intCont = intPrimeiraLinha To intUltimaLinha
      If (Not IsEmpty(wsPlanilhaAtual.Cells(intCont, intColunaAtivo))) Then
        
        Dim intPosAtivo As Integer
        intPosAtivo = GetPosAtivoDoArray(infoInvests, wsPlanilhaAtual.Cells(intCont, intColunaAtivo).Value)
        If (intPosAtivo < 0) Then
          ' Se ainda não existe a entrada, cria uma nova entrada no array
          Dim infoInvest As infoInvest
          infoInvest.strAtivo = wsPlanilhaAtual.Cells(intCont, intColunaAtivo).Value
          infoInvest.intQuantidade = wsPlanilhaAtual.Cells(intCont, intColunaQtde).Value
          If (intColunaCustoMedio > 0) Then
            infoInvest.dblCustoAnterior = wsPlanilhaAtual.Cells(intCont, intColunaCustoMedio).Value
          End If
          infoInvest.dblSaldo = wsPlanilhaAtual.Cells(intCont, intColunaSaldoFinal).Value
          
          intPosArray = intPosArray + 1
          ReDim Preserve infoInvests(1 To intPosArray)
          infoInvests(intPosArray) = infoInvest
        Else
          ' se já existe, atualiza a qtdade e o custo médio
          If (intColunaCustoMedio > 0) Then
            infoInvests(intPosAtivo).dblCustoAnterior = wsPlanilhaAtual.Cells(intCont, intColunaCustoMedio).Value
          End If
          infoInvests(intPosAtivo).intQuantidade = infoInvests(intPosAtivo).intQuantidade + wsPlanilhaAtual.Cells(intCont, intColunaQtde).Value
        End If
      End If
   Next intCont
   Exit Sub
   
ErrorColetarInformacoes:
  MostrarMsgErro ("ColetarInformacoes")
End Sub

Private Function GetPosAtivoDoArray(infoInvests() As infoInvest, strAtivo As String) As Integer
   Dim intCount As Integer
   On Error GoTo ErrorGetPosAtivoDoArray
   If (IsArrayEmpty(infoInvests) = True) Then
     GetPosAtivoDoArray = -1
     Exit Function
   End If
   For intCount = LBound(infoInvests) To UBound(infoInvests)
      Dim infoInvest As infoInvest
      infoInvest = infoInvests(intCount)
      If (infoInvest.strAtivo = strAtivo) Then
        GetPosAtivoDoArray = intCount
        Exit Function
      End If
   Next intCount
   GetPosAtivoDoArray = -1
   Exit Function
   
ErrorGetPosAtivoDoArray:
  MostrarMsgErro ("GetPosAtivoDoArray")
End Function


Private Function IsArrayEmpty(infoInvests() As infoInvest) As Boolean
   ' Determina se um array contém algum elemento
   Dim lngUBound As Long
   On Error Resume Next
   ' Se o array estiver vazio, um erro ocorrerá quando checar os limites do array
   lngUBound = UBound(infoInvests)
   If Err.Number <> 0 Then
      IsArrayEmpty = True
   Else
      IsArrayEmpty = False
   End If
End Function

Private Sub CopiarRendaVariavel(intPrimeiraLinha As Integer, intUltimaLinha As Integer, _
    intColunaAtivo As Integer, intColunaSaldoInicial As Integer, intColunaSaldoFinal As Integer, _
    intColunaQtde As Integer, _
    wsPlanilhaAtual As Worksheet, wsProxPlanilha As Worksheet, infoInvests() As infoInvest)
  '
  ' Sub CopiarRendaVariavel
  ' copiar saldos e descrições da carteira de renda variável
  '
  On Error GoTo ErrorCopiarRendaVariavel
  
  Dim intCount As Integer, intLinhaDestino As Integer
  For intCount = LBound(infoInvests) To UBound(infoInvests)
     Dim infoInvest As infoInvest
     infoInvest = infoInvests(intCount)
     If (infoInvest.intQuantidade > 0) Then
       intLinhaDestino = GetPrimeiraLinhaLivreDaCarteira(intPrimeiraLinha, intUltimaLinha, intColunaAtivo, wsProxPlanilha)
       With wsProxPlanilha
          .Cells(intLinhaDestino, intColunaAtivo).Value = infoInvest.strAtivo
          .Cells(intLinhaDestino, intColunaQtde).Value = infoInvest.intQuantidade
          .Cells(intLinhaDestino, intColunaSaldoInicial).Value = infoInvest.dblSaldo
          .Cells(intLinhaDestino, intColunaSaldoFinal).Value = infoInvest.dblSaldo
       End With
     End If
  Next intCount
  Exit Sub
  
ErrorCopiarRendaVariavel:
  MostrarMsgErro ("CopiarRendaVariavel")
End Sub

Private Function GetPrimeiraLinhaLivreDaCarteira(intPrimeiraLinhaResumo As Integer, intUltimaLinhaResumo As Integer, _
    intColunaDescricao As Integer, wsPlanilha As Worksheet) As Integer
  '
  ' Function GetPrimeiraLinhaLivreDaCarteira
  ' busca a primeira linha sem dados na carteira da próxima planilha
  '
  On Error GoTo ErrorGetPrimeiraLinhaLivreDaCarteira
  Dim intCont As Integer
  For intCont = intPrimeiraLinhaResumo To intUltimaLinhaResumo
    If (IsEmpty(wsPlanilha.Cells(intCont, intColunaDescricao))) Then
      GetPrimeiraLinhaLivreDaCarteira = intCont
      Exit Function
    End If
  Next intCont
  GetPrimeiraLinhaLivreDaCarteira = intCont
  Exit Function
  
ErrorGetPrimeiraLinhaLivreDaCarteira:
  MostrarMsgErro ("GetPrimeiraLinhaLivreDaCarteira")
End Function

Private Sub CopiarSaldosContaCorretora(wsPlanilhaAtual As Worksheet, wsProxPlanilha As Worksheet)
  '
  ' Sub CopiarSaldosContaCorretora
  ' copiar saldos e descrições da carteira
  '
  On Error GoTo ErrorCopiarSaldosContaCorretora
    
  Dim intCount As Integer, intPrimeiraLinha As Integer, intUltimaLinha As Integer
  Dim intColunaDescricao As Integer, intColunaDisponivel As Integer, intColunaBloqueado As Integer
  intPrimeiraLinha = RetornarPrimeiraLinha(Range(RANGE_COL_DESC_CONTA_CORRETORA))
  intUltimaLinha = RetornarUltimaLinha(Range(RANGE_COL_DESC_CONTA_CORRETORA))
  intColunaDescricao = RetornarPrimeiraColuna(Range(RANGE_COL_DESC_CONTA_CORRETORA))
  intColunaDisponivel = RetornarPrimeiraColuna(Range(RANGE_COL_SALDO_CONTA_CORRETORA))
  intColunaBloqueado = RetornarPrimeiraColuna(Range(RANGE_COL_BLOQUEADO_CONTA_CORRETORA))
  For intCount = intPrimeiraLinha To intUltimaLinha
    If (Not IsEmpty(wsPlanilhaAtual.Cells(intCount, RetornarPrimeiraColuna(Range(RANGE_COL_SALDO_CONTA_CORRETORA))))) Then
      With wsProxPlanilha
        .Cells(intCount, intColunaDescricao).Value = wsPlanilhaAtual.Cells(intCount, intColunaDescricao).Value
        .Cells(intCount, intColunaDisponivel).Value = wsPlanilhaAtual.Cells(intCount, intColunaDisponivel).Value
        .Cells(intCount, intColunaBloqueado).Value = wsPlanilhaAtual.Cells(intCount, intColunaBloqueado).Value
      End With
    End If
  Next intCount
  Exit Sub
  
ErrorCopiarSaldosContaCorretora:
  MostrarMsgErro ("CopiarSaldosContaCorretora")
End Sub

