Option Explicit

' ---------------------------------------------------------------------
' Módulo: FechaMes.bas
' Descrição: Módulo de fechamento do mês
' ---------------------------------------------------------------------

Private Const TITULO_COPIAR_RESUMO As String = "Copiar Resumo Mensal"
Private Const COLUNA_REFERENCIA_FORMULA As String = "D"
Private Const FONTE_SITUACAO As String = "Arial"
Private Const ESTILO_FONTE_SITUACAO As String = "Normal"
Private Const TAMANHO_FONTE_SITUACAO As Integer = 12
Private Const COR_SITUACAO_FECHADO As Integer = 3 'vermelho
Private Const COR_SITUACAO_ABERTO As Integer = 50 'verde

'
' Protege Macro
' Protege ou desprotege a planilha do mês.
'
' Atalho do teclado: Ctrl+b
'
Public Sub ProtegerPlanilha()  
  Dim wsPlanilha As Worksheet
  Dim strSituacao As String
  Dim blnCongelou As Boolean
  On Error GoTo ErroProtegerPlanilha
  Set wsPlanilha = ActiveSheet
  ' Verifica se é uma planilha de movimentação
  If IsError(wsPlanilha.Range(RANGE_SITUAC_PLANILHA).Value) Then
    Exit Sub
  End If
  strSituacao = CStr(wsPlanilha.Range(RANGE_SITUAC_PLANILHA).Value)
  If strSituacao <> SITUAC_ABERTO And strSituacao <> SITUAC_FECHADO Then
    Exit Sub
  End If
  CongelarCalculosPlanilha True
  blnCongelou = True
  If wsPlanilha.ProtectContents Then
    wsPlanilha.Unprotect
    AlterarSituacaoPlanilha wsPlanilha, SITUAC_ABERTO
  Else
    AlterarSituacaoPlanilha wsPlanilha, SITUAC_FECHADO
    CopiarSaldos wsPlanilha
    wsPlanilha.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
  End If

FimProtegerPlanilha:
  On Error Resume Next
  PosicionarTopo
  If blnCongelou Then
    CongelarCalculosPlanilha False
  End If
  Exit Sub

ErroProtegerPlanilha:
  MostrarMsgErro "ProtegerPlanilha"
  Resume FimProtegerPlanilha
End Sub

'
' Muda o texto do status da planilha
'
Private Sub AlterarSituacaoPlanilha(ByVal wsPlanilha As Worksheet, ByVal strSituacao As String)  
  On Error GoTo ErroAlterarSituacaoPlanilha
  With wsPlanilha.Range(RANGE_SITUAC_PLANILHA)
    .Value = strSituacao
    With .Font
      .Name = FONTE_SITUACAO
      .FontStyle = ESTILO_FONTE_SITUACAO
      .Size = TAMANHO_FONTE_SITUACAO
      .Strikethrough = False
      .Superscript = False
      .Subscript = False
      .OutlineFont = False
      .Shadow = False
      .Underline = xlUnderlineStyleNone
      If strSituacao = SITUAC_FECHADO Then
        .ColorIndex = COR_SITUACAO_FECHADO
      Else
        .ColorIndex = COR_SITUACAO_ABERTO
      End If
    End With
  End With
  RetornarUltimaCelulaMovimentacoes(wsPlanilha).Select
  Exit Sub

ErroAlterarSituacaoPlanilha:
  MostrarMsgErro "AlterarSituacaoPlanilha"
End Sub

'
' Sub copiarSaldos        De: 17/01/04
' Copia os saldos dos investimentos de um mês p/ o próximo
'
Private Sub CopiarSaldos(ByVal wsPlanilhaAtual As Worksheet)  
  Dim wsProxPlanilha As Worksheet
  On Error GoTo ErroCopiarSaldos

  ' Testa se existe planilha destino e se ela está aberta
  Set wsProxPlanilha = RetornarProximaPlanilha(wsPlanilhaAtual)
  If wsProxPlanilha Is Nothing Then
    Exit Sub
  End If
  If Not IsPlanilhaAberta(wsProxPlanilha.Range(RANGE_SITUAC_PLANILHA)) Then
    Exit Sub
  End If
  ' Verifica se existe saldo final preenchido no próximo mês do resumo de investimentos
  If HasSaldosCarteira(wsProxPlanilha) Then
    Exit Sub
  End If
  ' Verifica se existe dados a transferir
  If Not HasDadosCarteira(wsPlanilhaAtual) Then
    Exit Sub
  End If
  'Pede confirmação
  If MsgBox("Você deseja montar o Resumo Mensal de Carteiras do próximo mês?", _
        vbYesNo + vbQuestion, TITULO_COPIAR_RESUMO) = vbNo Then
    Exit Sub
  End If

  ' Reproduz o Resumo de Investimento Atual no próximo mês
  CopiarSaldosPortfolio wsPlanilhaAtual, wsProxPlanilha
  CopiarSaldosRendaVariavel wsPlanilhaAtual, wsProxPlanilha, "Ações", RANGE_COLUNA_ATIVO_ACOES, _
    RANGE_COLUNA_QTDE_ACOES, RANGE_COLUNA_SALDO_INICIAL_ACOES, RANGE_COLUNA_SALDO_FINAL_ACOES
  CopiarSaldosRendaVariavel wsPlanilhaAtual, wsProxPlanilha, "FII", RANGE_COLUNA_ATIVO_FII, _
    RANGE_COLUNA_QTDE_FII, RANGE_COLUNA_SALDO_INICIAL_FII, RANGE_COLUNA_SALDO_FINAL_FII
  CopiarSaldosRendaVariavel wsPlanilhaAtual, wsProxPlanilha, "Tesouro Direto", RANGE_COLUNA_ATIVO_TESOURO_DIRETO, _
    RANGE_COLUNA_QTDE_TESOURO_DIRETO, RANGE_COLUNA_SALDO_INICIAL_TESOURO_DIRETO, RANGE_COLUNA_SALDO_FINAL_TESOURO_DIRETO
  CopiarSaldosRendaVariavel wsPlanilhaAtual, wsProxPlanilha, "Tesouro Selic", RANGE_COLUNA_ATIVO_TESOURO_SELIC, _
    RANGE_COLUNA_QTDE_TESOURO_SELIC, RANGE_COLUNA_SALDO_INICIAL_TESOURO_SELIC, RANGE_COLUNA_SALDO_FINAL_TESOURO_SELIC
  CopiarSaldosRendaVariavel wsPlanilhaAtual, wsProxPlanilha, "ETF", RANGE_COLUNA_ATIVO_ETF, _
    RANGE_COLUNA_QTDE_ETF, RANGE_COLUNA_SALDO_INICIAL_ETF, RANGE_COLUNA_SALDO_FINAL_ETF
  CopiarSaldosRendaVariavel wsPlanilhaAtual, wsProxPlanilha, "Stock", RANGE_COLUNA_ATIVO_STOCK, _
    RANGE_COLUNA_QTDE_STOCK, RANGE_COLUNA_SALDO_INICIAL_STOCK, RANGE_COLUNA_SALDO_FINAL_STOCK
  CopiarSaldosRendaVariavel wsPlanilhaAtual, wsProxPlanilha, "REIT", RANGE_COLUNA_ATIVO_REIT, _
    RANGE_COLUNA_QTDE_REIT, RANGE_COLUNA_SALDO_INICIAL_REIT, RANGE_COLUNA_SALDO_FINAL_REIT
  CopiarSaldosRendaVariavel wsPlanilhaAtual, wsProxPlanilha, "Treasury", RANGE_COLUNA_ATIVO_TREASURY, _
    RANGE_COLUNA_QTDE_TREASURY, RANGE_COLUNA_SALDO_INICIAL_TREASURY, RANGE_COLUNA_SALDO_FINAL_TREASURY
  CopiarSaldosRendaVariavel wsPlanilhaAtual, wsProxPlanilha, "Ouro", RANGE_COLUNA_ATIVO_OURO, _
    RANGE_COLUNA_QTDE_OURO, RANGE_COLUNA_SALDO_INICIAL_OURO, RANGE_COLUNA_SALDO_FINAL_OURO
  CopiarSaldosRendaVariavel wsPlanilhaAtual, wsProxPlanilha, "Cripto", RANGE_COLUNA_ATIVO_CRIPTO, _
    RANGE_COLUNA_QTDE_CRIPTO, RANGE_COLUNA_SALDO_INICIAL_CRIPTO, RANGE_COLUNA_SALDO_FINAL_CRIPTO
  CopiarSaldosContaCorretora wsPlanilhaAtual, wsProxPlanilha

  ' Ordena os ativos de todas as carteiras do próximo mês
  OrdenarCarteiras wsProxPlanilha
  Exit Sub

ErroCopiarSaldos:
  MostrarMsgErro "CopiarSaldos"
End Sub

'
' Function HasSaldosCarteira
' verifica se existem saldos no próximo mês
'
Private Function HasSaldosCarteira(ByVal wsProxPlanilha As Worksheet) As Boolean  
  Dim varSaldos As Variant
  Dim lngLinha As Long
  varSaldos = wsProxPlanilha.Range(RANGE_COLUNA_SALDO_FINAL_PORTFOLIO).Value
  For lngLinha = 1 To UBound(varSaldos, 1)
    If RetornarValorNumerico(varSaldos(lngLinha, 1)) > 0 Then
      HasSaldosCarteira = True
      Exit Function
    End If
  Next lngLinha
End Function

'
' Function HasDadosCarteira
' verifica se existem dados a transferir
'
Private Function HasDadosCarteira(ByVal wsPlanilhaAtual As Worksheet) As Boolean  
  Dim varRanges As Variant, varRange As Variant
  varRanges = Array(RANGE_COLUNA_ATIVO_PORTFOLIO, RANGE_COLUNA_ATIVO_ACOES, RANGE_COLUNA_ATIVO_FII, _
    RANGE_COLUNA_ATIVO_TESOURO_DIRETO, RANGE_COLUNA_ATIVO_TESOURO_SELIC)
  For Each varRange In varRanges
    If Application.WorksheetFunction.CountA(wsPlanilhaAtual.Range(varRange)) > 0 Then
      HasDadosCarteira = True
      Exit Function
    End If
  Next varRange
End Function

'
' Sub CopiarSaldosPortfolio
' copia saldos e descrições do portfólio
'
Private Sub CopiarSaldosPortfolio(ByVal wsPlanilhaAtual As Worksheet, ByVal wsProxPlanilha As Worksheet)  
  Dim lngPrimeiraLinha As Long, lngUltimaLinha As Long, lngLinha As Long, lngLinhaDestino As Long
  Dim lngColunaDescricao As Long, lngColunaSaldoInicial As Long
  Dim lngColunaAuxiliarRend As Long, lngColunaSaldoFinal As Long
  Dim rgSaldoInicialAtual As Range, rgSaldoFinalAtual As Range
  On Error GoTo ErroCopiarSaldosPortfolio

  lngPrimeiraLinha = RetornarPrimeiraLinha(wsPlanilhaAtual.Range(RANGE_COLUNA_ATIVO_PORTFOLIO))
  lngUltimaLinha = RetornarUltimaLinha(wsPlanilhaAtual.Range(RANGE_COLUNA_ATIVO_PORTFOLIO))
  lngColunaDescricao = RetornarPrimeiraColuna(wsPlanilhaAtual.Range(RANGE_COLUNA_ATIVO_PORTFOLIO))
  lngColunaSaldoInicial = RetornarPrimeiraColuna(wsPlanilhaAtual.Range(RANGE_COLUNA_SALDO_INICIAL_PORTFOLIO))
  lngColunaAuxiliarRend = RetornarPrimeiraColuna(wsPlanilhaAtual.Range(RANGE_COLUNA_AUXILIAR_REND_PORTFOLIO))
  lngColunaSaldoFinal = RetornarPrimeiraColuna(wsPlanilhaAtual.Range(RANGE_COLUNA_SALDO_FINAL_PORTFOLIO))

  For lngLinha = lngPrimeiraLinha To lngUltimaLinha
    Set rgSaldoInicialAtual = wsPlanilhaAtual.Cells(lngLinha, lngColunaSaldoInicial)
    Set rgSaldoFinalAtual = wsPlanilhaAtual.Cells(lngLinha, lngColunaSaldoFinal)
    'se existir tipo de investimento na planilha atual
    If Not IsEmpty(wsPlanilhaAtual.Cells(lngLinha, lngColunaDescricao)) And _
       RetornarValorNumerico(rgSaldoFinalAtual.Value) <> 0 Then
      lngLinhaDestino = lngLinha
      If lngLinhaDestino > lngPrimeiraLinha Then
        If IsEmpty(wsProxPlanilha.Cells(lngLinhaDestino - 1, lngColunaDescricao)) Then
          lngLinhaDestino = RetornarPrimeiraLinhaLivre(wsProxPlanilha, lngPrimeiraLinha, lngUltimaLinha, lngColunaDescricao)
        End If
      End If
      If lngLinhaDestino > lngUltimaLinha Then
        Err.Raise ERRO_TABELA_CHEIA, , "Não há linhas livres no portfólio do próximo mês."
      End If
      With wsProxPlanilha
        .Cells(lngLinhaDestino, lngColunaDescricao).Value = wsPlanilhaAtual.Cells(lngLinha, lngColunaDescricao).Value
        If rgSaldoFinalAtual.HasFormula And rgSaldoInicialAtual.HasFormula And _
           InStr(rgSaldoInicialAtual.Formula, COLUNA_REFERENCIA_FORMULA) > 0 Then
          .Cells(lngLinhaDestino, lngColunaSaldoInicial).Formula = _
            AjustarLinhaFormula(rgSaldoInicialAtual.Formula, lngLinha, lngLinhaDestino)
        Else
          .Cells(lngLinhaDestino, lngColunaSaldoInicial).Value = rgSaldoFinalAtual.Value
        End If
        .Cells(lngLinhaDestino, lngColunaAuxiliarRend).Value = wsPlanilhaAtual.Cells(lngLinha, lngColunaAuxiliarRend).Value
        If rgSaldoFinalAtual.HasFormula Then
          .Cells(lngLinhaDestino, lngColunaSaldoFinal).Formula = _
            AjustarLinhaFormula(rgSaldoFinalAtual.Formula, lngLinha, lngLinhaDestino)
        Else
          .Cells(lngLinhaDestino, lngColunaSaldoFinal).Value = rgSaldoFinalAtual.Value
        End If
      End With
    End If
  Next lngLinha
  Exit Sub

ErroCopiarSaldosPortfolio:
  MostrarMsgErro "CopiarSaldosPortfolio"
End Sub

'
' Ajusta a referência à coluna de descrição quando o item muda de linha
'
Private Function AjustarLinhaFormula(ByVal strFormula As String, ByVal lngLinhaOrigem As Long, _
    ByVal lngLinhaDestino As Long) As String  
  If lngLinhaOrigem = lngLinhaDestino Then
    AjustarLinhaFormula = strFormula
  Else
    AjustarLinhaFormula = Replace(strFormula, COLUNA_REFERENCIA_FORMULA & lngLinhaOrigem, _
      COLUNA_REFERENCIA_FORMULA & lngLinhaDestino)
  End If
End Function

'
' Sub CopiarSaldosRendaVariavel
' copia saldos, quantidades e descrições de uma carteira, agrupando os ativos repetidos
'
Private Sub CopiarSaldosRendaVariavel(ByVal wsPlanilhaAtual As Worksheet, ByVal wsProxPlanilha As Worksheet, _
    ByVal strNomeCarteira As String, ByVal strRangeAtivo As String, ByVal strRangeQtde As String, _
    ByVal strRangeSaldoInicial As String, ByVal strRangeSaldoFinal As String)  
  Dim udtAtivos() As InfoAtivo
  Dim lngTotalAtivos As Long, lngIndice As Long
  Dim lngPrimeiraLinha As Long, lngUltimaLinha As Long, lngLinhaDestino As Long
  Dim lngColunaAtivo As Long, lngColunaQtde As Long
  Dim lngColunaSaldoInicial As Long, lngColunaSaldoFinal As Long
  On Error GoTo ErroCopiarSaldosRendaVariavel

  lngTotalAtivos = ColetarAtivos(wsPlanilhaAtual, strRangeAtivo, strRangeQtde, strRangeSaldoFinal, udtAtivos)
  If lngTotalAtivos = 0 Then
    Exit Sub
  End If
  lngPrimeiraLinha = RetornarPrimeiraLinha(wsProxPlanilha.Range(strRangeAtivo))
  lngUltimaLinha = RetornarUltimaLinha(wsProxPlanilha.Range(strRangeAtivo))
  lngColunaAtivo = RetornarPrimeiraColuna(wsProxPlanilha.Range(strRangeAtivo))
  lngColunaQtde = RetornarPrimeiraColuna(wsProxPlanilha.Range(strRangeQtde))
  lngColunaSaldoInicial = RetornarPrimeiraColuna(wsProxPlanilha.Range(strRangeSaldoInicial))
  lngColunaSaldoFinal = RetornarPrimeiraColuna(wsProxPlanilha.Range(strRangeSaldoFinal))

  lngLinhaDestino = lngPrimeiraLinha
  For lngIndice = 1 To lngTotalAtivos
    If udtAtivos(lngIndice).dblQuantidade >= QUANTIDADE_MINIMA_ATIVO Then
      ' as linhas anteriores já foram preenchidas: continua a busca a partir da última gravada
      lngLinhaDestino = RetornarPrimeiraLinhaLivre(wsProxPlanilha, lngLinhaDestino, lngUltimaLinha, lngColunaAtivo)
      If lngLinhaDestino > lngUltimaLinha Then
        Err.Raise ERRO_TABELA_CHEIA, , "Não há linhas livres na carteira " & strNomeCarteira & " do próximo mês."
      End If
      With wsProxPlanilha
        .Cells(lngLinhaDestino, lngColunaAtivo).Value = udtAtivos(lngIndice).strAtivo
        .Cells(lngLinhaDestino, lngColunaQtde).Value = udtAtivos(lngIndice).dblQuantidade
        .Cells(lngLinhaDestino, lngColunaSaldoInicial).Value = udtAtivos(lngIndice).dblSaldo
        .Cells(lngLinhaDestino, lngColunaSaldoFinal).Value = udtAtivos(lngIndice).dblSaldo
      End With
      lngLinhaDestino = lngLinhaDestino + 1
    End If
  Next lngIndice
  Exit Sub

ErroCopiarSaldosRendaVariavel:
  MostrarMsgErro "CopiarSaldosRendaVariavel (" & strNomeCarteira & ")"
End Sub

'
' Function RetornarPrimeiraLinhaLivre
' busca a primeira linha sem dados na carteira; retorna lngUltimaLinha + 1 se não houver
'
Private Function RetornarPrimeiraLinhaLivre(ByVal wsPlanilha As Worksheet, ByVal lngPrimeiraLinha As Long, _
    ByVal lngUltimaLinha As Long, ByVal lngColunaDescricao As Long) As Long  
  Dim lngLinha As Long
  For lngLinha = lngPrimeiraLinha To lngUltimaLinha
    If IsEmpty(wsPlanilha.Cells(lngLinha, lngColunaDescricao)) Then
      Exit For
    End If
  Next lngLinha
  RetornarPrimeiraLinhaLivre = lngLinha
End Function

'
' Sub CopiarSaldosContaCorretora
' copia os saldos das contas nas corretoras
'
Private Sub CopiarSaldosContaCorretora(ByVal wsPlanilhaAtual As Worksheet, ByVal wsProxPlanilha As Worksheet)  
  Dim varRanges As Variant, varRange As Variant
  On Error GoTo ErroCopiarSaldosContaCorretora
  varRanges = Array(RANGE_SALDO_CONTA_XP, RANGE_SALDO_CONTA_AVENUE_USD_TOTAL, _
    RANGE_SALDO_CONTA_AVENUE_USD_DO_BRASIL, RANGE_SALDO_CONTA_AVENUE_BR)
  For Each varRange In varRanges
    If Not IsEmpty(wsPlanilhaAtual.Range(varRange)) Then
      wsProxPlanilha.Range(varRange).Value = wsPlanilhaAtual.Range(varRange).Value
    End If
  Next varRange
  Exit Sub

ErroCopiarSaldosContaCorretora:
  MostrarMsgErro "CopiarSaldosContaCorretora"
End Sub
