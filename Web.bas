' Módulo de funções para buscar dados on line
Option Explicit

Private Type infoAtivo
  strAtivo As String
  intQuantidade As Double
End Type

Sub AtualizarDadosWeb()
  '
  ' AtualizarDadosWeb    Data: 31/07/23
  ' Atualiza tabela com fonte de dados na web
  '
  On Error GoTo ErroAtualizarDadosWeb
    
  If (Range(RANGE_SITUAC_PLANILHA).Value <> SITUAC_ABERTO) Then
    Exit Sub
  End If
  'Pede confirmação
  If MsgBox("Você deseja atualizar as cotações das carteiras?", _
        vbYesNo + vbQuestion, "Copiar Resumo Mensal") = vbNo Then
    Exit Sub
  End If
  Dim wsPlanilhaAtual As Worksheet, wsPlanilhaAcoes As Worksheet, wsPlanilhaTesouroDireto As Worksheet
  Dim blnOldStatusBar As Boolean
  Dim intPercentual As Integer
  Dim strAtivosNaoEncontrados As String
  blnOldStatusBar = Application.DisplayStatusBar
  Application.DisplayStatusBar = True
  CongelarCalculosPlanilha (True)
  Set wsPlanilhaAtual = ActiveSheet
  Set wsPlanilhaAcoes = Worksheets("Acoes")
  Set wsPlanilhaTesouroDireto = Worksheets("TesouroDireto")
  wsPlanilhaAcoes.Visible = xlSheetVisible
  wsPlanilhaTesouroDireto.Visible = xlSheetVisible
  strAtivosNaoEncontrados = ""
  
  intPercentual = 0
  Application.StatusBar = "Buscando cotações das bolsas de valores... " & intPercentual & "% completo."
  Application.Calculation = xlCalculationAutomatic
  Call AtualizarAcoesEMoedas
  Application.Calculation = xlCalculationManual
  
  intPercentual = 5
  Application.StatusBar = "Buscando cotações do Tesouro Direto... " & intPercentual & "% completo."
  'If Not AtualizarConsultasEConexoesTesouroDireto(wsPlanilhaTesouroDireto) Then
  '  MsgBox "Algumas consultas e conexões do TD não foram atualizadas...", vbExclamation
  'End If
  
  intPercentual = 10
  Application.StatusBar = "Transferindo dados de Ações Brasil... " & intPercentual & "% completo."
  Debug.Print "Transferindo ações Brasil..."
  Call AtualizarCotacaoAcoes(wsPlanilhaAtual, wsPlanilhaAcoes, strAtivosNaoEncontrados)
  
  intPercentual = 20
  Application.StatusBar = "Transferindo dados de FII... " & intPercentual & "% completo."
  Debug.Print "Transferindo FII..."
  Call AtualizarCotacaoFii(wsPlanilhaAtual, wsPlanilhaAcoes, strAtivosNaoEncontrados)
  
  intPercentual = 30
  'Application.StatusBar = "Transferindo dados de Tesouro Direto pré e indexado... " & intPercentual & "% completo."
  'Debug.Print "Transferindo TD..."
  'Call AtualizarCotacaoTd(wsPlanilhaAtual, wsPlanilhaTesouroDireto)
  
  intPercentual = 40
  'Application.StatusBar = "Transferindo dados de Tesouro Direto Selic... " & intPercentual & "% completo."
  'Debug.Print "Transferindo Selic..."
  'Call AtualizarCotacaoSelic(wsPlanilhaAtual, wsPlanilhaTesouroDireto)
  
  intPercentual = 50
  Application.StatusBar = "Transferindo dados de ETF... " & intPercentual & "% completo."
  Debug.Print "Transferindo ETF Brasil..."
  Call AtualizarCotacaoEtf(wsPlanilhaAtual, wsPlanilhaAcoes, strAtivosNaoEncontrados)
  
  intPercentual = 60
  Application.StatusBar = "Transferindo dados de Ações Internacionais... " & intPercentual & "% completo."
  Debug.Print "Transferindo ações USD..."
  Call AtualizarCotacaoStock(wsPlanilhaAtual, wsPlanilhaAcoes, strAtivosNaoEncontrados)
  
  intPercentual = 70
  Application.StatusBar = "Transferindo dados de REIT... " & intPercentual & "% completo."
  Debug.Print "Transferindo REIT..."
  Call AtualizarCotacaoReit(wsPlanilhaAtual, wsPlanilhaAcoes, strAtivosNaoEncontrados)
  
  intPercentual = 80
  Application.StatusBar = "Transferindo dados de Treasury... " & intPercentual & "% completo."
  Debug.Print "Transferindo Treasury..."
  Call AtualizarCotacaoTreasuries(wsPlanilhaAtual, wsPlanilhaAcoes, strAtivosNaoEncontrados)
  
  intPercentual = 90
  Application.StatusBar = "Transferindo dados de Ouro... " & intPercentual & "% completo."
  Debug.Print "Transferindo Ouro..."
  Call AtualizarCotacaoOuro(wsPlanilhaAtual, wsPlanilhaAcoes, strAtivosNaoEncontrados)
  
  intPercentual = 100
  Application.StatusBar = "Transferindo dados de Cripto... " & intPercentual & "% completo."
  Debug.Print "Transferindo Cripto..."
  Call AtualizarCotacaoCripto(wsPlanilhaAtual, wsPlanilhaAcoes, strAtivosNaoEncontrados)
  
  wsPlanilhaAtual.Range(RANGE_CELULA_DOLAR_FINAL_MES).Value = wsPlanilhaAcoes.Range(RANGE_CELULA_DOLAR).Value
  Debug.Print "Ativos não encontrados: " & strAtivosNaoEncontrados

FimAtualizarDadosWeb:
  wsPlanilhaAcoes.Visible = xlSheetHidden
  wsPlanilhaTesouroDireto.Visible = xlSheetHidden
  CongelarCalculosPlanilha (False)
  Application.StatusBar = False
  Application.DisplayStatusBar = blnOldStatusBar
  wsPlanilhaAtual.Activate
  Set wsPlanilhaAtual = Nothing
  Set wsPlanilhaAcoes = Nothing
  Set wsPlanilhaTesouroDireto = Nothing
  Exit Sub
   
ErroAtualizarDadosWeb:
  MostrarMsgErro ("AtualizarDadosWeb")
  Resume FimAtualizarDadosWeb
End Sub

Private Sub AtualizarAcoesEMoedas()
  On Error GoTo ErroAtualizarAcoesEMoedas
  
  Dim wsAcoes As Worksheet
  Debug.Print "Atualizando ações e moedas..."
  Set wsAcoes = Worksheets("Acoes")
  wsAcoes.Range(RANGE_TAB_ACOES_365).RefreshLinkedDataType
  wsAcoes.Range(RANGE_TAB_MOEDAS_365).RefreshLinkedDataType
  Debug.Print "Concluído atualização de ações e moedas."
  Exit Sub
ErroAtualizarAcoesEMoedas:
  Debug.Print "Erro ao atualizar ações e moedas. Erro: " & Err.Number & " - " & Err.Description
  MostrarMsgErro ("AtualizarAcoesEMoedas")
End Sub

Private Function AtualizarConsultasEConexoes() As Boolean
  On Error GoTo ErroAtualizarConsultasEConexoes
  
  Dim objConnection As Object
  Dim bBackground As Boolean
  Debug.Print "Atualizando todas as consultas e conexões..."
  AtualizarConsultasEConexoes = True
  For Each objConnection In ThisWorkbook.Connections
    'Busca o valor atual do background-refresh
    bBackground = objConnection.OLEDBConnection.BackgroundQuery
    'Temporariamente disativa o background-refresh
    'On Error GoTo Err_Control
    objConnection.OLEDBConnection.BackgroundQuery = False
    'Atualiza esta conexão
    'On Error GoTo Err_Control
    objConnection.Refresh
    'Ajusta o background-refresh de volta a seu valor original
    'On Error GoTo Err_Control
    objConnection.OLEDBConnection.BackgroundQuery = bBackground
  Next
  Debug.Print "Concluído atualização de todas as consultas e conexões."
ErroAtualizarConsultasEConexoes:
  If Err.Number <> 0 Then
    objConnection.OLEDBConnection.BackgroundQuery = bBackground
    AtualizarConsultasEConexoes = False
    Debug.Print "Erro ao atualizar consultas e conexões. Erro: " & Err.Number & " - " & Err.Description
    If Err.Number = 1004 Then
      Resume Next
    Else
      MostrarMsgErro ("AtualizarConsultasEConexoes")
    End If
  End If
End Function

Private Function AtualizarConsultasEConexoesTesouroDireto(wsPlanilhaTesouroDireto As Worksheet) As Boolean
  On Error GoTo ErroAtualizarConsultasEConexoesTesouroDireto
  
  Dim strHoje As String, strAtualiz As String
  
  Debug.Print "Atualizando consultas e conexões do Tesouro Direto..."
  strHoje = Format(Now, "dd/mm/yyyy")
  strAtualiz = Format(wsPlanilhaTesouroDireto.Range(RANGE_DATA_ATUALIZ_TD).Value, "dd/mm/yyyy")
  Debug.Print "Data atual: " & strHoje & " - última atualização: " & strAtualiz
  AtualizarConsultasEConexoesTesouroDireto = True
  If (strHoje = strAtualiz) Then
    Debug.Print "A cotação do Tesouro Direto já foi atualizada hoje."
    Exit Function
  End If
  ActiveWorkbook.Connections("Consulta - TesouroDireto").Refresh
  Debug.Print "Concluído atualização do Tesouro Direto."
  wsPlanilhaTesouroDireto.Range(RANGE_DATA_ATUALIZ_TD).Value = Now
ErroAtualizarConsultasEConexoesTesouroDireto:
  If Err.Number <> 0 Then
    Debug.Print "Erro ao atualizar Tesouro Direto. Erro: " & Err.Number & " - " & Err.Description
    AtualizarConsultasEConexoesTesouroDireto = False
    If Err.Number = 1004 Then
      Resume Next
    Else
      MostrarMsgErro ("AtualizarConsultasEConexoesTesouroDireto")
    End If
  End If
End Function

Private Sub AtualizarCotacaoAcoes(wsPlanilhaAtual As Worksheet, wsPlanilhaAcoes As Worksheet, ByRef strAtivosNaoEncontrados As String)
  ' Sub AtualizarCotacaoAcoes
  ' atualiza valor da cota final do ativo
  Dim intPrimeiraLinha As Integer, intUltimaLinha As Integer
  Dim intColunaAtivo As Integer, intColunaQtde As Integer
  Dim intColunaSaldoFinal As Integer
  Dim infoAtivos() As infoAtivo
  Dim strRetorno As String
  On Error GoTo ErrorAtualizarCotacaoAcoes
  intPrimeiraLinha = RetornarPrimeiraLinha(Range(RANGE_COLUNA_ATIVO_ACOES))
  intUltimaLinha = RetornarUltimaLinha(Range(RANGE_COLUNA_ATIVO_ACOES))
  intColunaAtivo = RetornarPrimeiraColuna(Range(RANGE_COLUNA_ATIVO_ACOES))
  intColunaQtde = RetornarPrimeiraColuna(Range(RANGE_COLUNA_QTDE_ACOES))
  intColunaSaldoFinal = RetornarPrimeiraColuna(Range(RANGE_COLUNA_SALDO_FINAL_ACOES))
    
  Call ColetarDados(intPrimeiraLinha, intUltimaLinha, _
   intColunaAtivo, intColunaQtde, _
   wsPlanilhaAtual, infoAtivos)
  If (IsArrayEmpty(infoAtivos) = True) Then
    Exit Sub
  End If
  'atualizar cada ativo com a a cotação atual
  strRetorno = AtualizarCotacao(intPrimeiraLinha, intUltimaLinha, _
    intColunaAtivo, intColunaQtde, intColunaSaldoFinal, _
    wsPlanilhaAtual, wsPlanilhaAcoes, infoAtivos)
  strAtivosNaoEncontrados = AnexarAtivo(strAtivosNaoEncontrados, strRetorno)
  
  Exit Sub
ErrorAtualizarCotacaoAcoes:
  MostrarMsgErro ("AtualizarCotacaoAcoes")
End Sub

Private Sub AtualizarCotacaoFii(wsPlanilhaAtual As Worksheet, wsPlanilhaAcoes As Worksheet, ByRef strAtivosNaoEncontrados As String)
  ' Sub AtualizarCotacaoFii
  ' atualiza valor da cota final do ativo
  Dim intPrimeiraLinha As Integer, intUltimaLinha As Integer
  Dim intColunaAtivo As Integer, intColunaQtde As Integer
  Dim intColunaSaldoFinal As Integer
  Dim infoAtivos() As infoAtivo
  Dim strRetorno As String
  On Error GoTo ErrorAtualizarCotacaoFii
  intPrimeiraLinha = RetornarPrimeiraLinha(Range(RANGE_COLUNA_ATIVO_FII))
  intUltimaLinha = RetornarUltimaLinha(Range(RANGE_COLUNA_ATIVO_FII))
  intColunaAtivo = RetornarPrimeiraColuna(Range(RANGE_COLUNA_ATIVO_FII))
  intColunaQtde = RetornarPrimeiraColuna(Range(RANGE_COLUNA_QTDE_FII))
  intColunaSaldoFinal = RetornarPrimeiraColuna(Range(RANGE_COLUNA_SALDO_FINAL_FII))
    
  Call ColetarDados(intPrimeiraLinha, intUltimaLinha, _
   intColunaAtivo, intColunaQtde, _
   wsPlanilhaAtual, infoAtivos)
  If (IsArrayEmpty(infoAtivos) = True) Then
    Exit Sub
  End If
  'atualizar cada ativo com a a cotação atual
  strRetorno = AtualizarCotacao(intPrimeiraLinha, intUltimaLinha, _
    intColunaAtivo, intColunaQtde, intColunaSaldoFinal, _
    wsPlanilhaAtual, wsPlanilhaAcoes, infoAtivos)
  strAtivosNaoEncontrados = AnexarAtivo(strAtivosNaoEncontrados, strRetorno)
  
  Exit Sub
ErrorAtualizarCotacaoFii:
  MostrarMsgErro ("AtualizarCotacaoFii")
End Sub

Private Sub AtualizarCotacaoTd(wsPlanilhaAtual As Worksheet, wsPlanilhaTesouroDireto As Worksheet)
  ' Sub AtualizarCotacaoTd
  ' atualiza valor da cota final do ativo
  Dim intPrimeiraLinha As Integer, intUltimaLinha As Integer
  Dim intColunaAtivo As Integer, intColunaQtde As Integer
  Dim intColunaSaldoFinal As Integer
  Dim infoAtivos() As infoAtivo
  On Error GoTo ErrorAtualizarCotacaoTd
  intPrimeiraLinha = RetornarPrimeiraLinha(Range(RANGE_COLUNA_ATIVO_TESOURO_DIRETO))
  intUltimaLinha = RetornarUltimaLinha(Range(RANGE_COLUNA_ATIVO_TESOURO_DIRETO))
  intColunaAtivo = RetornarPrimeiraColuna(Range(RANGE_COLUNA_ATIVO_TESOURO_DIRETO))
  intColunaQtde = RetornarPrimeiraColuna(Range(RANGE_COLUNA_QTDE_TESOURO_DIRETO))
  intColunaSaldoFinal = RetornarPrimeiraColuna(Range(RANGE_COLUNA_SALDO_FINAL_TESOURO_DIRETO))
    
  Call ColetarDados(intPrimeiraLinha, intUltimaLinha, _
   intColunaAtivo, intColunaQtde, _
   wsPlanilhaAtual, infoAtivos)
  If (IsArrayEmpty(infoAtivos) = True) Then
    Exit Sub
  End If
  'atualizar cada ativo com a a cotação atual
  Call AtualizarCotacao(intPrimeiraLinha, intUltimaLinha, _
    intColunaAtivo, intColunaQtde, intColunaSaldoFinal, _
    wsPlanilhaAtual, wsPlanilhaTesouroDireto, infoAtivos)
  
  Exit Sub
ErrorAtualizarCotacaoTd:
  MostrarMsgErro ("AtualizarCotacaoTd")
End Sub

Private Sub AtualizarCotacaoSelic(wsPlanilhaAtual As Worksheet, wsPlanilhaTesouroDireto As Worksheet)
  ' Sub AtualizarCotacaoSelic
  ' atualiza valor da cota final do ativo
  Dim intPrimeiraLinha As Integer, intUltimaLinha As Integer
  Dim intColunaAtivo As Integer, intColunaQtde As Integer
  Dim intColunaSaldoFinal As Integer
  Dim infoAtivos() As infoAtivo
  On Error GoTo ErrorAtualizarCotacaoSelic
  intPrimeiraLinha = RetornarPrimeiraLinha(Range(RANGE_COLUNA_ATIVO_TESOURO_SELIC))
  intUltimaLinha = RetornarUltimaLinha(Range(RANGE_COLUNA_ATIVO_TESOURO_SELIC))
  intColunaAtivo = RetornarPrimeiraColuna(Range(RANGE_COLUNA_ATIVO_TESOURO_SELIC))
  intColunaQtde = RetornarPrimeiraColuna(Range(RANGE_COLUNA_QTDE_TESOURO_SELIC))
  intColunaSaldoFinal = RetornarPrimeiraColuna(Range(RANGE_COLUNA_SALDO_FINAL_TESOURO_SELIC))
    
  Call ColetarDados(intPrimeiraLinha, intUltimaLinha, _
   intColunaAtivo, intColunaQtde, _
   wsPlanilhaAtual, infoAtivos)
  If (IsArrayEmpty(infoAtivos) = True) Then
    Exit Sub
  End If
  'atualizar cada ativo com a a cotação atual
  Call AtualizarCotacao(intPrimeiraLinha, intUltimaLinha, _
    intColunaAtivo, intColunaQtde, intColunaSaldoFinal, _
    wsPlanilhaAtual, wsPlanilhaTesouroDireto, infoAtivos)
  
  Exit Sub
ErrorAtualizarCotacaoSelic:
  MostrarMsgErro ("AtualizarCotacaoSelic")
End Sub

Private Sub AtualizarCotacaoEtf(wsPlanilhaAtual As Worksheet, wsPlanilhaAcoes As Worksheet, ByRef strAtivosNaoEncontrados As String)
  ' Sub AtualizarCotacaoEtf
  ' atualiza valor da cota final do ativo
  Dim intPrimeiraLinha As Integer, intUltimaLinha As Integer
  Dim intColunaAtivo As Integer, intColunaQtde As Integer
  Dim intColunaSaldoFinal As Integer
  Dim infoAtivos() As infoAtivo
  Dim strRetorno As String
  On Error GoTo ErrorAtualizarCotacaoEtf
  intPrimeiraLinha = RetornarPrimeiraLinha(Range(RANGE_COLUNA_ATIVO_ETF))
  intUltimaLinha = RetornarUltimaLinha(Range(RANGE_COLUNA_ATIVO_ETF))
  intColunaAtivo = RetornarPrimeiraColuna(Range(RANGE_COLUNA_ATIVO_ETF))
  intColunaQtde = RetornarPrimeiraColuna(Range(RANGE_COLUNA_QTDE_ETF))
  intColunaSaldoFinal = RetornarPrimeiraColuna(Range(RANGE_COLUNA_SALDO_FINAL_ETF))
    
  Call ColetarDados(intPrimeiraLinha, intUltimaLinha, _
   intColunaAtivo, intColunaQtde, _
   wsPlanilhaAtual, infoAtivos)
  If (IsArrayEmpty(infoAtivos) = True) Then
    Exit Sub
  End If
  'atualizar cada ativo com a a cotação atual
  strRetorno = AtualizarCotacao(intPrimeiraLinha, intUltimaLinha, _
    intColunaAtivo, intColunaQtde, intColunaSaldoFinal, _
    wsPlanilhaAtual, wsPlanilhaAcoes, infoAtivos)
  strAtivosNaoEncontrados = AnexarAtivo(strAtivosNaoEncontrados, strRetorno)
  
  Exit Sub
ErrorAtualizarCotacaoEtf:
  MostrarMsgErro ("AtualizarCotacaoEtf")
End Sub

Private Sub AtualizarCotacaoStock(wsPlanilhaAtual As Worksheet, wsPlanilhaAcoes As Worksheet, ByRef strAtivosNaoEncontrados As String)
  ' Sub AtualizarCotacaoStock
  ' atualiza valor da cota final do ativo
  Dim intPrimeiraLinha As Integer, intUltimaLinha As Integer
  Dim intColunaAtivo As Integer, intColunaQtde As Integer
  Dim intColunaSaldoFinal As Integer
  Dim infoAtivos() As infoAtivo
  Dim strRetorno As String
  On Error GoTo ErrorAtualizarCotacaoStock
  intPrimeiraLinha = RetornarPrimeiraLinha(Range(RANGE_COLUNA_ATIVO_STOCK))
  intUltimaLinha = RetornarUltimaLinha(Range(RANGE_COLUNA_ATIVO_STOCK))
  intColunaAtivo = RetornarPrimeiraColuna(Range(RANGE_COLUNA_ATIVO_STOCK))
  intColunaQtde = RetornarPrimeiraColuna(Range(RANGE_COLUNA_QTDE_STOCK))
  intColunaSaldoFinal = RetornarPrimeiraColuna(Range(RANGE_COLUNA_SALDO_FINAL_STOCK))
    
  Call ColetarDados(intPrimeiraLinha, intUltimaLinha, _
   intColunaAtivo, intColunaQtde, _
   wsPlanilhaAtual, infoAtivos)
  If (IsArrayEmpty(infoAtivos) = True) Then
    Exit Sub
  End If
  'atualizar cada ativo com a a cotação atual
  strRetorno = AtualizarCotacao(intPrimeiraLinha, intUltimaLinha, _
    intColunaAtivo, intColunaQtde, intColunaSaldoFinal, _
    wsPlanilhaAtual, wsPlanilhaAcoes, infoAtivos)
  strAtivosNaoEncontrados = AnexarAtivo(strAtivosNaoEncontrados, strRetorno)
  
  Exit Sub
ErrorAtualizarCotacaoStock:
  MostrarMsgErro ("AtualizarCotacaoStock")
End Sub

Private Sub AtualizarCotacaoReit(wsPlanilhaAtual As Worksheet, wsPlanilhaAcoes As Worksheet, ByRef strAtivosNaoEncontrados As String)
  ' Sub AtualizarCotacaoReit
  ' atualiza valor da cota final do ativo
  Dim intPrimeiraLinha As Integer, intUltimaLinha As Integer
  Dim intColunaAtivo As Integer, intColunaQtde As Integer
  Dim intColunaSaldoFinal As Integer
  Dim infoAtivos() As infoAtivo
  Dim strRetorno As String
  On Error GoTo ErrorAtualizarCotacaoReit
  intPrimeiraLinha = RetornarPrimeiraLinha(Range(RANGE_COLUNA_ATIVO_REIT))
  intUltimaLinha = RetornarUltimaLinha(Range(RANGE_COLUNA_ATIVO_REIT))
  intColunaAtivo = RetornarPrimeiraColuna(Range(RANGE_COLUNA_ATIVO_REIT))
  intColunaQtde = RetornarPrimeiraColuna(Range(RANGE_COLUNA_QTDE_REIT))
  intColunaSaldoFinal = RetornarPrimeiraColuna(Range(RANGE_COLUNA_SALDO_FINAL_REIT))
    
  Call ColetarDados(intPrimeiraLinha, intUltimaLinha, _
   intColunaAtivo, intColunaQtde, _
   wsPlanilhaAtual, infoAtivos)
  If (IsArrayEmpty(infoAtivos) = True) Then
    Exit Sub
  End If
  'atualizar cada ativo com a a cotação atual
  strRetorno = AtualizarCotacao(intPrimeiraLinha, intUltimaLinha, _
    intColunaAtivo, intColunaQtde, intColunaSaldoFinal, _
    wsPlanilhaAtual, wsPlanilhaAcoes, infoAtivos)
  strAtivosNaoEncontrados = AnexarAtivo(strAtivosNaoEncontrados, strRetorno)
  
  Exit Sub
ErrorAtualizarCotacaoReit:
  MostrarMsgErro ("AtualizarCotacaoReit")
End Sub

Private Sub AtualizarCotacaoTreasuries(wsPlanilhaAtual As Worksheet, wsPlanilhaAcoes As Worksheet, ByRef strAtivosNaoEncontrados As String)
  ' Sub AtualizarCotacaoTreasuries
  ' atualiza valor da cota final do ativo
  Dim intPrimeiraLinha As Integer, intUltimaLinha As Integer
  Dim intColunaAtivo As Integer, intColunaQtde As Integer
  Dim intColunaSaldoFinal As Integer
  Dim infoAtivos() As infoAtivo
  Dim strRetorno As String
  On Error GoTo ErrorAtualizarCotacaoTreasuries
  intPrimeiraLinha = RetornarPrimeiraLinha(Range(RANGE_COLUNA_ATIVO_TREASURY))
  intUltimaLinha = RetornarUltimaLinha(Range(RANGE_COLUNA_ATIVO_TREASURY))
  intColunaAtivo = RetornarPrimeiraColuna(Range(RANGE_COLUNA_ATIVO_TREASURY))
  intColunaQtde = RetornarPrimeiraColuna(Range(RANGE_COLUNA_QTDE_TREASURY))
  intColunaSaldoFinal = RetornarPrimeiraColuna(Range(RANGE_COLUNA_SALDO_FINAL_TREASURY))
    
  Call ColetarDados(intPrimeiraLinha, intUltimaLinha, _
   intColunaAtivo, intColunaQtde, _
   wsPlanilhaAtual, infoAtivos)
  If (IsArrayEmpty(infoAtivos) = True) Then
    Exit Sub
  End If
  'atualizar cada ativo com a a cotação atual
  strRetorno = AtualizarCotacao(intPrimeiraLinha, intUltimaLinha, _
    intColunaAtivo, intColunaQtde, intColunaSaldoFinal, _
    wsPlanilhaAtual, wsPlanilhaAcoes, infoAtivos)
  strAtivosNaoEncontrados = AnexarAtivo(strAtivosNaoEncontrados, strRetorno)
  
  Exit Sub
ErrorAtualizarCotacaoTreasuries:
  MostrarMsgErro ("AtualizarCotacaoTreasuries")
End Sub

Private Sub AtualizarCotacaoOuro(wsPlanilhaAtual As Worksheet, wsPlanilhaAcoes As Worksheet, ByRef strAtivosNaoEncontrados As String)
  ' Sub AtualizarCotacaoOuro
  ' atualiza valor da cota final do ativo
  Dim intPrimeiraLinha As Integer, intUltimaLinha As Integer
  Dim intColunaAtivo As Integer, intColunaQtde As Integer
  Dim intColunaSaldoFinal As Integer
  Dim infoAtivos() As infoAtivo
  Dim strRetorno As String
  On Error GoTo ErrorAtualizarCotacaoOuro
  intPrimeiraLinha = RetornarPrimeiraLinha(Range(RANGE_COLUNA_ATIVO_OURO))
  intUltimaLinha = RetornarUltimaLinha(Range(RANGE_COLUNA_ATIVO_OURO))
  intColunaAtivo = RetornarPrimeiraColuna(Range(RANGE_COLUNA_ATIVO_OURO))
  intColunaQtde = RetornarPrimeiraColuna(Range(RANGE_COLUNA_QTDE_OURO))
  intColunaSaldoFinal = RetornarPrimeiraColuna(Range(RANGE_COLUNA_SALDO_FINAL_OURO))
    
  Call ColetarDados(intPrimeiraLinha, intUltimaLinha, _
   intColunaAtivo, intColunaQtde, _
   wsPlanilhaAtual, infoAtivos)
  If (IsArrayEmpty(infoAtivos) = True) Then
    Exit Sub
  End If
  'atualizar cada ativo com a a cotação atual
  strRetorno = AtualizarCotacao(intPrimeiraLinha, intUltimaLinha, _
    intColunaAtivo, intColunaQtde, intColunaSaldoFinal, _
    wsPlanilhaAtual, wsPlanilhaAcoes, infoAtivos)
  strAtivosNaoEncontrados = AnexarAtivo(strAtivosNaoEncontrados, strRetorno)
  
  Exit Sub
ErrorAtualizarCotacaoOuro:
  MostrarMsgErro ("AtualizarCotacaoOuro")
End Sub

Private Sub AtualizarCotacaoCripto(wsPlanilhaAtual As Worksheet, wsPlanilhaAcoes As Worksheet, ByRef strAtivosNaoEncontrados As String)
  ' Sub AtualizarCotacaoCripto
  ' atualiza valor da cota final do ativo
  Dim intPrimeiraLinha As Integer, intUltimaLinha As Integer
  Dim intColunaAtivo As Integer, intColunaQtde As Integer
  Dim intColunaSaldoFinal As Integer
  Dim infoAtivos() As infoAtivo
  Dim strRetorno As String
  On Error GoTo ErrorAtualizarCotacaoCripto
  intPrimeiraLinha = RetornarPrimeiraLinha(Range(RANGE_COLUNA_ATIVO_CRIPTO))
  intUltimaLinha = RetornarUltimaLinha(Range(RANGE_COLUNA_ATIVO_CRIPTO))
  intColunaAtivo = RetornarPrimeiraColuna(Range(RANGE_COLUNA_ATIVO_CRIPTO))
  intColunaQtde = RetornarPrimeiraColuna(Range(RANGE_COLUNA_QTDE_CRIPTO))
  intColunaSaldoFinal = RetornarPrimeiraColuna(Range(RANGE_COLUNA_SALDO_FINAL_CRIPTO))
    
  Call ColetarDados(intPrimeiraLinha, intUltimaLinha, _
   intColunaAtivo, intColunaQtde, _
   wsPlanilhaAtual, infoAtivos)
  If (IsArrayEmpty(infoAtivos) = True) Then
    Exit Sub
  End If
  'atualizar cada ativo com a a cotação atual
  strRetorno = AtualizarCotacao(intPrimeiraLinha, intUltimaLinha, _
    intColunaAtivo, intColunaQtde, intColunaSaldoFinal, _
    wsPlanilhaAtual, wsPlanilhaAcoes, infoAtivos)
  strAtivosNaoEncontrados = AnexarAtivo(strAtivosNaoEncontrados, strRetorno)
  
  Exit Sub
ErrorAtualizarCotacaoCripto:
  MostrarMsgErro ("AtualizarCotacaoCripto")
End Sub

Private Function AtualizarCotacao(intPrimeiraLinha As Integer, intUltimaLinha As Integer, _
    intColunaAtivo As Integer, intColunaQtde As Integer, intColunaSaldoFinal As Integer, _
    wsPlanilhaAtual As Worksheet, wsPlanilhaOrigem As Worksheet, infoAtivos() As infoAtivo) As String
  ' Sub AtualizarCotacao
  ' Atualiza cada ativo com quantidade maior que zero com a cotação atual
  On Error GoTo ErrorAtualizarCotacao
  Dim intCount As Integer, intLinhaDestino As Integer
  Dim dblCotacaoAtual As Double
  Dim strAtivosNaoEncontrados As String
  strAtivosNaoEncontrados = ""
  For intCount = LBound(infoAtivos) To UBound(infoAtivos)
    Dim infoAtivo As infoAtivo
    infoAtivo = infoAtivos(intCount)
    If (infoAtivo.intQuantidade > 0) Then
      dblCotacaoAtual = BuscarCotacaoAtual(infoAtivo.strAtivo, wsPlanilhaOrigem)
      If (dblCotacaoAtual > 0) Then
        Call GravarValor(intPrimeiraLinha, intUltimaLinha, intColunaAtivo, intColunaQtde, _
                         intColunaSaldoFinal, dblCotacaoAtual, wsPlanilhaAtual, infoAtivo)
      Else
        strAtivosNaoEncontrados = AnexarAtivo(strAtivosNaoEncontrados, infoAtivo.strAtivo)
      End If
    End If
  Next intCount
  AtualizarCotacao = strAtivosNaoEncontrados
  Exit Function
ErrorAtualizarCotacao:
  MostrarMsgErro ("AtualizarCotacao")
End Function

Private Function BuscarCotacaoAtual(strAtivo As String, wsPlanilhaOrigem As Worksheet) As Double
  ' Function BuscarCotacaoAtual
  ' busca cotacao atual do ativo
  On Error GoTo ErrorBuscarCotacaoAtual
  Dim celula As Range
  Dim strTicketSimples As String
  BuscarCotacaoAtual = 0
  
  ' se for ativo do Tesouro Direto
  If InStr(strAtivo, ":") = 0 Then
    With wsPlanilhaOrigem.Range(RANGE_COLUNA_TICKETS_TD)
      Set celula = .Find(What:=strAtivo, LookIn:=xlValues)
      If Not celula Is Nothing Then
        BuscarCotacaoAtual = wsPlanilhaOrigem.Cells(celula.Row, RetornarPrimeiraColuna(Range(RANGE_COLUNA_PRECO_TD))).Value
      Else
        Debug.Print "Cotação do ativo do tesouro [" & strAtivo & "] não encontrada."
      End If
    End With
    GoTo FimCotacaoAtual
  End If
  
  strTicketSimples = Trim(Split(strAtivo, ":")(1))
  ' se for Bitcoin
  If strTicketSimples = "BTC" Then
    BuscarCotacaoAtual = wsPlanilhaOrigem.Range(RANGE_CELULA_BTC).Value
    GoTo FimCotacaoAtual
  End If
  ' se for Etherium
  If strTicketSimples = "ETH" Then
    BuscarCotacaoAtual = wsPlanilhaOrigem.Range(RANGE_CELULA_ETH).Value
    GoTo FimCotacaoAtual
  End If
  
  ' se for ações ou ETF
  With wsPlanilhaOrigem.Range(RANGE_COLUNA_TICKETS)
    Set celula = .Find(What:=strTicketSimples, LookIn:=xlValues)
    If Not celula Is Nothing Then
      BuscarCotacaoAtual = wsPlanilhaOrigem.Cells(celula.Row, RetornarPrimeiraColuna(Range(RANGE_COLUNA_PRECO))).Value
    Else
      Debug.Print "Cotação do ativo [" & strTicketSimples & "] não encontrada."
    End If
  End With
FimCotacaoAtual:
  Exit Function
ErrorBuscarCotacaoAtual:
  MostrarMsgErro ("BuscarCotacaoAtual")
End Function

Private Sub GravarValor(intPrimeiraLinha As Integer, intUltimaLinha As Integer, _
   intColunaAtivo As Integer, _
   intColunaQtde As Integer, _
   intColunaSaldoFinal As Integer, _
   dblCotacaoAtual As Double, _
   wsPlanilhaAtual As Worksheet, infoAtivo As infoAtivo)
  On Error GoTo ErrorGravarValor
  Dim intPosArray As Integer, intCont As Integer
  For intCont = intPrimeiraLinha To intUltimaLinha
     If (wsPlanilhaAtual.Cells(intCont, intColunaAtivo).Value = infoAtivo.strAtivo _
         And wsPlanilhaAtual.Cells(intCont, intColunaQtde) > 0) Then
       wsPlanilhaAtual.Cells(intCont, intColunaSaldoFinal).Value = dblCotacaoAtual
     End If
  Next intCont
  Exit Sub
ErrorGravarValor:
  MostrarMsgErro ("GravarValor")
End Sub

Private Sub ColetarDados(intPrimeiraLinha As Integer, intUltimaLinha As Integer, _
   intColunaAtivo As Integer, intColunaQtde As Integer, _
   wsPlanilhaAtual As Worksheet, ByRef infoAtivos() As infoAtivo)
   On Error GoTo ErrorColetarDados
   Dim intPosArray As Integer, intCont As Integer
   For intCont = intPrimeiraLinha To intUltimaLinha
      If (Not IsEmpty(wsPlanilhaAtual.Cells(intCont, intColunaAtivo))) Then
        Dim intPosAtivo As Integer
        intPosAtivo = GetPosAtivoDoArray(infoAtivos, wsPlanilhaAtual.Cells(intCont, intColunaAtivo).Value)
        If (intPosAtivo < 0) Then
          ' Se ainda não existe a entrada, cria uma nova entrada no array
          Dim infoAtivo As infoAtivo
          infoAtivo.strAtivo = wsPlanilhaAtual.Cells(intCont, intColunaAtivo).Value
          infoAtivo.intQuantidade = wsPlanilhaAtual.Cells(intCont, intColunaQtde).Value
          
          intPosArray = intPosArray + 1
          ReDim Preserve infoAtivos(1 To intPosArray)
          infoAtivos(intPosArray) = infoAtivo
        Else
          ' se já existe, atualiza a qtdade
          infoAtivos(intPosAtivo).intQuantidade = infoAtivos(intPosAtivo).intQuantidade + wsPlanilhaAtual.Cells(intCont, intColunaQtde).Value
        End If
      End If
   Next intCont
   Exit Sub
   
ErrorColetarDados:
  MostrarMsgErro ("ColetarDados")
End Sub

Private Function GetPosAtivoDoArray(infoAtivos() As infoAtivo, strAtivo As String) As Integer
   Dim intCount As Integer
   On Error GoTo ErrorGetPosAtivoDoArray
   If (IsArrayEmpty(infoAtivos) = True) Then
     GetPosAtivoDoArray = -1
     Exit Function
   End If
   For intCount = LBound(infoAtivos) To UBound(infoAtivos)
      Dim infoAtivo As infoAtivo
      infoAtivo = infoAtivos(intCount)
      If (infoAtivo.strAtivo = strAtivo) Then
        GetPosAtivoDoArray = intCount
        Exit Function
      End If
   Next intCount
   GetPosAtivoDoArray = -1
   Exit Function
   
ErrorGetPosAtivoDoArray:
  MostrarMsgErro ("GetPosAtivoDoArray")
End Function


Private Function IsArrayEmpty(infoAtivos() As infoAtivo) As Boolean
   ' Determina se um array contém algum elemento
   Dim lngUBound As Long
   On Error Resume Next
   ' Se o array estiver vazio, um erro ocorrerá quando checar os limites do array
   lngUBound = UBound(infoAtivos)
   If Err.Number <> 0 Then
      IsArrayEmpty = True
   Else
      IsArrayEmpty = False
   End If
End Function

Private Function AnexarAtivo(strAtivos As String, strNovoAtivo As String) As String
  On Error GoTo errorAnexarAtivo
    If (strAtivos = "") Then
      AnexarAtivo = strNovoAtivo
      Exit Function
    End If
    If (strNovoAtivo = "") Then
      AnexarAtivo = strAtivos
      Exit Function
    End If
    AnexarAtivo = strAtivos & ", " & strNovoAtivo
  Exit Function
errorAnexarAtivo:
  MostrarMsgErro ("AnexarAtivo")
End Function
