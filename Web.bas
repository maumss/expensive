Option Explicit

' ---------------------------------------------------------------------
' Módulo: Web.bas
' Descrição: Módulo de funções para buscar dados on line
' ---------------------------------------------------------------------

Private Const TITULO_ATUALIZAR_COTACOES As String = "Atualizar cotações"
Private Const SEPARADOR_ATIVOS As String = ", "

'
' AtualizarDadosWeb    Data: 31/07/23
' Atualiza as cotações das carteiras com base na planilha de ações (tipos de dados vinculados)
'
Public Sub AtualizarDadosWeb()  
  Dim wsPlanilhaAtual As Worksheet, wsPlanilhaAcoes As Worksheet
  Dim dicCotacoes As Object
  Dim lngVisibilidadeOriginal As XlSheetVisibility
  Dim blnOldStatusBar As Boolean, blnCongelou As Boolean, blnConcluiu As Boolean
  Dim strAtivosNaoEncontrados As String
  On Error GoTo ErroAtualizarDadosWeb

  If Not IsPlanilhaAberta(Range(RANGE_SITUAC_PLANILHA)) Then
    Exit Sub
  End If
  'Pede confirmação
  If MsgBox("Você deseja atualizar as cotações das carteiras?", _
        vbYesNo + vbQuestion, TITULO_ATUALIZAR_COTACOES) = vbNo Then
    Exit Sub
  End If
  Set wsPlanilhaAtual = ActiveSheet
  Set wsPlanilhaAcoes = ThisWorkbook.Worksheets(NOME_PLAN_ACOES365)
  lngVisibilidadeOriginal = wsPlanilhaAcoes.Visible
  blnOldStatusBar = Application.DisplayStatusBar
  Application.DisplayStatusBar = True
  CongelarCalculosPlanilha True
  blnCongelou = True
  wsPlanilhaAcoes.Visible = xlSheetVisible

  AtualizarStatus "Buscando cotações das bolsas de valores...", 0
  Application.Calculation = xlCalculationAutomatic
  AtualizarAcoesEMoedas wsPlanilhaAcoes
  Application.Calculation = xlCalculationManual
  Set dicCotacoes = CarregarCotacoesDoDia(wsPlanilhaAcoes)

  AtualizarCotacaoCarteira wsPlanilhaAtual, dicCotacoes, "Ações Brasil", 10, RANGE_COLUNA_ATIVO_ACOES, _
    RANGE_COLUNA_QTDE_ACOES, RANGE_COLUNA_SALDO_FINAL_ACOES, strAtivosNaoEncontrados
  AtualizarCotacaoCarteira wsPlanilhaAtual, dicCotacoes, "FII", 20, RANGE_COLUNA_ATIVO_FII, _
    RANGE_COLUNA_QTDE_FII, RANGE_COLUNA_SALDO_FINAL_FII, strAtivosNaoEncontrados
  AtualizarCotacaoCarteira wsPlanilhaAtual, dicCotacoes, "ETF", 50, RANGE_COLUNA_ATIVO_ETF, _
    RANGE_COLUNA_QTDE_ETF, RANGE_COLUNA_SALDO_FINAL_ETF, strAtivosNaoEncontrados
  AtualizarCotacaoCarteira wsPlanilhaAtual, dicCotacoes, "Ações Internacionais", 60, RANGE_COLUNA_ATIVO_STOCK, _
    RANGE_COLUNA_QTDE_STOCK, RANGE_COLUNA_SALDO_FINAL_STOCK, strAtivosNaoEncontrados
  AtualizarCotacaoCarteira wsPlanilhaAtual, dicCotacoes, "REIT", 70, RANGE_COLUNA_ATIVO_REIT, _
    RANGE_COLUNA_QTDE_REIT, RANGE_COLUNA_SALDO_FINAL_REIT, strAtivosNaoEncontrados
  AtualizarCotacaoCarteira wsPlanilhaAtual, dicCotacoes, "Treasury", 80, RANGE_COLUNA_ATIVO_TREASURY, _
    RANGE_COLUNA_QTDE_TREASURY, RANGE_COLUNA_SALDO_FINAL_TREASURY, strAtivosNaoEncontrados
  AtualizarCotacaoCarteira wsPlanilhaAtual, dicCotacoes, "Ouro", 90, RANGE_COLUNA_ATIVO_OURO, _
    RANGE_COLUNA_QTDE_OURO, RANGE_COLUNA_SALDO_FINAL_OURO, strAtivosNaoEncontrados
  AtualizarCotacaoCarteira wsPlanilhaAtual, dicCotacoes, "Cripto", 100, RANGE_COLUNA_ATIVO_CRIPTO, _
    RANGE_COLUNA_QTDE_CRIPTO, RANGE_COLUNA_SALDO_FINAL_CRIPTO, strAtivosNaoEncontrados

  wsPlanilhaAtual.Range(RANGE_CELULA_DOLAR_FINAL_MES_INDICADORES).Value = _
    wsPlanilhaAcoes.Range(RANGE_CELULA_DOLAR_ACOES365).Value
  Debug.Print "Atualização de cotações concluída."
  blnConcluiu = True

FimAtualizarDadosWeb:
  ' a restauração do ambiente não pode ser interrompida por novos erros
  On Error Resume Next
  If blnCongelou Then
    wsPlanilhaAcoes.Visible = lngVisibilidadeOriginal
    CongelarCalculosPlanilha False
    Application.StatusBar = False
    Application.DisplayStatusBar = blnOldStatusBar
    wsPlanilhaAtual.Activate
  End If
  If blnConcluiu And Len(strAtivosNaoEncontrados) > 0 Then
    Debug.Print "Ativos não encontrados ou desatualizados: " & strAtivosNaoEncontrados
    MsgBox "Não foi possível atualizar a cotação dos ativos abaixo (não encontrados ou sem cotação do dia):" & _
      vbNewLine & vbNewLine & strAtivosNaoEncontrados, vbInformation, TITULO_ATUALIZAR_COTACOES
  End If
  Exit Sub

ErroAtualizarDadosWeb:
  MostrarMsgErro "AtualizarDadosWeb"
  Resume FimAtualizarDadosWeb
End Sub

Private Sub AtualizarStatus(ByVal strMensagem As String, ByVal intPercentual As Integer)
  Application.StatusBar = strMensagem & " " & intPercentual & "% completo."
  Debug.Print strMensagem
End Sub

'
' Uma falha aqui não interrompe o processo: as cotações já carregadas ainda podem ser do dia
'
Private Sub AtualizarAcoesEMoedas(ByVal wsAcoes As Worksheet)  
  On Error GoTo ErroAtualizarAcoesEMoedas
  Debug.Print "Atualizando ações e moedas..."
  wsAcoes.Range(RANGE_TAB_ACOES_ACOES365).RefreshLinkedDataType
  wsAcoes.Range(RANGE_TAB_MOEDAS_ACOES365).RefreshLinkedDataType
  Debug.Print "Concluído atualização de ações e moedas."
  Exit Sub

ErroAtualizarAcoesEMoedas:
  MostrarMsgErro "AtualizarAcoesEMoedas"
End Sub

'
' Monta um dicionário ticket -> cotação a partir da planilha de ações.
' Cotações que não são do dia ficam com valor zero.
'
Private Function CarregarCotacoesDoDia(ByVal wsAcoes As Worksheet) As Object  
  Dim dicCotacoes As Object
  Dim varTickets As Variant, varPrecos As Variant, varDatas As Variant
  Dim lngLinha As Long
  Dim strTicket As String

  Set dicCotacoes = CreateObject("Scripting.Dictionary")
  dicCotacoes.CompareMode = vbTextCompare
  varTickets = wsAcoes.Range(RANGE_COLUNA_TICKETS_ACOES365).Value
  varPrecos = wsAcoes.Range(RANGE_COLUNA_PRECO_ACOES365).Value
  varDatas = wsAcoes.Range(RANGE_COLUNA_DATA_ATUAL_ACOES365).Value
  For lngLinha = 1 To UBound(varTickets, 1)
    If Not IsError(varTickets(lngLinha, 1)) Then
      strTicket = Trim$(CStr(varTickets(lngLinha, 1)))
      ' vale a primeira ocorrência do ticket
      If Len(strTicket) > 0 And Not dicCotacoes.Exists(strTicket) Then
        dicCotacoes.Add strTicket, RetornarCotacaoDoDia(varDatas(lngLinha, 1), varPrecos(lngLinha, 1))
      End If
    End If
  Next lngLinha
  Set CarregarCotacoesDoDia = dicCotacoes
End Function

Private Function RetornarCotacaoDoDia(ByVal varData As Variant, ByVal varPreco As Variant) As Double
  If IsError(varData) Or IsError(varPreco) Then
    Exit Function
  End If
  If Not IsDate(varData) Or Not IsNumeric(varPreco) Then
    Exit Function
  End If
  If DateValue(CDate(varData)) = Date Then
    RetornarCotacaoDoDia = CDbl(varPreco)
  End If
End Function

'
' Atualiza cada ativo da carteira com quantidade maior que zero com a cotação atual
'
Private Sub AtualizarCotacaoCarteira(ByVal wsPlanilhaAtual As Worksheet, ByVal dicCotacoes As Object, _
    ByVal strNomeCarteira As String, ByVal intPercentual As Integer, ByVal strRangeAtivo As String, _
    ByVal strRangeQtde As String, ByVal strRangeSaldoFinal As String, ByRef strAtivosNaoEncontrados As String)
  Dim udtAtivos() As InfoAtivo
  Dim lngTotalAtivos As Long, lngIndice As Long
  Dim dblCotacaoAtual As Double
  Dim strTicket As String
  On Error GoTo ErroAtualizarCotacaoCarteira

  AtualizarStatus "Transferindo dados de " & strNomeCarteira & "...", intPercentual
  lngTotalAtivos = ColetarAtivos(wsPlanilhaAtual, strRangeAtivo, strRangeQtde, strRangeSaldoFinal, udtAtivos)
  For lngIndice = 1 To lngTotalAtivos
    If udtAtivos(lngIndice).dblQuantidade >= QUANTIDADE_MINIMA_ATIVO Then
      dblCotacaoAtual = 0
      strTicket = Trim$(udtAtivos(lngIndice).strAtivo)
      If dicCotacoes.Exists(strTicket) Then
        dblCotacaoAtual = dicCotacoes(strTicket)
      End If
      If dblCotacaoAtual > 0 Then
        GravarCotacao wsPlanilhaAtual, strRangeAtivo, strRangeQtde, strRangeSaldoFinal, _
          udtAtivos(lngIndice).strAtivo, dblCotacaoAtual
      Else
        strAtivosNaoEncontrados = AnexarAtivo(strAtivosNaoEncontrados, udtAtivos(lngIndice).strAtivo)
      End If
    End If
  Next lngIndice
  Exit Sub

ErroAtualizarCotacaoCarteira:
  MostrarMsgErro "AtualizarCotacaoCarteira (" & strNomeCarteira & ")"
End Sub

'
' Grava a cotação em todas as linhas do ativo com quantidade maior que zero
'
Private Sub GravarCotacao(ByVal wsPlanilhaAtual As Worksheet, ByVal strRangeAtivo As String, _
    ByVal strRangeQtde As String, ByVal strRangeSaldoFinal As String, _
    ByVal strAtivo As String, ByVal dblCotacaoAtual As Double)  
  Dim varAtivos As Variant, varQtdes As Variant
  Dim rgSaldoFinal As Range
  Dim lngLinha As Long
  varAtivos = wsPlanilhaAtual.Range(strRangeAtivo).Value
  varQtdes = wsPlanilhaAtual.Range(strRangeQtde).Value
  Set rgSaldoFinal = wsPlanilhaAtual.Range(strRangeSaldoFinal)
  For lngLinha = 1 To UBound(varAtivos, 1)
    If Not IsError(varAtivos(lngLinha, 1)) Then
      If CStr(varAtivos(lngLinha, 1)) = strAtivo And RetornarValorNumerico(varQtdes(lngLinha, 1)) > 0 Then
        rgSaldoFinal.Cells(lngLinha, 1).Value = dblCotacaoAtual
      End If
    End If
  Next lngLinha
End Sub

Private Function AnexarAtivo(ByVal strAtivos As String, ByVal strNovoAtivo As String) As String
  If Len(strAtivos) = 0 Then
    AnexarAtivo = strNovoAtivo
  ElseIf Len(strNovoAtivo) = 0 Then
    AnexarAtivo = strAtivos
  Else
    AnexarAtivo = strAtivos & SEPARADOR_ATIVOS & strNovoAtivo
  End If
End Function
