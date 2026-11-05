Option Explicit

' --------------------------------------------------------------------------------
' Módulo: Ordenacao.bas
' Descrição: Módulo de ordenamento das tabelas de movimentos, cartões e carteiras
' --------------------------------------------------------------------------------

' Data usada temporariamente para manter no topo os ativos sem compra/venda no mês
Private Const DATA_FALSA As Date = #12/31/1980#

'
' OrdenarMovimentos Macro
' Ordena tabelas de Movimentos e Cartões em ordem crescente de data.
'
' Atalho do teclado: Ctrl+o
'
Public Sub OrdenarMovimentos()
  Dim wsPlanilha As Worksheet
  Dim blnCongelou As Boolean
  On Error GoTo ErroOrdenarMovimentos
  If Not IsPlanilhaAberta(Range(RANGE_SITUAC_PLANILHA)) Then
    Exit Sub
  End If
  Set wsPlanilha = ActiveSheet
  Debug.Print "Ordenando movimentos..."
  CongelarCalculosPlanilha True
  blnCongelou = True
  OrdenarIntervalo wsPlanilha, RANGE_COLUNA_DATA_MOVIMENTACAO, RANGE_TAB_MOVIMENTACAO
  OrdenarIntervalo wsPlanilha, RANGE_COLUNA_DATA_CARTOES, RANGE_TAB_CARTOES

FimOrdenarMovimentos:
  On Error Resume Next
  If blnCongelou Then
    CongelarCalculosPlanilha False
  End If
  ' seleciona a coluna de datas
  If Not wsPlanilha Is Nothing Then
    RetornarUltimaCelulaMovimentacoes(wsPlanilha).Select
  End If
  Debug.Print "Ordenação completada."
  Exit Sub

ErroOrdenarMovimentos:
  MostrarMsgErro "OrdenarMovimentos"
  Resume FimOrdenarMovimentos
End Sub

'
' Ordena todas as carteiras da planilha ativa
'
Public Sub OrdenarCarteirasPlanilhaAtiva()
  On Error GoTo ErroOrdenarCarteirasPlanilhaAtiva
  OrdenarCarteiras ActiveSheet
  Exit Sub

ErroOrdenarCarteirasPlanilhaAtiva:
  MostrarMsgErro "OrdenarCarteirasPlanilhaAtiva"
End Sub

'
' Ordena as carteiras de Ações, Fii, Tesouro Direto, Tesouro Selic, ETF, Stock, Reit,
' Treasury, Ouro e Cripto em ordem crescente de ticket,
' mantendo agrupados por data os ativos com compra ou venda no mês.
'
Public Sub OrdenarCarteiras(ByVal wsPlanilha As Worksheet)
  Dim blnCongelou As Boolean
  On Error GoTo ErroOrdenarCarteiras
  If Not IsPlanilhaAberta(wsPlanilha.Range(RANGE_SITUAC_PLANILHA)) Then
    Exit Sub
  End If
  CongelarCalculosPlanilha True
  blnCongelou = True
  OrdenarCarteira wsPlanilha, "ações", RANGE_COLUNA_DATA_ACOES, RANGE_COLUNA_ATIVO_ACOES, RANGE_TAB_ACOES
  OrdenarCarteira wsPlanilha, "fii", RANGE_COLUNA_DATA_FII, RANGE_COLUNA_ATIVO_FII, RANGE_TAB_FII
  OrdenarCarteira wsPlanilha, "tesouro direto", RANGE_COLUNA_DATA_TESOURO_DIRETO, RANGE_COLUNA_ATIVO_TESOURO_DIRETO, _
    RANGE_TAB_TESOURO_DIRETO
  OrdenarCarteira wsPlanilha, "tesouro selic", RANGE_COLUNA_DATA_TESOURO_SELIC, RANGE_COLUNA_ATIVO_TESOURO_SELIC, _
    RANGE_TAB_TESOURO_SELIC
  OrdenarCarteira wsPlanilha, "etf", RANGE_COLUNA_DATA_ETF, RANGE_COLUNA_ATIVO_ETF, RANGE_TAB_ETF
  OrdenarCarteira wsPlanilha, "stock", RANGE_COLUNA_DATA_STOCK, RANGE_COLUNA_ATIVO_STOCK, RANGE_TAB_STOCK
  OrdenarCarteira wsPlanilha, "reit", RANGE_COLUNA_DATA_REIT, RANGE_COLUNA_ATIVO_REIT, RANGE_TAB_REIT
  OrdenarCarteira wsPlanilha, "treasury", RANGE_COLUNA_DATA_TREASURY, RANGE_COLUNA_ATIVO_TREASURY, RANGE_TAB_TREASURY
  OrdenarCarteira wsPlanilha, "ouro", RANGE_COLUNA_DATA_OURO, RANGE_COLUNA_ATIVO_OURO, RANGE_TAB_OURO
  OrdenarCarteira wsPlanilha, "cripto", RANGE_COLUNA_DATA_CRIPTO, RANGE_COLUNA_ATIVO_CRIPTO, RANGE_TAB_CRIPTO
  ' a ordenação deixa selecionada a última tabela ordenada
  ReposicionarSelecao wsPlanilha

FimOrdenarCarteiras:
  If blnCongelou Then
    CongelarCalculosPlanilha False
  End If
  Exit Sub

ErroOrdenarCarteiras:
  MostrarMsgErro "OrdenarCarteiras"
  Resume FimOrdenarCarteiras
End Sub

'
' Seleciona a primeira linha livre das movimentações, mesmo que a planilha não seja a ativa
'
Private Sub ReposicionarSelecao(ByVal wsPlanilha As Worksheet)
  Dim objPlanilhaAtiva As Object
  ' ajuste apenas visual: não deve interromper a ordenação já concluída
  On Error Resume Next
  Set objPlanilhaAtiva = ActiveSheet
  If Not objPlanilhaAtiva Is wsPlanilha Then
    wsPlanilha.Activate
  End If
  RetornarUltimaCelulaMovimentacoes(wsPlanilha).Select
  If Not objPlanilhaAtiva Is wsPlanilha Then
    objPlanilhaAtiva.Activate
  End If
End Sub

Private Sub OrdenarCarteira(ByVal wsPlanilha As Worksheet, ByVal strNomeCarteira As String, _
    ByVal strRangeData As String, ByVal strRangeAtivo As String, ByVal strRangeTabela As String)
  Dim rgData As Range
  Dim blnHasCompraOuVenda As Boolean
  On Error GoTo ErroOrdenarCarteira
  Debug.Print "Ordenando " & strNomeCarteira & "..."
  Set rgData = wsPlanilha.Range(strRangeData)
  blnHasCompraOuVenda = HasCompraOuVenda(rgData)
  If blnHasCompraOuVenda Then
    InserirDataFalsa rgData
  End If
  OrdenarIntervalo wsPlanilha, strRangeAtivo, strRangeTabela
  If blnHasCompraOuVenda Then
    OrdenarIntervalo wsPlanilha, strRangeData, strRangeTabela
    RestaurarDataFalsa rgData
  End If
  Exit Sub

ErroOrdenarCarteira:
  MostrarMsgErro "OrdenarCarteira (" & strNomeCarteira & ")"
  Resume RestaurarCarteira

RestaurarCarteira:
  ' não deixa a data falsa na planilha caso a ordenação tenha sido interrompida
  On Error Resume Next
  If blnHasCompraOuVenda Then
    RestaurarDataFalsa rgData
  End If
End Sub

'
' Ordena a tabela em ordem crescente pela coluna chave (as tabelas não possuem cabeçalho)
'
Private Sub OrdenarIntervalo(ByVal wsPlanilha As Worksheet, ByVal strRangeChave As String, _
    ByVal strRangeTabela As String)  
  With wsPlanilha.Sort
    .SortFields.Clear
    .SortFields.Add Key:=wsPlanilha.Range(strRangeChave), _
      SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    .SetRange wsPlanilha.Range(strRangeTabela)
    .Header = xlNo
    .MatchCase = False
    .Orientation = xlTopToBottom
    .SortMethod = xlPinYin
    .Apply
  End With
End Sub

Private Function IsValorEmBranco(ByVal varValor As Variant) As Boolean
  If IsError(varValor) Then
    Exit Function
  End If
  IsValorEmBranco = (Len(Trim$(CStr(varValor))) = 0)
End Function

Private Function HasCompraOuVenda(ByVal rgData As Range) As Boolean
  Dim varDatas As Variant
  Dim lngLinha As Long
  varDatas = rgData.Value
  For lngLinha = 1 To UBound(varDatas, 1)
    If Not IsValorEmBranco(varDatas(lngLinha, 1)) Then
      HasCompraOuVenda = True
      Exit Function
    End If
  Next lngLinha
End Function

'
' Preenche com a data falsa as linhas com ativo e sem data
'
Private Sub InserirDataFalsa(ByVal rgData As Range)  
  Dim varDatas As Variant, varAtivos As Variant
  Dim lngLinha As Long
  varDatas = rgData.Value
  varAtivos = rgData.Offset(0, 1).Value
  For lngLinha = 1 To UBound(varDatas, 1)
    If IsValorEmBranco(varDatas(lngLinha, 1)) And Not IsValorEmBranco(varAtivos(lngLinha, 1)) Then
      rgData.Cells(lngLinha, 1).Value = DATA_FALSA
    End If
  Next lngLinha
End Sub

'
' Remove a data falsa inserida antes da ordenação
'
Private Sub RestaurarDataFalsa(ByVal rgData As Range)  
  Dim varDatas As Variant
  Dim lngLinha As Long
  varDatas = rgData.Value
  For lngLinha = 1 To UBound(varDatas, 1)
    If VarType(varDatas(lngLinha, 1)) = vbDate Then
      If varDatas(lngLinha, 1) = DATA_FALSA Then
        rgData.Cells(lngLinha, 1).ClearContents
      End If
    End If
  Next lngLinha
End Sub
