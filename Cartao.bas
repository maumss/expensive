Option Explicit

' -------------------------------------------------------------------------
' Módulo: Cartao.bas
' Descrição: Módulo de transferência de lançamentos de cartão p/ outro mês
' -------------------------------------------------------------------------

Private Const TITULO_MOVER_LANCAMENTO As String = "Mover lançamento"
Private Const TOTAL_COLUNAS_TEXTO_CARTAO As Integer = 4
Private Const DESLOCAMENTO_COLUNA_VALOR_CARTAO As Integer = 4

Private Type MovimentoCartao
  datDataCartao As Date
  strDescCartao As String
  strTipoCartao As String
  strNomeCartao As String
  dblValorCartao As Double
  strFormulaValorCartao As String
End Type

'
' MoveCartao Macro
' Move um lançamento de cartão de um mês para o próximo.
'
' Atalho do teclado: Ctrl+m
'
Public Sub MoverCartao()
  Dim wsPlanilhaAtual As Worksheet, wsProxPlanilha As Worksheet
  Dim udtMovimentoCartao As MovimentoCartao
  Dim lngLinhaAtual As Long, lngColunaInicioCartao As Long
  Dim blnCongelou As Boolean
  On Error GoTo ErroMoverCartao

  If Not TypeOf Selection Is Range Then
    Exit Sub
  End If
  Set wsPlanilhaAtual = ActiveSheet
  lngLinhaAtual = Selection.Cells(1).Row
  If IsMoverInvalido(wsPlanilhaAtual, lngLinhaAtual) Then
    Exit Sub
  End If
  'Pede confirmação
  If MsgBox("Você deseja mover o lançamento para o próximo mês?", _
        vbYesNo + vbQuestion, TITULO_MOVER_LANCAMENTO) = vbNo Then
    Exit Sub
  End If

  'Começa o processo de mover o lançamento
  CongelarCalculosPlanilha True
  blnCongelou = True
  Set wsProxPlanilha = RetornarProximaPlanilha(wsPlanilhaAtual)
  lngColunaInicioCartao = wsPlanilhaAtual.Range(RANGE_PRIMEIRA_DATA_CARTOES).Column
  udtMovimentoCartao = RetornarDadosMovCartao(wsPlanilhaAtual, lngLinhaAtual, lngColunaInicioCartao)
  ' Grava primeiro no destino: se falhar, o lançamento original é preservado.
  ' Habilita os eventos p/ marcar que houve mudança
  Application.EnableEvents = True
  JogarValoresCartaoNaPlanilha wsProxPlanilha, udtMovimentoCartao
  Application.EnableEvents = False
  ' Apaga a linha original
  wsPlanilhaAtual.Range(wsPlanilhaAtual.Cells(lngLinhaAtual, lngColunaInicioCartao), _
    wsPlanilhaAtual.Cells(lngLinhaAtual, wsPlanilhaAtual.Range(RANGE_ULTIMO_VALOR_CARTAO).Column)).ClearContents
  wsPlanilhaAtual.Cells(lngLinhaAtual, lngColunaInicioCartao).Select

FimMoverCartao:
  If blnCongelou Then
    CongelarCalculosPlanilha False
  End If
  Exit Sub

ErroMoverCartao:
  MostrarMsgErro "MoverCartao"
  Resume FimMoverCartao
End Sub

Private Function IsMoverInvalido(ByVal wsPlanilha As Worksheet, ByVal lngLinhaAtual As Long) As Boolean
  Dim wsProxPlanilha As Worksheet
  IsMoverInvalido = True
  'Testa se é planilha mensal válida e se está aberta
  If Not IsPlanilhaAberta(wsPlanilha.Range(RANGE_SITUAC_PLANILHA)) Then
    Exit Function
  End If
  'Testa se está dentro do lançamento do cartão
  If Application.Intersect(Selection, wsPlanilha.Range(RANGE_TAB_CARTOES)) Is Nothing Then
    Exit Function
  End If
  'Testa se já está na planilha de Dezembro
  Set wsProxPlanilha = RetornarProximaPlanilha(wsPlanilha)
  If wsProxPlanilha Is Nothing Then
    MsgBox "Não existe planilha posterior a essa para mover valores", vbCritical, TITULO_MOVER_LANCAMENTO
    Exit Function
  End If
  'Testa se a planilha destino está aberta
  If Not IsPlanilhaAberta(wsProxPlanilha.Range(RANGE_SITUAC_PLANILHA)) Then
    MsgBox "A planilha destino está Fechada para alterações", vbCritical, TITULO_MOVER_LANCAMENTO
    Exit Function
  End If
  'Verifica se existe algo para mover
  If IsEmpty(wsPlanilha.Cells(lngLinhaAtual, wsPlanilha.Range(RANGE_PRIMEIRA_DATA_CARTOES).Column)) Then
    MsgBox "Você não se posicionou em uma célula com dados", vbCritical, TITULO_MOVER_LANCAMENTO
    Exit Function
  End If
  IsMoverInvalido = False
End Function

Private Function RetornarDadosMovCartao(ByVal wsPlanilha As Worksheet, ByVal lngLinhaAtual As Long, _
    ByVal lngColunaInicioCartao As Long) As MovimentoCartao
  Dim udtMovimentoCartao As MovimentoCartao
  Dim varLinha As Variant
  Dim rgValor As Range
  ' leitura em bloco das colunas de texto: data, descrição, tipo e cartão
  varLinha = wsPlanilha.Cells(lngLinhaAtual, lngColunaInicioCartao).Resize(1, TOTAL_COLUNAS_TEXTO_CARTAO).Value
  udtMovimentoCartao.datDataCartao = varLinha(1, 1)
  udtMovimentoCartao.strDescCartao = CStr(varLinha(1, 2))
  udtMovimentoCartao.strTipoCartao = CStr(varLinha(1, 3))
  udtMovimentoCartao.strNomeCartao = CStr(varLinha(1, 4))
  Set rgValor = wsPlanilha.Cells(lngLinhaAtual, lngColunaInicioCartao + DESLOCAMENTO_COLUNA_VALOR_CARTAO)
  If rgValor.HasFormula Then
    udtMovimentoCartao.strFormulaValorCartao = rgValor.Formula
  Else
    udtMovimentoCartao.dblValorCartao = RetornarValorNumerico(rgValor.Value)
  End If
  RetornarDadosMovCartao = udtMovimentoCartao
End Function

'
' Joga valores na próxima planilha, na primeira linha livre da tabela de cartões
'
Private Sub JogarValoresCartaoNaPlanilha(ByVal wsPlanilhaDestino As Worksheet, ByRef udtMovimentoCartao As MovimentoCartao)
  Dim lngLinhaDestino As Long, lngColunaDestino As Long
  lngLinhaDestino = RetornarLinhaLivreCartao(wsPlanilhaDestino)
  lngColunaDestino = wsPlanilhaDestino.Range(RANGE_PRIMEIRA_DATA_CARTOES).Column
  With udtMovimentoCartao
    wsPlanilhaDestino.Cells(lngLinhaDestino, lngColunaDestino).Resize(1, TOTAL_COLUNAS_TEXTO_CARTAO).Value = _
      Array(.datDataCartao, .strDescCartao, .strTipoCartao, .strNomeCartao)
    If Len(.strFormulaValorCartao) > 0 Then
      wsPlanilhaDestino.Cells(lngLinhaDestino, lngColunaDestino + DESLOCAMENTO_COLUNA_VALOR_CARTAO).Formula = .strFormulaValorCartao
    Else
      wsPlanilhaDestino.Cells(lngLinhaDestino, lngColunaDestino + DESLOCAMENTO_COLUNA_VALOR_CARTAO).Value = .dblValorCartao
    End If
  End With
End Sub
