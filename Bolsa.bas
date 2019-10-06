' Módulo de funções para baixar cotações da bolsa
Option Explicit
Sub BaixarCotacoes()
  '
  ' Sub BaixarCotacoes    Criado por: MSS  Em: 25.04.16
  ' baixa cotações do Yahoo Finanças
  '
  ' variáveis
  Dim dsDataSheet As Worksheet
  Dim strUrl As String
  Dim intLinha As Integer
  Dim intColunaSimbolo As Integer
  Dim intTamanhoOriginal As Integer
  
  ' principal
  On Error GoTo EndMacro
  CongelarCalculosPlanilha (True)
  
  Set dsDataSheet = ActiveSheet
  If dsDataSheet.ProtectContents Then
    dsDataSheet.Unprotect
  End If
  'Apaga a região onde os dados serão atualizados
  range(RANGE_CELULA_INICIO_QUOTACAO_SIMBOLOATIVO).CurrentRegion.ClearContents
  intLinha = range(RANGE_CELULA_INICIO_QUOTACAO_SIMBOLOATIVO).Row
  intColunaSimbolo = range(RANGE_CELULA_INICIO_QUOTACAO_SIMBOLO).Column
  'Cria uma url no formato <http://download.finance.yahoo.com/d/quotes.csv?s=^BVSP+^GSPC+PETR4.SA+VALE5.SA&f=snd1t1c1ol1ghv&e=.csv>
  strUrl = YAHOO_FINANCE_URL + Cells(intLinha, intColunaSimbolo)
  intLinha = intLinha + 1
  While Cells(intLinha, intColunaSimbolo) <> ""
    strUrl = strUrl + "+" + Cells(intLinha, intColunaSimbolo)
    intLinha = intLinha + 1
  Wend
  strUrl = strUrl + "&f=" + YAHOO_TAG_DADOS + "&e=" + YAHOO_TAG_FORMATO
  'Cria uma QueryTable para conter os dados de retorno da URL
  intTamanhoOriginal = range(RANGE_CELULA_INICIO_QUOTACAO_SIMBOLOATIVO).ColumnWidth
  With ActiveSheet.QueryTables.Add(Connection:="URL;" & strUrl, Destination:=dsDataSheet.range(RANGE_CELULA_INICIO_QUOTACAO_SIMBOLOATIVO))
    .BackgroundQuery = True
    .TablesOnlyFromHTML = False
    .Refresh BackgroundQuery:=False
    .SaveData = True
  End With
  range(RANGE_CELULA_INICIO_QUOTACAO_SIMBOLOATIVO).CurrentRegion.TextToColumns Destination:=range(RANGE_CELULA_INICIO_QUOTACAO_SIMBOLOATIVO), DataType:=xlDelimited, _
    TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
    Semicolon:=False, Comma:=True, Space:=False, other:=False
  Columns(range(RANGE_CELULA_INICIO_QUOTACAO_SIMBOLOATIVO).Column).ColumnWidth = intTamanhoOriginal
    
  dsDataSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
  Set dsDataSheet = Nothing
  
EndMacro:
  CongelarCalculosPlanilha (False)
  Exit Sub
  
ErroBaixarCotacoes:
  MostrarMsgErro ("BaixarCotacoes")
  Resume EndMacro
End Sub
