Option Explicit

' ---------------------------------------------------------------------
' Módulo: Relatorio.bas
' Descrição: Módulo de funções sobre a geração de relatório
' ---------------------------------------------------------------------


Private Const TITULO_RELATORIO As String = "Investimentos"
Private Const TITULO_RETRATO As String = "Retrato do mês"
Private Const MSG_AJUSTANDO_IMPRESSAO As String = "Ajustando área de impressão. Por favor, aguarde..."
Private Const PASTA_COFRE As String = "Cofre"
Private Const SUFIXO_RETRATO As String = "-snapshot"
Private Const EXTENSAO_PDF As String = ".pdf"
Private Const CARACTERES_INVALIDOS_ARQUIVO As String = "\/:*?""<>|"
Private Const LIMITE_COLUNAS_RETRATO As Long = 6

Private Type Cabecalho
  strHeaderEsquerdo As String
  strHeaderCentro As String
  strHeaderDireito As String
End Type

Private Type Rodape
  strFooterEsquerdo As String
  strFooterCentro As String
  strFooterDireito As String
End Type

'
' Sub relatRend        De: 22/05/04
' Cria relatório sobre Rendimentos
'
Public Sub GerarRelatRend()  
  Dim udtCabecalho As Cabecalho
  Dim udtRodape As Rodape
  Dim blnAlterouImpressao As Boolean
  On Error GoTo ErroGerarRelatRend

  If MsgBox("Gostaria de gerar relatório da posição " & vbLf & _
            "de investimentos atual?", vbQuestion + vbYesNo, TITULO_RELATORIO) = vbNo Then
    Exit Sub
  End If
  Application.StatusBar = MSG_AJUSTANDO_IMPRESSAO
  udtCabecalho = RetornarCabecalho
  udtRodape = RetornarRodape
  '============= ajusta configurações de impressão ===========
  Application.ScreenUpdating = False
  blnAlterouImpressao = True
  With ActiveSheet.PageSetup
    .PrintArea = RANGE_AREA_RELATORIO
    .LeftHeader = udtCabecalho.strHeaderEsquerdo
    .CenterHeader = udtCabecalho.strHeaderCentro
    .RightHeader = udtCabecalho.strHeaderDireito
    .LeftFooter = udtRodape.strFooterEsquerdo
    .CenterFooter = udtRodape.strFooterCentro
    .RightFooter = udtRodape.strFooterDireito
    .LeftMargin = Application.CentimetersToPoints(1.9)
    .RightMargin = Application.CentimetersToPoints(1.9)
    .TopMargin = Application.CentimetersToPoints(2.5)
    .BottomMargin = Application.CentimetersToPoints(2.5)
    .HeaderMargin = Application.CentimetersToPoints(1.3)
    .FooterMargin = Application.CentimetersToPoints(1.3)
    .PrintHeadings = False
    .PrintGridlines = False
    .PrintNotes = False
    .CenterHorizontally = True
    .CenterVertically = False
    .Orientation = RetornarOrientacao(ActiveSheet.Range(RANGE_AREA_RELATORIO)) ' Paisagem ou retrato
    .Draft = False
    .PaperSize = xlPaperA4
    .FirstPageNumber = xlAutomatic
    .Order = xlDownThenOver
    .BlackAndWhite = False
    .Zoom = False
    .FitToPagesWide = 1     'força largura de uma página
    .FitToPagesTall = False 'ainda 1 de largura mas ilimitada p/ baixo
  End With
  ActiveWindow.SelectedSheets.PrintPreview

FimGerarRelatRend:
  On Error Resume Next
  If blnAlterouImpressao Then
    ActiveSheet.PageSetup.PrintArea = ""
  End If
  Application.StatusBar = False
  Application.ScreenUpdating = True
  Exit Sub

ErroGerarRelatRend:
  MostrarMsgErro "GerarRelatRend"
  Resume FimGerarRelatRend
End Sub

Private Function RetornarOrientacao(ByVal rgAreaRelatorio As Range) As XlPageOrientation
  If (RetornarUltimaColuna(rgAreaRelatorio) - RetornarPrimeiraColuna(rgAreaRelatorio)) < LIMITE_COLUNAS_RETRATO Then
    RetornarOrientacao = xlPortrait
  Else
    RetornarOrientacao = xlLandscape
  End If
End Function

Private Function RetornarCabecalho() As Cabecalho
  Dim udtCabecalho As Cabecalho
  udtCabecalho.strHeaderEsquerdo = "Posição de " & RetornarNomeMesCompleto(RetornarNumeroMesPlanilha(ActiveSheet.Name))
  udtCabecalho.strHeaderCentro = ""
  udtCabecalho.strHeaderDireito = Format$(Now, "dd/mm/yyyy hh:mm:ss")
  RetornarCabecalho = udtCabecalho
End Function

Private Function RetornarRodape() As Rodape
  Dim udtRodape As Rodape
  Dim strNomePasta As String
  strNomePasta = RetornarNomeBasePasta()
  strNomePasta = UCase$(Left$(strNomePasta, 1)) & LCase$(Mid$(strNomePasta, 2))
  udtRodape.strFooterEsquerdo = "&8" & strNomePasta & vbLf & _
                   "Última atualização em: " & ActiveSheet.Range(RANGE_DATA_POSICAO).Value & vbLf & _
                   Chr$(169) & Year(Now) & _
                   " Propriedade Confidencial de Mauricio Soares"
  udtRodape.strFooterCentro = "Página &P de &N"
  udtRodape.strFooterDireito = "&8" & _
                  "Mês Líquido = diferença entre saldos" & vbLf & _
                  "Mês Real = Mês Líquido - IGPM" & vbLf & _
                  "Outros, fonte: " & Chr$(34) & "HSBC Bank Brasil S.A." & Chr$(34)
  RetornarRodape = udtRodape
End Function

'
' Sub relatRend        De: 31/03/18
' Cria relatório sobre Situação no Mês
'
Public Sub GerarRelatRetrato()  
  Dim strNomeArquivo As String, strPastaDestino As String, strCaminhoArquivo As String
  Dim strMesFechado As String
  Dim blnAlterouImpressao As Boolean
  On Error GoTo ErroGerarRelatRetrato

  ' define o nome do PDF
  strMesFechado = CStr(RetornarRangeNomeado(RANGE_PLAN_FECHADA).Value)
  strNomeArquivo = RetornarNomeBasePasta() & SUFIXO_RETRATO & _
    Format$(Month(DateValue(strMesFechado & " 1")), "00") & EXTENSAO_PDF
  strPastaDestino = ThisWorkbook.Path
  If InStr(strPastaDestino, PASTA_COFRE) > 0 Then
    strPastaDestino = RetornarPastaMeusDocumentos()
  End If
  strCaminhoArquivo = DefinirCaminhoArquivo(strPastaDestino, strNomeArquivo)
  If Len(strCaminhoArquivo) = 0 Then
    ' cancelado pelo usuário
    Exit Sub
  End If
  Debug.Print "Arquivo destino: " & strCaminhoArquivo

  Application.StatusBar = MSG_AJUSTANDO_IMPRESSAO
  Application.ScreenUpdating = False
  blnAlterouImpressao = True
  With ActiveSheet.PageSetup
    .PrintArea = RANGE_RELAT_RETRAT
    .LeftMargin = Application.CentimetersToPoints(0.64)
    .RightMargin = Application.CentimetersToPoints(0.64)
    .TopMargin = Application.CentimetersToPoints(1.91)
    .BottomMargin = Application.CentimetersToPoints(1.91)
    .HeaderMargin = Application.CentimetersToPoints(0.76)
    .FooterMargin = Application.CentimetersToPoints(0.76)
    .PrintHeadings = False
    .PrintGridlines = False
    .PrintNotes = False
    .CenterHorizontally = True
    .CenterVertically = False
    .Orientation = xlPortrait
    .Draft = False
    .PaperSize = xlPaperA4
    .BlackAndWhite = False
    .Zoom = False
    .FitToPagesWide = 1 'força largura de uma página
    .FitToPagesTall = False 'não força altura
  End With
  ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, _
       Filename:=strCaminhoArquivo, _
       Quality:=xlQualityStandard, _
       IncludeDocProperties:=True, _
       IgnorePrintAreas:=False, _
       OpenAfterPublish:=True

FimGerarRelatRetrato:
  On Error Resume Next
  If blnAlterouImpressao Then
    ActiveSheet.PageSetup.PrintArea = ""
  End If
  Application.StatusBar = False
  Application.ScreenUpdating = True
  Exit Sub

ErroGerarRelatRetrato:
  MostrarMsgErro "GerarRelatRetrato"
  Resume FimGerarRelatRetrato
End Sub

'
' Monta o caminho do arquivo; se já existir, pergunta se deve sobrescrever ou renomear.
' Retorna vazio se o usuário cancelar.
'
Private Function DefinirCaminhoArquivo(ByVal strPasta As String, ByVal strNomeArquivo As String) As String  
  Dim strCaminho As String
  Dim varResposta As Variant
  Do
    strCaminho = strPasta & Application.PathSeparator & strNomeArquivo
    If Len(Dir(strCaminho)) = 0 Then
      Exit Do
    End If
    Select Case MsgBox("Arquivo já existe! Clique [Sim] para sobrescrever. Clique [Não] para renomear.", _
                       vbYesNoCancel + vbQuestion, TITULO_RETRATO)
      Case vbYes
        Exit Do
      Case vbNo
        Do
          'Recupera novo nome de arquivo
          varResposta = Application.InputBox("Digite um novo nome de arquivo " & _
            "(irá perguntá-lo novamente se você digitar um nome de arquivo inválido)", TITULO_RETRATO, _
            strNomeArquivo, Type:=2)
          'sai se o usuário quiser
          If VarType(varResposta) = vbBoolean Then
            Exit Function
          End If
          strNomeArquivo = Trim$(CStr(varResposta))
          If Len(strNomeArquivo) = 0 Then
            Exit Function
          End If
        Loop While Not IsNomeArquivoValido(strNomeArquivo)
        If LCase$(Right$(strNomeArquivo, Len(EXTENSAO_PDF))) <> EXTENSAO_PDF Then
          strNomeArquivo = strNomeArquivo & EXTENSAO_PDF
        End If
        ' volta ao início do laço para verificar se o novo nome também já existe
      Case Else
        Exit Function
    End Select
  Loop
  DefinirCaminhoArquivo = strCaminho
End Function

Private Function RetornarPastaMeusDocumentos() As String
  Dim objShell As Object
  Set objShell = CreateObject("WScript.Shell")
  RetornarPastaMeusDocumentos = objShell.SpecialFolders("MyDocuments")
End Function

'
' Nome da pasta de trabalho sem a extensão
'
Private Function RetornarNomeBasePasta() As String  
  Dim strNome As String
  Dim lngPosPonto As Long
  strNome = ThisWorkbook.Name
  lngPosPonto = InStrRev(strNome, ".")
  If lngPosPonto > 1 Then
    strNome = Left$(strNome, lngPosPonto - 1)
  End If
  RetornarNomeBasePasta = strNome
End Function

'
' Determina se um nome de arquivo é válido no Windows (sem caracteres reservados)
'
Private Function IsNomeArquivoValido(ByVal strNomeArquivo As String) As Boolean  
  Dim lngPosicao As Long
  If Len(Trim$(strNomeArquivo)) = 0 Then
    Exit Function
  End If
  For lngPosicao = 1 To Len(CARACTERES_INVALIDOS_ARQUIVO)
    If InStr(strNomeArquivo, Mid$(CARACTERES_INVALIDOS_ARQUIVO, lngPosicao, 1)) > 0 Then
      Exit Function
    End If
  Next lngPosicao
  IsNomeArquivoValido = True
End Function
