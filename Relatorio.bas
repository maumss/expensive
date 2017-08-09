Attribute VB_Name = "Relatorio"
' Módulo de funções sobre a geração de relatório
Option Explicit

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

Sub GerarRelatRend()
  '
  ' Sub relatRend        De: 22/05/04
  ' Cria relatório sobre Rendimentos
  '
  On Error GoTo erroRelatRend
    
  If MsgBox("Gostaria de gerar relatório da posição " & vbLf & _
            "de investimentos atual?", vbQuestion + vbYesNo, "Investimentos") = vbNo Then
    Exit Sub
  End If
  Application.StatusBar = "Ajustando área de impressão. Por favor, aguarde..."
  '============= define cabeçalho ==================
  Dim udtCabecalho As Cabecalho
  udtCabecalho = RetornarCabecalho
  '============= define rodapé =====================
  Dim udtRodape As Rodape
  udtRodape = RetornarRodape
  '============= ajusta configurações de impressão ===========
  Application.ScreenUpdating = False
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
      .Orientation = RetornarOrientacao(Range(RANGE_AREA_RELATORIO)) ' Paisagem ou retrato
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
  
fimRelat:
  ActiveSheet.PageSetup.PrintArea = ""
  Application.StatusBar = "Pronto"
  Application.ScreenUpdating = True
  Exit Sub
erroRelatRend:

  MostrarMsgErro ("GerarRelatRend")
  Resume fimRelat
End Sub

Private Function RetornarOrientacao(rgAreaRelatorio As Range) As Integer
  Const XL_PORTRAIT As Integer = 1 ' retrato
  Const XL_LANDSCAPE As Integer = 2 ' paisagem
  Dim intPageSettg, intLastCol As Integer
  intLastCol = RetornarUltimaLinha(rgAreaRelatorio) - RetornarPrimeiraLinha(rgAreaRelatorio)
  If intLastCol < 6 Then
    RetornarOrientacao = XL_PORTRAIT
  Else
    RetornarOrientacao = XL_LANDSCAPE
  End If
End Function

Private Function RetornarCabecalho() As Cabecalho
  Dim udtCabecalho As Cabecalho
  Dim strDeMes As String
  strDeMes = RetornarMesPlanilha
  udtCabecalho.strHeaderEsquerdo = "Posição de " & Trim(strDeMes)
  udtCabecalho.strHeaderCentro = ""
  udtCabecalho.strHeaderDireito = Application.Text(Now(), "dd/mm/yyyy HH:mm:ss")
  RetornarCabecalho = udtCabecalho
End Function

Private Function RetornarMesPlanilha() As String
  Select Case ActiveSheet.Name
     Case "Jan."
       RetornarMesPlanilha = "Janeiro"
     Case "Fev."
       RetornarMesPlanilha = "Fevereiro"
     Case "Mar."
       RetornarMesPlanilha = "Março"
     Case "Abril"
       RetornarMesPlanilha = "Abril"
     Case "Mai."
       RetornarMesPlanilha = "Maio"
     Case "Jun."
       RetornarMesPlanilha = "Junho"
     Case "Jul."
       RetornarMesPlanilha = "Julho"
     Case "Ago."
       RetornarMesPlanilha = "Agosto"
     Case "Set."
       RetornarMesPlanilha = "Setembro"
     Case "Out."
       RetornarMesPlanilha = "Outubro"
     Case "Nov."
       RetornarMesPlanilha = "Novembro"
     Case Else
       RetornarMesPlanilha = "Dezembro"
  End Select
End Function

Private Function RetornarRodape() As Rodape
  Dim udtRodape As Rodape
  Dim strNmPlan As String
  Dim intPos As Integer
  strNmPlan = ThisWorkbook.Name
  intPos = InStr(strNmPlan, ".")
  strNmPlan = Left(strNmPlan, (intPos - 1))
  strNmPlan = UCase(Left(strNmPlan, 1)) & LCase(Mid(strNmPlan, 2, (Len(strNmPlan) - 1)))
  udtRodape.strFooterEsquerdo = "&8" & strNmPlan & Chr(10) & _
                   "Última atualização em: " & Range(RANGE_DATA_POSICAO).Value & Chr(10) & _
                   Chr(169) & Year(Now()) & _
                   " Propriedade Confidencial de Mauricio Soares"
  udtRodape.strFooterCentro = "Página &P de &N"
  udtRodape.strFooterDireito = "&8" & _
                  "Mês Líquido = diferença entre saldos" & Chr(10) & _
                  "Mês Real = Mês Líquido - IGPM" & Chr(10) & _
                  "Outros, fonte: " & Chr(34) & "HSBC Bank Brasil S.A." & Chr(34)
  RetornarRodape = udtRodape
End Function


