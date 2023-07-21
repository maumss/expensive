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
     Case "Jan"
       RetornarMesPlanilha = "Janeiro"
     Case "Fev"
       RetornarMesPlanilha = "Fevereiro"
     Case "Mar"
       RetornarMesPlanilha = "Março"
     Case "Abr"
       RetornarMesPlanilha = "Abril"
     Case "Mai"
       RetornarMesPlanilha = "Maio"
     Case "Jun"
       RetornarMesPlanilha = "Junho"
     Case "Jul"
       RetornarMesPlanilha = "Julho"
     Case "Ago"
       RetornarMesPlanilha = "Agosto"
     Case "Set"
       RetornarMesPlanilha = "Setembro"
     Case "Out"
       RetornarMesPlanilha = "Outubro"
     Case "Nov"
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

Sub GerarRelatRetrato()

  '
  ' Sub relatRend        De: 31/03/18
  ' Cria relatório sobre Situação no Mês
  '
  On Error GoTo erroRelatRetrato
  
  Dim fileName As String
  Dim dirFile As String
  Dim uniqueName As Boolean
  Dim userAnswer As VbMsgBoxResult
  Dim monthNumber As String
  Const XL_PORTRAIT As Integer = 1 ' retrato
  ' define o nome do PDF
  'MsgBox RetornarFileName()
  monthNumber = Month(DateValue(Range(RANGE_PLAN_FECHADA).Value & " 1"))
  If (Len(monthNumber) < 2) Then
    monthNumber = "0" & monthNumber
  End If
  fileName = RetornarFileName() & "-" & "snapshot" & monthNumber & ".pdf"
  dirFile = RetornarCurrentFolder() & fileName
  Do While uniqueName = False
    If Len(Dir(dirFile)) <> 0 Then
      userAnswer = MsgBox("Arquivo já existe! Click " & _
        "[Sim] para sobreescrever. Clique [Não] para Renomear.", vbYesNoCancel)
      If userAnswer = vbYes Then
        uniqueName = True
      ElseIf userAnswer = vbNo Then
        Do
          'Recupera novo nome de arquivo
          fileName = Application.InputBox("Digite um novo nome de arquivo " & _
            "(irá perguntá-lo novamente se você digitar um nome de arquivo inválido)", , _
            fileName, Type:=2)
          'sai se o usuário quiser
          If fileName = "False" Or fileName = "" Then
            Exit Sub
          End If
        Loop While ValidFileName(fileName) = False
        uniqueName = True
       Else
         Exit Sub 'cancela
       End If
     Else
       uniqueName = True
    End If
  Loop
  
  Application.StatusBar = "Ajustando área de impressão. Por favor, aguarde..."
  Application.ScreenUpdating = False
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
    .Orientation = XL_PORTRAIT
    .Draft = False
    .PaperSize = xlPaperA4
    .BlackAndWhite = False
    .Zoom = False
    .FitToPagesWide = 1 'força largura de uma página
    .FitToPagesTall = False 'não força altura
   End With
   ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, _
        fileName:=fileName, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=True
fimRetrato:
  ActiveSheet.PageSetup.PrintArea = ""
  Application.StatusBar = "Pronto"
  Application.ScreenUpdating = True
  Exit Sub
  
erroRelatRetrato:
  MostrarMsgErro ("GerarRelatRetrato")
  Resume fimRetrato
End Sub

Private Function RetornarCurrentFolder() As String
    RetornarCurrentFolder = ActiveWorkbook.Path & "\"
End Function

Private Function RetornarFileName() As String
  Dim fileName As String
  Dim myPath As String
  Dim uniqueName As Boolean

  myPath = ActiveWorkbook.FullName
  fileName = Mid(myPath, InStrRev(myPath, "\") + 1, _
    InStrRev(myPath, ".") - InStrRev(myPath, "\") - 1)
  RetornarFileName = fileName
End Function

Private Function ValidFileName(fileName As String) As Boolean
  'Determina se um dado nome de arquivo excel é válido

  Dim tempPath As String
  Dim wb As Workbook

  'Determina a pasta onde arquivos temporários são gravados
  tempPath = Environ("TEMP")

  'Cria um arquivo XLSM temporário file (XLSM no caso de ter macros)
  On Error GoTo InvalidFileName
  CongelarCalculosPlanilha (True)
  Set wb = Workbooks.Add
  wb.SaveAs tempPath & "\" & fileName & ".xlsm", xlExcel8
  
  On Error Resume Next
  'Fecha o arquivo temporário
  wb.Close (False)
  'Apaga o arquivo temporário
  Kill tempPath & "\" & fileName & ".xlsm"
  CongelarCalculosPlanilha (False)
  'O nome do arquivo é válido
  ValidFileName = True
  Exit Function

  'ERROR HANDLERS
InvalidFileName:
  wb.Close (False)
  CongelarCalculosPlanilha (False)
  'O nome do arquivo é inválido
  ValidFileName = False
End Function
