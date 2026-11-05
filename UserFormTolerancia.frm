VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormTolerancia 
   Caption         =   "Toler�ncia ao Risco"
   ClientHeight    =   5145
   ClientLeft      =   50
   ClientTop       =   440
   ClientWidth     =   7810
   OleObjectBlob   =   "UserFormTolerancia.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserFormTolerancia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

' ---------------------------------------------------------------------
' Módulo: UserFormTolerancia.bas
' Descrição: Formulário de avaliação de tolerância ao risco
' ---------------------------------------------------------------------

Private Const TOTAL_PERGUNTAS As Integer = 10
Private Const PONTOS_PERFIL_AGRESSIVO As Integer = 22
Private Const PONTOS_PERFIL_MODERADO As Integer = 12

Private Sub CommandButtonAvalia_Click()
  Dim intPontos As Integer
  Dim intRespostas As Integer
  On Error GoTo ErroAvaliar

  '1) idade: > 55 anos | de 35 a 55 anos | < 35 anos
  PontuarPergunta OptionButton011, OptionButton012, OptionButton013, 1, 2, 3, intPontos, intRespostas
  '2) bens financeiros em relação ao patrimônio total: > 70% | de 40% a 70% | < 40%
  PontuarPergunta OptionButton021, OptionButton022, OptionButton023, 3, 2, 1, intPontos, intRespostas
  '3) atitude perante perda: levaria numa boa | jamais investiria nesse novamente | ficaria desolado
  PontuarPergunta OptionButton031, OptionButton032, OptionButton033, 3, 1, 2, intPontos, intRespostas
  '4) dívidas totais: < 10% | de 10% a 15% | > 15%
  PontuarPergunta OptionButton041, OptionButton042, OptionButton043, 3, 2, 1, intPontos, intRespostas
  '5) fundo de emergência: < 3 meses | de 3 a 6 meses | > 6 meses
  PontuarPergunta OptionButton051, OptionButton052, OptionButton053, 1, 2, 3, intPontos, intRespostas
  '6) prognóstico nos próx. 12 meses de economia: incerto | otimista | pessimista
  PontuarPergunta OptionButton061, OptionButton062, OptionButton063, 2, 3, 1, intPontos, intRespostas
  '7) prazo sem resgatar: > 5 anos | de 1 a 5 anos | até 1 ano
  PontuarPergunta OptionButton071, OptionButton072, OptionButton073, 3, 2, 1, intPontos, intRespostas
  '8) experiência em investimentos: pouco | boa | razoável
  PontuarPergunta OptionButton081, OptionButton082, OptionButton083, 1, 3, 2, intPontos, intRespostas
  '9) atual carteira de investimentos: renda fixa | fundos | ações
  PontuarPergunta OptionButton091, OptionButton092, OptionButton093, 2, 3, 1, intPontos, intRespostas
  '10) tempo até a aposentadoria: > 10 anos | entre 5 e 10 anos | < 5 anos
  PontuarPergunta OptionButton101, OptionButton102, OptionButton103, 3, 2, 1, intPontos, intRespostas

  If intRespostas < TOTAL_PERGUNTAS Then
    MsgBox "Preenchimento incorreto, assinale sua opção em cada pergunta", vbCritical, Me.Caption
    Exit Sub
  End If
  MsgBox RetornarPerfilRisco(intPontos), vbInformation, Me.Caption
  Unload Me
  Exit Sub

ErroAvaliar:
  MostrarMsgErro "UserFormTolerancia.CommandButtonAvalia_Click"
End Sub

'
' Soma os pontos da opção assinalada e conta a pergunta como respondida
'
Private Sub PontuarPergunta(ByVal optOpcao1 As MSForms.OptionButton, ByVal optOpcao2 As MSForms.OptionButton, _
    ByVal optOpcao3 As MSForms.OptionButton, ByVal intPontosOpcao1 As Integer, ByVal intPontosOpcao2 As Integer, _
    ByVal intPontosOpcao3 As Integer, ByRef intPontos As Integer, ByRef intRespostas As Integer)
  If optOpcao1.Value Then
    intPontos = intPontos + intPontosOpcao1
  ElseIf optOpcao2.Value Then
    intPontos = intPontos + intPontosOpcao2
  ElseIf optOpcao3.Value Then
    intPontos = intPontos + intPontosOpcao3
  Else
    Exit Sub
  End If
  intRespostas = intRespostas + 1
End Sub

Private Function RetornarPerfilRisco(ByVal intPontos As Integer) As String
  If intPontos >= PONTOS_PERFIL_AGRESSIVO Then
    RetornarPerfilRisco = "Situação confortável para eventualmente ampliar sua atual tolerância a riscos: AGRESSIVO"
  ElseIf intPontos >= PONTOS_PERFIL_MODERADO Then
    RetornarPerfilRisco = "Tolerância moderada a novos riscos: MODERADO"
  Else
    RetornarPerfilRisco = "Baixa tolerância a assumir novos riscos: CONSERVADOR"
  End If
End Function