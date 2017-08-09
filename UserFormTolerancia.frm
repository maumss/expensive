VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormTolerancia 
   Caption         =   "Tolerância ao Risco"
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

Private Sub CommandButtonAvalia_Click()
  'variáveis
  Dim intPontos As Integer
  Dim intRespostas As Integer
 
  intPontos = 0
  intRespostas = 0
  '1) idade
  If OptionButton011 Then     '> 55 anos
    intPontos = intPontos + 1
    intRespostas = intRespostas + 1
  ElseIf OptionButton012 Then 'de 35 a 55 anos
    intPontos = intPontos + 2
    intRespostas = intRespostas + 1
  ElseIf OptionButton013 Then '< 35 anos
    intPontos = intPontos + 3
    intRespostas = intRespostas + 1
  End If
  '2) bens financeiros em relação ao patrimônio total
  If OptionButton021 Then     '> 70%
    intPontos = intPontos + 3
    intRespostas = intRespostas + 1
  ElseIf OptionButton022 Then 'de 40% a 70%
    intPontos = intPontos + 2
    intRespostas = intRespostas + 1
  ElseIf OptionButton023 Then '< 40%
    intPontos = intPontos + 1
    intRespostas = intRespostas + 1
  End If
  '3) atitude perante perda
  If OptionButton031 Then     'levaria numa boa
    intPontos = intPontos + 3
    intRespostas = intRespostas + 1
  ElseIf OptionButton032 Then 'jamais investiria nesse novamente
    intPontos = intPontos + 1
    intRespostas = intRespostas + 1
  ElseIf OptionButton033 Then 'ficaria desolado
    intPontos = intPontos + 2
    intRespostas = intRespostas + 1
  End If
  '4) dívidas totais
  If OptionButton041 Then     '< 10%
    intPontos = intPontos + 3
    intRespostas = intRespostas + 1
  ElseIf OptionButton042 Then 'de 10% a 15%
    intPontos = intPontos + 2
    intRespostas = intRespostas + 1
  ElseIf OptionButton043 Then '> 15%
    intPontos = intPontos + 1
    intRespostas = intRespostas + 1
  End If
  '5) fundo de emergência
  If OptionButton051 Then     '< 3 meses
    intPontos = intPontos + 1
    intRespostas = intRespostas + 1
  ElseIf OptionButton052 Then 'de 3 a 6 meses
    intPontos = intPontos + 2
    intRespostas = intRespostas + 1
  ElseIf OptionButton053 Then '> 6 meses
    intPontos = intPontos + 3
    intRespostas = intRespostas + 1
  End If
  '6) prognóstico nos próx. 12 meses de economia
  If OptionButton061 Then     'incerto
    intPontos = intPontos + 2
    intRespostas = intRespostas + 1
  ElseIf OptionButton062 Then 'otimista
    intPontos = intPontos + 3
    intRespostas = intRespostas + 1
  ElseIf OptionButton063 Then 'pessimista
    intPontos = intPontos + 1
    intRespostas = intRespostas + 1
  End If
  '7) prazo sem resgatar
  If OptionButton071 Then     '> 5 anos
    intPontos = intPontos + 3
    intRespostas = intRespostas + 1
  ElseIf OptionButton072 Then 'de 1 a 5 anos
    intPontos = intPontos + 2
    intRespostas = intRespostas + 1
  ElseIf OptionButton073 Then 'até 1 ano
    intPontos = intPontos + 1
    intRespostas = intRespostas + 1
  End If
  '8) experiência em investimentos
  If OptionButton081 Then     'pouco
    intPontos = intPontos + 1
    intRespostas = intRespostas + 1
  ElseIf OptionButton082 Then 'boa
    intPontos = intPontos + 3
    intRespostas = intRespostas + 1
  ElseIf OptionButton083 Then 'razoável
    intPontos = intPontos + 2
    intRespostas = intRespostas + 1
  End If
  '9) atual carteira de investimentos
  If OptionButton091 Then     'renda fixa
    intPontos = intPontos + 2
    intRespostas = intRespostas + 1
  ElseIf OptionButton092 Then 'fundos
    intPontos = intPontos + 3
    intRespostas = intRespostas + 1
  ElseIf OptionButton093 Then 'ações
    intPontos = intPontos + 1
    intRespostas = intRespostas + 1
  End If
  '10) tempo até a aposentadoria
  If OptionButton101 Then     '> 10 anos
    intPontos = intPontos + 3
    intRespostas = intRespostas + 1
  ElseIf OptionButton102 Then 'entre 5 e 10 anos
    intPontos = intPontos + 2
    intRespostas = intRespostas + 1
  ElseIf OptionButton103 Then '< 5 anos
    intPontos = intPontos + 1
    intRespostas = intRespostas + 1
  End If
  If intRespostas < 10 Then
    MsgBox "Preenchimento incorreto, assinale sua opção em cada pergunta", vbCritical
  Else
    If intPontos >= 22 Then
      MsgBox "Situação confortável para eventualmente ampliar sua atual tolerância a riscos: AGRESSIVO", vbInformation
    ElseIf (intPontos >= 12) And (intPontos < 22) Then
      MsgBox "Tolerância moderada a novos riscos: MODERADO", vbInformation
    Else
      MsgBox "Baixa tolerância a assumir novos riscos: CONSERVADOR", vbInformation
    End If
    UserFormTolerancia.Hide
    Unload Me
  End If
End Sub

