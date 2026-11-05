Option Explicit

' ---------------------------------------------------------------------
' Módulo: Avaliacao.bas
' Descrição: Módulo referente ao formulário de avaliação
' ---------------------------------------------------------------------

'
' Avalia Macro
' Avalia tolerância ao risco nos investimentos
'
' Atalho do teclado: Ctrl+a
'
Public Sub AvaliarRisco()  
  On Error GoTo ErroAvalia
  UserFormTolerancia.Show False
  Exit Sub
  
ErroAvalia:
  MostrarMsgErro ("AvaliarRisco")
End Sub
