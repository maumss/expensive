Option Explicit
'Módulo referente ao formulário de avaliação

Public Sub AvaliarRisco()
  '
  ' Avalia Macro
  ' Avalia tolerância ao risco nos investimentos
  '
  ' Atalho do teclado: Ctrl+a
  '
  On Error GoTo erroAvalia
  UserFormTolerancia.Show False
  Exit Sub
  
erroAvalia:
  MostrarMsgErro ("AvaliarRisco")
End Sub
