Attribute VB_Name = "Avaliacao"
Option Explicit
'M�dulo referente ao formul�rio de avalia��o

Public Sub AvaliarRisco()
Attribute AvaliarRisco.VB_Description = "Avalia toler�ncia ao risco nos investimentos"
Attribute AvaliarRisco.VB_ProcData.VB_Invoke_Func = "a\n14"
  '
  ' Avalia Macro
  ' Avalia toler�ncia ao risco nos investimentos
  '
  ' Atalho do teclado: Ctrl+a
  '
  On Error GoTo erroAvalia
  UserFormTolerancia.Show False
  Exit Sub
  
erroAvalia:
  MostrarMsgErro ("AvaliarRisco")
End Sub
