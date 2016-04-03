Option Explicit

'Oracle
Dim ORA

Public Sub Ora_Init(ByRef ostr_ReturnCode,ByRef ostr_Message)
  Dim llng_Index
  
  Set ORA = New Ora_Oracle
  ostr_ReturnCode = "N"
  ostr_Message    = ""
End Sub
