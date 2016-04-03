Option Explicit

Class Sys_MessageAPI
  '=== ATTRIBUTE ===
  
  '=== PROPERTY ===
  Public Property Get CLASSNAME                     (): CLASSNAME                      = "MessageAPI"  : End Property
  
  '=== PROCEDURE ===
  Public Sub Class_Initialize()
  End Sub
  
  Public Sub Class_Terminate()
  End Sub
  
  Public Function GetMessageText(ByVal istr_ApplicationShortName, ByVal istr_MessageCode)
    GetMessageText = istr_ApplicationShortName & "." & istr_MessageCode
  End Function
End Class
