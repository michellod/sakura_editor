Option Explicit

Class Vbs_VBS
  '=== ATTRIBUTE ===
  Private mobj_Locale
  
  '=== PROPERTY ===
  Public Property Get CLASSNAME(): CLASSNAME  = "VBS"       : End Property
  Public Property Get Locale   (): Set Locale = mobj_Locale : End Property
  
  '=== PROCEDURE ===
  Public Sub Class_Initialize()
    Set mobj_Locale = New Vbs_Locale
  End Sub
  
  Public Sub Class_Terminate()
    Set mobj_Locale = Nothing
  End Sub
End Class
