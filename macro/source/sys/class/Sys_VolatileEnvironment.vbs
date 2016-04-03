Option Explicit

Class Sys_VolatileEnvironment
  '=== ATTRIBUTE ===
  Private MSTR_CLASSNAME
  Private MLNG_NUMBEROFFIELDS
  
  Private mobj_Shell
  Private mobj_Env
  
  '=== PROPERTY ===
  Public Property Get HOMEDRIVE                     (): HOMEDRIVE              = mobj_Env.Item("HOMEDRIVE"             ): End Property
  Public Property Get HOMEPATH                      (): HOMEPATH               = mobj_Env.Item("HOMEPATH"              ): End Property
  
  '=== PROCEDURE ===
  Public Sub Class_Initialize()
    MSTR_CLASSNAME      = "Sys_VolatileEnvironment"
    MLNG_NUMBEROFFIELDS = 16
    
    Set mobj_Shell = CreateObject("WScript.Shell")
    Set mobj_Env   = mobj_Shell.Environment("Volatile")
  End Sub
  
  Public Sub Class_Terminate()
    Set mobj_Env   = Nothing
    Set mobj_Shell = Nothing
  End Sub
  
  Public Function GetItem(ByVal istr_ItemName)
    GetItem = mobj_Env.Item(istr_ItemName)
  End Function
  Public Function SetItem(ByVal istr_ItemName, ByVal istr_ItemValue)
    mobj_Env.Item(istr_ItemName) = istr_ItemValue
    SetItem = mobj_Env.Item(istr_ItemName)
  End Function
End Class
