Option Explicit

Class Sys_UserEnvironment
  '=== ATTRIBUTE ===
  Private MSTR_CLASSNAME
  Private MLNG_NUMBEROFFIELDS
  
  Private mobj_Shell
  Private mobj_Env
  
  '=== PROPERTY ===
  Public Property Get TEMP                          (): TEMP                   = mobj_Env.Item("TEMP"                  ): End Property
  Public Property Get TMP                           (): TMP                    = mobj_Env.Item("TMP"                   ): End Property
  
  Public Sub Class_Initialize()
    MSTR_CLASSNAME      = "Sys_UserEnvironment"
    MLNG_NUMBEROFFIELDS = 16
    
    Set mobj_Shell = CreateObject("WScript.Shell")
    Set mobj_Env   = mobj_Shell.Environment("User")
  End Sub
  
  Public Sub Class_Terminate()
    Set mobj_Env   = Nothing
    Set mobj_Shell = Nothing
  End Sub
End Class
