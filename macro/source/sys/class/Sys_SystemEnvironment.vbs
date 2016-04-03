Option Explicit

Class Sys_SystemEnvironment
  '=== ATTRIBUTE ===
  Private MSTR_CLASSNAME
  Private MLNG_NUMBEROFFIELDS
  
  Private mobj_Shell
  Private mobj_Env
  
  '=== PROPERTY ===
  Public Property Get NUMBER_OF_PROCESSES           (): NUMBER_OF_PROCESSES    = mobj_Env.Item("NUMBER_OF_PROCESSES"   ): End Property
  Public Property Get PROCESSOR_ARCHITECTURE        (): PROCESSOR_ARCHITECTURE = mobj_Env.Item("PROCESSOR_ARCHITECTURE"): End Property
  Public Property Get PROCESSOR_IDENTIFIER          (): PROCESSOR_IDENTIFIER   = mobj_Env.Item("PROCESSOR_IDENTIFIER"  ): End Property
  Public Property Get PROCESSOR_LEVEL               (): PROCESSOR_LEVEL        = mobj_Env.Item("PROCESSOR_LEVEL"       ): End Property
  Public Property Get PROCESSOR_REVISION            (): PROCESSOR_REVISION     = mobj_Env.Item("PROCESSOR_REVISION"    ): End Property
  Public Property Get OS                            (): OS                     = mobj_Env.Item("OS"                    ): End Property
  Public Property Get COMSPEC                       (): COMSPEC                = mobj_Env.Item("COMSPEC"               ): End Property
  Public Property Get PATH                          (): PATH                   = mobj_Env.Item("PATH"                  ): End Property
  Public Property Get PATHEXT                       (): PATHEXT                = mobj_Env.Item("PATHEXT"               ): End Property
  Public Property Get WINDIR                        (): WINDIR                 = mobj_Env.Item("WINDIR"                ): End Property
  
  Public Sub Class_Initialize()
    MSTR_CLASSNAME      = "Sys_SystemEnvironment"
    MLNG_NUMBEROFFIELDS = 16
    
    Set mobj_Shell = CreateObject("WScript.Shell")
    Set mobj_Env   = mobj_Shell.Environment("System")
  End Sub
  
  Public Sub Class_Terminate()
    Set mobj_Env   = Nothing
    Set mobj_Shell = Nothing
  End Sub
End Class
