Option Explicit

Class Sys_ProcessEnvironment
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
  Public Property Get HOMEDRIVE                     (): HOMEDRIVE              = mobj_Env.Item("HOMEDRIVE"             ): End Property
  Public Property Get HOMEPATH                      (): HOMEPATH               = mobj_Env.Item("HOMEPATH"              ): End Property
  Public Property Get PATH                          (): PATH                   = mobj_Env.Item("PATH"                  ): End Property
  Public Property Get PATHEXT                       (): PATHEXT                = mobj_Env.Item("PATHEXT"               ): End Property
  Public Property Get PROMPT                        (): PROMPT                 = mobj_Env.Item("PROMPT"                ): End Property
  Public Property Get SYSTEMDRIVE                   (): SYSTEMDRIVE            = mobj_Env.Item("SYSTEMDRIVE"           ): End Property
  Public Property Get SYSTEMROOT                    (): SYSTEMROOT             = mobj_Env.Item("SYSTEMROOT"            ): End Property
  Public Property Get WINDIR                        (): WINDIR                 = mobj_Env.Item("WINDIR"                ): End Property
  Public Property Get TEMP                          (): TEMP                   = mobj_Env.Item("TEMP"                  ): End Property
  Public Property Get TMP                           (): TMP                    = mobj_Env.Item("TMP"                   ): End Property
  Public Property Let NUMBER_OF_PROCESSES           (ByVal istr_NUMBER_OF_PROCESSES   ): mobj_Env.Item("NUMBER_OF_PROCESSES"   ) = istr_NUMBER_OF_PROCESSES   : End Property
  Public Property Let PROCESSOR_ARCHITECTURE        (ByVal istr_PROCESSOR_ARCHITECTURE): mobj_Env.Item("PROCESSOR_ARCHITECTURE") = istr_PROCESSOR_ARCHITECTURE: End Property
  Public Property Let PROCESSOR_IDENTIFIER          (ByVal istr_PROCESSOR_IDENTIFIER  ): mobj_Env.Item("PROCESSOR_IDENTIFIER"  ) = istr_PROCESSOR_IDENTIFIER  : End Property
  Public Property Let PROCESSOR_LEVEL               (ByVal istr_PROCESSOR_LEVEL       ): mobj_Env.Item("PROCESSOR_LEVEL"       ) = istr_PROCESSOR_LEVEL       : End Property
  Public Property Let PROCESSOR_REVISION            (ByVal istr_PROCESSOR_REVISION    ): mobj_Env.Item("PROCESSOR_REVISION"    ) = istr_PROCESSOR_REVISION    : End Property
  Public Property Let OS                            (ByVal istr_OS                    ): mobj_Env.Item("OS"                    ) = istr_OS                    : End Property
  Public Property Let COMSPEC                       (ByVal istr_COMSPEC               ): mobj_Env.Item("COMSPEC"               ) = istr_COMSPEC               : End Property
  Public Property Let HOMEDRIVE                     (ByVal istr_HOMEDRIVE             ): mobj_Env.Item("HOMEDRIVE"             ) = istr_HOMEDRIVE             : End Property
  Public Property Let HOMEPATH                      (ByVal istr_HOMEPATH              ): mobj_Env.Item("HOMEPATH"              ) = istr_HOMEPATH              : End Property
  Public Property Let PATH                          (ByVal istr_PATH                  ): mobj_Env.Item("PATH"                  ) = istr_PATH                  : End Property
  Public Property Let PATHEXT                       (ByVal istr_PATHEXT               ): mobj_Env.Item("PATHEXT"               ) = istr_PATHEXT               : End Property
  Public Property Let PROMPT                        (ByVal istr_PROMPT                ): mobj_Env.Item("PROMPT"                ) = istr_PROMPT                : End Property
  Public Property Let SYSTEMDRIVE                   (ByVal istr_SYSTEMDRIVE           ): mobj_Env.Item("SYSTEMDRIVE"           ) = istr_SYSTEMDRIVE           : End Property
  Public Property Let SYSTEMROOT                    (ByVal istr_SYSTEMROOT            ): mobj_Env.Item("SYSTEMROOT"            ) = istr_SYSTEMROOT            : End Property
  Public Property Let WINDIR                        (ByVal istr_WINDIR                ): mobj_Env.Item("WINDIR"                ) = istr_WINDIR                : End Property
  Public Property Let TEMP                          (ByVal istr_TEMP                  ): mobj_Env.Item("TEMP"                  ) = istr_TEMP                  : End Property
  Public Property Let TMP                           (ByVal istr_TMP                   ): mobj_Env.Item("TMP"                   ) = istr_TMP                   : End Property
  
  '=== PROCEDURE ===
  Public Sub Class_Initialize()
    MSTR_CLASSNAME      = "Sys_ProcessEnvironment"
    MLNG_NUMBEROFFIELDS = 16
    
    Set mobj_Shell = CreateObject("WScript.Shell")
    Set mobj_Env   = mobj_Shell.Environment("Process")
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
