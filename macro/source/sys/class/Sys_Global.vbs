Option Explicit

Class Sys_Global
  '=== ATTRIBUTE ===
  Private mdat_StartDate
  Private mdat_EndDate
  Private mstr_Language
  Private mobj_MessageAPI
  Private mobj_WMI
  
  '=== PROPERTY ===
  Public Property Get CLASSNAME                     (): CLASSNAME                      = "Global"        : End Property
  Public Property Get StartDate                     (): StartDate                      = mdat_StartDate  : End Property
  Public Property Get Language                      (): Language                       = mstr_Language   : End Property
  Public Property Get Message                       (): Set Message                    = mobj_MessageAPI : End Property
  Public Property Get WMI                           (): Set WMI                        = mobj_WMI        : End Property
  
  '=== PROCEDURE ===
  Public Sub Class_Initialize()
    mdat_StartDate = CDate(CStr(Date) & " " & CStr(Time))
    
    mstr_Language = VBS.Locale.ConvertLanguageCode(GetLocale())
    
    Set mobj_MessageAPI = New Sys_MessageAPI
    Set mobj_WMI = New Sys_WindowsManagementInstrumentation
  End Sub
  
  Public Sub Class_Terminate()
    Set mobj_MessageAPI = Nothing
    Set mobj_WMI = Nothing
  End Sub
End Class
