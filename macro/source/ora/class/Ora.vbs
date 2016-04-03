Option Explicit

Class Ora_Oracle
  '=== ATTRIBUTE ===
  Private mobj_Ora_DBA
  Private mobj_Ora_Command
  Private mobj_Ora_SQLCommand
  Private mobj_Ora_SQLInsCommand
  
  '=== PROPERTY ===
  Public Property Get CLASSNAME                     (): CLASSNAME                      = "Oracle"                 : End Property
  Public Property Get DBA                           (): Set DBA                        = mobj_Ora_DBA             : End Property
  Public Property Get Command                       (): Set Command                    = mobj_Ora_Command         : End Property
  Public Property Get SQLCommand                    (): Set SQLCommand                 = mobj_Ora_SQLCommand      : End Property
  Public Property Get SQLInsCommand                 (): Set SQLInsCommand              = mobj_Ora_SQLInsCommand   : End Property
  
  '=== PROCEDURE ===
  Public Sub Class_Initialize()
    Set mobj_Ora_DBA           = New Ora_DBA
    Set mobj_Ora_Command       = New Ora_Command
    Set mobj_Ora_SQLCommand    = New Ora_SQLCommand
    Set mobj_Ora_SQLInsCommand = New Ora_SQLInsCommand
  End Sub
  
  Public Sub Class_Terminate()
    Set mobj_Ora_DBA           = Nothing
    Set mobj_Ora_Command    = Nothing
    Set mobj_Ora_SQLCommand = Nothing
    Set mobj_Ora_SQLInsCommand = Nothing
  End Sub
  
  Public Function GetConnectionString(ByVal istr_sid, ByVal istr_uid, ByVal istr_pwd)
    GetConnectionString = "Driver={Microsoft ODBC for Oracle}; Server=" & istr_sid & "; UID=" & istr_uid & "; PWD=" & istr_pwd & ";"
  End Function
  
  Public Sub GetDBConnectionInfo(ByRef ostr_uid, ByRef ostr_pwd, ByRef ostr_sid, ByVal istr_Text)
    Dim lobj_RE
    Dim lobj_Matches
    Dim lobj_Submatches
    
    Set lobj_RE = CreateObject("VBScript.RegExp")
    lobj_RE.IgnoreCase = False
    lobj_RE.Global     = False
    lobj_RE.Pattern = "([a-zA-Z0-9\_]+)\/([a-zA-Z0-9\_\#]+)\@([a-zA-Z0-9\_]+)"
    If lobj_RE.Test(istr_Text) Then
      Set lobj_Matches = lobj_RE.Execute(istr_Text)
      If lobj_Matches.Count > 0 Then
        Set lobj_Submatches = lobj_Matches(0).Submatches
        ostr_uid     = lobj_Submatches.Item(0)
        ostr_pwd     = lobj_Submatches.Item(1)
        ostr_sid     = lobj_Submatches.Item(2)
      End If
    End If
    
    Set lobj_Submatches = Nothing
    Set lobj_Matches    = Nothing
    Set lobj_RE         = Nothing
  End Sub
  
End Class
