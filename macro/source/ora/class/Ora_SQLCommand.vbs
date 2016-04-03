Option Explicit

Class Ora_SQLCommand
  '=== ATTRIBUTE ===
  
  '=== PROPERTY ===
  Public Property Get CLASSNAME                     (): CLASSNAME                      = "SQLCommand"   : End Property
  
  '=== PROCEDURE ===
  Public Sub Class_Initialize()
  End Sub
  
  Public Sub Class_Terminate()
  End Sub
  
  Public Sub Operate(ByRef ostr_ReturnCode, ByRef ostr_Message, ByVal istr_CommandText)
    Dim lobj_RE
    Dim lobj_Matches
    Dim lobj_Submatches
    Dim lstr_Subcommand
    Dim lstr_uid
    Dim lstr_pwd
    Dim lstr_sid
    Dim lstr_options
    
    Set lobj_RE = CreateObject("VBScript.RegExp")
    lobj_RE.IgnoreCase = False
    lobj_RE.Global     = False
    lobj_RE.Pattern    = "SQL\.([A-Z\_]+) *(.*)"
    
    If Not lobj_RE.Test(istr_CommandText) Then
      Set lobj_RE     = Nothing
      ostr_ReturnCode = "E"
      ostr_Message    = SYS.MSGMNG.GetMessageText("ora","SQL-000-01001")
      Exit Sub
    End If
    
    Set lobj_Matches = lobj_RE.Execute(istr_CommandText)
    If lobj_Matches.Count > 0 Then
      Set lobj_Submatches = lobj_Matches(0).Submatches
      lstr_Subcommand     = lobj_Submatches.Item(0)
      lstr_options        = lobj_Submatches.Item(1)
      
      Select Case lstr_Subcommand
        Case "INS"
          Call ORA.SQLInsCommand.Operate(ostr_ReturnCode, ostr_Message, istr_CommandText, lstr_options)
      End Select
    End If
    
    Set lobj_Submatches = Nothing
    Set lobj_Matches = Nothing
    Set lobj_RE      = Nothing
    
    ostr_ReturnCode = "N"
  End Sub
End Class
