Option Explicit

Class Ora_Command
  '=== ATTRIBUTE ===
  
  '=== PROPERTY ===
  Public Property Get CLASSNAME                     (): CLASSNAME                      = "Command"   : End Property
  
  '=== PROCEDURE ===
  Public Sub Class_Initialize()
  End Sub
  
  Public Sub Class_Terminate()
  End Sub
  
  Public Sub Operate(ByRef ostr_ReturnCode, ByRef ostr_Message, ByVal istr_CommandText)
    Dim lobj_RE
    Dim lobj_Matches
    Dim lobj_Submatches
    Dim lstr_Command
    
    Set lobj_RE = CreateObject("VBScript.RegExp")
    lobj_RE.IgnoreCase = False
    lobj_RE.Global     = True
    lobj_RE.Pattern    = "\r\n$"
    istr_CommandText   = lobj_RE.Replace(istr_CommandText, "")

    lobj_RE.Global     = False
    lobj_RE.Pattern    = "([a-zA-Z0-9\_]+)"
    
    If Not lobj_RE.Test(istr_CommandText) Then
      Set lobj_RE     = Nothing
      ostr_ReturnCode = "E"
'      ostr_Message    = SYS.MSGMNG.GetMessageText("ora","CMD-000-00001")
      Exit Sub
    End If
    
    Set lobj_Matches = lobj_RE.Execute(istr_CommandText)
    If lobj_Matches.Count > 0 Then
      Set lobj_Submatches = lobj_Matches(0).Submatches
      lstr_Command = lobj_Submatches.Item(0)
      
      Select Case lstr_Command
        Case "SQL"
          Call Ora.SQLCommand.Operate(ostr_ReturnCode, ostr_Message, istr_CommandText)
      End Select
    End If
    
    Set lobj_Matches = Nothing
    Set lobj_RE      = Nothing
    
    ostr_ReturnCode = "N"
  End Sub
  
End Class
