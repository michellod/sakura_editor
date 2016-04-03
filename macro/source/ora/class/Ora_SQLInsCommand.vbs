Option Explicit

Class Ora_SQLInsCommand
  '=== ATTRIBUTE ===
  Private mstr_uid
  Private mstr_pwd
  Private mstr_sid
  Private mstr_Delim
  Private mstr_DestOwner
  Private mstr_DestTable
  
  Private mbln_ValidateFlag
  Private mbln_ArrangeFlag
  Private mbln_ExecuteFlag
  Private mstr_Method
  Private mbln_ConnectionStringFlag
  Private mstr_DelimCode
  
  Private mstr_Commands
  Private mlng_CommandsNumber
  Private mlng_CommandDataLineTop
  
  Private mobj_Connector
  Private mobj_Columns
  Private mlng_ColumnsNumber
  
  Private mlng_ArrangedLinesNumber
  Private mstr_ArrangedLines
  
  '=== PROPERTY ===
  Public Property Get CLASSNAME                     (): CLASSNAME                      = "SQLInsCommand"   : End Property
  
  '=== PROCEDURE ===
  Public Sub Class_Initialize()
  End Sub
  
  Public Sub Class_Terminate()
    Set mobj_Connector = Nothing
  End Sub
  
  'SQL.INS DEST=[OWNER].[TABLE] (DELIM=TAB|COMMA) ([UID]/[PWD]@[SID]) (METH=FULL|PART) (/VALID) (/ARRANGE) (/EXEC)
  Public Sub Operate(ByRef ostr_ReturnCode, ByRef ostr_Message, ByVal istr_Command, ByVal istr_Options)
    Dim lstr_Command
    
    lstr_Command = istr_Command
    Call ExtractHeader_(ostr_ReturnCode, ostr_Message, istr_Options)
    
    Call GenerateColumns_(ostr_ReturnCode, ostr_Message, mstr_DestOwner, mstr_DestTable, mstr_Method, lstr_Command)
    
    If mbln_ValidateFlag Then
      Call ValidateDataLines_(ostr_ReturnCode, ostr_Message)
    End If
    
    If mbln_ArrangeFlag Then
      Call Arrange_(ostr_ReturnCode, ostr_Message)
    End If
    
ostr_Message = "SQL.INS " & mstr_uid & "/" & mstr_pwd & "@" & mstr_sid & vbCrLf _
     & "Delim = " & mstr_Delim & vbCrLf _
     & "Owner = " & mstr_DestOwner & ", Table = " & mstr_DestTable & vbCrLf _
     & "Validate = " & mbln_ValidateFlag & ", Arrange = " & mbln_ArrangeFlag & ", Execute = " & mbln_ExecuteFlag & vbCrLf _
     & "Method   = " & mstr_Method
    
  End Sub
  
  Private Sub ExtractHeader_(ByRef ostr_ReturnCode, ByRef ostr_Message, ByVal istr_Options)
    Dim lobj_RE
    Dim lobj_Matches
    Dim lobj_Submatches
    Dim lstr_Subcommand
    Dim lstr_Options
    Dim lstr_Option
    Dim llng_Index
    Dim llng_OptionNum
    
    mbln_ConnectionStringFlag = False
    mbln_ArrangeFlag = False
    mbln_ValidateFlag = False
    mbln_ExecuteFlag = False
    
    Set lobj_RE = CreateObject("VBScript.RegExp")
    lobj_RE.IgnoreCase = False
    lobj_RE.Global     = False
    lobj_RE.Pattern    = " +"
    
    istr_Options = lobj_RE.Replace(istr_Options, " ")
    lstr_Options = Split(istr_Options, " ")
    llng_OptionNum = UBound(lstr_Options)
    
    For llng_Index = 0 To llng_OptionNum
      lstr_Option = lstr_Options(llng_Index)
      
      If mstr_uid = "" Then
        Call ORA.GetDBConnectionInfo(mstr_uid, mstr_pwd, mstr_sid, lstr_Option)
        If mstr_uid <> "" Then mbln_ConnectionStringFlag = True
      End If
      
      lobj_RE.Pattern = "DELIM\=(TAB|COMMA)"
      If lobj_RE.Test(lstr_Option) Then
        Set lobj_Matches = lobj_RE.Execute(lstr_Option)
        If lobj_Matches.Count > 0 Then
          Set lobj_Submatches = lobj_Matches(0).Submatches
          mstr_Delim = lobj_Submatches.Item(0)
        End If
      Else
        lobj_RE.Pattern = "DEST\=([A-Z0-9_]+)\.([A-Z0-9_]+)"
        If lobj_RE.Test(lstr_Option) Then
          Set lobj_Matches = lobj_RE.Execute(lstr_Option)
          If lobj_Matches.Count > 0 Then
            Set lobj_Submatches = lobj_Matches(0).Submatches
            mstr_DestOwner = lobj_Submatches.Item(0)
            mstr_DestTable = lobj_Submatches.Item(1)
          End If
        Else
          lobj_RE.Pattern = "\/VALID"
          If lobj_RE.Test(lstr_Option) Then
            Set lobj_Matches = lobj_RE.Execute(lstr_Option)
            If lobj_Matches.Count > 0 Then
              mbln_ValidateFlag = True
            End If
          Else
            lobj_RE.Pattern = "\/ARRANGE"
            If lobj_RE.Test(lstr_Option) Then
              Set lobj_Matches = lobj_RE.Execute(lstr_Option)
              If lobj_Matches.Count > 0 Then
                mbln_ArrangeFlag = True
              End If
            Else
              lobj_RE.Pattern = "\/EXEC"
              If lobj_RE.Test(lstr_Option) Then
                Set lobj_Matches = lobj_RE.Execute(lstr_Option)
                If lobj_Matches.Count > 0 Then
                  mbln_ExecuteFlag = True
                End If
              Else
                lobj_RE.Pattern = "METH\=(FULL|PART)"
                If lobj_RE.Test(lstr_Option) Then
                  Set lobj_Matches = lobj_RE.Execute(lstr_Option)
                  If lobj_Matches.Count > 0 Then
                    Set lobj_Submatches = lobj_Matches(0).Submatches
                    mstr_Method = lobj_Submatches.Item(0)
                  End If
                End If
              End If
            End If
          End If
        End If
      End If
    Next
    
    If mstr_Delim = "" Then mstr_Delim = "TAB"
    If mstr_Method = "" Then mstr_Method = "FULL"
    If mstr_DestOwner = "" Then
      ostr_ReturnCode = "E"
      ostr_Message    = GLOBAL.Message.GetMessageText("ora", "SQL-INS-00001")
      Set lobj_Submatches = Nothing
      Set lobj_Matches = Nothing
      Set lobj_RE      = Nothing
      Exit Sub
    ElseIf mstr_DestTable = "" Then
      ostr_ReturnCode = "E"
      ostr_Message    = GLOBAL.Message.GetMessageText("ora", "SQL-INS-00001")
      Set lobj_Submatches = Nothing
      Set lobj_Matches = Nothing
      Set lobj_RE      = Nothing
      Exit Sub
    End If
    
    Select Case mstr_Delim
      Case "TAB"
        mstr_DelimCode = Chr(9)
      Case "COMMA"
        mstr_DelimCode = ","
      Case Else
        mstr_DelimCode = mstr_Delim
    End Select
    
    Set lobj_Submatches = Nothing
    Set lobj_Matches = Nothing
    Set lobj_RE      = Nothing
    
    ostr_ReturnCode = "N"
  End Sub
  
  Private Sub GenerateColumns_(ByRef ostr_ReturnCode, ByRef ostr_Message, ByVal istr_DestOwner, ByVal istr_DestTable, ByVal istr_Method, ByVal istr_Command)
    Dim llng_ColumnIndex
    Dim lobj_Column
    
    If mbln_ConnectionStringFlag Then
      Set mobj_Connector = CreateObject("ADODB.Connection")
      Call mobj_Connector.Open(ORA.GetConnectionString(mstr_uid, mstr_pwd, mstr_sid))
    End If
    
    Select Case istr_Method
      Case "FULL"
        mstr_Commands           = Split(istr_Command, vbCrLf)
        mlng_CommandDataLineTop = 1
        mlng_CommandsNumber     = UBound(mstr_Commands)
        
        If mlng_CommandDataLineTop > mlng_CommandsNumber Then
          Erase mstr_Commands
          ostr_ReturnCode = "E"
          ostr_Message    = GLOBAL.Message.GetMessageText("ora", "SQL-INS-000002")
          Exit Sub
        End If
        
        Call ORA.DBA.GenerateAllTabColumn(ostr_ReturnCode, ostr_Message, mobj_Columns, mlng_ColumnsNumber, mobj_Connector, istr_DestOwner, istr_DestTable)
      Case "PART"
        mstr_Commands           = Split(istr_Command, vbCrLf)
        mlng_CommandDataLineTop = 2
        mlng_CommandsNumber     = UBound(mstr_Commands)
        
        If mlng_CommandDataLineTop > mlng_CommandsNumber Then
          Erase mstr_Commands
          ostr_ReturnCode = "E"
          ostr_Message    = GLOBAL.Message.GetMessageText("ora","SQL-INS-000002")
          Exit Sub
        End If
        
        Dim lstr_Columns
        
        lstr_Columns       = Split(mstr_Commands(1), mstr_DelimCode)
        mlng_ColumnsNumber = UBound(lstr_Columns)
        ReDim mobj_Columns(mlng_ColumnsNumber)
        
        For llng_ColumnIndex = 0 To mlng_ColumnsNumber
          Call ORA.DBA.GenerateSingleTabColumn(ostr_ReturnCode, ostr_Message, mobj_Columns(llng_ColumnIndex), mobj_Connector, istr_DestOwner, istr_DestTable, lstr_Columns(llng_ColumnIndex), mbln_ConnectionStringFlag)
        Next
    End Select
  End Sub
  
  Private Sub ValidateDataLines_(ByRef ostr_ReturnCode, ByRef ostr_Message)
    Dim lstr_ReturnCode
    Dim lstr_Message
    
    Dim llng_LineIndex
    Dim llng_ColumnsNumber
    Dim lstr_Fields
    Dim llng_FieldsNumber
    
    ostr_ReturnCode = "N"
    
    For llng_LineIndex = mlng_CommandDataLineTop To mlng_CommandsNumber
      lstr_Fields = Split(mstr_Commands(llng_LineIndex), mstr_DelimCode)
      llng_FieldsNumber = UBound(lstr_Fields)
      
      Call ValidateColumns_(lstr_ReturnCode, lstr_Message, lstr_Fields, llng_FieldsNumber)
      If lstr_ReturnCode <> "N" Then
        ostr_ReturnCode = lstr_ReturnCode
        ostr_Message    = lstr_Message
        Exit For
      End If
    Next
  End Sub
  
  Private Sub ValidateColumns_(ByRef ostr_ReturnCode, ByRef ostr_Message, ByVal istr_Fields, ByVal ilng_FieldsNumber)
    If mlng_ColumnsNumber <> ilng_FieldsNumber Then
      ostr_ReturnCode = "E"
      ostr_Message    = GLOBAL.Message.GetMessageText("ora", "SQL-INS-000002")
      Exit Sub
    End If
    
    If mbln_ConnectionStringFlag Then
      ostr_ReturnCode = "N"
      
      Dim llng_FieldIndex
      
      For llng_FieldIndex = 0 To ilng_FieldsNumber
        Call ORA.DBA.ValidateColumnValue(ostr_ReturnCode, ostr_Message, mobj_Columns(llng_FieldIndex), istr_Fields(llng_FieldIndex))
        If ostr_ReturnCode <> "N" Then
          Exit For
        End If
      Next
    Else
      ostr_ReturnCode = "N"
    End If
  End Sub
  
  Private Sub Arrange_(ByRef ostr_ReturnCode, ByRef ostr_Message)
    
    mlng_ArrangedLinesNumber = mlng_CommandsNumber - mlng_CommandDataLineTop
    ReDim mstr_ArrangedLines(mlng_ArrangedLinesNumber)
    
    Dim llng_LineIndex
    Dim lstr_ArrangedLine
    Dim lstr_Fields
    Dim llng_FieldsNumber
    Dim llng_ColumnIndex
    
    Select Case mstr_Method
      Case "PART"
        For llng_LineIndex = mlng_CommandDataLineTop To mlng_CommandsNumber
          lstr_ArrangedLine = "INSERT INTO " & mstr_DestOwner & "." & mstr_DestTable & " ("
          
          For llng_ColumnIndex = 0 To mlng_ColumnsNumber
            lstr_ArrangedLine = lstr_ArrangedLine & mobj_Columns(llng_ColumnIndex).ColumnName & ","
          Next
          lstr_ArrangedLine = Mid(lstr_ArrangedLine, 1, Len(lstr_ArrangedLine) - 1)
          lstr_ArrangedLine = lstr_ArrangedLine & ") VALUES ("
          
          lstr_Fields = Split(mstr_Commands(llng_LineIndex), mstr_DelimCode)
          For llng_ColumnIndex = 0 To mlng_ColumnsNumber
            lstr_ArrangedLine = lstr_ArrangedLine & ORA.DBA.ConvertField(lstr_Fields(llng_ColumnIndex)) & ","
          Next
          lstr_ArrangedLine = Mid(lstr_ArrangedLine, 1, Len(lstr_ArrangedLine) - 1)
          lstr_ArrangedLine = lstr_ArrangedLine & ");"
          
          mstr_ArrangedLines(llng_LineIndex - mlng_CommandDataLineTop) = lstr_ArrangedLine
          
          Editor.InsText(lstr_ArrangedLine & vbCrLf)
        Next
      Case "FULL"
        For llng_LineIndex = mlng_CommandDataLineTop To mlng_CommandsNumber
          lstr_ArrangedLine = "INSERT INTO " & mstr_DestOwner & "." & mstr_DestTable & " VALUES ("
          
          lstr_Fields = Split(mstr_Commands(llng_LineIndex), mstr_DelimCode)
          llng_FieldsNumber = UBound(lstr_Fields)
          For llng_ColumnIndex = 0 To llng_FieldsNumber
            lstr_ArrangedLine = lstr_ArrangedLine & ORA.DBA.ConvertField(lstr_Fields(llng_ColumnIndex)) & ","
          Next
          
          lstr_ArrangedLine = Mid(lstr_ArrangedLine, 1, Len(lstr_ArrangedLine) - 1)
          lstr_ArrangedLine = lstr_ArrangedLine & ");"
          
          mstr_ArrangedLines(llng_LineIndex - mlng_CommandDataLineTop) = lstr_ArrangedLine
          
          Editor.InsText(lstr_ArrangedLine & vbCrLf)
        Next
    End Select
    
  End Sub
End Class
