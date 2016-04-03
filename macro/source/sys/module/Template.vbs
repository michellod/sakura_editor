Option Explicit

Dim SAKURA_OPENED_FILE_PATH       : SAKURA_OPENED_FILE_PATH       = Editor.ExpandParameter("$F")
Dim SAKURA_OPENED_FILE_NAME       : SAKURA_OPENED_FILE_NAME       = Editor.ExpandParameter("$f")
Dim SAKURA_OPENED_FILE_BODY       : SAKURA_OPENED_FILE_BODY       = Editor.ExpandParameter("$g")
Dim SAKURA_OPENED_FILE_PATH_SLASH : SAKURA_OPENED_FILE_PATH_SLASH = Editor.ExpandParameter("$/")
Dim SAKURA_EXEC_PATH              : SAKURA_EXEC_PATH              = Editor.ExpandParameter("$S")
Dim SAKURA_MACRO_PATH             : SAKURA_MACRO_PATH             = Editor.ExpandParameter("$M")
Dim SAKURA_EXEC_DIR               : SAKURA_EXEC_DIR               = Mid(SAKURA_EXEC_PATH, 1, Len(SAKURA_EXEC_PATH) - Len("sakura.exe") - 1)
Dim SAKURA_TEMP_TOP               : SAKURA_TEMP_TOP               = SAKURA_EXEC_DIR & "\macro\temp"
Dim SAKURA_SOURCE_TOP             : SAKURA_SOURCE_TOP             = SAKURA_EXEC_DIR & "\macro\source"
Dim SAKURA_RESOURCE_TOP           : SAKURA_RESOURCE_TOP           = SAKURA_EXEC_DIR & "\macro\resource"
Dim LOGGER

Call EntryPoint()

Public Sub EntryPoint()
  Dim lstr_ResultCode
  Dim lstr_Message
  Dim lobj_FSO
  Dim lstr_ResultPath
  
  Set lobj_FSO = CreateObject("Scripting.FileSystemObject")
  If Err > 0 Then
    Exit Sub
  End If
  
  lstr_ResultPath = SAKURA_TEMP_TOP & "\result.txt"
  Set LOGGER = lobj_FSO.OpenTextFile(lstr_ResultPath, 2, True, 0)
  If Not lobj_FSO.FileExists(lstr_ResultPath) Then
    Set lobj_FSO = Nothing
    Exit Sub
  End If
  
  Call LoadLibraries_(lstr_ResultCode,lstr_Message)
  If lstr_ResultCode <> "N" Then
    Call MsgBox(lstr_Message, vbOKOnly + vbExclamation)
    Exit Sub
  End If

  Call Sys_Init(lstr_ResultCode,lstr_Message)
  
  '### Customization Start Point ###
  
  
  '### Customization End Point ###
  
  Call LOGGER.Close()
  
  Set LOGGER = Nothing
  Set lobj_FSO   = Nothing
End Sub

Private Sub LoadLibraries_(ByRef ostr_ResultCode,ByRef ostr_Message)
  ostr_ResultCode = ""
  ostr_Message    = ""
  
  Const lstr_LIBRARY_LIST_FILE_NAME = "libraries.lst"
  
  Dim lstr_LibraryListPath
  Dim lobj_FSO
  Dim lobj_TS
  Dim lstr_Line
  
  lstr_LibraryListPath = SAKURA_SOURCE_TOP & "\" & lstr_LIBRARY_LIST_FILE_NAME
  LOGGER.WriteLine("[" & CStr(Date) & " " & CStr(Time) & "] Load libraries...")
  
  On Error Resume Next
    Set lobj_FSO = CreateObject("Scripting.FileSystemObject")
    If Err > 0 Then
      ostr_ResultCode = "E"
      ostr_Message    = "The FSO object could not be created."
      Exit Sub
    End If
    
    If Not lobj_FSO.FileExists(lstr_LibraryListPath) Then
      ostr_ResultCode = "E"
      ostr_Message    = "The file could not be found." & vbCrLf & lstr_LibraryListPath
      Exit Sub
    End If
    
    Set lobj_TS = lobj_FSO.OpenTextFile(lstr_LibraryListPath, 1, False, 0)
    If Err > 0 Then
      ostr_ResultCode = "E"
      ostr_Message    = "The file could not be opened." & vbCrLf & lstr_LibraryListPath
      Exit Sub
    End If
    
    Do Until lobj_TS.AtEndOfStream
      lstr_Line = lobj_TS.ReadLine()
      Call LoadLibrary_(ostr_ResultCode, ostr_Message, lobj_FSO, lstr_Line, " ")
      If ostr_ResultCode = "E" Then Exit Do
    Loop
    
    lobj_TS.Close()
  On Error GoTo 0
  
  Set lobj_TS  = Nothing
  Set lobj_FSO = Nothing
End Sub

Private Sub LoadLibrary_(ByRef ostr_ResultCode,ByRef ostr_Message,ByRef iobj_FSO,ByVal istr_Line,ByVal istr_Indent)
  Dim lstr_ResultCode
  Dim lstr_Message
  Dim lstr_Key
  Dim lstr_Path
  Dim lobj_TS

  lstr_Key  = Mid(istr_Line,1,3)
  lstr_Path = SAKURA_SOURCE_TOP & "\" & Mid(istr_Line, 5)
  
  On Error Resume Next
    Select Case lstr_Key
      Case "LST"
        LOGGER.WriteLine(istr_Indent & lstr_Path)
        
        If Not iobj_FSO.FileExists(lstr_Path) Then
          ostr_ResultCode = "E"
          ostr_Message    = "The file could not be found." & vbCrLf & lstr_Path
          Exit Sub
        End If
        
        Set lobj_TS = iobj_FSO.OpenTextFile(lstr_Path, 1, False, 0)
        If Err > 0 Then
          ostr_ResultCode = "E"
          ostr_Message    = "The file could not be opened." & vbCrLf & lstr_Path
          Exit Sub
        End If
        
        Do Until lobj_TS.AtEndOfStream
          Call LoadLibrary_(lstr_ResultCode,lstr_Message,iobj_FSO,lobj_TS.ReadLine(), istr_Indent & " ")
          If lstr_ResultCode = "E" Then
            ostr_ResultCode = "E"
            ostr_Message    = ostr_Message & vbCrLf & lstr_Message
            Exit Do
          ElseIf lstr_ResultCode = "W" Then
            ostr_ResultCode = "W"
          End If
        Loop
        
        lobj_TS.Close()
        
        If ostr_ResultCode = "" Then ostr_ResultCode = "N"
      Case "SRC"
        If Not iobj_FSO.FileExists(lstr_Path) Then
          ostr_ResultCode = "E"
          ostr_Message    = "The file could not be found." & vbCrLf & lstr_Path
          Exit Sub
        End If
        
        Set lobj_TS = iobj_FSO.OpenTextFile(lstr_Path, 1, False, 0)
        If Err > 0 Then
          ostr_ResultCode = "E"
          ostr_Message    = "The file could not be opened." & vbCrLf & lstr_Path
          Exit Sub
        End If
        
        ExecuteGlobal lobj_TS.ReadAll()
        If Err > 0 Then
          ostr_ResultCode = "E"
          ostr_Message    = Err.Description & vbCrLf & lstr_Path
          Call lobj_TS.Close()
          Exit Sub
        End If
        
        lobj_TS.Close()
        
        LOGGER.WriteLine(istr_Indent & lstr_Path)
'        Call WScript.StdOut.WriteLine(" " & lstr_Path)
        ostr_ResultCode = "N"
      Case Else
        ostr_ResultCode = "E"
        ostr_Message    = "Invalid item type."
    End Select
  On Error GoTo 0
End Sub
