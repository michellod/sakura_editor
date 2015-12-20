Option Explicit

Public Const GSTR_MACRO_FOLDER_NAME = "macro"
Public Const GSTR_MACRO_FILE_NAME   = "vbs_comment.vbs"

Dim glng_CurrentLineNo
Dim gdat_MacroStartDate
Dim gstr_FileName
Dim gstr_MacroDir
Dim gstr_MacroPath
Dim glng_SelectedLineFrom
Dim glng_SelectedColumnFrom

Call Initialize()
Call Main()

Public Sub Initialize()
  glng_CurrentLineNo      = ExpandParameter("$y")
  gdat_MacroStartDate     = Now()
  gstr_FileName           = ExpandParameter("$f")
  gstr_MacroPath          = ExpandParameter("$M")
  gstr_MacroDir           = Mid(gstr_MacroPath, 1, Len(gstr_MacroPath) - Len(GSTR_MACRO_FILE_NAME) - 1)
  glng_SelectedLineFrom   = GetSelectLineFrom
  glng_SelectedColumnFrom = GetSelectColmFrom
End Sub

Public Sub Main()
  Dim lobj_RE
  Dim lstr_Line
  Dim lstr_CommentTemplate
  
  lstr_CommentTemplate = ""
  
  'Get the line string either selected currently or in the current line.
  lstr_Line = GetSelectedString(0)
  If lstr_Line = "" Then lstr_Line = GetLineStr(glng_CurrentLineNo)
  
  On Error Resume Next
    Set lobj_RE = CreateObject("VBScript.RegExp")
    lobj_RE.Global     = False
    lobj_RE.IgnoreCase = False
    lobj_RE.Pattern    = "\r\n$"
    
    lstr_Line = lobj_RE.Replace(lstr_Line, "")
  On Error GoTo 0
  
  If lstr_Line = "" Then
    If glng_CurrentLineNo = 1 Then
      Call ReadComment_(lstr_CommentTemplate,"header_comment")
    End If
  Else
    'Function
    Call GetFunctionComment_(lstr_CommentTemplate, lstr_Line)
    
    'Sub
    Call GetSubComment_(lstr_CommentTemplate, lstr_Line)
    
    'Class
    Call GetClassComment_(lstr_CommentTemplate, lstr_Line)
  End If
  
  If lstr_CommentTemplate <> "" Then
    If glng_SelectedLineFrom = 0 Then
      Up
      GoLineEnd
    Else
      Call MoveCursor(glng_SelectedLineFrom - 1, glng_SelectedColumnFrom, 0)
      GoLineEnd
    End If
    InsText(lstr_CommentTemplate)
  End If
  
  Set lobj_RE = Nothing
End Sub

Private Sub ReadComment_(ByRef ostr_Comment,ByVal istr_TemplateName)
  Dim lstr_Path
  Dim lobj_FSO
  Dim lobj_TS
  
  lstr_Path = gstr_MacroDir & "\vbs\template\" & istr_TemplateName & ".txt"
  
  On Error Resume Next
    Set lobj_FSO = CreateObject("Scripting.FileSystemObject")
    If Err > 0 Then
      ErrorMsg("Unexpected error has occured.")
      Exit Sub
    End If
    
    If Not lobj_FSO.FileExists(lstr_Path) Then
      ErrorMsg("The template file could not be found." & vbCrLf _
             & " " & lstr_Path)
      Set lobj_FSO = Nothing
      Exit Sub
    End If
    
    Set lobj_TS = lobj_FSO.OpenTextFile(lstr_Path, 1, False, 0)
    ostr_Comment = lobj_TS.ReadAll
    
    Call lobj_TS.Close
    
    Set lobj_TS  = Nothing
    Set lobj_FSO = Nothing
  On Error GoTo 0
End Sub

Private Function GetFunctionComment_(ByRef ostr_CommentTemplate,ByVal istr_Line)
  GetFunctionComment_ = False
  
  Dim lstr_Indent
  Dim lstr_VisibleQualifier
  Dim lstr_FunctionName
  Dim lstr_CommentTemplate
  Dim lobj_RE
  Dim lobj_Matches
  Dim lobj_SubMatches
  
  Set lobj_RE = CreateObject("VBScript.RegExp")
  lobj_RE.Global     = False
  lobj_RE.IgnoreCase = False
  lobj_RE.Pattern    = "( *)(Private|Public) +Function +([a-zA-Z0-9_]+)"
  
  Set lobj_Matches = lobj_RE.Execute(istr_Line)
  If lobj_Matches.Count > 0 Then
    Set lobj_SubMatches   = lobj_Matches.Item(0).SubMatches
    lstr_Indent           = lobj_SubMatches.Item(0)
    lstr_VisibleQualifier = lobj_SubMatches.Item(1)
    lstr_FunctionName     = lobj_SubMatches.Item(2)
    Set lobj_SubMatches = Nothing
    Set lobj_Matches    = Nothing
    
    Call ReadComment_(lstr_CommentTemplate,"function_comment")
    
    lobj_RE.Global  = True
    lobj_RE.Pattern = "\r\n"
    lstr_CommentTemplate = vbCrLf & lstr_Indent & lobj_RE.Replace(lstr_CommentTemplate, vbCrLf & lstr_Indent)
    lobj_RE.Pattern = "\r\n" & lstr_Indent & "$"
    lstr_CommentTemplate = lobj_RE.Replace(lstr_CommentTemplate, "")
    lobj_RE.Pattern = "__VISIBLE_QUALIFIER__"
    lstr_CommentTemplate = lobj_RE.Replace(lstr_CommentTemplate, lstr_VisibleQualifier)
    lobj_RE.Pattern = "__FUNCTION_NAME__"
    lstr_CommentTemplate = lobj_RE.Replace(lstr_CommentTemplate, lstr_FunctionName)
    
    ostr_CommentTemplate = lstr_CommentTemplate
    GetFunctionComment_ = True
  End If
  
  Set lobj_RE = Nothing
End Function

Private Function GetSubComment_(ByRef ostr_CommentTemplate,ByVal istr_Line)
  GetSubComment_ = False
  
  Dim lstr_Indent
  Dim lstr_VisibleQualifier
  Dim lstr_SubName
  Dim lstr_CommentTemplate
  Dim lobj_RE
  Dim lobj_Matches
  Dim lobj_SubMatches
  
  Set lobj_RE = CreateObject("VBScript.RegExp")
  lobj_RE.Global     = False
  lobj_RE.IgnoreCase = False
  lobj_RE.Pattern    = "( *)(Private|Public) +Sub +([a-zA-Z0-9_]+)"
  
  Set lobj_Matches = lobj_RE.Execute(istr_Line)
  If lobj_Matches.Count > 0 Then
    Set lobj_SubMatches   = lobj_Matches.Item(0).SubMatches
    lstr_Indent           = lobj_SubMatches.Item(0)
    lstr_VisibleQualifier = lobj_SubMatches.Item(1)
    lstr_SubName          = lobj_SubMatches.Item(2)
    Set lobj_SubMatches = Nothing
    Set lobj_Matches    = Nothing
    
    Call ReadComment_(lstr_CommentTemplate,"sub_comment")
    
    lobj_RE.Global  = True
    lobj_RE.Pattern = "\r\n"
    lstr_CommentTemplate = vbCrLf & lstr_Indent & lobj_RE.Replace(lstr_CommentTemplate, vbCrLf & lstr_Indent)
    lobj_RE.Pattern = "\r\n" & lstr_Indent & "$"
    lstr_CommentTemplate = lobj_RE.Replace(lstr_CommentTemplate, "")
    lobj_RE.Pattern = "__VISIBLE_QUALIFIER__"
    lstr_CommentTemplate = lobj_RE.Replace(lstr_CommentTemplate, lstr_VisibleQualifier)
    lobj_RE.Pattern = "__SUB_NAME__"
    lstr_CommentTemplate = lobj_RE.Replace(lstr_CommentTemplate, lstr_SubName)
    
    ostr_CommentTemplate = lstr_CommentTemplate
    GetSubComment_ = True
  End If
  
  Set lobj_RE = Nothing
End Function

Private Function GetClassComment_(ByRef ostr_CommentTemplate,ByVal istr_Line)
  GetClassComment_ = False
  
  Dim lstr_Indent
  Dim lstr_ClassName
  Dim lstr_CommentTemplate
  Dim lobj_RE
  Dim lobj_Matches
  Dim lobj_SubMatches
  
  Set lobj_RE = CreateObject("VBScript.RegExp")
  lobj_RE.Global     = False
  lobj_RE.IgnoreCase = False
  lobj_RE.Pattern    = "( *)Class +([a-zA-Z0-9_]+)"
  
  Set lobj_Matches = lobj_RE.Execute(istr_Line)
  If lobj_Matches.Count > 0 Then
    Set lobj_SubMatches = lobj_Matches.Item(0).SubMatches
    lstr_Indent         = lobj_SubMatches.Item(0)
    lstr_ClassName      = lobj_SubMatches.Item(1)
    Set lobj_SubMatches = Nothing
    Set lobj_Matches    = Nothing
    
    Call ReadComment_(lstr_CommentTemplate,"class_comment")
    
    lobj_RE.Global  = True
    lobj_RE.Pattern = "\r\n"
    lstr_CommentTemplate = vbCrLf & lstr_Indent & lobj_RE.Replace(lstr_CommentTemplate, vbCrLf & lstr_Indent)
    lobj_RE.Pattern = "\r\n" & lstr_Indent & "$"
    lstr_CommentTemplate = lobj_RE.Replace(lstr_CommentTemplate, "")
    lobj_RE.Pattern = "__CLASS_NAME__"
    lstr_CommentTemplate = lobj_RE.Replace(lstr_CommentTemplate, lstr_ClassName)
    
    ostr_CommentTemplate = lstr_CommentTemplate
    GetClassComment_ = True
  End If
  
  Set lobj_RE = Nothing
End Function
