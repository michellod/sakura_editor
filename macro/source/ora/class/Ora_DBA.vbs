Option Explicit

Class Ora_DBA
  '=== ATTRIBUTE ===
  Private mobj_RE
  
  '=== PROPERTY ===
  Public Property Get CLASSNAME                     (): CLASSNAME                      = "DBA"                 : End Property
  Public Property Get TAB_COLUMNS_PATH              (): TAB_COLUMNS_PATH               = "dba\tab_columns.sql" : End Property

  '=== PROCEDURE ===
  Public Sub Class_Initialize()
    Set mobj_RE = CreateObject("VBScript.RegExp")
    mobj_RE.IgnoreCase = False
    mobj_RE.Global     = False
  End Sub
  
  Public Sub Class_Terminate()
  End Sub
  
  Public Sub GenerateAllTabColumn(ByRef ostr_ReturnCode, ByRef ostr_Message, ByRef oobj_Columns, ByRef olng_ColumnsNumber, ByRef iobj_Connector, ByVal istr_DestOwner, ByVal istr_DestTable)
    Dim lstr_Path
    Dim lobj_FSO
    DIm lobj_TS
    Dim lstr_SQL
    Dim lobj_RS
    Dim lobj_Column
    Dim lobj_Fields
    
    lstr_Path = SAKURA_RESOURCE_TOP & "\ora\" & TAB_COLUMNS_PATH
    
    Set lobj_FSO = CreateObject("Scripting.FileSystemObject")
    If Not lobj_FSO.FileExists(lstr_Path) Then
      ostr_ReturnCode = "E"
      ostr_Message    = GLOBAL.Message.GetMessage("ora","DBA-SQL-000001")
      Set lobj_FSO = Nothing
      Exit Sub
    End If
    
    Set lobj_TS = lobj_FSO.OpenTextFile(lstr_Path, 1, False, 0)
    lstr_SQL = lobj_TS.ReadAll()
    
    Call lobj_TS.Close()
    Set lobj_TS = Nothing
    Set lobj_FSO = Nothing
    
    mobj_RE.Pattern = "\:OWNER"
    lstr_SQL = mobj_RE.Replace(lstr_SQL, "'" & istr_DestOwner & "'")
    
    mobj_RE.Pattern = "\:TABLE_NAME"
    lstr_SQL = mobj_RE.Replace(lstr_SQL, "'" & istr_DestTable & "'")
    
    mobj_RE.Pattern = "\:COLUMN_NAME"
    lstr_SQL = mobj_RE.Replace(lstr_SQL, "NULL")
    
    Set lobj_RS = CreateObject("ADODB.RecordSet")
    Call lobj_RS.Open(lstr_SQL, iobj_Connector, 1, 1)
    Call lobj_RS.MoveFirst()
    
    olng_ColumnsNumber = 0
    Do Until lobj_RS.EOF
      Set lobj_Fields = lobj_RS.Fields
      
      Set lobj_Column = New Ora_Column
      lobj_Column.Owner         = lobj_Fields("OWNER"         ).Value
      lobj_Column.TableName     = lobj_Fields("TABLE_NAME"    ).Value
      lobj_Column.ColumnName    = lobj_Fields("COLUMN_NAME"   ).Value
      lobj_Column.DataType      = lobj_Fields("DATA_TYPE"     ).Value
      lobj_Column.DataLength    = lobj_Fields("DATA_LENGTH"   ).Value
      lobj_Column.DataPrecision = lobj_Fields("DATA_SCALE"    ).Value
      lobj_Column.DataScale     = lobj_Fields("DATA_PRECISION").Value
      lobj_Column.Nullable      = lobj_Fields("NULLABLE"      ).Value
      lobj_Column.ColumnID      = lobj_Fields("COLUMN_ID"     ).Value
      
      ReDim Preserve oobj_Columns(olng_ColumnsNumber)
      Set oobj_Columns(olng_ColumnsNumber) = lobj_Fields
      olng_ColumnsNumber = olng_ColumnsNumber + 1
      
      Call lobj_RS.MoveNext()
    Loop
    
    Call lobj_RS.Close()
    
    Set lobj_RS = Nothing
  End Sub
  
  Public Sub GenerateSingleTabColumn(ByRef ostr_ReturnCode, ByRef ostr_Message, ByRef oobj_Column, ByRef iobj_Connector, ByVal istr_DestOwner, ByVal istr_DestTable, ByVal istr_DestColumn, ByVal ibln_IsDynamic)
    If ibln_IsDynamic Then
      Set oobj_Column = New Ora_Column
    Else
      Set oobj_Column = New Ora_Column
      oobj_Column.ColumnName = istr_DestColumn
    End If
  End Sub
  
  Public Sub ValidateColumnValue(ByRef ostr_ReturnCode, ByRef ostr_Message, ByRef iobj_Column, ByVal istr_Field)
    ostr_ReturnCode = "N"
  End Sub
  
  Public Function ConvertField(ByVal istr_Field)
    mobj_RE.Pattern = "^__(.+)__$"
    If mobj_RE.Test(istr_Field) Then
      ConvertField = Mid(istr_Field, 3, Len(istr_Field) - 4)
    Else
      ConvertField = "'" & istr_Field & "'"
    End If
  End Function
End Class
