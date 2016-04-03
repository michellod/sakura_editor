Option Explicit

Class Ora_Column
  '=== ATTRIBUTE ===
  Private mstr_Owner          '
  Private mstr_TableName      '
  Private mstr_ColumnName     '
  Private mstr_DataType       '
  Private mlng_DataLength     '
  Private mlng_DataPrecision  '
  Private mlng_DataScale      '
  Private mbln_Nullable       '
  Private mlng_ColumnID       '
  
  '=== PROPERTY ===
  Public Property Get CLASSNAME                     (): CLASSNAME                      = "Column"                 : End Property
  Public Property Get Owner                         (): Owner                          = mstr_Owner               : End Property
  Public Property Get TableName                     (): TableName                      = mstr_TableName           : End Property
  Public Property Get ColumnName                    (): ColumnName                     = mstr_ColumnName          : End Property
  Public Property Get DataType                      (): DataType                       = mstr_DataType            : End Property
  Public Property Get DataLength                    (): DataLength                     = mlng_DataLength          : End Property
  Public Property Get DataPrecision                 (): DataPrecision                  = mlng_DataPrecision       : End Property
  Public Property Get DataScale                     (): DataScale                      = mlng_DataScale           : End Property
  Public Property Get Nullable                      (): Nullable                       = mbln_Nullable            : End Property
  Public Property Get ColumnID                      (): ColumnID                       = mlng_ColumnID            : End Property
  
  Public Property Let Owner                         (ByVal istr_Owner        ): mstr_Owner                 = istr_Owner               : End Property
  Public Property Let TableName                     (ByVal istr_TableName    ): mstr_TableName             = istr_TableName           : End Property
  Public Property Let ColumnName                    (ByVal istr_ColumnName   ): mstr_ColumnName            = istr_ColumnName          : End Property
  Public Property Let DataType                      (ByVal istr_DataType     ): mstr_DataType              = istr_DataType            : End Property
  Public Property Let DataLength                    (ByVal ilng_DataLength   ): mlng_DataLength            = ilng_DataLength          : End Property
  Public Property Let DataPrecision                 (ByVal ilng_DataPrecision): mlng_DataPrecision         = ilng_DataPrecision       : End Property
  Public Property Let DataScale                     (ByVal ilng_DataScale    ): mlng_DataScale             = ilng_DataScale           : End Property
  Public Property Let Nullable                      (ByVal ibln_Nullable     ): mbln_Nullable              = ibln_Nullable            : End Property
  Public Property Let ColumnID                      (ByVal ilng_ColumnID     ): mlng_ColumnID              = ilng_ColumnID            : End Property

  '=== PROCEDURE ===
  Public Sub Class_Initialize()
    mstr_Owner         = ""
    mstr_TableName     = ""
    mstr_ColumnName    = ""
    mstr_DataType      = ""
    mlng_DataLength    = 0
    mlng_DataPrecision = 0
    mlng_DataScale     = 0
    mbln_Nullable      = True
    mlng_ColumnID      = 0
  End Sub
  
  Public Sub Class_Terminate()
  End Sub
  
End Class
