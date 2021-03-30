VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_Open()
    'Clear cell content and formating
    Range("A19:AC100").Clear
    Range("A19:AC100").ClearFormats
    Range("L1:AC18").Clear
    Range("L1:AC18").ClearFormats
    
    'Remove all previous arrows
    Dim Shp As Shape
    For Each Shp In ActiveSheet.Shapes
        If Shp.Type = msoAutoShape Then Shp.Delete
    Next Shp
    
    'Reset row height and column width
    Rows.UseStandardHeight = True
    Columns.UseStandardWidth = True
End Sub
