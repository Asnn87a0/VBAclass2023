Attribute VB_Name = "Module1"
Option Explicit

Sub 口罩小到大()
Attribute 口罩小到大.VB_Description = "123"
Attribute 口罩小到大.VB_ProcData.VB_Invoke_Func = "a\n14"
'
' 口罩小到大 巨集
'
' 快速鍵: Ctrl+a
'
    Range("B1").Select
    ActiveWorkbook.Worksheets("工作表1").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("工作表1").Sort.SortFields.Add2 Key:=Range("B2:B414") _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("工作表1").Sort
        .SetRange Range("A1:B414")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub

Sub 口罩大到小()
Attribute 口罩大到小.VB_Description = "321"
Attribute 口罩大到小.VB_ProcData.VB_Invoke_Func = "b\n14"
'
' 口罩大到小 巨集
'
' 快速鍵: Ctrl+b
'
    Range("B1").Select
    ActiveWorkbook.Worksheets("工作表1").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("工作表1").Sort.SortFields.Add2 Key:=Range("B2:B414") _
        , SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("工作表1").Sort
        .SetRange Range("A1:B414")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub
