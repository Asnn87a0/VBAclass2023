Attribute VB_Name = "Module1"
Option Explicit

Sub �f�n�p��j()
Attribute �f�n�p��j.VB_Description = "123"
Attribute �f�n�p��j.VB_ProcData.VB_Invoke_Func = "a\n14"
'
' �f�n�p��j ����
'
' �ֳt��: Ctrl+a
'
    Range("B1").Select
    ActiveWorkbook.Worksheets("�u�@��1").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("�u�@��1").Sort.SortFields.Add2 Key:=Range("B2:B414") _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("�u�@��1").Sort
        .SetRange Range("A1:B414")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub

Sub �f�n�j��p()
Attribute �f�n�j��p.VB_Description = "321"
Attribute �f�n�j��p.VB_ProcData.VB_Invoke_Func = "b\n14"
'
' �f�n�j��p ����
'
' �ֳt��: Ctrl+b
'
    Range("B1").Select
    ActiveWorkbook.Worksheets("�u�@��1").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("�u�@��1").Sort.SortFields.Add2 Key:=Range("B2:B414") _
        , SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("�u�@��1").Sort
        .SetRange Range("A1:B414")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub
