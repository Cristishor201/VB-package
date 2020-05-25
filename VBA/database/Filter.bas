Attribute VB_Name = "Module1"
Public Sub Filtrare()
    If Range("M1").Value = "Lista" Then
        Call Sort
    ElseIf Range("M1").Value = "Familie" Then
        Call Sort_Fam
    End If
End Sub

Sub Sort()
Attribute Sort.VB_Description = "sort button"
Attribute Sort.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Sort Macro
' sort button
'

'
    ActiveWorkbook.Worksheets("Baza date IDEI").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Baza date IDEI").Sort.SortFields.Add Key:=Range( _
        "J3:J100"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:= _
        xlSortNormal
    ActiveWorkbook.Worksheets("Baza date IDEI").Sort.SortFields.Add Key:=Range( _
        "K3:K100"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:= _
        xlSortNormal
    ActiveWorkbook.Worksheets("Baza date IDEI").Sort.SortFields.Add Key:=Range( _
        "H3:H100"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("Baza date IDEI").Sort
        .SetRange Range("A3:N100")
        .Header = xlGuess
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub

Sub Sort_Fam()
'
' Sort_Fam Macro
' Sorteaza pe Familie si Nivel.
'

'
    ActiveWorkbook.Worksheets("Baza date IDEI").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Baza date IDEI").Sort.SortFields.Add Key:=Range( _
        "L3:L75"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    ActiveWorkbook.Worksheets("Baza date IDEI").Sort.SortFields.Add Key:=Range( _
        "M3:M75"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("Baza date IDEI").Sort
        .SetRange Range("A2:N100")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub
