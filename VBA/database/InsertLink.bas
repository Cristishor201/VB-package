Attribute VB_Name = "Module2"
Sub InsertLink()
'
' InsertLink Macro
' Insert a desired link in Baze_date_idei_proiecte
'

'
    Dim Name, Link As String
    Dim cell As String
    Name = "Btn"
    
    cell = ActiveCell.Address ' stochez adresa
    
    ActiveCell.Value = ""
    Application.Dialogs(xlDialogInsertHyperlink).Show
    If Range(cell).Hyperlinks.Count > 0 Then
        Link = ActiveCell.Hyperlinks(1).Address
        ActiveCell.Hyperlinks(1).delete
        ActiveCell.FormulaR1C1 = _
           "=IF(RC[1]=""Ready"",HYPERLINK(""" & Link & """,""" & Name & """),"""")"
    
        'Range(cell).Formula = "=IF(ADDRESS(ROW();COLUMN((G2));3)=" & "'Ready'" _
            & ";HYPERLINK('" & Link & "';'" & Name & "');'')"
    Else
        ActiveCell.Hyperlinks(1).delete
        ActiveCell.Value = ""
    End If
    
    With Selection.Font
        .Color = -4165632
        .TintAndShade = 0
        .Underline = xlUnderlineStyleSingle
    End With
    Link = ""
End Sub


