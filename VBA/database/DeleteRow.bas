Attribute VB_Name = "Module3"
Public Sub delete()
    Dim line As Integer
    Dim raspuns As Integer
    line = ActiveCell.Row
    
    raspuns = MsgBox("Esti sigur ca vrei sa stergi proiectul """ & Range("$C" & line).Value & """", vbYesNo + vbQuestion, "Intrebare")
    
    If raspuns = vbYes Then
        Rows(line & ":" & line).Select 'Selecteaza toata linia
        Selection.delete Shift:=xlUp
    End If
End Sub
