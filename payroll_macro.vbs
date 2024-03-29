Sub future_dates()
'
' future_dates Macro
'
'
    Dim price As Integer
    
    With ActiveDocument
        If .FormFields("Check1").CheckBox.Value = True Then
            price = 100
        End If
    End With
    
    With ActiveDocument
        If .FormFields("Check2").CheckBox.Value = True Then
            price = 400
        End If
    End With
    
    With ActiveDocument
        If .FormFields("Check3").CheckBox.Value = True Then
            price = 200
        End If
    End With
    
    With ActiveDocument
        If .FormFields("Check4").CheckBox.Value = True Then
            price = 200
        End If
    End With
    
    With ActiveDocument
        If .FormFields("Check5").CheckBox.Value = True Then
            price = 400
        End If
    End With
    
    With ActiveDocument
        If .FormFields("Check6").CheckBox.Value = True Then
            price = 100
        End If
    End With
    
    With ActiveDocument
        If .FormFields("Check7").CheckBox.Value = True Then
            price = 400
        End If
    End With
    
    With ActiveDocument
        If .FormFields("Check8").CheckBox.Value = True Then
            price = 600
        End If
    End With
    
    With ActiveDocument
        If .FormFields("Check9").CheckBox.Value = True Then
            price = 400
        End If
    End With
        
    
    Dim newDate, newDate2, newDate3, newDate4 As Date
    
    If Format(Date, "d") >= 15 Then
        newDate = DateSerial(Year(Date), Month(Date) + 1, 0)
        Else
            newDate = DateSerial(Year(Date), Month(Date), 15)
    End If
    
    If Format(newDate, "d") >= 15 And Format(newDate, "d") < 32 Then
        newDate2 = DateSerial(Year(newDate), Month(newDate) + 1, 15)
        Else
            newDate2 = DateSerial(Year(newDate), Month(newDate) + 1, 0)
    End If
    
    If Format(newDate2, "d") >= 15 And Format(newDate2, "d") < 32 Then
        newDate3 = DateSerial(Year(newDate2), Month(newDate2) + 1, 0)
        Else
            newDate3 = DateSerial(Year(newDate2), Month(newDate2), 15)
    End If
    
    If Format(newDate3, "d") >= 15 And Format(newDate3, "d") < 32 Then
        newDate4 = DateSerial(Year(newDate3), Month(newDate3) + 1, 15)
        Else
            newDate4 = DateSerial(Year(newDate3), Month(newDate3) + 1, 0)
    End If
    
    With Selection.Find
        .ClearFormatting
        .Text = "<<today>>"
        .Replacement.ClearFormatting
        .Replacement.Text = Format(Date, "mmmm dd yyyy")
        .Execute Replace:=wdReplaceAll, Forward:=True, Wrap:=wdFindContinue
    End With
    
    With Selection.Find
        .ClearFormatting
        .Text = "<<price1>>"
        .Replacement.ClearFormatting
        .Replacement.Text = "$" + Str(price)
        .Execute Replace:=wdReplaceAll, Forward:=True, Wrap:=wdFindContinue
    End With
    
    With Selection.Find
        .ClearFormatting
        .Text = "<<1 payment>>"
        .Replacement.ClearFormatting
        .Replacement.Text = Format(newDate, "mmmm dd, yyyy")
        .Execute Replace:=wdReplaceAll, Forward:=True, Wrap:=wdFindContinue
    End With
    
    With Selection.Find
        .ClearFormatting
        .Text = "<<price2>>"
        .Replacement.ClearFormatting
        .Replacement.Text = "$" + Str((price / 2))
        .Execute Replace:=wdReplaceAll, Forward:=True, Wrap:=wdFindContinue
    End With
    
    With Selection.Find
        .ClearFormatting
        .Text = "<<2 payment>>"
        .Replacement.ClearFormatting
        .Replacement.Text = Format(newDate, "mmmm dd, yyyy") + " + " + Format(newDate2, "mmmm dd, yyyy")
        .Execute Replace:=wdReplaceAll, Forward:=True, Wrap:=wdFindContinue
    End With
    
    With Selection.Find
        .ClearFormatting
        .Text = "<<price4>>"
        .Replacement.ClearFormatting
        .Replacement.Text = "$" + Str((price / 4))
        .Execute Replace:=wdReplaceAll, Forward:=True, Wrap:=wdFindContinue
    End With
    
    With Selection.Find
        .ClearFormatting
        .Text = "<<4 payment>>"
        .Replacement.ClearFormatting
        .Replacement.Text = Format(newDate, "mmmm dd, yyyy") + " + " + Format(newDate2, "mmmm dd, yyyy") + " + " + Format(newDate3, "mmmm dd, yyyy") + " + " + Format(newDate4, "mmmm dd, yyyy")
        .Execute Replace:=wdReplaceAll, Forward:=True, Wrap:=wdFindContinue
    End With
End Sub

