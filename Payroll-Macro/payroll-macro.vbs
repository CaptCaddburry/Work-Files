Sub future_dates()
'
' future_dates Macro
'
'
    Dim price, shipping As Integer
    Dim total(2) As Double ' total is set as an array, that accepts Double values
    Dim newDates(3) As Date ' newDates is set as an array, that accepts Date values
    Dim pricePlaceholders, datePlaceholders As Variant ' The placeholder variables are going to be used as arrays to run through the for loop
    Dim member As Integer ' member will be used as a step counter for the for loop

    pricePlaceholders = Array("<<price1>>", "<<price2>>", "<<price4>>")
    datePlaceholders = Array("<<1 payment>>", "<<2 payment>>", "<<4 payment>>")
    shipping = 0
    
    ' When you run the macro, it checks which checkboxes are selected and updates price/shipping variables
    '
    If ActiveDocument.FormFields("Check1").CheckBox.Value = True Then ' Check1 = HP Elitebook G5 Laptop
        price = 100
    End If
    If ActiveDocument.FormFields("Check2").CheckBox.Value = True Then ' Check2 = MacBook Pro A1989
        price = 400
    End If
    If ActiveDocument.FormFields("Check3").CheckBox.Value = True Then ' Check3 = Surface Pro 5
        price = 100
    End If
    If ActiveDocument.FormFields("Check4").CheckBox.Value = True Then ' Check4 = HP Elitebook G7
        price = 200
    End If
    If ActiveDocument.FormFields("Check5").CheckBox.Value = True Then ' Check5 = HP Elitebook G8
        price = 400
    End If
    If ActiveDocument.FormFields("Check6").CheckBox.Value = True Then ' Check6 = Surface Pro 7
        price = 300
    End If
    If ActiveDocument.FormFields("Check7").CheckBox.Value = True Then ' Check7 = MacBook Pro A1990
        price = 400
    End If
    If ActiveDocument.FormFields("Check8").CheckBox.Value = True Then ' Check8 = MacBook Pro A2141
        price = 400
    End If
    If ActiveDocument.FormFields("Check9").CheckBox.Value = True Then ' Check9 = MacBook Pro A2338
        price = 400
    End If
    If ActiveDocument.FormFields("Check10").CheckBox.Value = True Then ' Check10 = Shipping Cost
        shipping = 25
    End If

    ' Next, the macro will get the totals for the outputs
    total(0) = (price + shipping) ' 1 Total Payment
    total(1) = (total(0) / 2) ' 2 Total Payments
    total(2) = (total(0) / 4) ' 4 Total Payments

    ' This will take today's date and calculate the following dates to populate on the form
    '
    ' If today's day value is greater than or equal to 15, newDates is going to be the last date of this month
    ' Since every month ends on different days, add 1 to the month value and set the day value to 0
    ' 0 in the day value will output the last day of the previous month
    ' Since we increased the month value by 1, it will give us the last day of the current month
    ' If today's day value is anything else, newDate will be the 15th of this current month
    '
    ' Example:
    ' If today's date is April 18, 2024
    '     newDates(x) = 2024, (April + 1 = May), (Last day of previous month, current month is now May, so last day of April)
    '     newDates(x) = 2024, April, 30
    ' If today's date is April 3, 2024
    '     newDates(x) = 2024, April, 15
    '
    ' Following newDates will use the most recent newDates variable to verify what the next one will be using the same rules

    If Format(Date, "d") >= 15 Then
        newDates(0) = DateSerial(Year(Date), Month(Date) + 1, 0)
        Else
            newDates(0) = DateSerial(Year(Date), Month(Date), 15)
    End If
    
    If Format(newDates(0), "d") = 15 Then
        newDates(1) = DateSerial(Year(newDates(0)), Month(newDates(0)) + 1, 0)
        Else
            newDates(1) = DateSerial(Year(newDates(0)), Month(newDates(0)) + 1, 15)
    End If
    
    If Format(newDates(1), "d") = 15 Then
        newDates(2) = DateSerial(Year(newDates(1)), Month(newDates(1)) + 1, 0)
        Else
            newDates(2) = DateSerial(Year(newDates(1)), Month(newDates(1)) + 1, 15)
    End If
    
    If Format(newDates(2), "d") = 15 Then
        newDates(3) = DateSerial(Year(newDates(2)), Month(newDates(2)) + 1, 0)
        Else
            newDates(3) = DateSerial(Year(newDates(2)), Month(newDates(2)) + 1, 15)
    End If

    ' The following section will find the price and date placeholders and replace them with the updated variables
    ' NewDates will be formatted as (long month, day, long year)
    '
    ' Example:
    ' 4/18/24 -> April 18, 2024
    '
    ' The pricePlaceholder will get replaced everytime, but depending on which value of member is active will replace the proper datePlaceholder

    For member = 0 To 2
        With Selection.Find
            .ClearFormatting
            .Text = pricePlaceholders(member)
            .Replacement.ClearFormatting
            .Replacement.Text = "$" + Str(total(member))
            .Execute Replace:=wdReplaceAll, Forward:=True, Wrap:=wdFindContinue
        End With

        If member = 0 Then
            With Selection.Find
                .ClearFormatting
                .Text = datePlaceholders(member)
                .Replacement.ClearFormatting
                .Replacement.Text = Format(Str(newDates(member)), "mmmm dd, yyyy")
                .Execute Replace:=wdReplaceAll, Forward:=True, Wrap:=wdFindContinue
            End With
        ElseIf member = 1 Then
            With Selection.Find
                .ClearFormatting
                .Text = datePlaceholders(member)
                .Replacement.ClearFormatting
                .Replacement.Text = Format(Str(newDates(member - 1)), "mmmm dd, yyyy") + " + " + Format(Str(newDates(member)), "mmmm dd, yyyy")
                .Execute Replace:=wdReplaceAll, Forward:=True, Wrap:=wdFindContinue
            End With
        ElseIf member = 2 Then
            With Selection.Find
                .ClearFormatting
                .Text = datePlaceholders(member)
                .Replacement.ClearFormatting
                .Replacement.Text = Format(Str(newDates(member - 2)), "mmmm dd, yyyy") + " + " + Format(Str(newDates(member - 1)), "mmmm dd, yyyy") + " + " + Format(Str(newDates(member)), "mmmm dd, yyyy") + " + " + Format(Str(newDates(member + 1)), "mmmm dd, yyyy")
                .Execute Replace:=wdReplaceAll, Forward:=True, Wrap:=wdFindContinue
            End With
        End If
    Next

End Sub
