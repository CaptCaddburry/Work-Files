Sub future_dates()
'
' future_dates Macro
'
'
    Dim total(2) As Double ' total is set as an array, that accepts Double values
    Dim newDates(3) As Date ' newDates is set as an array, that accepts Date values
    Dim pricePlaceholders, datePlaceholders, priceList As Variant ' The placeholder variables are going to be used as arrays to run through the for loop
    Dim member As Integer ' member will be used as a step counter for the for loop

    pricePlaceholders = Array("<<price1>>", "<<price2>>", "<<price4>>")
    datePlaceholders = Array("<<1 payment>>", "<<2 payment>>", "<<4 payment>>")
    ' priceList = (HP Elitebook G5 Laptop, MacBook Pro A1989, Surface Pro 5, HP Elitebook G7, HP Elitebook G8, Surface Pro 7, MacBook Pro A1990, MacBook Pro A2141, MacBook Pro A2338, Shipping Cost)
    priceList = Array(100, 400, 100, 200, 400, 300, 400, 400, 400, 25)

    ' When you run the macro, it checks which checkboxes are selected and updates the total variable
    For member = 0 To 9
        If ActiveDocument.FormFields("Check" + CStr(member + 1)).CheckBox.Value = True Then
            total(0) = total(0) + priceList(member)
        End If
    Next

    ' Next, the macro will get the totals for the outputs
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
    
    For member = 0 To 2
        If Format(newDates(member), "d") = 15 Then
            newDates(member + 1) = DateSerial(Year(newDates(member)), Month(newDates(member)) + 1, 0)
        Else
            newDates(member + 1) = DateSerial(Year(newDates(member)), Month(newDates(member)) + 1, 15)
        End If
    Next

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
            .Replacement.Text = "$" + CStr(total(member))
            .Execute Replace:=wdReplaceAll, Forward:=True, Wrap:=wdFindContinue
        End With

        If member = 0 Then
            With Selection.Find
                .ClearFormatting
                .Text = datePlaceholders(member)
                .Replacement.ClearFormatting
                .Replacement.Text = Format(CStr(newDates(member)), "mmmm dd, yyyy")
                .Execute Replace:=wdReplaceAll, Forward:=True, Wrap:=wdFindContinue
            End With
        ElseIf member = 1 Then
            With Selection.Find
                .ClearFormatting
                .Text = datePlaceholders(member)
                .Replacement.ClearFormatting
                .Replacement.Text = Format(CStr(newDates(member - 1)), "mmmm dd, yyyy") + " + " + Format(CStr(newDates(member)), "mmmm dd, yyyy")
                .Execute Replace:=wdReplaceAll, Forward:=True, Wrap:=wdFindContinue
            End With
        ElseIf member = 2 Then
            With Selection.Find
                .ClearFormatting
                .Text = datePlaceholders(member)
                .Replacement.ClearFormatting
                .Replacement.Text = Format(CStr(newDates(member - 2)), "mmmm dd, yyyy") + " + " + Format(CStr(newDates(member - 1)), "mmmm dd, yyyy") + " + " + Format(CStr(newDates(member)), "mmmm dd, yyyy") + " + " + Format(CStr(newDates(member + 1)), "mmmm dd, yyyy")
                .Execute Replace:=wdReplaceAll, Forward:=True, Wrap:=wdFindContinue
            End With
        End If
    Next

End Sub
