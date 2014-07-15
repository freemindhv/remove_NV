Attribute VB_Name = "delete_NV"
Sub delete_NV()
'by dgreiner @ NITS


''SETUP
'''sheet to work with
wsheet = 2
'''column to work with
wcol = 10
'''row to start at
x = 2
'''get list end
sheetend = Sheets(wsheet).Cells(Rows.Count, wcol).End(xlUp).Row

'''Spin that shit


For x = 2 To sheetend
    If Sheets(wsheet).Cells(x, wcol).Text = "#NV" Then
        Rows(x).Delete Shift:=xlUp
        x = x - 1
    End If
Next

End Sub

