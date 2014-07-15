'uncomment the following line when you want to import this into a excel module(file extension must be .bas)
'Attribute VB_Name = "delete_NV"
Sub delete_NV()
'by dgreiner @ NITS


''SETUP
Dim wsheet, wcol, sheetend, x, i As Integer

''CONFIG
'''sheet to work with
wsheet = 2
'''column to work with
wcol = 10
'''row to start at
x = 2
'''get list end
sheetend = Sheets(wsheet).Cells(Rows.Count, wcol).End(xlUp).Row

'''Spin that shit


For i = x To sheetend Step 1
    If Sheets(wsheet).Cells(i, wcol).Text = "#NV" Then
        Rows(i).Delete Shift:=xlUp
        i = i - 1
    End If
Next

End Sub