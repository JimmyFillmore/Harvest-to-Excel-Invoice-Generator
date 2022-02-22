Sub Harvest_To_Invoice_Converter()
'
' Harvest_To_Invoice_Converter Macro
'
' Keyboard Shortcut: Ctrl+Shift+Y


'FOR FUTURE REVIEWERS OF THIS CODE
'The data moves around and rows are added in certain places.
'The order of the code is important because of this.
'Example: Lastrow calculates the length of the table and provides a number
'   This number is then referenced to add to the bottom row. But this number is no longer useful when empty rows are added to the top
'
'Useful data is copy pasted to the right in order, the useless columns are then deleted
'The code will combine notes and Project columns, and then delete the extra column later
'Then it checks for duplicate dates and combines description and hours
'Duplicate rows are then removed
'Top of the invoice is added
'
'At the end of the invoice the Excel Invoice is autosaved as "Month, #invoice number"
' Then the option to save as a pdf is given



'Format Columns and delete unneccessary
    Columns("A:A").Select
    Application.CutCopyMode = False
    Selection.Copy
    Columns(20).Select
    ActiveSheet.Paste
    
    Columns("C:C").Select
    Application.CutCopyMode = False
    Selection.Copy
    Columns(17).Select
    ActiveSheet.Paste
    
    Columns("F:F").Select
    Application.CutCopyMode = False
    Selection.Copy
    Columns(18).Select
    ActiveSheet.Paste
     
    Columns("G:G").Select
    Application.CutCopyMode = False
    Selection.Copy
    Columns(19).Select
    ActiveSheet.Paste
    
    Columns("A:P").Select
    Selection.EntireColumn.Delete


'initialize lastrow to count how long table is
Dim lastrow As Integer
lastrow = Range("A1").End(xlDown).Row + 1

'Hourly rate input
Dim rate As Variant
rate = InputBox("Enter hourly rate (15 or 18):")

'input validation for hourly rate
    While rate <> 15 And rate <> 18
    MsgBox ("Error: Please input rate as either 15 or 18")
    rate = InputBox("Enter hourly rate (15 or 18):")
    Wend
   
'Personal Info Inputs and input validation
Dim name As Variant
name = InputBox("Enter your First and Last name:")
While name = ""
    MsgBox ("Please input a name")
    name = InputBox("Enter your First and Last name:")
Wend

Dim addressPtTwo As Variant
addressPtTwo = InputBox("Enter your City, State, and Zip (ex. Clemson, SC 29631)")
While addressPtTwo = ""
    MsgBox ("Please input your City, State, and Zip")
    addressPtTwo = InputBox("Enter your City, State, and Zip (ex. Clemson, SC 29631)")
Wend

Dim addressPtOne As Variant
addressPtOne = InputBox("Enter your street address")
While addressPtOne = ""
    MsgBox ("Please input your street address")
    addressPtOne = InputBox("Enter your street address")
Wend

Dim invoiceNumber As Variant
invoiceNumber = InputBox("What invoice # is this? (ex. 8)")
While invoiceNumber = ""
    MsgBox ("Please input an invoice #")
    invoiceNumber = InputBox("What invoice # is this? (ex. 8)")
Wend
While IsNumeric(invoiceNumber) = False
    MsgBox ("Please input invoice # as a number")
    invoiceNumber = InputBox("What invoice # is this? (ex. 8)")
Wend

'Input month for title
Dim MyMonth As Variant
MyMonth = InputBox("What is the month for the invoice?")

'Did you submit your hours?
Dim answer As Integer
answer = MsgBox("Did you submit your Harvest weeks for approval?", vbYesNo + vbQuestion, "Harvest Approval")
If answer = vbNo Then
MsgBox ("Please submit your weeks for approval in Harvest")
End If



'Concatenate Description and Notes along with duplicate date combining
Dim i As Integer
i = 2
Do Until i >= lastrow
Cells(i, 1).Value = Cells(i, 1).Value & " " & "(" & Cells(i, 2).Value & ")"
Cells(i, 5).Value = "$" & rate & ".00/hr"

    'if date is the same as next
    If Cells(i, 4).Value = Cells(i + 1, 4).Value Then
    'combine description and notes
    Cells(i + 1, 1).Value = Cells(i + 1, 1).Value & " " & "(" & Cells(i + 1, 2).Value & ")"
    ' combine hours
    Cells(i, 3).Value = Cells(i, 3).Value + Cells(i + 1, 3).Value
    ' combine description with other dates
    Cells(i, 1).Value = Cells(i, 1).Value & vbNewLine & vbNewLine & Cells(i + 1, 1).Value
    ' set dupe hours to 0
    Cells(i + 1, 3).Value = 0
    
        'if date #2
        If Cells(i, 4).Value = Cells(i + 2, 4).Value Then
        'combine description and notes
        Cells(i + 2, 1).Value = Cells(i + 2, 1).Value & " " & "(" & Cells(i + 2, 2).Value & ")"
        ' combine hours
        Cells(i, 3).Value = Cells(i, 3).Value + Cells(i + 2, 3).Value
        ' combine description with other dates
        Cells(i, 1).Value = Cells(i, 1).Value & vbNewLine & vbNewLine & Cells(i + 2, 1).Value
        ' set dupe hours to 0
        Cells(i + 2, 3).Value = 0
    
            'if date #3
            If Cells(i, 4).Value = Cells(i + 3, 4).Value Then
            'combine description and notes
            Cells(i + 3, 1).Value = Cells(i + 3, 1).Value & " " & "(" & Cells(i + 3, 2).Value & ")"
            ' combine hours
            Cells(i, 3).Value = Cells(i, 3).Value + Cells(i + 3, 3).Value
            ' combine description with other dates
            Cells(i, 1).Value = Cells(i, 1).Value & vbNewLine & vbNewLine & Cells(i + 3, 1).Value
            ' set dupe hours to 0
            Cells(i + 3, 3).Value = 0
        
                'if date #4
                If Cells(i, 4).Value = Cells(i + 4, 4).Value Then
                'combine description and notes
                Cells(i + 4, 1).Value = Cells(i + 4, 1).Value & " " & "(" & Cells(i + 4, 2).Value & ")"
                ' combine hours
                Cells(i, 3).Value = Cells(i, 3).Value + Cells(i + 4, 3).Value
                ' combine description with other dates
                Cells(i, 1).Value = Cells(i, 1).Value & vbNewLine & vbNewLine & Cells(i + 4, 1).Value
                ' set dupe hours to 0
                Cells(i + 4, 3).Value = 0
                
                i = i + 1
                End If
                
            i = i + 1
            End If
            
        i = i + 1
        End If
        
    i = i + 1
    End If
    
i = i + 1
Loop


'Calculate total for each day
i = 2
Do Until i >= lastrow

Cells(i, 6).Value = rate * Cells(i, 3).Value

i = i + 1
Loop


'delete excess notes column now that everything is transferred over
Columns("B:B").Select
Selection.EntireColumn.Delete



'Setting Description column to be a certain width and text wrap
Range("A1:A100").WrapText = True
Columns("A").ColumnWidth = 27

'label headers
Range("D1").Value = "Rate"
Range("D1").Font.Bold = True
Range("E1").Value = "Total for Day"
Range("E1").Font.Bold = True
Range("A1") = "Description"
Range("A1").Font.Bold = True
Range("B1").Font.Bold = True
Range("C1").Font.Bold = True

'sum of hours
Dim hours As Variant
hours = WorksheetFunction.Sum(Range("B2:B300"))
Cells(lastrow, 2).Value = hours


'sum of cash
Dim cash As Variant
cash = rate * Cells(lastrow, 2).Value
Cells(lastrow, 5).Value = cash

'bottom row titles
Cells(lastrow, 1).Value = "Total for Invoice"
Cells(lastrow, 1).Font.Bold = True
Cells(lastrow, 3).Value = date_created
Cells(lastrow, 4).Value = "$" & rate & ".00/hr"

'border around all current cells
i = 1
Do Until i >= lastrow + 1
Cells(i, 1).Borders.Weight = xlMedium
Cells(i, 2).Borders.Weight = xlMedium
Cells(i, 3).Borders.Weight = xlMedium
Cells(i, 4).Borders.Weight = xlMedium
Cells(i, 5).Borders.Weight = xlMedium
i = i + 1
Loop

'Autofit rows
Rows("1:300").AutoFit



'delete extra date rows
i = 2
Do Until i >= lastrow
    
    'if date is the same as next
    If Cells(i, 3).Value = Cells(i + 1, 3).Value Then
        
        'if date #2
        If Cells(i, 3).Value = Cells(i + 2, 3).Value Then
        
            'if date #3
            If Cells(i, 3).Value = Cells(i + 3, 3).Value Then
            
                'if date #4
                If Cells(i, 3).Value = Cells(i + 4, 3).Value Then
                Rows(i + 4).Select
                Selection.EntireRow.Delete
                End If
                
            Rows(i + 3).Select
            Selection.EntireRow.Delete
            End If
            
        Rows(i + 2).Select
        Selection.EntireRow.Delete
        End If
        
    Rows(i + 1).Select
    Selection.EntireRow.Delete
    End If
i = i + 1
Loop



'format as currency
Range("E2:E300").NumberFormat = "$#,##0.00"

'Top of Invoice inputs
Rows("1:8").Insert Shift:=x1Down
Range("A1").Value = "Attn: Clayton Survance"
Range("A3").Value = name
Range("A4").Value = addressPtOne
Range("A5").Value = addressPtTwo
Range("D1").Value = "Date Created:"
Range("D3").Value = "Invoice:"
Range("D4").Value = "Terms:"
Range("E4").Value = "Due Upon Receipt"
Range("E1").Value = Date

' invoice number format: #001 or #010 or #100
If invoiceNumber < 10 Then
Range("E3").Value = "#00" & invoiceNumber
ElseIf invoiceNumber < 100 Then Range("E3").Value = "#0" & invoiceNumber
Else
Range("E3").Value = "#" & invoiceNumber
End If

'Total for Invoice top table entry
Range("A7").Value = "Total for Invoice"
Range("A7").Font.Bold = True
Range("B7").Value = hours
Range("D7").Value = "$" & rate & ".00/hr"
Range("E7").Value = cash
Range("E7").NumberFormat = "$#,##0.00"

'border for top table entry
Cells(7, 1).Borders.Weight = xlMedium
Cells(7, 2).Borders.Weight = xlMedium
Cells(7, 3).Borders.Weight = xlMedium
Cells(7, 4).Borders.Weight = xlMedium
Cells(7, 5).Borders.Weight = xlMedium

'align everything to center
Range("A7:E300").HorizontalAlignment = xlCenter

'Alignment of invoice to right
Range("E3").HorizontalAlignment = xlRight

'column width setting
Columns("D").ColumnWidth = 15
Columns("E").ColumnWidth = 15
Columns("B").ColumnWidth = 8.25
Columns("C").ColumnWidth = 12


'Current month is found from Date
'Dim mon As Variant
'mon = Month(Date)
'Dim MyMonth As Variant
'If mon = 1 Then
'    MyMonth = "January"
'ElseIf mon = 2 Then
'    MyMonth = "February"
'ElseIf mon = 3 Then
'    MyMonth = "March"
'ElseIf mon = 4 Then
'    MyMonth = "April"
'ElseIf mon = 5 Then
'    MyMonth = "May"
'ElseIf mon = 6 Then
'    MyMonth = "June"
'ElseIf mon = 7 Then
'    MyMonth = "July"
'ElseIf mon = 8 Then
'    MyMonth = "August"
'ElseIf mon = 9 Then
'    MyMonth = "September"
'ElseIf mon = 10 Then
'    MyMonth = "October"
'ElseIf mon = 11 Then
'    MyMonth = "November"
'ElseIf mon = 12 Then
'    MyMonth = "December"
'End If

'Saves File as "Month, invoice #"
Dim INVnum As Variant
INVnum = Range("E3")
Dim FileTitle As Variant
FileTitle = MyMonth & ", Invoice " & INVnum
ActiveWorkbook.SaveAs (FileTitle)

' select folder for file
Dim myFile As Variant
myFile = Application.GetSaveAsFilename _
    (InitialFileName:=FileTitle, _
        FileFilter:="PDF Files (*.pdf), *.pdf", _
        Title:="Select Folder and FileName to save")

'export to PDF if a folder was selected
If myFile <> "False" Then
    ActiveWorkbook.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        Filename:=myFile, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=False
    'confirmation message with file info
    MsgBox "PDF file has been created: " _
      & vbCrLf _
      & myFile
End If


'Review works
MsgBox ("Please review total money, hours, dates and personal information to verify all are correct and there are no formatting errors.")
End Sub


