Attribute VB_Name = "Module1"
Option Explicit

' Validate Cell Values
' of Trainee Schedule 

Function Validate() As Boolean

    Dim frm As Worksheet
    
    Set frm = ThisWorkbook.Sheets("Trainee Schedule")
    
    Validate = True
    
    With frm
    
        Range("D2:D5").Interior.Color = xlNone
        Range("F2:F5").Interior.Color = xlNone
    
    End With
    
    ' Validating Employee Name 
    If Trim(frm.Range("D2").Value) = "" Then
        MsgBox "Employee Name is blank.", vbOKOnly + vbInformation, "Employee Name"
        frm.Range("D2").Select
        frm.Range("D2").Interior.Color = vbRed
        Validate = False
        Exit Function
    End If

    ' Validating Employee Number
     If Trim(frm.Range("D3").Value) = "" Then
        MsgBox "Employee Number is blank.", vbOKOnly + vbInformation, "Employee Number"
        frm.Range("D3").Select
        frm.Range("D3").Interior.Color = vbRed
        Validate = False
        Exit Function
    End If

    ' Validating Hire Date
    If Trim(frm.Range("D4").Value) = "" Or Not IsDate(Trim(frm.Range("D4").Value)) Then
        MsgBox "Please Enter Valid Hire Date.", vbOKOnly + vbInformation, "Hire Date"
        frm.Range("D4").Select
        frm.Range("D4").Interior.Color = vbRed
        Validate = False
        Exit Function
    End If

End Function

' Validate Employee Number
' Used only Once
Function ValEmpNum(TstEmpNum)
    msgbox "TEST"
    ' Open Sheet TraineeDatabase

    ' Get Len of Employees in TraineeDatabase

    ' For each employee in TraineeDatabase
        
        ' if TstEmpNum = Employee Number
            ' Valid = False
            ' Send Error Message
            ' Dont Save
        
        ' ELSE PASS

End Function

' Validating Cells
' of Trainer Schedule

Function ValTnR(TnrRow, i As Integer, frm As Worksheet)

If TnrRow = 0 Then
    
    MsgBox "No such Trainer record found!", vbOKOnly + vbCritical, "No Record"
            
    frm.Range("F" & i).Select
    frm.Range("F" & i).Interior.Color = vbRed
                
    Exit Function
        
End If

End Function

' Clear Cells 
' Trainee Schedule

Sub TraineeReset()

    With Sheets("Trainee Schedule")
    
        .Range("D2:D5").Interior.Color = xlNone
        .Range("D2:D5").Value = ""
        
        .Range("F2:F5").Interior.Color = xlNone
        .Range("F2:F5").Value = ""
        
        .Range("B11:B71").Interior.Color = xlNone
        .Range("B11:B71").Value = ""
        
        .Range("D11:F71").Interior.Color = xlNone
        .Range("D11:F71").Value = ""
        
        .Range("M2:N2").Value = ""
        .Range("H13").Value = ""
        .Range("H15").Value = ""
    
    End With

End Sub

' Clear Cells
' Trainer Schedule

Sub TrnRReset()

    With Sheets("Trainer Schedule")
    
        .Range("C2").Interior.Color = xlNone
        .Range("C2").Value = ""
        
        .Range("F2").Interior.Color = xlNone
        .Range("F2").Value = ""
        
        .Range("B8:B71").Interior.Color = xlNone
        .Range("B8:B71").Value = ""
        
        .Range("D8:F71").Interior.Color = xlNone
        .Range("D8:F71").Value = ""
        
        .Range("I4:I5").Value = ""
    
    End With

End Sub

' Save Employee Info
' from Trainee Schedule

Sub Save()

    ' Declare Worksheets
    Dim frm As Worksheet
    Dim TnEdb As Worksheet
    ' Declare numeric values to serialize data
    Dim iRow As Long
    Dim iSerial As Long
    ' Point to worksheets
    Set frm = ThisWorkbook.Sheets("Trainee Schedule")
    Set TnEdb = ThisWorkbook.Sheets("Trainee Database")
    
    ' Check for existing iRow and iSerial values
    ' Else assign new iRow and iSerial
    If Trim(frm.Range("N2").Value) = "" Then
        iRow = TnEdb.Range("A" & Application.Rows.Count).End(xlUp).Row + 1
        iSerial = TnEdb.Cells(iRow - 1, 1).Value + 1
        Call ValEmpNum(frm.Range("D3").Value)
    Else
        iRow = frm.Range("M2").Value
        iSerial = frm.Range("N2").Value
    End If
    
    ' Populate Trainee Database with Data from Form
    With TnEdb
        .Cells(iRow, 1).Value = iSerial
        'EE Number
        .Cells(iRow, 2).Value = frm.Range("D3").Value
        'EE Name
        .Cells(iRow, 3).Value = frm.Range("D2").Value
        'Hire Date
        .Cells(iRow, 4).Value = frm.Range("D4").Value
        'Locker Number
        .Cells(iRow, 5).Value = frm.Range("D5").Value
        'Computer Start
        .Cells(iRow, 6).Value = frm.Range("F2").Value
        'Computer End
        .Cells(iRow, 7).Value = frm.Range("F3").Value
        'Dallas Training
        .Cells(iRow, 8).Value = frm.Range("F4").Value
        'Training Completion
        .Cells(iRow, 9).Value = frm.Range("F5").Value
    End With

End Sub

' Save Schedule
' from Trainee Schedule Form

Sub Schedule()

    ' Declare and point to Worksheet
    Dim frm As Worksheet
    Set frm = ThisWorkbook.Sheets("Trainee Schedule")
    ' Declare Month and Day Values for reference
    Dim Month As String
    Dim Day As Long
    ' Declare Values to reference cells for where to find schedule
    Dim TrnStart As Long
    Dim TrnLen As Long

    TrnStart = 11
    TrnLen = frm.Range("B" & Application.Rows.Count).End(xlUp).Row
    
    ' Declare Variables 
    Dim TnEname As String
    Dim TnEnum As Long
    Dim Hours As String
    Dim JbDuty As String
    Dim TnRname As String
    Dim TnRnum As Long
    Dim TnRtime As String

    TnEname = frm.Range("D2")
    TnEnum = frm.Range("D3")
    
    'Counter for workdays on form
    Dim i As Integer
    'Counter for Trainees on a Date
    Dim x As Integer
    'Counter for Trainee info position
    Dim y As Integer
    
    ' Declare Variables to find cells
    Dim FndSheet As Worksheet
    Dim FndDay As Long
    Dim Daylen As Long
    Dim TnrRow As Long
    
    'Evaluate every workday on form
    For i = TrnStart To TrnLen
    
        'Find Needed SHEET and DAY for each workday
        Month = Format(Range("B" & i), "mmmm yyyy")
        Day = Format(Range("B" & i), "d") + 1
        
        Hours = frm.Range("D" & i)
        JbDuty = frm.Range("E" & i)
        
        'Evaluate Trainer Info
        
        TnRname = frm.Range("F" & i)
        TnRtime = "PLACE HOLDER"
    
        TnrRow = Application.WorksheetFunction.IfError(Application.Match(TnRname, Sheets("Trainer Database").Range("B:B"), 0), 0)
        Call ValTnR(TnrRow, i, frm)
        
        TnRnum = ThisWorkbook.Sheets("Trainer Database").Cells(TnrRow, 1).Value
        
        'Find Needed SHEET
        Set FndSheet = ThisWorkbook.Sheets(Month)
        'Open Needed SHEET
        With FndSheet
        
            Daylen = FndSheet.Cells(Application.Rows.Count, Day).End(xlUp).Row + 2
            x = 3
            
            'Evaluate Each Trainee for that workday
            Do While x <= Daylen
                If Trim(FndSheet.Cells(x, Day).Value) = "" Or Trim(FndSheet.Cells(x, Day).Value) = TnEname Then    
                    y = x
                    .Cells(y, Day).Value = TnEname
                    y = x + 1
                    .Cells(y, Day).Value = TnEnum
                    y = x + 2
                    .Cells(y, Day).Value = Hours
                    y = x + 3
                    .Cells(y, Day).Value = JbDuty
                    y = x + 4
                    .Cells(y, Day).Value = TnRname
                    y = x + 5
                    .Cells(y, Day).Value = TnRnum
                    y = x + 6
                    .Cells(y, Day).Value = TnRtime
                    Exit Do
                End If
                ' Check next Trainee
                x = x + 8
            Loop
        End With
    Next i
End Sub





Sub Update()

    Dim iRow As Long
    Dim iSerial As Long
    Dim EEnum As Long
    
    EEnum = Application.InputBox("Please enter Trainee Employee Number.", "Select Employee", , , , , , 1)
    
    'iSerial = Application.InputBox("Please enter Serial Number to make modification.", "Modify", , , , , , 1)
    
    On Error Resume Next
    
    iSerial = Application.WorksheetFunction.IfError(Application.Match(EEnum, ThisWorkbook.Sheets("Trainee Database").Range("B:B"), 0), 0) - 1
    
    iRow = Application.WorksheetFunction.IfError(Application.WorksheetFunction.Match(iSerial, Sheets("Trainee Database").Range("A:A")), 0)
    
    If iRow = 0 Then
    
        MsgBox "No record found.", vbOKOnly + vbCritical, "No Record"
        Exit Sub
        
    End If
        
    Sheets("Trainee Schedule").Range("M2").Value = iRow
    Sheets("Trainee Schedule").Range("N2").Value = iSerial
    
    Sheets("Trainee Schedule").Range("D3").Value = Sheets("Trainee Database").Cells(iRow, 2).Value
    'Employee Name
    Sheets("Trainee Schedule").Range("D2").Value = Sheets("Trainee Database").Cells(iRow, 3).Value
    'Employee Hire Date
    Sheets("Trainee Schedule").Range("D4").Value = Sheets("Trainee Database").Cells(iRow, 4).Value
    'Locker Number
    Sheets("Trainee Schedule").Range("D5").Value = Sheets("Trainee Database").Cells(iRow, 5).Value
    'Comp Start
    Sheets("Trainee Schedule").Range("F2").Value = Sheets("Trainee Database").Cells(iRow, 6).Value
    'Comp End
    Sheets("Trainee Schedule").Range("F3").Value = Sheets("Trainee Database").Cells(iRow, 7).Value
    'Dallas
    Sheets("Trainee Schedule").Range("F4").Value = Sheets("Trainee Database").Cells(iRow, 8).Value
    'All Training
    Sheets("Trainee Schedule").Range("F5").Value = Sheets("Trainee Database").Cells(iRow, 9).Value
    
End Sub





Sub LoadSchedule()

Dim st As Range
Dim en As Range
Dim x As Integer
Dim stDate As Date
Dim enDate As Date
Dim d As Date
Dim numRows As Long

numRows = Range("F2", Range("F2").End(xlDown)).Rows.Count

Dim SheetMonth As String
Dim SheetDay As Long
Dim TnEname As String
Dim FndSheet As Worksheet
Dim FrmRow As Long

Dim Hours As String
Dim JbDuty As String
Dim TnRname As String

Dim Daylen As Long
Dim y As Integer
Dim z As Integer

TnEname = Range("D2").Value

Set st = Range("I13").Offset(x, 0)
Set en = Range("I15").Offset(x, 0)
stDate = DateSerial(Year(st), Month(st), Day(st))
enDate = DateSerial(Year(en), Month(en), Day(en))

'loop through the dates as necessary
For d = stDate To enDate + 1

    SheetMonth = Format(d, "mmmm yyyy")
    SheetDay = Format(d, "d")
    'MsgBox SheetMonth & "   " & SheetDay
    
    'Find Needed SHEET
    Set FndSheet = ThisWorkbook.Sheets(SheetMonth)
        
    'Open Needed SHEET
    With FndSheet
    
        Daylen = FndSheet.Cells(Application.Rows.Count, SheetDay).End(xlUp).Row + 2
        y = 3
            
            'Evaluate Each Trainee for that workday
            Do While y <= Daylen
                
                If Trim(FndSheet.Cells(y, SheetDay).Value) = TnEname Then
                
                    'Get hours jobduty and trainer
                    z = y + 2
                    Hours = .Cells(z, SheetDay).Value
                    z = z + 1
                    JbDuty = .Cells(z, SheetDay).Value
                    z = z + 1
                    TnRname = .Cells(z, SheetDay).Value
                    
                    FrmRow = ThisWorkbook.Sheets("Trainee Schedule").Range("B" & Application.Rows.Count).End(xlUp).Row + 1
                    
                    ThisWorkbook.Sheets("Trainee Schedule").Range("B" & FrmRow).Value = d - 1
                    ThisWorkbook.Sheets("Trainee Schedule").Range("D" & FrmRow).Value = Hours
                    ThisWorkbook.Sheets("Trainee Schedule").Range("E" & FrmRow).Value = JbDuty
                    ThisWorkbook.Sheets("Trainee Schedule").Range("F" & FrmRow).Value = TnRname
                    
                    Exit Do
                
                End If
                
                y = y + 8
                
            Loop
        
    End With
        
Next

End Sub





Sub LoadTnR()

Dim st As Range
Dim en As Range
Dim x As Integer
Dim stDate As Date
Dim enDate As Date
Dim d As Date
Dim numRows As Long

numRows = Range("F2", Range("F2").End(xlDown)).Rows.Count

Dim SheetMonth As String
Dim SheetDay As Long
Dim TnRname As String
Dim FndSheet As Worksheet
Dim FrmRow As Long

Dim Hours As String
Dim JbDuty As String
Dim TnEname As String

Dim Daylen As Long
Dim y As Integer
Dim z As Integer

TnRname = Range("C2").Value

Set st = Range("I4").Offset(x, 0)
Set en = Range("I5").Offset(x, 0)
stDate = DateSerial(Year(st), Month(st), Day(st))
enDate = DateSerial(Year(en), Month(en), Day(en))

'loop through the dates as necessary
For d = stDate To enDate + 1

    SheetMonth = Format(d, "mmmm yyyy")
    SheetDay = Format(d, "d")
    'MsgBox SheetMonth & "   " & SheetDay
    
    'Find Needed SHEET
    Set FndSheet = ThisWorkbook.Sheets(SheetMonth)
        
    'Open Needed SHEET
    With FndSheet
    
        Daylen = FndSheet.Cells(Application.Rows.Count, SheetDay).End(xlUp).Row + 2
        y = 7
            
            'Evaluate Each Trainee for that workday
            Do While y <= Daylen
                
                If Trim(FndSheet.Cells(y, SheetDay).Value) = TnRname Then
                
                    'Get hours jobduty and trainer
                    z = y - 2
                    Hours = .Cells(z, SheetDay).Value
                    z = z + 1
                    JbDuty = .Cells(z, SheetDay).Value
                    z = y - 4
                    TnEname = .Cells(z, SheetDay).Value
                    
                    FrmRow = ThisWorkbook.Sheets("Trainer Schedule").Range("B" & Application.Rows.Count).End(xlUp).Row + 1
                    
                    ThisWorkbook.Sheets("Trainer Schedule").Range("B" & FrmRow).Value = d - 1
                    ThisWorkbook.Sheets("Trainer Schedule").Range("D" & FrmRow).Value = Hours
                    ThisWorkbook.Sheets("Trainer Schedule").Range("E" & FrmRow).Value = JbDuty
                    ThisWorkbook.Sheets("Trainer Schedule").Range("F" & FrmRow).Value = TnEname
                    
                    Exit Do
                
                End If
                
                y = y + 8
                
            Loop
        
    End With
        
Next

End Sub





Sub AddMonth()

    Dim ws As Worksheet
    Dim wsNew As Worksheet
    
    Set ws = Sheets("Monthly Training Schedule")
    ws.Copy After:=Sheets("Sheet3")
    Set wsNew = Sheets(Sheets("Sheet3").Index + 1)
    wsNew.Name = "Test"

End Sub
