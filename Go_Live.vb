Option Explicit

' Main procedure to set up the task schedule
Sub CreateTaskSchedule()
    With ThisWorkbook.Sheets("GanttChart")
        .Range("B6").Value = "TASK"
        .Range("E6").Value = "START"
        .Range("F6").Value = "END"
        .Range("H6").Value = "DURATION (DAYS)"
        
        ' Add headers for weeks
        Dim i As Integer
        For i = 0 To 42 ' 43 weeks (about 10 months)
            .Cells(6, 10 + i).Value = "Week " & (i + 1)
        Next i
        
        ' Format headers
        .Range("B6:H6").Font.Bold = True
        .Range("J6:AZ6").Font.Bold = True
        .Range("B6:AZ6").Borders.Weight = xlThin
    End With
    
    ' Add initial task data to specific rows
    AddTaskData "Knowledge Sharing", #3/25/2025#, #5/14/2025#, 9
    AddTaskData "Shadowing", #5/6/2025#, #5/30/2025#, 10
    AddTaskData "Reverse Shadowing", #5/6/2025#, #5/30/2025#, 11
    AddTaskData "Go Live", #5/30/2025#, #12/31/2025#, 12
    
    ' Auto-fit columns
    ThisWorkbook.Sheets("GanttChart").Columns("B:H").AutoFit
    
    ' Now color the cells based on the dates in the cells
    ColorTaskRows
    
    MsgBox "Task schedule created successfully!", vbInformation
End Sub

' Add a task to the table
Sub AddTaskData(taskName As String, startDate As Date, endDate As Date, rowNum As Integer)
    Dim ws As Worksheet
    Dim businessDays As Integer
    
    Set ws = ThisWorkbook.Sheets("GanttChart")
    
    ' Add task data to the specified row
    ws.Cells(rowNum, 2).Value = taskName                ' Column B
    ws.Cells(rowNum, 5).Value = startDate               ' Column E
    ws.Cells(rowNum, 6).Value = endDate                 ' Column F
    ' ws.Cells(rowNum, 8).Formula = "=F" & rowNum & "-E" & rowNum & "+1" ' Column H - Calculate duration
    
    businessDays = NetworkDays(startDate, endDate)
    ws.Cells(rowNum, 8).Value = businessDays
    
    ' Format date cells
    ws.Cells(rowNum, 5).NumberFormat = "dd-mm-yyyy"     ' Format as dd-mm-yyyy
    ws.Cells(rowNum, 6).NumberFormat = "dd-mm-yyyy"     ' Format as dd-mm-yyyy
    
    ' Add borders to the task row
    'ws.Range(ws.Cells(rowNum, 2), ws.Cells(rowNum, 52)).Borders.LineStyle = xlContinuous
    'ws.Range(ws.Cells(rowNum, 2), ws.Cells(rowNum, 52)).Borders.Weight = xlThin
End Sub

' Color task rows based on the dates in the cells
Sub ColorTaskRows()
    Dim ws As Worksheet
    Dim startDate As Date
    Dim endDate As Date
    Dim taskName As String
    Dim startCol As Long
    Dim endCol As Long
    Dim taskColor As Long
    Dim i As Long
    Dim rowNum As Long
    
    Set ws = ThisWorkbook.Sheets("GanttChart")
    
    ' Process each task row (9-12)
    For rowNum = 9 To 12
        ' Get task name and dates from the cells
        taskName = ws.Cells(rowNum, 2).Value
        
        ' Check if there's a valid date in the cell
        If IsDate(ws.Cells(rowNum, 5).Value) And IsDate(ws.Cells(rowNum, 6).Value) Then
            startDate = ws.Cells(rowNum, 5).Value
            endDate = ws.Cells(rowNum, 6).Value
            
            ' Calculate start and end columns based on dates
            startCol = CalculateWeekColumn(startDate)
            endCol = CalculateWeekColumn(endDate)
            
            ' Set color based on task name
            Select Case taskName
                Case "Knowledge Sharing"
                    taskColor = RGB(186, 184, 108) ' Olive green
                Case "Shadowing", "Reverse Shadowing"
                    taskColor = RGB(210, 180, 140) ' Tan
                Case "Go Live"
                    taskColor = RGB(0, 0, 255) ' Blue
                Case Else
                    taskColor = RGB(200, 200, 200) ' Default light gray
            End Select
            
            ' Fill the cells with the appropriate color
            For i = startCol To endCol
                ws.Cells(rowNum, i).Interior.Color = taskColor
            Next i
            
            ' Debug information
            Debug.Print "Task: " & taskName & " (Row " & rowNum & ")"
            Debug.Print "Start Date: " & startDate & ", End Date: " & endDate
            Debug.Print "Start Column: " & startCol & ", End Column: " & endCol
        End If
    Next rowNum
End Sub

' Calculate the column number for a given date
Function CalculateWeekColumn(inputDate As Date) As Long
    Dim referenceDate As Date
    Dim weekNumber As Long
    
    ' Reference date is March 1, 2025
    referenceDate = #3/1/2025#
    
    ' Calculate week number (0-based)
    weekNumber = Int((inputDate - referenceDate) / 7)
    
    ' Column J (10) is Week 1, so add 10 to the week number
    CalculateWeekColumn = 10 + weekNumber
End Function

Function NetworkDays(startDate As Date, endDate As Date) As Integer
    Dim currentDate As Date
    Dim days As Integer
    
    days = 0
    currentDate = startDate
    
    ' Loop through each day
    Do While currentDate <= endDate
        ' Check if it's not a weekend (1=Sunday, 7=Saturday in VBA)
        If Weekday(currentDate) <> 1 And Weekday(currentDate) <> 7 Then
            days = days + 1
        End If
        currentDate = currentDate + 1
    Loop
    
    NetworkDays = days
End Function