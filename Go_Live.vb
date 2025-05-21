Option Explicit

' Main procedure to set up the task schedule
Sub CreateTaskSchedule()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("GanttChart")
    
    With ws
        .Range("B6").Value = "TASK"
        .Range("E6").Value = "START"
        .Range("F6").Value = "END"
        .Range("G6").Value = "% COMPLETE"
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
    
    ' Check if data already exists in any of the row sets
    If ws.Range("B9").Value = "" And ws.Range("B20").Value = "" And ws.Range("B33").Value = "" Then
        ' Add initial task data to specific rows if cells are empty
        ' First set (rows 9-12)
        AddTaskData "Knowledge Sharing", #3/25/2025#, #5/14/2025#, 9
        AddTaskData "Shadowing", #5/6/2025#, #5/30/2025#, 10
        AddTaskData "Reverse Shadowing", #5/6/2025#, #5/30/2025#, 11
        AddTaskData "Go Live", #5/30/2025#, #12/31/2025#, 12
        
        ' Second set (rows 20-23)
        AddTaskData "Knowledge Sharing", #3/25/2025#, #5/14/2025#, 20
        AddTaskData "Shadowing", #5/6/2025#, #5/30/2025#, 21
        AddTaskData "Reverse Shadowing", #5/6/2025#, #5/30/2025#, 22
        AddTaskData "Go Live", #5/30/2025#, #12/31/2025#, 23
        
        ' Third set (rows 33-36)
        AddTaskData "Knowledge Sharing", #3/25/2025#, #5/14/2025#, 33
        AddTaskData "Shadowing", #5/6/2025#, #5/30/2025#, 34
        AddTaskData "Reverse Shadowing", #5/6/2025#, #5/30/2025#, 35
        AddTaskData "Go Live", #5/30/2025#, #12/31/2025#, 36
    Else
        ' Read existing data from cells and process it
        ProcessExistingData
    End If
    
    ' Auto-fit columns
    ws.Columns("B:H").AutoFit
    
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
    
    ' Calculate percentage of work done based on weeks
    percentageWorkDone = CalculatePercentage(startDate, endDate)
    ws.Cells(rowNum, 7).Value = percentageWorkDone & "%"
    
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
    Dim rowSets As Variant
    Dim rowSet As Variant
    Dim startRow As Integer
    Dim endRow As Integer
    
    Set ws = ThisWorkbook.Sheets("GanttChart")
    
    ' Define the sets of rows to process
    rowSets = Array(Array(9, 12), Array(20, 23), Array(33, 36))
    
    ' Process each set of rows
    For Each rowSet In rowSets
        startRow = rowSet(0)
        endRow = rowSet(1)
        
        ' Process rows in this set
        For rowNum = startRow To endRow
            ' Get task name and dates from the cells
            taskName = ws.Cells(rowNum, 2).Value
            
            ' Check if there's a valid date in the cell
            If IsDate(ws.Cells(rowNum, 5).Value) And IsDate(ws.Cells(rowNum, 6).Value) Then
                startDate = ws.Cells(rowNum, 5).Value
                endDate = ws.Cells(rowNum, 6).Value
                
                ' Calculate start and end columns based on dates
                startCol = CalculateWeekColumn(startDate)
                endCol = CalculateWeekColumn(endDate)
                
                ' Special handling for dates that fall on the 31st of a month
                If Day(endDate) = 31 Then
                    endCol = endCol - 1
                    Debug.Print "Adjusted end column for date " & Format(endDate, "dd-mm-yyyy") & " from " & (endCol + 1) & " to " & endCol
                End If
                
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
                Debug.Print "Start Date: " & Format(startDate, "dd-mm-yyyy") & ", End Date: " & Format(endDate, "dd-mm-yyyy")
                Debug.Print "Start Column: " & startCol & ", End Column: " & endCol
            End If
        Next rowNum
    Next rowSet
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
    Debug.Print ("weekNumber" & weekNumber & inputDate)
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

' Process existing data in the cells
Sub ProcessExistingData()
    Dim ws As Worksheet
    Dim rowNum As Integer
    Dim taskName As String
    Dim startDate As Date
    Dim endDate As Date
    Dim percentageWorkDone As Double
    Dim rowSets As Variant
    Dim rowSet As Variant
    Dim startRow As Integer
    Dim endRow As Integer
    
    Set ws = ThisWorkbook.Sheets("GanttChart")
    
    ' Define the sets of rows to process
    rowSets = Array(Array(9, 12), Array(20, 23), Array(33, 36))
    
    ' Process each set of rows
    For Each rowSet In rowSets
        startRow = rowSet(0)
        endRow = rowSet(1)
        
        ' Process rows in this set
        For rowNum = startRow To endRow
            ' Check if there's data in this row
            If ws.Cells(rowNum, 2).Value <> "" Then
                taskName = ws.Cells(rowNum, 2).Value
                
                ' Check if there are valid dates
                If IsDate(ws.Cells(rowNum, 5).Value) And IsDate(ws.Cells(rowNum, 6).Value) Then
                    startDate = ws.Cells(rowNum, 5).Value
                    endDate = ws.Cells(rowNum, 6).Value
                    
                    ' Calculate business days (excluding weekends)
                    ws.Cells(rowNum, 8).Formula = "=NETWORKDAYS(E" & rowNum & ",F" & rowNum & ")"
                    
                    ' Calculate percentage of work done based on weeks
                    percentageWorkDone = CalculatePercentage(startDate, endDate)
                    ws.Cells(rowNum, 7).Value = percentageWorkDone & "%"
                    
                    ' Format date cells
                    ws.Cells(rowNum, 5).NumberFormat = "dd-mm-yyyy"
                    ws.Cells(rowNum, 6).NumberFormat = "dd-mm-yyyy"
                    
                    ' Format percentage cell
                    ws.Cells(rowNum, 7).NumberFormat = "0%"
                    
                    ' Debug information
                    Debug.Print "Processing existing data: " & taskName & ", " & startDate & " to " & endDate & " (Row " & rowNum & ")"
                End If
            End If
        Next rowNum
    Next rowSet
End Sub

Function CalculatePercentage(startDate As Date, endDate As Date) As Double
    Dim currentDate As Date
    Dim totalWeeks As Integer
    Dim completedWeeks As Integer
    Dim percentage As Double
    
    ' Get current date
    currentDate = Date
    
    ' Calculate total weeks between start and end dates
    totalWeeks = Int((endDate - startDate) / 7) + 1
    
    ' If the task hasn't started yet
    If currentDate < startDate Then
        percentage = 0
    ' If the task is already completed
    ElseIf currentDate >= endDate Then
        percentage = 100
    ' If the task is in progress
    Else
        ' Calculate completed weeks (from start date to current date)
        completedWeeks = Int((currentDate - startDate) / 7) + 1
        
        ' Calculate percentage
        percentage = (completedWeeks / totalWeeks) * 100
        
        ' Ensure percentage doesn't exceed 100%
        If percentage > 100 Then
            percentage = 100
            Debug.Print (percentage)
        End If
    End If
    
    ' Round to nearest integer
    CalculatePercentage = Round(percentage, 0)
End Function
