VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FillMeIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Const TableStartRowIndex = 8
Dim dateColumns(1 To 2) As String, weekdayColumns(1 To 2) As String
Const MaxNumberRowsInTable = 16 ' Maximum number of days you want to automate filling cells
Const MaxDaysInAMonth = 31 ' This is for reseting cells if the month doesnt have the maximum number of days

Sub Init()
    ' Specify these column names in your worksheet, so that this macro will exactly fill the right cells based on your customized file
    dateColumns(1) = "C"
    dateColumns(2) = "I"
    
    weekdayColumns(1) = "B"
    weekdayColumns(2) = "H"
End Sub
Function GetDropdownItems(ByVal cell As range) As String()
    Dim validationList As String
    Dim options() As String
    'Check if the cell has validation (dropdown)
    If cell.Validation.Type = xlValidateList Then
        'Get the validation formula
        validationList = cell.Validation.Formula1
        
        'If the validation formula is a Range, convert it to an array
        If Left(validationList, 1) = "=" Then
            validationList = Mid(validationList, 2)
            options = Split(Evaluate(validationList), ",")
        Else 'If the validation formula is a list of items separated by commas
            options = Split(validationList, ",")
        End If
    Else
        ReDim options(0)
    End If
    GetDropdownItems = options
End Function

Function DateString(year As Integer, month As Integer, day As Integer)
    DateString = year & "/" & IIf(month < 10, "0" & month, month) & "/" & IIf(day < 10, "0" & day, day)
End Function

Function DaysInMonth(month As Integer, year As Integer)
    DaysInMonth = IIf(month < 7, 31, _
        IIf(month < 12, 30, _
            IIf(year Mod 4 <> 0, 29, 30) _
        ) _
    )
    
End Function

Sub ClearCells()
    Dim dateCell As range, weekdayRange As range
    Dim i As Integer, c As Integer
    For c = 1 To UBound(dateColumns)
        Set dateCell = Me.range(dateColumns(c) & TableStartRowIndex)
        Set weekdayRange = Me.range(weekdayColumns(c) & TableStartRowIndex)
        For i = IIf(c = 1, 1, 0) To MaxDaysInAMonth
            dateCell.Offset(i, 0).Value = Empty
            weekdayRange.Offset(i, 0).Value = Empty
        Next i
    Next c
End Sub
Private Sub Worksheet_Change(ByVal Target As range)
    Dim dateCell As range
    Call Init
    Dim i As Integer, currentTable As Integer
    Dim startDate As Date
    Dim dateDay As Integer, dateMonth As Integer, dateYear As Integer
    
    ' Define the Range where you want this to happen
    Set dateCell = Me.range(dateColumns(LBound(dateColumns)) & TableStartRowIndex)
    ' Check if the changed cell is within the defined Range and is not empty
    If Not Intersect(Target, dateCell) Is Nothing And Target.Value <> "" Then
        ' Check if the entered value is a date
        If IsDate(dateCell.Value) Then
            Call ClearCells

            Dim loopingTableIndexOffset As Integer, monthDaysCount As Integer
            loopingTableIndexOffset = 0
            currentTable = 1
            
            startDate = CDate(dateCell.Value)
            dateDay = day(startDate)
            dateMonth = month(startDate)
            dateYear = year(startDate)
            monthDaysCount = DaysInMonth(dateMonth, dateYear)
            ' Loop through the cells below Specified one and fill them with next day dates
            For i = 1 To monthDaysCount - dateDay ' Adjust this number as per your requirement
                dateDay = dateDay + 1
                If dateDay > monthDaysCount Then
                    dateDay = dateDay Mod monthDaysCount
                    dateMonth = dateMonth + CInt(dateDay / monthDaysCount)
                    If dateMonth > 12 Then
                        dateMonth = dateMonth Mod 12
                        dateYear = dateYear + CInt(dateMonth / 12)
                    End If
                End If
                
                If i = currentTable * MaxNumberRowsInTable Then
                    loopingTableIndexOffset = currentTable * MaxNumberRowsInTable
                    currentTable = currentTable + 1
                    Set dateCell = Me.range(dateColumns(currentTable) & TableStartRowIndex)
                End If
                If Not IsError(startDate) Then
                    dateCell.Offset(i - loopingTableIndexOffset, 0).Value = DateString(dateYear, dateMonth, dateDay)
                Else
                    Exit For ' Exit the loop if there's an error in date calculation
                End If
            Next i

        End If
    End If
    
    Dim weekdayCell As range
    Set weekdayCell = Me.range(weekdayColumns(LBound(weekdayColumns)) & TableStartRowIndex)

    If Not Intersect(Target, weekdayCell) Is Nothing And Target.Value <> "" _
            & IsDate(dateCell.Value) Then
        Dim weekdayNames() As String

        weekdayNames = GetDropdownItems(weekdayCell)
        If UBound(weekdayNames) > 1 Then
            Dim weekday As String
            Dim weekdayIndex As Integer
            weekday = ""
            For weekdayIndex = LBound(weekdayNames) To UBound(weekdayNames)
                If weekdayNames(weekdayIndex) = weekdayCell.Value Then
                    weekday = weekdayNames(weekdayIndex)
                    Exit For
                End If
            Next weekdayIndex
            
            If weekday <> "" Then
                
                loopingTableIndexOffset = 0
                currentTable = 1
                startDate = CDate(dateCell.Value)
                dateDay = day(startDate)
                dateMonth = month(startDate)
                dateYear = year(startDate)
                monthDaysCount = DaysInMonth(dateMonth, dateYear)
                For i = 1 To monthDaysCount - dateDay ' Adjust this number as per your requirement
                    If i = currentTable * MaxNumberRowsInTable Then
                        loopingTableIndexOffset = currentTable * MaxNumberRowsInTable
                        currentTable = currentTable + 1
                        Set weekdayCell = Me.range(weekdayColumns(currentTable) & TableStartRowIndex)
                    End If
                    weekdayIndex = (weekdayIndex + 1) Mod UBound(weekdayNames)
                    weekday = weekdayNames(weekdayIndex)
                    weekdayCell.Offset(i - loopingTableIndexOffset, 0).Value = weekday
                Next i
            Else
                MsgBox "The selected weekday is not a valid one.", vbOKOnly + vbExclamation, "Invalid Weekday"
            End If
        Else
            MsgBox "The selected cell does not have a dropdown list.", vbOKOnly + vbExclamation, "Invalid Value"
        End If
    End If
End Sub



