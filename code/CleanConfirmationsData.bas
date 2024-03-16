Attribute VB_Name = "CleanConfirmationsData"
Option Explicit

Private Sub Format()
' Write summary

    Columns("A").ColumnWidth = 60
    Columns("A").HorizontalAlignment = xlLeft
    Columns("B").ColumnWidth = 15
    Columns("B").HorizontalAlignment = xlRight
    
    ' Add rows for Congress, Session, Start Date and End Date
    Rows("1:5").EntireRow.Insert Shift:=xlDown
        ' Rename columns
    Cells(1, 1).Value = "Labels"
    Cells(1, 2).Value = "Values"
    Cells(2, 1).Value = "Congress"
    Cells(3, 1).Value = "Session"
    Cells(4, 1).Value = "Start Date"
    Cells(5, 1).Value = "End Date"
        
End Sub

Private Sub DeleteEmptyCells()
' Write summary

    Dim lastCell As Range
    Dim r As Range
    Dim c As Range
    Dim toDelete As Range
    
    Set toDelete = ActiveSheet.UsedRange.SpecialCells(xlCellTypeBlanks)
    toDelete.Delete Shift:=xlShiftToLeft
    
End Sub


Private Sub SectionHeadings()
' Write summary

    Dim summaryCell As Range
    Dim rowRange As Range
    Dim r As Range
    Dim lastCell As Range
    Dim c As Range
    
    ' Use the Summary label as an end row for our loop
    Set summaryCell = Columns("A").Find(What:="Summary")
    If Not summaryCell Is Nothing Then
        Set rowRange = Range(Cells(1, 1), summaryCell.Offset(-1, 0))
            
        ' Loop through each cell in column A until the Summary label
        For Each r In rowRange
        
            ' Use nominations as a flag for section headings
            If InStr(r.Value, "nominations") Then
                If Not IsEmpty(r.Offset(0, 1)) Then
                    Set lastCell = Cells(r.Row, ActiveSheet.Columns.Count).End(xlToLeft)
                    
                    ' If column b of the current row is not empty, then loop through all cells in the row with data
                    ' and concatenate the values into the current cell in column a
                    For Each c In Range(r.Offset(0, 1), lastCell)
                        r.Value = r.Value & c.Value & " "
                        Next
                    
                    ' Clear the cells that contained portions of the section heading
                    Range(r.Offset(0, 1), lastCell).Delete Shift:=xlShiftToLeft
                End If
                
                ' Check to see if the section heading spans a second row, using the : character as a flag
                If InStrRev(r.Value, ":") = 0 Then
                    r.Value = r.Value & r.Offset(1, 0).Value
                    r.Value = Replace(r.Value, "- ", "")
                    r.Value = Replace(r.Value, "-", "")
                    r.Offset(1, 0).EntireRow.Delete Shift:=xlUp
                End If
            Else
                r.Value = "     " & r.Value
            End If
            Next
    End If
                
End Sub

Private Sub TotalNominations()
' Write summary

    Dim totalCell As Range
    Dim totalNoms As Integer
    Dim carryoverNoms As Integer
    
    Set totalCell = Columns("A").Find(What:="totaling")
    If Not totalCell Is Nothing Then
        Do
            Rows(totalCell.Row + 1).EntireRow.Insert Shift:=xlDown
            'Scenario 1
            If InStr(totalCell.Value, "including") > 0 Then
                Rows(totalCell.Row + 1).EntireRow.Insert Shift:=xlDown
                totalCell.Offset(1, 0).Value = "     New nominations"
                totalCell.Offset(2, 0).Value = "     Carryover nominations"
                carryoverNoms = Split(Split(totalCell.Value, "including ")(1), " carried")(0)
                totalNoms = Split(Split(totalCell, "totaling ")(1), ", disposed")(0)
                totalCell.Offset(1, 1).Value = totalNoms - carryoverNoms
                totalCell.Offset(2, 1).Value = carryoverNoms
                totalCell.Value = Split(totalCell, "nominations")(0)
                
            'Scenario 2
            ElseIf InStr(totalCell.Value, "nominations, totaling") Or InStr(totalCell.Value, "nominations,totaling") Or InStr(totalCell.Value, "), totaling") Then
                totalCell.Offset(1, 0).Value = "     New nominations"
                totalCell.Offset(1, 1).Value = Split(Split(totalCell, "totaling")(1), ", disposed")(0)
                totalCell.Value = Split(totalCell, "nominations")(0)
                
            'Scenario

            End If
            
            Set totalCell = Columns("A").FindNext(totalCell.Offset(1, 0))
            Loop Until totalCell Is Nothing
    End If
 
End Sub


Private Sub SetCongressAndSession()
' Write Summary

    Dim fileName As String
        
    fileName = ActiveWorkbook.Name
    ActiveSheet.Name = Left(fileName, Len(fileName) - 5)
    Cells(2, 2).Value = Left(fileName, InStr(1, fileName, "_") - 1)
    Cells(3, 2).Value = Left(Right(fileName, 6), 1)
    
End Sub

Private Sub RemovePeriods()
' Write summary

    Dim r As Range
    Dim cont_array() As String
    Dim qty As String

    For Each r In Range("A:A")
        If InStr(1, r, ".") Then
            cont_array = Split(r.Value, ".", 2)
            r.Value = cont_array(0)
            If cont_array(1) Like "*[0-9]*" Then
                qty = Split(cont_array(1), " ", 2)(1)
                r.Offset(0, 1).Value = qty
            End If
         End If
         Next
           
End Sub


Private Sub PromptForDates()
' Write summary

    Cells(4, 2).Value = InputBox("Enter Session Start Date", "Start Date")
    Cells(5, 2).Value = InputBox("Enter Session End Date", "End Date")
    
End Sub

Sub CleanConfirmationsData()
' Write summary
    
    'Call DeleteEmptyCells
    Call SectionHeadings
    Call TotalNominations
    Call RemovePeriods
    Call Format
    Call SetCongressAndSession
    Call PromptForDates

End Sub

