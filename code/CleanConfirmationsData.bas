Attribute VB_Name = "CleanConfirmationsData"
Option Explicit

' This module contains code to clean the most common formatting errors encountered during the conversion of Resumes of
' Congressional Activity from PDF to CSV. This module is not intended to cover every possible issue, but to automate as
' much of the common cleanup as is reasonable.

Private Sub Format()
' This subroutine adjusts column widths and alignments, and adds rows Congress, Session, Start Date and End Date

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
' This subroutine removes all empty cells to simplify merging of content

    Dim lastCell As Range
    Dim r As Range
    Dim c As Range
    Dim toDelete As Range
    
    ' Since the special cells function returns an errror if no empty cells are found, we have to add
    ' an error handler to prevent the macro from crashing.
    On Error GoTo errorhandler
        Set toDelete = ActiveSheet.UsedRange.SpecialCells(xlCellTypeBlanks)
        toDelete.Delete Shift:=xlShiftToLeft
errorhandler:
    Exit Sub
    
End Sub


Private Sub sectionHeadings()
' This subroutine identifies and merges the text of section headings. Detail lines for each section are indented

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
                        r.Value = r.Value & " " & c.Value
                        Next
                    
                    ' Clear the cells that contained portions of the section heading
                    Range(r.Offset(0, 1), lastCell).Delete Shift:=xlShiftToLeft
                End If
                
                ' Check to see if the section heading spans a second or third row, using the : character as a flag
                If InStrRev(r.Value, ":") = 0 Then
                Do
                    ' Loop through all cells in the next row with data and concatenate the values into the
                    ' current cell in column a
                    Set lastCell = Cells(r.Offset(1, 0).Row, ActiveSheet.Columns.Count).End(xlToLeft)
                    For Each c In Range(r.Offset(1, 0), lastCell)
                        r.Value = r.Value & " " & c.Value
                        Next
                    ' Remove any hyphens
                    r.Value = Replace(r.Value, "- ", "")
                    r.Value = Replace(r.Value, "-", "")
                    
                    ' Delete the row
                    r.Offset(1, 0).EntireRow.Delete Shift:=xlUp
                    Loop Until InStrRev(r.Value, ":") > 0
                End If
            Else
                r.Value = "     " & r.Value
            End If
            Next
    End If
                
End Sub

Private Sub TotalNominations()
' This subroutine extracts new and carryover nominations from section headings

    Dim totalCell As Range
    Dim totalNoms As Integer
    Dim carryoverNoms As Integer
    
    Set totalCell = Columns("A:A").Find(What:="totaling")
    If Not totalCell Is Nothing Then
        Do
            ' Scenario example: "Army nominations, totaling X (including Y nominations carried over from the...
            '   ... first session), disposed of as follows:)
            If InStr(totalCell.Value, "including") > 0 Then
                Rows(totalCell.Row + 1).EntireRow.Resize(2).Insert Shift:=xlDown
                totalCell.Offset(1, 0).Value = "     New nominations"
                totalCell.Offset(2, 0).Value = "     Carryover nominations"
                carryoverNoms = Split(Split(totalCell.Value, "including")(1), " nominations")(0)
                totalNoms = Split(Split(totalCell, "totaling ")(1), "(including")(0)
                totalCell.Offset(1, 1).Value = totalNoms - carryoverNoms
                totalCell.Offset(2, 1).Value = carryoverNoms
                totalCell.Value = Split(totalCell, "nominations")(0)
                
            ' Scenario example: "Army nonminations, totaling X (and Y nominations carried over from the...
            '   ... first session), disposed of as follows:
            ' Scenario example: "Civilian nominations (FS, PHG), totaling X (and Y nominations carried over...
            '   ... from the first session), disposed of as follows:
            ElseIf InStr(totalCell.Value, "(and") > 0 Then
                Rows(totalCell.Row + 1).EntireRow.Resize(2).Insert Shift:=xlDown
                totalCell.Offset(1, 0).Value = "     New nominations"
                totalCell.Offset(2, 0).Value = "     Carryover nominations"
                carryoverNoms = Split(Split(totalCell.Value, " (and")(1), " nominations")(0)
                totalNoms = Split(Split(totalCell, " (and ")(1), " nominations")(0)
                totalCell.Offset(1, 1).Value = totalNoms
                totalCell.Offset(2, 1).Value = carryoverNoms
                totalCell.Value = Split(totalCell, "nominations")(0)
                
            ' Scenario example: "Army nominations, totaling X, disposed of as follows:"
            ' Scenario example: "Civilian (lists) nominations, totaling X, disposed of as follows:"
            ' Scenario example: "Civilian nominations (lists), totaling X, disposed of as follows:"
            Else
                Rows(totalCell.Row + 1).EntireRow.Insert Shift:=xlDown
                totalCell.Offset(1, 0).Value = "     New nominations"
                totalCell.Offset(1, 1).Value = Split(Split(totalCell, "totaling")(1), ", disposed")(0)
                totalCell.Value = Split(totalCell, ",")(0)
                If InStr(totalCell.Value, "{") > 0 And InStr(totalCell.Value, "nominations (") Then
                    totalCell.Value = Split(totalCell.Value, ",")(0)
                    totalCell.Value = Replace(totalCell.Value, " nominations ", "")
                    
                Else
                    totalCell.Value = Replace(totalCell.Value, " nominations", "")
                End If
            End If
            
            Set totalCell = Columns("A").FindNext(totalCell.Offset(1, 0))
            Loop Until totalCell Is Nothing
    End If
 
End Sub


Private Sub SetCongressAndSession()
' This subroutine extracts the Congress and Session data from the file name

    Dim fileName As String
        
    fileName = ActiveWorkbook.Name
    ActiveSheet.Name = Left(fileName, Len(fileName) - 5)
    Cells(2, 2).Value = Left(fileName, InStr(1, fileName, "_") - 1)
    Cells(3, 2).Value = Left(Right(fileName, 6), 1)
    
End Sub

Private Sub RemovePeriods()
' This subroutine removes periods from data labels, and extracts data values that have not been properly separated from
' their labels

    Dim lastCell As Range
    Dim r As Range
    Dim startPos As Integer
    Dim qty As String

    Set lastCell = Columns("A").SpecialCells(xlCellTypeLastCell)
    For Each r In Range(Cells(1, 1), Cells(lastCell.Row, 1))
        r.Value = Replace(r.Value, ".", "")
        If r.Value Like "*[0-9]*" Then
            startPos = InStrRev(r.Value, " ")
            r.Offset(0, 1).Value = Mid(r.Value, startPos + 1)
            r.Value = Mid(r.Value, 1, startPos)
        End If
    Next r
           
End Sub


Private Sub PromptForDates()
' This subroutine prompts for Start and End dates for the session

    Cells(4, 2).Value = InputBox("Enter Session Start Date", "Start Date")
    Cells(5, 2).Value = InputBox("Enter Session End Date", "End Date")
    
End Sub

Sub CleanConfirmations()
' This subroutine executes all private subroutines that perform individual cleanup activities
    
    Call DeleteEmptyCells
    Call sectionHeadings
    Call TotalNominations
    Call RemovePeriods
    Call Format
    Call SetCongressAndSession
    Call PromptForDates

End Sub

