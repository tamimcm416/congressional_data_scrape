Attribute VB_Name = "CleanGenLegislationData"
Option Explicit

' This module contains code to clean the most common formatting errors encountered during the conversion of Resumes of
' Congressional Activity from PDF to CSV. This module is not intended to cover every possible issue, but to automate as
' much of the common cleanup as is reasonable.

Private Sub Format()
' This subroutine adjusts column widths and alignments, standardizes the column headings,
' deletes the Congressional Record row, and corrects the case of the Extensions of Remarks row label

    ' Adjust column widths
    Columns("A").ColumnWidth = 60
    Columns("B:D").ColumnWidth = 15
    Columns("B:D").HorizontalAlignment = xlRight
    
    ' Rename columns
    Cells(1, 2).Value = "Senate"
    Cells(1, 3).Value = "House"
    Cells(1, 4).Value = "Total"
    
    ' Delete row for Congressional Record:
    Columns("A").Find(What:="Congressional Record").EntireRow.Delete Shift:=xlUp
    
    ' Correct Case for the Extensions of Remarks label
    Columns("A").Replace What:="Remarks", Replacement:="remarks"

End Sub

Private Sub RemoveSpecialCharacters()
' This subroutine removes the period and semicolon characters from labels column, as well as any asterisks that have
' been added as footnote indicators.

    ' Remove periods and semicolons from the labels column
    Columns("A").Replace What:=".", Replacement:=""
    Columns("A").Replace What:=";", Replacement:=""
    
    ' Remove Footnote Markers
    Columns("A:D").Replace What:="~*", Replacement:=""
    
End Sub

Private Sub MergeSplitLines()
' This subroutine looks for labels that have been split across two rows using the hyphen as an indicator. If
' a split label is found, the contents of the two cells are concatenated and the extra row is deleted.

    Dim foundCell As Range
    
    ' Look for hyphens in the labels column
    Set foundCell = Columns("A").Find(What:="-")
    If (Not foundCell Is Nothing) And (InStr(foundCell.Value, "Yea-any-nay") > 0) Then
        Do
            ' Replace the hyphen in the first cell with the text from the cell below
            foundCell.Replace What:="-", Replacement:=foundCell.Offset(1, 0).Value
            
            ' Move the values in columns B:D up one row to align with the merged label
            foundCell.Offset(0, 1).Value = foundCell.Offset(1, 1).Value
            foundCell.Offset(0, 2).Value = foundCell.Offset(1, 2).Value
            foundCell.Offset(0, 3).Value = foundCell.Offset(1, 3).Value
            
            ' Delete the row that contained the second half of the label
            foundCell.Offset(1, 0).EntireRow.Delete
            
            ' Search for the next cell that contains a hyphen
            Set foundCell = Columns("A").FindNext(foundCell)
            
            ' Loop until the search reaches the "Yea-and-nay" cell. This cell is not a split heading, and the few rows after
            ' are short enough to not be split
            Loop Until InStr(foundCell.Value, "Yea-and-nay")
    End If
    
End Sub

Private Sub SetCongressAndSession()
' This subroutine adds rows for Congress and Session. These values are extracted from the file name added to the
' appropriate cells. The worksheet is renamed to match the filename.

    Dim fileName As String
    
    ' Add rows for Congress, and Session
    Rows("2:3").EntireRow.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Cells(2, 1).Value = "Congress"
    Cells(3, 1).Value = "Session"
        
    ' Extract Congress and Session from the name of the file, and populate the proper rows
    fileName = ActiveWorkbook.Name
    Range("B2:D2").Value = Left(fileName, InStr(1, fileName, "_") - 1)
    Range("B3:D3").Value = Left(Right(fileName, 6), 1)
    
    ' Rename the worksheet to match the filename
    ActiveSheet.Name = Left(fileName, Len(fileName) - 5)
    
End Sub



Private Sub PromptForDates()
' This subroutine adds rows for the session Start Date and End Date. The user is prompted to enter both dates, which are
' added to the appropriate cells.

    ' Add rows for Congress, and Session
    Rows("4:5").EntireRow.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Cells(4, 1).Value = "Start Date"
    Cells(5, 1).Value = "End Date"

    ' Prompt the user to enter a Start Date and End Date
    Range("B4:D4").Value = InputBox("Enter Session Start Date", "Start Date")
    Range("B5:D5").Value = InputBox("Enter Session End Date", "End Date")
    
End Sub

Private Sub IndentLabels()
' This subroutine indents each detail line for measure types and prepends a heading. This is necessary to prevent
' duplicate values during later processing.

    Dim sectionHeadings(2) As String
    Dim foundCell As Range
    Dim curCell As Range
    Dim i As Integer
    
    ' Define headings
    sectionHeadings(0) = "Measures passed"
    sectionHeadings(1) = "Measures reported"
    sectionHeadings(2) = "Measures introduced"
    
    ' Loop through each section with detail lines
    For i = LBound(sectionHeadings) To UBound(sectionHeadings)
        
        'Find the next section heading
        Set foundCell = Columns("A").Find(What:=sectionHeadings(i))
        If Not foundCell Is Nothing Then
            
            ' Prepend the section heading and indent each detail line, ending with Simple Resolutions
            Do
                Set foundCell = foundCell.Offset(1, 0)
                foundCell.Value = "     " & sectionHeadings(i) & ", " & foundCell.Value
            Loop Until InStr(foundCell.Value, "Simple resolutions")
        
        End If
    Next i
    
End Sub

Sub CleanLegislativeActivity()
' This subroutine executes all private subroutines that perform individual cleanup activities

    Call Format
    Call RemoveSpecialCharacters
    Call MergeSplitLines
    Call SetCongressAndSession
    Call PromptForDates
    Call IndentLabels
    
End Sub
