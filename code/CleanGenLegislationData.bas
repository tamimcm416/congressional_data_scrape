Attribute VB_Name = "CleanGenLegislationData"
Option Explicit

Private Sub Format()
' Write summary

    ' Adjust column widths
    Columns("A").ColumnWidth = 60
    Columns("B:D").ColumnWidth = 15
    Columns("B:D").HorizontalAlignment = xlRight
    
    ' Rename columns
    Cells(1, 2).Value = "Senate"
    Cells(1, 3).Value = "House"
    Cells(1, 4).Value = "Total"
    
    ' Add rows for Congress, Session, Start Date and End Date
    Rows("2:5").EntireRow.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Cells(2, 1).Value = "Congress"
    Cells(3, 1).Value = "Session"
    Cells(4, 1).Value = "Start Date"
    Cells(5, 1).Value = "End Date"
       
    ' Delete row for Congressional Record:
    Columns("A").Find(What:="Congressional Record").EntireRow.Delete Shift:=xlUp
    
    ' Correct Case for the Extensions of Remarks label
    Columns("A").Replace What:="Remarks", Replacement:="remarks"

End Sub

Private Sub RemoveSpecialCharacters()
' Strip special characters from the end of the label column

    Columns("A").Replace What:=".", Replacement:=""
    Columns("A").Replace What:=";", Replacement:=""
    
    ' Remove Footnote Markers
    Columns("A:D").Replace What:="~*", Replacement:=""
    
End Sub

Private Sub MergeSplitLines()
' Merge lines that are hyphonated

    Dim foundCell As Range
    
    Set foundCell = Columns("A").Find(What:="-")
    If Not foundCell Is Nothing Then
        Do
           foundCell.Replace What:="-", Replacement:=foundCell.Offset(1, 0).Value
           foundCell.Offset(0, 1).Value = foundCell.Offset(1, 1).Value
           foundCell.Offset(0, 2).Value = foundCell.Offset(1, 2).Value
           foundCell.Offset(1, 0).EntireRow.Delete
           Set foundCell = Columns("A").FindNext(foundCell)
           Loop Until InStr(1, foundCell.Value, "Yea-and-nay")
    End If
    
End Sub

Private Sub SetCongressAndSession()
' Write Summary

    Dim fileName As String
        
    fileName = ActiveWorkbook.Name
    ActiveSheet.Name = Left(fileName, Len(fileName) - 5)
    Range("B2:D2").Value = Left(fileName, InStr(1, fileName, "_") - 1)
    Range("B3:D3").Value = Left(Right(fileName, 6), 1)
    
End Sub



Private Sub PromptForDates()
' Write summary

    Range("B4:D4").Value = InputBox("Enter Session Start Date", "Start Date")
    Range("B5:D5").Value = InputBox("Enter Session End Date", "End Date")
    
End Sub

Private Sub IndentLabels()
' Write summary
' There may be a better way to construct this...

    Dim foundCell As Range
    Dim i As Integer
    
    ' Indent and label the subcagetories of Measures passed
    i = 1
    Set foundCell = Columns("A").Find(What:="Measures passed")
    If Not foundCell Is Nothing Then
        Do
            foundCell.Offset(i, 0).Value = "     Measures passed, " + foundCell.Offset(i, 0)
            i = i + 1
            Loop Until InStr(1, foundCell.Offset(i, 0).Value, "Measures reported")
    End If
        
    ' Indent and label the subcagetories of Measures reported
    i = 1
    Set foundCell = Columns("A").Find(What:="Measures reported")
    If Not foundCell Is Nothing Then
        Do
            foundCell.Offset(i, 0).Value = "     Measures reported, " + foundCell.Offset(i, 0)
            i = i + 1
            Loop Until InStr(1, foundCell.Offset(i, 0).Value, "Special reports")
    End If
    
    ' Indent and label the subcagetories of Measures reported
    i = 1
    Set foundCell = Columns("A").Find(What:="Measures introduced")
    If Not foundCell Is Nothing Then
        Do
            foundCell.Offset(i, 0).Value = "     Measures introduced, " + foundCell.Offset(i, 0)
            i = i + 1
            Loop Until InStr(1, foundCell.Offset(i, 0).Value, "Quorum calls")
    End If
    
End Sub

Sub CleanLegislativeActivity()
' Write summary

    Call Format
    Call RemoveSpecialCharacters
    Call MergeSplitLines
    Call SetCongressAndSession
    Call PromptForDates
    Call IndentLabels
    
End Sub
