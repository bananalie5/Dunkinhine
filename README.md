# Dunkinhine
Macro text change
Sub UpdateFormulaDateRange()
    Dim ws As Worksheet
    Dim targetCell As Range
    Dim dateRangeCell As Range
    Dim currentFormula As String
    Dim newDateRange As String
    Dim pattern As String
    Dim updatedFormula As String
    
    ' Set your worksheet (change "Sheet1" to your actual sheet name)
    Set ws = ThisWorkbook.Sheets("Sheet1")
    
    ' Set the cell where the formula exists (change C2 to your actual target cell)
    Set targetCell = ws.Range("C2")
    
    ' Set the cell containing the new date range (B2)
    Set dateRangeCell = ws.Range("B2")
    
    ' Ensure B2 has a valid value
    If dateRangeCell.Value = "" Then
        MsgBox "B2 is empty. Please enter a valid date range.", vbExclamation, "Error"
        Exit Sub
    End If
    
    ' Get the current formula in the target cell
    currentFormula = targetCell.Formula
    
    ' Get the new date range from B2
    newDateRange = dateRangeCell.Value
    
    ' Define the pattern to look for (dates in format YYYY-MM-DD to YYYY-MM-DD)
    pattern = "\d{4}-\d{2}-\d{2} to \d{4}-\d{2}-\d{2}"
    
    ' Replace old date range with the new one
    updatedFormula = ReplaceWithPattern(currentFormula, pattern, newDateRange)
    
    ' Update the formula only if it changed
    If updatedFormula <> currentFormula Then
        targetCell.Formula = updatedFormula
        MsgBox "Formula updated successfully!", vbInformation, "Success"
    Else
        MsgBox "No changes were made. Check the existing formula.", vbExclamation, "Info"
    End If
End Sub

Function ReplaceWithPattern(ByVal inputText As String, ByVal pattern As String, ByVal replacement As String) As String
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    
    With regex
        .Pattern = pattern
        .Global = True
        .IgnoreCase = True
    End With
    
    If regex.Test(inputText) Then
        ReplaceWithPattern = regex.Replace(inputText, replacement)
    Else
        ReplaceWithPattern = inputText
    End If
End Function

