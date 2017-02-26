
' Code -2

Sub Macro5()
'select and merge
Dim sCell As Range
    Dim CopyText1 As String
    Dim DataObj As MSForms.DataObject
    Set DataObj = New MSForms.DataObject
  
    Selection.Copy
     'Set sCell = Application.InputBox(Prompt:="Select a Cell", Type:=8)
    'sCell.Copy
  
 DataObj.GetFromClipboard

strpaste = DataObj.GetText(1)


    Application.CutCopyMode = False
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    ActiveCell.FormulaR1C1 = strpaste
  
  
End Sub
