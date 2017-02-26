'OPtion  1

' Code 1  + Please make sure that you will need to have a reference to the Forms library in Tools. Refer document

' Macro4 Macro - prompt to select and merge

Sub Macro4()
'
    Dim sCell As Range
    Dim CopyText1 As String
    Dim DataObj As MSForms.DataObject
  
    Set DataObj = New MSForms.DataObject
     Set sCell = Application.InputBox(Prompt:="Select a Cell", Type:=8)
    sCell.Copy
  
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
    sCell.Merge
    ActiveCell.FormulaR1C1 = strpaste
  
End Sub

