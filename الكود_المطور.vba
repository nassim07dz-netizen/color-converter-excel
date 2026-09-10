Private Sub Worksheet_Change(ByVal Target As Range)
    On Error GoTo ErrorHandler
    
    Dim cell As Range
    Dim colors() As String
    Dim i As Integer
    Dim numValue As Integer
    Dim startRow As Integer
    Dim inputValue As String
    
    colors = Split("VJ,BLEU,MARRON,GRIS,NOIR", ",")
    
    For Each cell In Target
        inputValue = CStr(cell.Value)
        
        ' التحقق من الأرقام (1-5)
        If IsNumeric(inputValue) Then
            numValue = CInt(Val(inputValue))
            
            If numValue >= 1 And numValue <= 5 Then
                Application.EnableEvents = False
                
                ' احذف الخلايا القديمة تحت هذه الخلية
                Dim clearRange As Range
                Set clearRange = Me.Range(cell.Offset(1, 0), cell.Offset(5, 0))
                clearRange.ClearContents
                
                ' اكتب الألوان في خلايا منفصلة عمودياً
                startRow = cell.Row
                For i = 0 To numValue - 1
                    Me.Cells(startRow + i + 1, cell.Column).Value = colors(i)
                Next i
                
                ' امسح الخلية الأصلية
                cell.Value = ""
                
                Application.EnableEvents = True
            End If
        
        ' التحقق من الحرف t أو T
        ElseIf UCase(inputValue) = "T" Then
            Application.EnableEvents = False
            cell.Value = "SUR TOURET"
            Application.EnableEvents = True
        
        ' التحقق من الحرف B أو b
        ElseIf UCase(inputValue) = "B" Then
            Application.EnableEvents = False
            cell.Value = "SUR BOBINE"
            Application.EnableEvents = True
        End If
        
    Next cell
    
    Exit Sub
ErrorHandler:
    Application.EnableEvents = True
End Sub
