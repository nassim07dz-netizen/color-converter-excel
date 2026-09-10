Private Sub Worksheet_Change(ByVal Target As Range)
    On Error GoTo ErrorHandler
    
    Dim cell As Range
    Dim colors() As String
    Dim result As String
    Dim i As Integer
    Dim numValue As Integer
    
    ' حدد النطاق الذي تريد تطبيق الخاصية عليه
    If Intersect(Target, Me.Range("A:Z")) Is Nothing Then
        Exit Sub
    End If
    
    colors = Split("VJ,BLEU,MARRON,GRIS,NOIR", ",")
    
    For Each cell In Target
        If IsNumeric(cell.Value) Then
            numValue = CInt(Val(cell.Value))
            
            If numValue >= 1 And numValue <= 5 Then
                Application.EnableEvents = False
                
                result = ""
                For i = 0 To numValue - 1
                    If i > 0 Then result = result & vbNewLine
                    result = result & colors(i)
                Next i
                
                cell.Value = result
                cell.WrapText = True
                
                Application.EnableEvents = True
            End If
        End If
    Next cell
    
    Exit Sub
ErrorHandler:
    Application.EnableEvents = True
End Sub
