Attribute VB_Name = "Ä£¿é1"
Function EDITDIST(str1 As String, str2 As String) As Integer
    Dim len1 As Integer
    Dim len2 As Integer
    Dim i As Integer
    Dim j As Integer
    Dim cost() As Integer
    
    len1 = Len(str1)
    len2 = Len(str2)
    ReDim cost(len1, len2)
    
    For i = 0 To len1
        cost(i, 0) = i
    Next i
    
    For j = 0 To len2
        cost(0, j) = j
    Next j
    
    For i = 1 To len1
        For j = 1 To len2
            If Mid(str1, i, 1) = Mid(str2, j, 1) Then
                cost(i, j) = cost(i - 1, j - 1)
            Else
                cost(i, j) = WorksheetFunction.Min(cost(i - 1, j) + 1, cost(i, j - 1) + 1, cost(i - 1, j - 1) + 1)
            End If
        Next j
    Next i
    
    EDITDIST = cost(len1, len2)
End Function
