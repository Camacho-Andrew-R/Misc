Type pet
    Weight As Integer
    Name As String
End Type
Function uniqueItems(list_var() As pet, Length_Var As Integer)

    Dim unique_list() As String
    Dim index_item As Integer
    Dim i As Integer
    
    Application.ScreenUpdating = False
    
    index_item = -1
    
    For i = 0 To Length_Var

        If i = Length_Var Then
        
            If list_var(i).Name <> unique_list(index_item) Then
            
                index_item = index_item + 1
                
                ReDim Preserve unique_list(index_item) As String
                
                unique_list(index_item) = list_var(i).Name
            
            End If
            
        Else
        
            If list_var(i).Name <> list_var(i + 1).Name Then
            
                index_item = index_item + 1
                
                ReDim Preserve unique_list(index_item) As String
                
                unique_list(index_item) = list_var(i).Name
            
            End If
            
        End If
        
    Next i
    
    uniqueItems = unique_list

End Function
Sub bubblesort()

    Dim list_var() As pet
    Dim Length_Var As Integer
    Dim i As Integer, cycles As Integer, counter As Integer, j As Integer, k As pet
    Dim BenchMark As Double
    Dim unique_var() As String
    
    Application.ScreenUpdating = False
    BenchMark = Timer
    
    Length_Var = Range(Range("A1"), Range("A1").End(xlDown)).Rows.Count - 1
    
    Debug.Print "Length_Var: " & Length_Var

    ReDim list_var(Length_Var) As pet
    
    'Declare array elements
    
    For i = 0 To Length_Var
    
        list_var(i).Weight = Range("A1").Offset(i, 0).Value
        list_var(i).Name = Range("A1").Offset(i, 1).Value
    
    Next i
    
    'Sort with basic bubble sort algorithm

    For i = 0 To Length_Var
        Debug.Print list_var(i).Weight & list_var(i).Name
    Next i

    For i = 0 To (Length_Var - 1)
        counter = 0
        For j = 0 To (Length_Var - 1)
            If list_var(j).Weight > list_var(j + 1).Weight Then
                k = list_var(j)
                list_var(j) = list_var(j + 1)
                list_var(j + 1) = k
                counter = counter + 1
            End If
        Next j
        'Debug.Print "counter per cycle: " & counter
        If counter = 0 Then
            Exit For
        End If
    Next i
    
    'Print the new sorted array
    For i = 0 To Length_Var
        Debug.Print "new: " & list_var(i).Weight & list_var(i).Name
    Next i
    
    unique_var = uniqueItems(list_var(), Length_Var)
    
    Debug.Print unique_var(0)
    
    MsgBox Timer - BenchMark
End Sub
