' ###############################################################################################################################
' Реализация алгоритма быстрой сортировки. На основе http://www.rsdn.ru/forum/src/312064.1.aspx
'###############################################################################################################################

Option Explicit
'###############################################################################################################################
' Процедура быстрой сортировки массива (QuickSort)
' [in,out]    aArray - подлежащий сортировке массив
' [in]        aCompareFunction - функция-делегат для вычисления позиции элемента в результирующем массиве
'             должна иметь прототип SomeFunction(a, b) = <0, 0, >0 для сравнения a, b
Sub QuickSortArray(ByRef aArray, aCompare)

    If Not IsArray(aArray) Then Exit Sub
    
    If (UBound(aArray) < LBound(aArray)) Then Exit Sub
    
    ' А теперь отсортируем
    QuickSortArrayPartial aArray, _
            aCompare, _
            LBound(aArray), _
            UBound(aArray), _
            IsObject( aArray( LBound( aArray)))
End Sub

'###############################################################################################################################
' Процедура быстрой сортировки части массива (QuickSort)
' [in,out]    aArray - подлежащий сортировке массив
' [in]        aCompareFunction - функция-делегат для вычисления позиции элемента в результирующем массиве
'             должна иметь прототип SomeFunction(a, b) = <0, 0, >0 для сравнения a, b
' [in]        nLeft - первый элемент границ сортировки
' [in]        nRight - последний элемент границ сортировки
' [in]        bIsObject - признак работы с массивом объектов
Sub QuickSortArrayPartial(ByRef aArray, aCompare, nLeft, nRight, bIsObject)
    
    Dim I, J, P, L, R, T
    
    L = nLeft
    R = nRight
    
    Do
        I = L
        J = R
        If bIsObject Then
            Set P = aArray((L + R) \ 2)
        Else
            P = aArray((L + R) \ 2)
        End If
        
        Do
            While (aCompare(aArray(I), P) < 0)
                I = I + 1
            Wend
            While (aCompare(aArray(J), P) > 0)
                J = J - 1
            Wend
            
            If I <= J Then
                If bIsObject Then
                    Set T = aArray(I)
                    Set aArray(I) = aArray(J)
                    Set aArray(J) = T
                    Set T = Nothing
                Else
                    T = aArray(I)
                    aArray(I) = aArray(J)
                    aArray(J) = T
                    T = Null
                End If
                I = I + 1
                J = J - 1
            End If
        Loop Until I > J
        
        If L < J Then 
            QuickSortArrayPartial aArray, aCompare, L, J, bIsObject
        End If    
        L = I
    Loop Until I >= R
End Sub

'###############################################################################################################################
' Функция - делегат для сравнения произвольных скалярных данных
Function Cmp_Any(a, b)
    If a < b Then
      Cmp_Any = -1
    ElseIf a > b Then
      Cmp_Any = 1
    Else
      Cmp_Any = 0
    End If      
End Function

'###############################################################################################################################
' Функция - делегат для сравнения произвольных строковых данных
Function Cmp_String(a, b)
    Cmp_String = StrComp(a, b)
End Function


' Пример
Dim A, S, X

A = Array(2, 3, 2, 1)

Call QuickSortArray(A, GetRef("Cmp_Any"))

For Each X in A
  S = S & X & "; "
Next

MsgBox(S)