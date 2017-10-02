Attribute VB_Name = "h_aditional_functions"
'Funcion que filtra un array dejando solo NUMEROS, "/" y "-"
Public Function CleanNonNumbers(fila As Variant) As Variant
    If Not IsEmpty(fila) Then
        fila = Trim(fila)
        For u = Len(fila) To 1 Step -1
            MyChar = Mid(fila, u, 1)
            If Not MyChar Like "[0-9/-]" Then
                CleanNonNumbers = Replace(fila, MyChar, "")
                fila = CleanNonNumbers
            Else
                CleanNonNumbers = fila
            End If
        Next u
    End If
End Function

'Funcion que filtra un array dejando solo NUMEROS, "$" y ","
Public Function CleanNonNumbers_Price(fila As Variant) As Variant
    If Not IsEmpty(fila) Then
        fila = Trim(fila)
        For u = Len(fila) To 1 Step -1
            MyChar = Mid(fila, u, 1)
            If Not MyChar Like "[0-9$,]" Then
                CleanNonNumbers_Price = Replace(fila, MyChar, "")
                fila = CleanNonNumbers_Price
            Else
                CleanNonNumbers_Price = fila
            End If
        Next u
    End If
End Function

'Funcion que filtra un array dejando solo NUMEROS y ","
Public Function CleanNonNumbers_Price2(fila As Variant) As Variant
    If Not IsEmpty(fila) Then
        fila = Trim(fila)
        For u = Len(fila) To 1 Step -1
            MyChar = Mid(fila, u, 1)
            If Not MyChar Like "[0-9,]" Then
                CleanNonNumbers_Price2 = Replace(fila, MyChar, "")
                fila = CleanNonNumbers_Price2
            Else
                CleanNonNumbers_Price2 = fila
            End If
        Next u
    End If
End Function

'Funcion que reemplaza "," por "." para que el sistema tome el numero como FLOAT
Public Function ReplaceDecimal(fila As Variant) As Variant
    If Not IsEmpty(fila) Then
        fila = Trim(fila)
        For u = Len(fila) To 1 Step -1
            MyChar = Mid(fila, u, 1)
            If Not MyChar Like "[0-9]" Then
                ReplaceDecimal = Replace(fila, MyChar, ".")
                fila = ReplaceDecimal
            Else
                ReplaceDecimal = fila
            End If
        Next u
    End If
End Function

'Funcion que reemplaza "-" por "A"
Public Function ReplaceGuionToA(fila As Variant) As Variant
    If Not IsEmpty(fila) Then
        fila = Trim(fila)
        For u = Len(fila) To 1 Step -1
            MyChar = Mid(fila, u, 1)
            If Not MyChar Like "[0-9]" Then
                ReplaceGuionToA = Replace(fila, MyChar, "A")
                fila = ReplaceGuionToA
            Else
                ReplaceGuionToA = fila
            End If
        Next u
    End If
End Function

