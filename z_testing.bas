Attribute VB_Name = "z_testing"
'Option Explicit

Sub testing_FUNCIONES()

'numeroFactura
    Set nFactura = numeroFactura("edesur")

End Sub



Sub probando_MAIN()

'testing MAIN
    Call plainTextReader("EDESUR efactura@edesur.com.ar Mail 1 Fecha 31-01-2017 Link 62")

End Sub


Sub TestCarClass()
    Dim Car As CarClass
    Set Car = New CarClass
    Car.LicensePlate = "34344W"
    Car.Speed = 100 'set speed to 100 mph
    
    Car.DriveBack 'set speed to -100 mph
 
    Car.DriveForward 'set speed to 100 mph
End Sub


Sub testFacturaClass()


End Sub




Public Function papdsad() As Dictionary
'    Dim dict As Dictionary
    'Create the dictionary
    Set numeroCliente_diccionario = New Dictionary
    'Add some (key, value) Key=texto clave : value=numero que el proveedor le asigna a TASA o TMA
    numeroCliente_diccionario.Add "mero de cliente es", "80050225-1" 'edesur
    numeroCliente_diccionario.Add "Jane", 30655116202# 'edenor
    numeroCliente_diccionario.Add "Ted", 30659014056# 'edesal
    
    MsgBox numeroCliente_diccionario.Item("Ted")
End Function





Sub lasdadedzz()
Dim diccionario As Dictionary

    Set diccionario = numeroCliente_diccionario()
    MsgBox diccionario.Item("Ted") & "  " & diccionario.count

End Sub
Sub asdwwwwknknkmm()
    Dim myArray
    ReDim myArray(1 To 5, 1 To 2)
    Dim col As Long, fila As Long
        For col = 1 To 5
            For fila = 1 To 2
               myArray(col, fila) = 1 'Rng.Value
            Next
        Next
    
End Sub
Sub ReDimPreserve2D_real()
Dim myArray() As String, i As Integer, j As Integer
ReDim myArray(1, 3)
'put your code to populate your array here
For i = LBound(myArray, 1) To UBound(myArray, 1)
    For j = LBound(myArray, 2) To UBound(myArray, 2)
        myArray(i, j) = i & "," & j
    Next j
Next i
'ReDim Preserve MyArray(1, 5)
'Stop
End Sub



Sub efe()
ThisWorkbook.Close True
End Sub


Function countOccurencesOf(needle As String, s As String)
    Dim count As Integer, i As Integer
    count = 0
    For i = 0 To Len(s) - 1
        If s.Substring(i).Startswith(needle) Then
            count = count + 1
        End If
    Next
    countOccurencesOf = count
End Function

