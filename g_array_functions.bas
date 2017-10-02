Attribute VB_Name = "g_array_functions"
Function f_numeroCliente(ByVal fila As String)
    If Not IsEmpty(fila) Then
        f_numeroCliente = CleanNonNumbers(fila)
    End If
End Function

Public Function f_nroFactura(ByVal proveedor As String, ByVal fila As String, ByVal x As Integer, ByVal diccionario_numero_factura As Collection)
    Dim functionName As String
    functionName = "f_nroFactura_" & proveedor
    functionName = Application.Run(functionName, fila, x, diccionario_numero_factura)
    f_nroFactura = functionName
End Function

Function f_fechaVencimiento(ByVal fila As String)
    If Not IsEmpty(fila) Then
        f_fechaVencimiento = Mid(fila, InStr(1, fila, "vence el"), Len(fila))
        filaAUX = Right(f_fechaVencimiento, Len(f_fechaVencimiento) - InStrRev(f_fechaVencimiento, "(*)") + 1)
        f_fechaVencimiento = Mid(f_fechaVencimiento, 1, InStr(1, f_fechaVencimiento, filaAUX) - 1)
        f_fechaVencimiento = CleanNonNumbers(f_fechaVencimiento)
        filaAUX = Empty
    End If
End Function

Function f_fecha2Vencimiento(ByVal fila As String)
    If Not IsEmpty(fila) Then
        f_fecha2Vencimiento = Mid(fila, InStr(1, fila, "aprox el"), Len(fila))
        f_fecha2Vencimiento = CleanNonNumbers(f_fecha2Vencimiento)
    End If
End Function

Function f_Subtotal(ByVal fila As String)
    If Not IsEmpty(fila) Then
        'Caso tipo de Factura 1
        If (InStr(1, fila, "subtotal por servicio el")) > 0 Then
            f_Subtotal = Mid(fila, InStr(1, fila, "subtotal por servicio el"), Len(fila))
            f_Subtotal = CleanNonNumbers_Price(f_Subtotal)
            f_Subtotal = ReplaceDecimal(f_Subtotal)
        End If
        'Caso tipo de Factura 2
        If (InStr(1, fila, "subtotal cargos netos del mes")) > 0 Then
            f_Subtotal = Mid(fila, InStr(1, fila, "subtotal cargos netos del mes"), Len(fila))
            f_Subtotal = CleanNonNumbers_Price(f_Subtotal)
            f_Subtotal = ReplaceDecimal(f_Subtotal)
        End If
    End If
End Function

Function f_Total(ByVal fila As String)
    If Not IsEmpty(fila) Then
        f_Total = Mid(fila, InStr(1, fila, "total a pagar hasta"), Len(fila))
        f_Total = CleanNonNumbers_Price(f_Total)
        f_Total = Mid(fila, InStr(1, fila, "$"), Len(fila))
        f_Total = CleanNonNumbers_Price2(f_Total)
        f_Total = ReplaceDecimal(f_Total)
    End If
End Function

Function f_TotalVencido(ByVal fila As String)
    If Not IsEmpty(fila) Then
        f_TotalVencido = Mid(fila, InStr(1, fila, "s del vencimiento"), Len(fila))
        f_TotalVencido = CleanNonNumbers_Price(f_TotalVencido)
        f_TotalVencido = CleanNonNumbers_Price2(f_TotalVencido)
        f_TotalVencido = ReplaceDecimal(f_TotalVencido)
    End If
End Function

Function f_IVA(ByVal proveedor As String, ByVal fila As String, ByVal x As Integer, ByVal diccionario_numero_factura As Collection)
    Dim functionName As String
    functionName = "f_IVA_" & proveedor
    functionName = Application.Run(functionName, fila, x, diccionario_numero_factura)
    f_IVA = functionName
End Function

