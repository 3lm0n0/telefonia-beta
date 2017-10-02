Attribute VB_Name = "g_array_functions_edesur"
Public Function f_dic_numeroCliente_edesur(ByVal fila As String)
    If Not IsEmpty(fila) Then
        f_numeroCliente = CleanNonNumbers(fila)
    End If
End Function

Public Function f_nroFactura_edesur(ByVal fila As String, ByVal x As Integer, ByVal diccionario_numero_factura As Collection)
'supongo que cada linea del diccionario de facturas siempre tiene 2 palabras y estan separadas por "|".
    If Not IsEmpty(fila) Then
        f_nroFactura_edesur = Mid(fila, InStr(1, fila, Split(diccionario_numero_factura(x), "|")(0)), Len(fila))
        If (InStr(1, f_nroFactura_edesur, Split(diccionario_numero_factura(x), "|")(1))) > 0 Then
            f_nroFactura_edesur = Mid(f_nroFactura_edesur, 1, InStr(1, f_nroFactura_edesur, Split(diccionario_numero_factura(x), "|")(1)) - 1)
        End If
        filaAUX = Right(f_nroFactura_edesur, Len(f_nroFactura_edesur) - InStrRev(f_nroFactura_edesur, " "))
        f_nroFactura_edesur = Mid(f_nroFactura_edesur, 1, InStr(1, f_nroFactura_edesur, filaAUX) - 1)
        f_nroFactura_edesur = CleanNonNumbers(f_nroFactura_edesur)
        f_nroFactura_edesur = ReplaceGuionToA(f_nroFactura_edesur)
        filaAUX = Empty
    End If
End Function

Public Function f_fechaVencimiento_edesur(ByVal fila As String)
    If Not IsEmpty(fila) Then
        f_fechaVencimiento = Mid(fila, InStr(1, fila, "vence el"), Len(fila))
        filaAUX = Right(f_fechaVencimiento, Len(f_fechaVencimiento) - InStrRev(f_fechaVencimiento, "(*)") + 1)
        f_fechaVencimiento = Mid(f_fechaVencimiento, 1, InStr(1, f_fechaVencimiento, filaAUX) - 1)
        f_fechaVencimiento = CleanNonNumbers(f_fechaVencimiento)
        filaAUX = Empty
    End If
End Function

Public Function f_fecha2Vencimiento_edesur(ByVal fila As String)
    If Not IsEmpty(fila) Then
        f_fecha2Vencimiento = Mid(fila, InStr(1, fila, "aprox el"), Len(fila))
        f_fecha2Vencimiento = CleanNonNumbers(f_fecha2Vencimiento)
    End If
End Function

Public Function f_Subtotal_edesur(ByVal fila As String)
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

Public Function f_Total_edesur(ByVal fila As String)
    If Not IsEmpty(fila) Then
        f_Total = Mid(fila, InStr(1, fila, "total a pagar hasta"), Len(fila))
        f_Total = CleanNonNumbers_Price(f_Total)
        f_Total = Mid(fila, InStr(1, fila, "$"), Len(fila))
        f_Total = CleanNonNumbers_Price2(f_Total)
        f_Total = ReplaceDecimal(f_Total)
    End If
End Function

Public Function f_TotalVencido_edesur(ByVal fila As String)
    If Not IsEmpty(fila) Then
        f_TotalVencido = Mid(fila, InStr(1, fila, "s del vencimiento"), Len(fila))
        f_TotalVencido = CleanNonNumbers_Price(f_TotalVencido)
        f_TotalVencido = CleanNonNumbers_Price2(f_TotalVencido)
        f_TotalVencido = ReplaceDecimal(f_TotalVencido)
    End If
End Function

Public Function f_IVA_edesur(ByVal fila As String, ByVal x As Integer, ByVal diccionario_numero_factura As Collection)
    Dim filaMonto As Double
    Dim fila_percent As Variant
    If Not IsEmpty(fila) Then
        If InStr(1, fila, "%") > 0 Then
            fila_percent = Split(fila, "%")(0)
            filaMonto = Split(fila, "%")(1)
        End If
        f_IVA_edesur = CleanNonNumbers_Price2(filaMonto)
        f_IVA_edesur = CleanNonNumbers_Price2(fila_percent)

    End If
End Function
