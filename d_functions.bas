Attribute VB_Name = "d_functions"
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'debe haber una funcion para dato que debemos extraer de la factura.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Option Explicit
'numeroClientte
Public Function numeroCliente(ByVal arrFileLines)
    Dim myArray As Collection
    Set myArray = New Collection
    numeroCliente = diccionario_numeroCliente(myArray)
    For i = 1 To myArray.count
        countSplitters = UBound(Split(myArray(i), "|"))
        If countSplitters > 0 Then
            numeroCliente = Split(myArray(i), "|")(0)
        End If
        If InStr(1, LCase(arrFileLines), LCase(numeroCliente), vbBinaryCompare) = 0 Then
            numeroCliente = ""
        Else
            If countSplitters > 0 Then
                numeroCliente = Split(myArray(i), "|")(2)
                Exit For
            Else
                numeroCliente = ""
            End If
        End If
    Next
End Function
'numeroFactura
Public Function numeroFactura(ByVal proveedor As String) As Collection
    Dim functionName As String
    Dim myArray As Collection
    Set myArray = New Collection
    functionName = "diccionario_numeroFactura_" & proveedor
    functionName = Application.Run(functionName, myArray)
    Set numeroFactura = myArray
End Function
'periodoDeFacturacion
Public Function periodoDeFacturacion(ByVal proveedor As String)
    Dim functionName As String
    Dim myArray As Collection
    Set myArray = New Collection
    functionName = "diccionario_periodoDeFacturacion_" & proveedor
    functionName = Application.Run(functionName, myArray)
    Set periodoDeFacturacion = myArray
End Function
'fechaVencimiento
Public Function fechaVencimiento(ByVal proveedor As String)
    Dim functionName As String
    Dim myArray As Collection
    Set myArray = New Collection
    functionName = "diccionario_fechaVencimiento_" & proveedor
    functionName = Application.Run(functionName, myArray)
    Set fechaVencimiento = myArray
End Function
'fecha2vencimiento
Public Function fecha2vencimiento(ByVal proveedor As String)
    Dim functionName As String
    Dim myArray As Collection
    Set myArray = New Collection
    functionName = "diccionario_fecha2vencimiento_" & proveedor
    functionName = Application.Run(functionName, myArray)
    Set fecha2vencimiento = myArray
End Function
'importeSubtotal
Public Function importeSubtotal(ByVal proveedor As String)
    Dim functionName As String
    Dim myArray As Collection
    Set myArray = New Collection
    functionName = "diccionario_importeSubtotal_" & proveedor
    functionName = Application.Run(functionName, myArray)
    Set importeSubtotal = myArray
End Function
'importeTotal
Public Function importeTotal(ByVal proveedor As String)
    Dim functionName As String
    Dim myArray As Collection
    Set myArray = New Collection
    functionName = "diccionario_importeTotal_" & proveedor
    functionName = Application.Run(functionName, myArray)
    Set importeTotal = myArray
End Function
'importeTotalVencido
Public Function importeTotalVencido(ByVal proveedor As String)
    Dim functionName As String
    Dim myArray As Collection
    Set myArray = New Collection
    functionName = "diccionario_importeTotalVencido_" & proveedor
    functionName = Application.Run(functionName, myArray)
    Set importeTotalVencido = myArray
End Function
'viaPago
Public Function viaPago(ByVal proveedor As String)
    Dim functionName As String
    Dim myArray As Collection
    Set myArray = New Collection
    functionName = "diccionario_viaPago_" & proveedor
    functionName = Application.Run(functionName, myArray)
    Set viaPago = myArray
End Function
'categoriaDeCosto
Public Function categoriaDeCosto(ByVal proveedor As String)
    Dim functionName As String
    Dim myArray As Collection
    Set myArray = New Collection
    functionName = "diccionario_categoriaDeCosto_" & proveedor
    functionName = Application.Run(functionName, myArray)
    Set categoriaDeCosto = myArray
End Function
'bancoPropio
Public Function bancoPropio(ByVal proveedor As String)
    Dim functionName As String
    Dim myArray As Collection
    Set myArray = New Collection
    functionName = "diccionario_bancoPropio_" & proveedor
    functionName = Application.Run(functionName, myArray)
    Set bancoPropio = myArray
End Function
'numeroMedidor
Public Function numeroMedidor(ByVal proveedor As String)
    Dim functionName As String
    Dim myArray As Collection
    Set myArray = New Collection
    functionName = "diccionario_numeroMedidor_" & proveedor
    functionName = Application.Run(functionName, myArray)
    Set numeroMedidor = myArray
End Function
'provincia
Public Function provincia(ByVal proveedor As String)
    Dim functionName As String
    Dim myArray As Collection
    Set myArray = New Collection
    functionName = "diccionario_provincia_" & proveedor
    functionName = Application.Run(functionName, myArray)
    Set provincia = myArray
End Function
'domicilio
Public Function domicilio(ByVal proveedor As String)
    Dim functionName As String
    Dim myArray As Collection
    Set myArray = New Collection
    functionName = "diccionario_domicilio_" & proveedor
    functionName = Application.Run(functionName, myArray)
    Set domicilio = myArray
End Function
'IVA
Public Function IVA(ByVal proveedor As String)
    Dim functionName As String
    Dim myArray As Collection
    Set myArray = New Collection
    functionName = "diccionario_IVA_" & proveedor
    functionName = Application.Run(functionName, myArray)
    Set IVA = myArray
End Function
'IIBB
Public Function IIBB(ByVal proveedor As String)
    Dim functionName As String
    Dim myArray As Collection
    Set myArray = New Collection
    functionName = "diccionario_IIBB_" & proveedor
    functionName = Application.Run(functionName, myArray)
    Set IIBB = myArray
End Function
