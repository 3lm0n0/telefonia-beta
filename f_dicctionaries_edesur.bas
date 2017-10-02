Attribute VB_Name = "f_dicctionaries_edesur"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'debe haber un diccionario para dato que debemos extraer de la factura.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit 'EDESUR

Public Function diccionario_numeroCliente_edesur(ByVal myArray As Collection)
    
    myArray.Add LCase("mero de cliente es")
    
End Function

Public Function diccionario_numeroFactura_edesur(ByVal myArray As Collection)
    'myArray.Add LCase("L.S.P|capital federal")
    myArray.Add LCase("LSP|capital federal")
    myArray.Add LCase("L.S.P|cap fed")
    myArray.Add LCase("LSP|cap fed")
    myArray.Add LCase("L.S.P|caba")
    myArray.Add LCase("LSP|caba")
    myArray.Add LCase("L.S.P|capital federal")
End Function

Public Function diccionario_periodoDeFacturacion_edesur(ByVal myArray As Collection)
    
    
End Function

Public Function diccionario_fechaVencimiento_edesur(ByVal myArray As Collection)
    
    myArray.Add LCase("vence el")
    
End Function

Public Function diccionario_fecha2vencimiento_edesur(ByVal myArray As Collection)
    
    myArray.Add LCase("aprox el")
    
End Function

Public Function diccionario_importeSubtotal_edesur(ByVal myArray As Collection)
    
    myArray.Add LCase("subtotal por servicio el")
    myArray.Add LCase("subtotal cargos netos del mes")
    
End Function

Public Function diccionario_importeTotal_edesur(ByVal myArray As Collection)
    
    myArray.Add LCase("total a pagar hasta")
    
End Function

Public Function diccionario_importeTotalVencido_edesur(ByVal myArray As Collection)
    
    myArray.Add LCase("s del vencimiento")
    
End Function

Public Function diccionario_viaPago_edesur(ByVal myArray As Collection)
    
    myArray.Add LCase("Se debitará de su Cuenta Corriente Nº")
    'myArray.Add LCase("")
    
End Function

Public Function diccionario_categoriaDeCosto_edesur(ByVal myArray As Collection)
    
    myArray.Add LCase("")
    
End Function

Public Function diccionario_bancoPropio_edesur(ByVal myArray As Collection)
    
    myArray.Add LCase("")
    
End Function

Public Function diccionario_numeroMedidor_edesur(ByVal myArray As Collection)
    
    myArray.Add LCase("")
    
End Function

Public Function diccionario_provincia_edesur(ByVal myArray As Collection)
    
    myArray.Add LCase("capital federal")
    
End Function

Public Function diccionario_domicilio_edesur(ByVal myArray As Collection)
    myArray.Add LCase("iguazú 341")
    myArray.Add LCase("iguazu 341")
End Function

Public Function diccionario_IVA_edesur(ByVal myArray As Collection)
    myArray.Add LCase("Imp. Valor Agregado 27%")
    myArray.Add LCase("Imp. Valor Agregado 23%")
End Function

Public Function diccionario_IIBB_edesur(ByVal myArray As Collection)
    
    myArray.Add LCase("")
    
End Function

