Attribute VB_Name = "b_main"
'Option Explicit
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'este procedimiento es llamado por AA.
'espera recibir como parametro el nombre del file que tiene que recorrer.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub plainTextReader(ByVal fileName As String)
    '~~> definiciones
    Dim wb As Excel.Workbook
    Dim wsBase As Excel.Worksheet, wsImpuestos As Excel.Worksheet
    Dim forReading As Integer, i As Integer, x As Integer
    Dim filePath As String
    Dim fso As FileSystemObject
    Dim arrFileLines()
    Dim numero_proveedor, diccionario_numero_factura As Collection, diccionario_periodo_facturacion As Collection, diccionario_fecha_vencimiento As Collection, diccionario_fecha_2vencimiento As Collection
    Dim diccionario_importe_subtotal As Collection, diccionario_importe_total As Collection, via_pago As Collection, categoria_costo As Collection, banco_propio As Collection, numero_medidor As Collection
    Dim provincia_proveedor As Collection, domicilio_proveedor As Collection, diccionario_impuesto_IVA As Collection, impuesto_IIBB As Collection
    Dim diccionario_importe_total_vencido As Collection
    Dim impuesto_IVA
    Dim objTextFileInternos
    '~~> seteo de objetos.
    Set wb = ThisWorkbook
    Set wsBase = wb.Sheets("base_facturas")
    Set wsImpuestos = wb.Sheets("impuestos")
    '~~> genera un array con cada linea del archivo de texto.
        forReading = 1
        filePath = invoicesTxtPath & fileName & ".txt"
        Set fso = New FileSystemObject
        Set objTextFileInterno = fso.OpenTextFile(filePath, forReading, False)
        '~~> recorro cada linea y la asigno a una arrFileLines(i) dentro del array.
            Do Until objTextFileInterno.AtEndOfStream
                ReDim Preserve arrFileLines(i)
                arrFileLines(i) = objTextFileInterno.ReadLine
                If IsEmpty(numero_proveedor) Or numero_proveedor = "" Then
                    numero_proveedor = numeroCliente(arrFileLines(i)) '~~> con el numero que el proveedor le asigna a Telefonica vamos a saber que cliente es.
                End If
                If Not IsEmpty(numero_proveedor) And numero_proveedor <> "" Then
                    '~~> diccionarios para cada campo y popula la data.
                    If diccionario_numero_factura Is Nothing Or IsEmpty(diccionario_numero_factura) Then Set diccionario_numero_factura = numeroFactura(numero_proveedor)
                    If diccionario_periodo_facturacion Is Nothing Or IsEmpty(diccionario_periodo_facturacion) Then Set diccionario_periodo_facturacion = periodoDeFacturacion(numero_proveedor)
                    If diccionario_fecha_vencimiento Is Nothing Or IsEmpty(diccionario_fecha_vencimiento) Then Set diccionario_fecha_vencimiento = fechaVencimiento(numero_proveedor)
                    If diccionario_fecha_2vencimiento Is Nothing Or IsEmpty(diccionario_fecha_2vencimiento) Then Set diccionario_fecha_2vencimiento = fecha2vencimiento(numero_proveedor)
                    If diccionario_importe_subtotal Is Nothing Or IsEmpty(diccionario_importe_subtotal) Then Set diccionario_importe_subtotal = importeSubtotal(numero_proveedor)
                    If diccionario_importe_total Is Nothing Or IsEmpty(diccionario_importe_total) Then Set diccionario_importe_total = importeTotal(numero_proveedor)
                    If diccionario_importe_total_vencido Is Nothing Or IsEmpty(diccionario_importe_total_vencido) Then Set diccionario_importe_total_vencido = importeTotalVencido(numero_proveedor)
'                    Set via_pago = viaPago(numero_proveedor)
'                    Set categoria_costo = categoriaDeCosto(numero_proveedor)
'                    Set banco_propio = bancoPropio(numero_proveedor)
'                    Set numero_medidor = numeroMedidor(numero_proveedor)
'                    Set provincia_proveedor = provincia(numero_proveedor)
'                    Set domicilio_proveedor = domicilio(numero_proveedor) 'domicilio de TASA o TMA.
                    If diccionario_impuesto_IVA Is Nothing Or IsEmpty(diccionario_impuesto_IVA) Then Set diccionario_impuesto_IVA = IVA(numero_proveedor)
'                    Set impuesto_IIBB = IIBB(numero_proveedor) '~~> el porcentaje varia segun cliente y provincia
                
                    'comparo cada linea de la factura con los parametros y cuando coincide
                    'controlo que tenga la estructura esperada.
                    If Not IsEmpty(diccionario_numero_factura) Then
                        For x = 1 To (diccionario_numero_factura.count)
                            ocurrencia = 0
                            For j = 0 To UBound(Split(diccionario_numero_factura(x), "|"))
                                If InStr(1, LCase(arrFileLines(i)), Split(diccionario_numero_factura(x), "|")(j), vbBinaryCompare) > 0 Then
                                    ocurrencia = ocurrencia + 1
                                End If
                            Next
                            If ocurrencia = (UBound(Split(diccionario_numero_factura(x), "|")) + 1) Then
                                'numeroDeFactura = LCase(arrFileLines(i))
                                numeroDeFactura = f_nroFactura(numero_proveedor, LCase(arrFileLines(i)), x, diccionario_numero_factura)
                                Exit For
                            End If
                        Next
                    End If
                    If Not diccionario_periodo_facturacion Is Nothing Then
                        For x = 1 To (diccionario_periodo_facturacion.count)
                            If InStr(1, LCase(arrFileLines(i)), LCase(diccionario_periodo_facturacion(x)), vbBinaryCompare) > 0 Then
                                periodoFacturacion = LCase(arrFileLines(i))
                                Exit For
                            End If
                        Next
                    End If
                    If Not diccionario_fecha_vencimiento Is Nothing Then
                        For x = 1 To (diccionario_fecha_vencimiento.count)
                            If InStr(1, LCase(arrFileLines(i)), LCase(diccionario_fecha_vencimiento(x)), vbBinaryCompare) > 0 Then
                                'fechaDeVencimiento = LCase(arrFileLines(i))
                                fechaDeVencimiento = f_fechaVencimiento(LCase(arrFileLines(i)))
                                Exit For
                            End If
                        Next
                    End If
                    If Not diccionario_fecha_2vencimiento Is Nothing Then
                        For x = 1 To (diccionario_fecha_2vencimiento.count)
                            If InStr(1, LCase(arrFileLines(i)), LCase(diccionario_fecha_2vencimiento(x)), vbBinaryCompare) > 0 Then
                                'fechaDe2vencimiento = LCase(arrFileLines(i))
                                fechaDe2Vencimiento = f_fecha2Vencimiento(LCase(arrFileLines(i)))
                                Exit For
                            End If
                        Next
                    End If
                    If Not diccionario_importe_subtotal Is Nothing Then
                        For x = 1 To (diccionario_importe_subtotal.count)
                            If InStr(1, LCase(arrFileLines(i)), LCase(diccionario_importe_subtotal(x)), vbBinaryCompare) > 0 Then
                                'Subtotal = LCase(arrFileLines(i))
                                Subtotal = f_Subtotal(LCase(arrFileLines(i)))
                                Exit For
                            End If
                        Next
                    End If
                    If Not diccionario_importe_total Is Nothing Then
                        For x = 1 To (diccionario_importe_total.count)
                            If InStr(1, LCase(arrFileLines(i)), LCase(diccionario_importe_total(x)), vbBinaryCompare) > 0 Then
                                'Total = LCase(arrFileLines(i))
                                Total = f_Total(LCase(arrFileLines(i)))
                                Exit For
                            End If
                        Next
                    End If
                    If Not diccionario_importe_total_vencido Is Nothing Then
                        For x = 1 To (diccionario_importe_total_vencido.count)
                            If InStr(1, LCase(arrFileLines(i)), LCase(diccionario_importe_total_vencido(x)), vbBinaryCompare) > 0 Then
                                'TotalVencido = LCase(arrFileLines(i))
                                TotalVencido = f_TotalVencido(LCase(arrFileLines(i)))
                                Exit For
                            End If
                        Next
                    End If
                    If Not via_pago Is Nothing Then
                        For x = 1 To (via_pago.count)
                            If InStr(1, LCase(arrFileLines(i)), LCase(via_pago(x)), vbBinaryCompare) > 0 Then
                                'viaDePago = LCase(arrFileLines(i))
                                viaDePago = viaPago(LCase(arrFileLines(i)))
                                Exit For
                            End If
                        Next
                    End If
                    If Not categoria_costo Is Nothing Then
                        For x = 1 To (categoria_costo.count)
                            If InStr(1, LCase(arrFileLines(i)), LCase(categoria_costo(x)), vbBinaryCompare) > 0 Then
                                categoriaCosto = LCase(arrFileLines(i))
                                Exit For
                            End If
                        Next
                    End If
                    If Not banco_propio Is Nothing Then
                        For x = 1 To (banco_propio.count)
                            If InStr(1, LCase(arrFileLines(i)), LCase(banco_propio(x)), vbBinaryCompare) > 0 Then
                                banco = LCase(arrFileLines(i))
                                Exit For
                            End If
                        Next
                    End If
                    If Not numero_medidor Is Nothing Then
                        For x = 1 To (numero_medidor.count)
                            If InStr(1, LCase(arrFileLines(i)), LCase(numero_medidor(x)), vbBinaryCompare) > 0 Then
                                numeroDeMedidor = LCase(arrFileLines(i))
                                Exit For
                            End If
                        Next
                    End If
                    If Not provincia_proveedor Is Nothing Then
                        For x = 1 To (provincia_proveedor.count)
                            If InStr(1, LCase(arrFileLines(i)), LCase(provincia_proveedor(x)), vbBinaryCompare) > 0 Then
                                provinciaDeProveedor = LCase(arrFileLines(i))
                                Exit For
                            End If
                        Next
                    End If
                    If Not domicilio_proveedor Is Nothing Then
                        For x = 1 To (domicilio_proveedor.count)
                            If InStr(1, LCase(arrFileLines(i)), LCase(domicilio_proveedor(x)), vbBinaryCompare) > 0 Then
                                domicilioDeCliente = LCase(arrFileLines(i))
                                Exit For
                            End If
                        Next
                    End If
                    If Not diccionario_impuesto_IVA Is Nothing Then
                        For x = 1 To (diccionario_impuesto_IVA.count)
                            ocurrencia = 0
                            For j = 0 To UBound(Split(diccionario_impuesto_IVA(x), "|"))
                                If InStr(1, LCase(arrFileLines(i)), Split(diccionario_impuesto_IVA(x), "|")(j), vbBinaryCompare) > 0 Then
                                    ocurrencia = ocurrencia + 1
                                End If
                            Next
                            If ocurrencia = (UBound(Split(diccionario_impuesto_IVA(x), "|")) + 1) Then
                                'numeroDeFactura = LCase(arrFileLines(i))
                                impuesto_IVA = f_IVA(numero_proveedor, LCase(arrFileLines(i)), x, diccionario_impuesto_IVA)
                                Exit For
                            End If
                        Next
                    End If
                    If Not impuesto_IIBB Is Nothing Then
                        For x = 1 To (impuesto_IIBB.count)
                            If InStr(1, LCase(arrFileLines(i)), LCase(impuesto_IIBB(x)), vbBinaryCompare) > 0 Then
                                impuestoIIBB = LCase(arrFileLines(i))
                                Exit For
                            End If
                        Next
                    End If
                End If
                i = i + 1
            Loop
            
            '~~> con el numero de proveedor hago un vlookup en la hoja "hoja_impuestos_" para obtener los siguientes datos:
            hoja_impuestos_sociedadSAP = Excel.Application.VLookup(numero_proveedor, wsImpuestos, 2, 0) '~~> numero_proveedor se cruza con tabla correspondiente
            hoja_impuestos_N°ProveedorSAP = Excel.Application.VLookup(numero_proveedor, wsImpuestos, 3, 0) '~~> numero_proveedor se cruza con tabla correspondiente
            hoja_impuestos_viaPago = Excel.Application.VLookup(numero_proveedor, wsImpuestos, 7, 0)
            hoja_impuestos_viaPagoSuplementaria = Excel.Application.VLookup(numero_proveedor, wsImpuestos, 8, 0) '~~> numero_proveedor se cruza con tabla correspondiente
            hoja_impuestos_NIF = Excel.Application.VLookup(numero_proveedor, wsImpuestos, 4, 0) '~~> CUIT
            hoja_impuestos_IVA = Excel.Application.VLookup(numero_proveedor, wsImpuestos, 11, 0) '~~> IVA / exento
            hoja_impuestos_IVA_1% = Excel.Application.VLookup(numero_proveedor, wsImpuestos, 11, 0) '~~> IVA 1%
            hoja_impuestos_IVA_2% = Excel.Application.VLookup(numero_proveedor, wsImpuestos, 11, 0) '~~> IVA 2%
            hoja_impuestos_IVAmas% = Excel.Application.VLookup(numero_proveedor, wsImpuestos, 11, 0) '~~> IVA% + 3%
            hoja_impuestos_IIBB = Excel.Application.VLookup(numero_proveedor, wsImpuestos, , 0) '~~> IIBB / exento
            hoja_impuestos_IIBB_1% = Excel.Application.VLookup(numero_proveedor, wsImpuestos, , 0) '~~> IIBB 1%
            hoja_impuestos_IIBB_2% = Excel.Application.VLookup(numero_proveedor, wsImpuestos, , 0) '~~> IIBB 2%
            hoja_impuestos_IIBB_3% = Excel.Application.VLookup(numero_proveedor, wsImpuestos, , 0) '~~> IIBB 3%
            hoja_impuestos_provincia = Excel.Application.VLookup(numero_proveedor, wsImpuestos, , 0)
            
            

            
    '~~>
        'HES
        
End Sub
