Attribute VB_Name = "y_paths"
'SETEO DE DIRECTIORIOS DE ARCHIVOS

Public Function generalPath() As String
    generalPath = "C:\Users\" & Environ("USERNAME") & "\Documents\telefonica\"
End Function

Public Function invoicesPdfPath() As String
    invoicesPdfPath = "C:\Users\" & Environ("USERNAME") & "\Documents\telefonica\facturas_pdf\"
End Function

Public Function invoicesTxtPath() As String
    invoicesTxtPath = "C:\Users\" & Environ("USERNAME") & "\Documents\telefonica\facturas_txt\"
End Function

Public Function scriptsPath() As String
    scriptsPath = "C:\Users\" & Environ("USERNAME") & "\Documents\telefonica\scripts\"
End Function

Public Function Destinatario() As String
    Destinatario = "PSInvoices@bhpbilliton.com"
End Function
