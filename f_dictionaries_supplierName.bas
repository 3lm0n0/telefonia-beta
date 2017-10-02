Attribute VB_Name = "f_dictionaries_supplierName"
Option Explicit
'este diccionario contiene las palabaras clave para identificar al proveedor.
Public Function diccionario_numeroCliente(ByVal myArray As Collection)
    myArray.Add "mero de cliente es" + "|" + "80050225-1" + "|" + "edesur"  'edesur
    myArray.Add "su número de cliente es" + "|" + "missing" + "|" + "edenor"  'edenor
End Function
