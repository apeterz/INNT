' Devuelve Fecha en formato dd-mm-yyyy + "DifDias" días. 
'
'
DifDias = CInt(Wscript.Arguments.Item(0))
Dim fechaHoy 
fecha = DateAdd("d", DifDias, Now())

Dim result
result = Right("0" & Day(fecha), 2) & "-" & Right("0" & Month(fecha), 2) _
                              & "-" & Year(fecha)

WScript.StdOut.WriteLine(result)

