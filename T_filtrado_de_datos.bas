Attribute VB_Name = "Módulo1"
Option Explicit
'Busca correos y celulares repetidos y elimina las filas asociadas a estos
Sub borrar_repetido_correo()

Dim celda_correo, celda_correo2 As Range
For Each celda_correo In Worksheets(1).Range("D2:D577")

    For Each celda_correo2 In Worksheets(2).Range("I2:I2131")

        If celda_correo2.Value = celda_correo.Value Then

            celda_correo2.EntireRow.Delete

        End If

    Next celda_correo2

Next celda_correo

Dim celda_celular, celda_celular2 As Range
For Each celda_celular In Worksheets(1).Range("F2:F577")

    For Each celda_celular2 In Worksheets(2).Range("J2:J1532")

        If celda_celular2.Value = celda_celular.Value Then

            celda_celular2.EntireRow.Delete

        End If

    Next celda_celular2

Next celda_celular

End Sub

'Busca los correos repetidos dentro de un rango y los marca de color rojo
Sub buscar_correo_repetido()

Dim correo, correo2 As Range
Dim contador As Integer

For Each correo In Worksheets(1).Range("I2:I1532")
    
    contador = 0

    For Each correo2 In Worksheets(1).Range("J2:J1532")
    
        If correo2.Value = correo.Value Then
        
            contador = contador + 1
            
        End If
        
        If contador > 1 Then
        
            correo.Font.Color = vbRed
        
        End If
    
    Next correo2
    
Next correo

End Sub

'Busca los numero de celular repetidos dentro de un rango y los marca de color rojo
Sub buscar_celular_repetido()

Dim celular, celular2 As Range
Dim contador As Integer

For Each celular In Worksheets(1).Range("J2:J1532")
    
    contador = 0

    For Each celular2 In Worksheets(1).Range("K2:K1532")
    
        If celular2.Value = celular.Value Then
        
            contador = contador + 1
            
        End If
        
        If contador > 1 Then
        
            celular.Font.Color = vbRed
        
        End If
    
    Next celular2
    
Next celular

End Sub
