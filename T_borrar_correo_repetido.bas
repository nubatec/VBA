Attribute VB_Name = "Módulo11"
Option Explicit
'Busca correos marcados de rojo y elimina la fila asociada
Sub borrar_correo_repetido()

For Each celda In Worksheets(2).Range("I2:I1434")

    If celda.Font.Color = vbRed Then

        celda.EntireRow.Delete

    End If

Next celda

End Sub
