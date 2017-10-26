Attribute VB_Name = "Módulo1"
Option Explicit

Sub matrices_dos_dimensiones()

Dim mimatriz(3, 4) As Integer 'Declaracion de una matriz de dos dimensiones (columna, fila)

Dim i As Integer, z As Integer

For i = 0 To 2 Step 1 'Primer bucle para recorrer las columnas

    For z = 0 To 3 Step 1 'Segundo bucle para recorreo las filas
    
        mimatriz(i, z) = Math.Round(Math.Rnd * 100) 'Asignando informacion a la posicion de la matriz
        
        Debug.Print mimatriz(i, z)
        
    Next z
Next i

End Sub
