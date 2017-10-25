Attribute VB_Name = "Módulo5"
Option Explicit

Sub declaracion_matrices()

'Primera forma de declarar una matriz
Dim mirango(4) As Integer

mirango(0) = 55 'Asignación de datos a la posición cero de la matriz
mirango(1) = 52 'Asignación de datos a la posición uno de la matriz
mirango(2) = 335 'Asignación de datos a la posición dos de la matriz
mirango(3) = 35 'Asignación de datos a la posición tres de la matriz
mirango(4) = 15 'Asignación de datos a la posición cuatro de la matriz

MsgBox mirango(4) 'Mostrando el valor de la posición 4

'Segunda forma de declarar una matriz
Dim mimatriz(0 To 4) As Integer

mimatriz(0) = 31
mimatriz(1) = 33
mimatriz(2) = 43
mimatriz(3) = 23
mimatriz(4) = 53

End Sub

Sub introducir_valores()

Dim mirango(5) As Integer 'Defiendo la matriz

Dim celda_seleccionada As Range 'Defiendo una variable de tipo posicion

Dim indice As Integer 'Definiendo el contador, por defecto su valor sera cero

Range("a1").Select 'Posicionando el cursor en la celda a1

Selection.CurrentRegion.Select 'seleccionando las celdas contiguas (con valores) a la celda a1

For Each celda_seleccionada In Selection 'Recorriendo el grupo de celdas seleccionadas y asignandolas a
                                         'a la variable celda_seleccionada

mirango(indice) = celda_seleccionada.Value 'llenando las posiciones de la matriz con el valor de la celda seleccionada

indice = indice + 1 'Aumentando el valor del indice

Next celda_seleccionada 'Fin del bucle for

'Recorriendo los valores almacenados en la matriz y mostrandolos

Dim i As Integer

For i = 0 To 5 Step 1

    Debug.Print mirango(i)
    
Next i

End Sub

