Attribute VB_Name = "Módulo1"
Option Explicit

Sub variable_objeto()

Dim micelda As Range 'Declarando una variable de tipo objeto

'Defiendo propiedades en una celda sin utilizar variables de tipo objeto
'Worksheets(1).Range("b5").Value = 124
'Worksheets(1).Range("b5").Font.Bold = True
'Worksheets(1).Range("b5").Font.Italic = True

'Defiendo propiedades de una celda utilizando variables de tipo objeto
'Set micelda = Worksheets(1).Range("c5")
'micelda.Value = 134
'micelda.Font.Bold = True
'micelda.Font.Italic = True

'Defiendo propiedades de una celda utilizando variables de tipo objeto y un bucle
Set micelda = Worksheets(1).Range("c5")
With micelda
    .Value = 229
    .Font.Bold = True
    .Font.Italic = True
End With

End Sub

'Recorriendo las hojas dentro del libro y sacando su nombre
Sub leyendo_hojas()

Dim mihoja As Worksheet 'declarando un variable de tipo hoja

For Each mihoja In Worksheets 'recorriendo cada hoja del libro

    MsgBox mihoja.Name 'mostrando el nombre de la hoja
    
Next mihoja

End Sub

'Analizando un grupo de celdas seleccionadas y marcando con rojo las superiores a 300
Sub mi_formato()

Dim micelda As Range 'Declarando una variable de tipo "rango de celdas"

For Each micelda In Selection 'Recorremos cada celda dentro de lo seleccionado

    If micelda.Value >= 300 Then 'Si el valor de la celda es mayor a 300 se marcara de rojo
    
        micelda.Interior.Color = vbRed
        
    End If
    
Next micelda

End Sub

Sub localizar()
Attribute localizar.VB_ProcData.VB_Invoke_Func = "k\n14"

Dim mihoja As Worksheet

For Each mihoja In Worksheets

    If mihoja.Name = "Hoja12" Then
        mihoja.Select
        
    End If
    
Next mihoja

End Sub

