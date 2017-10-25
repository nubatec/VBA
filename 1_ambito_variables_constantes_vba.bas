Attribute VB_Name = "Módulo4"
Option Explicit
Public Const tipo_cambio As Double = 3.25 'declaración de constante pública
Public nombre As String 'declaración de variable ública

Sub declaracion_constantes()

Dim monto As Double 'declaración de variable local

Const valor As Integer = 7 'declaración de constante local

nombre = "Diego" 'asignación de valor a variable local

monto = tipo_cambio + valor

MsgBox "El monto a cobrar por " & nombre & " es : " & monto

End Sub
