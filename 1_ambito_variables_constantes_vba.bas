Attribute VB_Name = "M�dulo4"
Option Explicit
Public Const tipo_cambio As Double = 3.25 'declaraci�n de constante p�blica
Public nombre As String 'declaraci�n de variable �blica

Sub declaracion_constantes()

Dim monto As Double 'declaraci�n de variable local

Const valor As Integer = 7 'declaraci�n de constante local

nombre = "Diego" 'asignaci�n de valor a variable local

monto = tipo_cambio + valor

MsgBox "El monto a cobrar por " & nombre & " es : " & monto

End Sub
