Attribute VB_Name = "Módulo1"
Option Explicit

'copiando informacion de una hoja a otra, filtrando por correo hotmail
Sub buscar_correo_hotmail()

Dim celda As Range
Dim f As Integer
f = 578
Dim f2 As Integer

For Each celda In Worksheets(2).Range("I2:I1434")

    If celda.Value Like "*hotmail.com" Then
    
        celda.Font.Color = vbRed
        
        celda.Copy
        
        Worksheets(1).Cells(f, 4).PasteSpecial xlPasteValues
        
        f2 = celda.Row
        
        Worksheets(2).Cells(f2, 7).Copy 'copia el nombre
        
        Worksheets(1).Cells(f, 2).PasteSpecial xlPasteValues
        
        Worksheets(2).Cells(f2, 5).Copy 'copia telefono fijo
        
        Worksheets(1).Cells(f, 5).PasteSpecial xlPasteValues
        
        Worksheets(2).Cells(f2, 10).Copy 'copia celular
        
        Worksheets(1).Cells(f, 6).PasteSpecial xlPasteValues
          
        Worksheets(2).Cells(f2, 19).Copy
        
        Worksheets(1).Cells(f, 8).PasteSpecial xlPasteValues
              
        Application.CutCopyMode = False
        
        
        f = f + 1
        
    End If
    
    
Next celda

End Sub

'copiando informacion de una hoja a otra, filtrando por correo gmail
Sub buscar_correo_gmail()

Dim celda As Range
Dim f As Integer
f = 768
Dim f2 As Integer

For Each celda In Worksheets(2).Range("I2:I1434")

    If celda.Value Like "*gmail.com" Then
    
        celda.Font.Color = vbRed
        
        celda.Copy
        
        Worksheets(1).Cells(f, 4).PasteSpecial xlPasteValues
        
        f2 = celda.Row
        
        Worksheets(2).Cells(f2, 7).Copy 'copia el nombre
        
        Worksheets(1).Cells(f, 2).PasteSpecial xlPasteValues
        
        Worksheets(2).Cells(f2, 5).Copy 'copia telefono fijo
        
        Worksheets(1).Cells(f, 5).PasteSpecial xlPasteValues
        
        Worksheets(2).Cells(f2, 10).Copy 'copia celular
        
        Worksheets(1).Cells(f, 6).PasteSpecial xlPasteValues
          
        Worksheets(2).Cells(f2, 19).Copy
        
        Worksheets(1).Cells(f, 8).PasteSpecial xlPasteValues
              
        Application.CutCopyMode = False
        
        
        f = f + 1
        
    End If
    
    
Next celda

End Sub

'copiando informacion de una hoja a otra, filtrando por correo yahoo
Sub buscar_correo_yahoo()

Dim celda As Range
Dim f As Integer
f = 933
Dim f2 As Integer

For Each celda In Worksheets(2).Range("I2:I1434")

    If celda.Value Like "*yahoo.com" Then
    
        celda.Font.Color = vbRed
        
        celda.Copy
        
        Worksheets(1).Cells(f, 4).PasteSpecial xlPasteValues
        
        f2 = celda.Row
        
        Worksheets(2).Cells(f2, 7).Copy 'copia el nombre
        
        Worksheets(1).Cells(f, 2).PasteSpecial xlPasteValues
        
        Worksheets(2).Cells(f2, 5).Copy 'copia telefono fijo
        
        Worksheets(1).Cells(f, 5).PasteSpecial xlPasteValues
        
        Worksheets(2).Cells(f2, 10).Copy 'copia celular
        
        Worksheets(1).Cells(f, 6).PasteSpecial xlPasteValues
          
        Worksheets(2).Cells(f2, 19).Copy
        
        Worksheets(1).Cells(f, 8).PasteSpecial xlPasteValues
              
        Application.CutCopyMode = False
        
        
        f = f + 1
        
    End If
    
    
Next celda

End Sub

