Attribute VB_Name = "search"

Sub buscar()
Dim finNouveauTableau, encontrado As Long
Dim Item, celda As Variant
Dim palabras(37) As Variant
Dim numColumnpago, numColumnVacia, numColumnDate, i, mes, ano As Integer
Dim work1 As Workbook
Dim ws As Worksheet
Dim sociedad As String

Set work1 = ActiveWorkbook
Set ws = work1.Worksheets("BANCARIOS")
numColumnpago = 14
numColumnVacia = 6
numColumnDate = 4

finNouveauTableau = work1.Worksheets("BANCARIOS").Range("A" & Rows.count).End(xlUp).Row

palabras(1) = Array("SAISIE", "SAISIES", "saisie", "saisies", "PENSIONS ALIM", "PR OVENCE")
palabras(2) = Array("monnaie", "MONNAIE")
palabras(3) = Array("retraite", "RETRAITE")
palabras(4) = Array("salaire", "SALAIRE", "VIREME NT SALAIRE", "STC", "VIRT S TC", "SALARIES", "SALAIRES", "ALAIRES", "ALAIRE", "F02 SA LAIRES", "LAIRE", "LAIRES", "POPULAIRE")
palabras(5) = Array("acompte", "ACOMPTE", "A COMPTE")
palabras(6) = Array("Interessement", "INTERESSEMENT")
palabras(7) = Array("IMPOT", "impot", "impôt", "IMPÔT", "impôts", "IMPÔTS", "PASDSN")
palabras(8) = Array("DGFIPIMPOT")
palabras(9) = Array(" DIRECTION GENERALE DES FINANCES PUBLIQUES", "D.G.F.I.P")
palabras(10) = Array("Mutuelle", "MUTUELLE", "mutuele", "muttuelle", "mutuelles", "MUTUELLES", "MUTUEL LE")


For Each celda In Range("I2:I" & finNouveauTableau)

 If celda.Offset(, numColumnVacia) = "" Then
 
mes = Month(celda.Offset(, numColumnVacia - 10))
ano = Year(celda.Offset(, numColumnVacia - 10))

    For Each Item In palabras(1)
    encontrado = InStr(celda, Item)
        If encontrado > 0 Then
        celda.Offset(, numColumnVacia) = "X"
        celda.Offset(, numColumnVacia + 1) = "4658001"
        End If
    Next Item

    For Each Item In palabras(2)
    encontrado = InStr(celda, Item)
        If encontrado > 0 Then
        celda.Offset(, numColumnVacia) = "X"
        celda.Offset(, numColumnVacia + 1) = "7780000"
        celda.Offset(, numColumnVacia + 2) = "F99V9"
        celda.Offset(, numColumnVacia + 3) = "V990"
        celda.Offset(, numColumnVacia + 4) = "CV990"
        celda.Offset(, numColumnVacia + 5) = "MONNAIE DE PARIS SVF SAZIAS"
        End If
    Next Item
    
    For Each Item In palabras(3)
    'La función InStr devuelve la posición de un carácter dentro de la cadena
    encontrado = InStr(celda, Item)
        If encontrado > 0 Then
        sociedad = celda.Offset(, numColumnVacia - 13)
        sociedad = Left(sociedad, 3)
            If sociedad = "F00" Then
            celda.Offset(, numColumnVacia + 2) = "Comprobar si es de F02 en el fichero"
            End If
        celda.Offset(, numColumnVacia) = "X"
        celda.Offset(, numColumnVacia + 1) = "4765014"
        'La función Mid se moverá a la posición "encontrado" de la cadena y a partir de ahí contará 20 caracteres
        celda.Offset(, numColumnVacia + 5) = Mid(celda, encontrado, 40)
        End If
    Next Item
    
    'enlever populaire
    For Each Item In palabras(4)
    encontrado = InStr(celda, Item)
        If encontrado > 0 Then
            celda.Offset(, numColumnVacia) = "X"
            celda.Offset(, numColumnVacia + 1) = "4650000"
                If Item = "STC" Then
                celda.Offset(, numColumnVacia + 5) = "STC " & mes & "/" & ano
                Else
                    If celda.Offset(, numColumnVacia - 1) - Fix(celda.Offset(, numColumnVacia - 1)) = 0 Then
                    celda.Offset(, numColumnVacia + 5) = "Podria ser un acompte / 4600000"
                    Else
                    celda.Offset(, numColumnVacia + 2) = "VIREMENT SALAIRE " & mes & "/" & ano
                    End If
                End If
                   If Item = "POPULAIRE" Then
                   celda.Offset(, numColumnVacia) = ""
                   celda.Offset(, numColumnVacia + 1) = ""
                   celda.Offset(, numColumnVacia + 2) = ""
                   celda.Offset(, numColumnVacia + 5) = ""
                   End If
        End If
    Next Item
    
    For Each Item In palabras(5)
    encontrado = InStr(celda, Item)
        If encontrado > 0 Then
        celda.Offset(, numColumnVacia) = "X"
        celda.Offset(, numColumnVacia + 1) = "4600000"
        celda.Offset(, numColumnVacia + 5) = "ACOMPTE " & mes & "/" & ano
        End If
    Next Item
    
    For Each Item In palabras(6)
    encontrado = InStr(celda, Item)
        If encontrado > 0 Then
        celda.Offset(, numColumnVacia) = "X"
        celda.Offset(, numColumnVacia + 1) = "4651001"
        celda.Offset(, numColumnVacia + 5) = "INTERESSEMENT " & mes & "/" & ano
        End If
    Next Item
    
    For Each Item In palabras(7)
    encontrado = InStr(celda, Item)
        If encontrado > 0 Then
        celda.Offset(, numColumnVacia) = "X"
        celda.Offset(, numColumnVacia + 1) = "4751000"
        celda.Offset(, numColumnVacia + 5) = "IMPÔTS S/REVENUS " & mes & "/" & ano
        End If
    Next Item
    
    For Each Item In palabras(8)
    encontrado = InStr(celda, Item)
        If encontrado > 0 Then
        celda.Offset(, numColumnVacia) = ""
        celda.Offset(, numColumnVacia + 1) = "auto TAYA"
        celda.Offset(, numColumnVacia + 5) = "No contabilizar (auto TAYA)"
        End If
    Next Item
    
    For Each Item In palabras(9)
    encontrado = InStr(celda, Item)
        If encontrado > 0 Then
        celda.Offset(, numColumnVacia) = "X"
        celda.Offset(, numColumnVacia + 1) = "6310007"
        celda.Offset(, numColumnVacia + 2) = "F06T0"
        celda.Offset(, numColumnVacia + 3) = "T990"
        celda.Offset(, numColumnVacia + 4) = "AT204"
        celda.Offset(, numColumnVacia + 5) = "RGLMT SOLDE CFE " & mes & "/" & ano
        End If
    Next Item
    
    For Each Item In palabras(10)
    encontrado = InStr(celda, Item)
        If encontrado > 0 Then
        celda.Offset(, numColumnVacia) = "X"
        celda.Offset(, numColumnVacia + 1) = "4765023"
        celda.Offset(, numColumnVacia + 5) = "MUTUELLE " & mes & "/" & ano
        End If
    Next Item

    
Next celda
End Sub
