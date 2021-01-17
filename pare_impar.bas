Attribute VB_Name = "pare_impar"
'**************
'fecha creacion 09/2019
'Abrir fichero PARES. La macro indica si el importe es pare o impare.
'**************
Sub Intercia()
Dim work1 As Workbook
Dim ws As Worksheet
Dim i, j, z, t As Integer
Dim contador As Integer
Dim ultimaLinea As Variant
Dim datos() As Variant
Dim valorSolucion() As Variant
Dim va, vas As Variant
Dim vaNegativa, vasNeg, VaNegativaLeida As Double
Dim res As Double
'when turned on - will visibly update the Excel worksheet on your screen
Application.ScreenUpdating = False
res = 0

Set work1 = ActiveWorkbook
Set ws = work1.Worksheets("pares")
ultimaLinea = ws.Cells(Rows.count, "C").End(xlUp).Row
ws.Range("A1:D" & ultimaLinea).Sort Key1:=Range("C1"), Order1:=xlDescending, Header:=xlYes

ReDim datos(ultimaLinea)
'UBound para determinar el mayor subíndice disponible para la dimensión indicada de una matriz.
For i = 2 To (UBound(datos))
    contador = 1
    ReDim valorSolucion(ultimaLinea)
    datos(i) = ws.Cells(i, 3).Value
    va = ws.Cells(i, 3).Value

If va <> 0 Then
valorSolucion(contador) = i

    For j = 1 To (UBound(datos))
        vas = ws.Cells(i + j, 3).Value
        If vas <> va Then
             If res = 0 Then
                res = va
             End If
        Exit For
        End If
    
        If vas = va Then
            contador = contador + 1
            valorSolucion(contador) = i + j
            If res <> 0 Then
            res = res + vas
            Else
            res = va + vas
            End If
        End If
    Next 'j
    
    vaNegativa = ws.Cells(i, 3).Value * -1
    
        For z = 0 To (UBound(datos)) - 2
            VaNegativaLeida = ws.Cells(ultimaLinea - z, 3).Value
    
            If VaNegativaLeida = vaNegativa Then
            contador = contador + 1
            valorSolucion(contador) = ultimaLinea - z
            res = res + vaNegativa
            
           End If
       Next 'z

ReDim Preserve valorSolucion(contador)
Dim DerLigC As Integer
res = Round(res, 2)
 If NumerosPares(contador) And SumaCero(res) Then
     For j = 1 To (UBound(valorSolucion))
        ws.Range(ws.Cells(valorSolucion(j), 1), ws.Cells(valorSolucion(j), 4)).Copy
        DerLigC = ws.Range("F" & Rows.count).End(xlUp).Row
        ws.Range("F" & DerLigC + 1).PasteSpecial Paste:=xlPasteValues
        ws.Range(ws.Cells(valorSolucion(j), 1), ws.Cells(valorSolucion(j), 4)).ClearContents
    Next
       Else ' si impair copia columna "M"
    For j = 1 To (UBound(valorSolucion))
        ws.Range(ws.Cells(valorSolucion(j), 1), ws.Cells(valorSolucion(j), 4)).Copy
        DerLigC = ws.Range("K" & Rows.count).End(xlUp).Row
        ws.Range("K" & DerLigC + 1).PasteSpecial Paste:=xlPasteValues
        ws.Range(ws.Cells(valorSolucion(j), 1), ws.Cells(valorSolucion(j), 4)).ClearContents
    Next
 End If
    End If
    res = 0
Next 'i
ws.Columns("H").NumberFormat = "#,##0.00_ ;[Red]-#,##0.00 "
ws.Columns("M").NumberFormat = "#,##0.00_ ;[Red]-#,##0.00 "
End Sub

Function NumerosPares(Number As Integer) As Boolean
    NumerosPares = (Number Mod 2 = 0)
End Function
Function SumaCero(Number1 As Double) As Boolean
    SumaCero = Number1 = 0
End Function

