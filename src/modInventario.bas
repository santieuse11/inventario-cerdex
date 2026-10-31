Attribute VB_Name = "modInventario"
Option Explicit

Public Sub AgregarMovimiento()
    Dim wsForm As Worksheet, wsMov As Worksheet, wsProd As Worksheet
    Dim codigo As String, tipo As String, tipoNormalizado As String, responsable As String
    Dim documento As String, tercero As String, observacion As String
    Dim fecha As Date, cantidad As Double, existencia As Double
    Dim unidad As String, factorTubo As Double, equivalencia As Variant
    Dim filaDestino As Long, filaFormula As Long, filaActual As Long
    Dim encontrado As Range

    Set wsForm = ThisWorkbook.Worksheets("Registro")
    Set wsMov = ThisWorkbook.Worksheets("Movimientos")
    Set wsProd = ThisWorkbook.Worksheets("Productos")

    If Not IsDate(wsForm.Range("B5").Value) Then
        MsgBox "Ingresa una fecha valida.", vbExclamation, "Fecha requerida"
        Exit Sub
    End If
    fecha = CDate(wsForm.Range("B5").Value)
    tipo = Trim$(CStr(wsForm.Range("B6").Value))
    tipoNormalizado = LCase$(tipo)
    tipoNormalizado = Replace(tipoNormalizado, ChrW(243), "o")
    codigo = Trim$(CStr(wsForm.Range("B7").Value))
    responsable = Trim$(CStr(wsForm.Range("B12").Value))
    documento = Trim$(CStr(wsForm.Range("B13").Value))
    tercero = Trim$(CStr(wsForm.Range("B14").Value))
    observacion = Trim$(CStr(wsForm.Range("B15").Value))

    If tipo = "" Or codigo = "" Then
        MsgBox "Selecciona el tipo de movimiento y el producto.", vbExclamation, "Faltan datos"
        Exit Sub
    End If
    Select Case tipoNormalizado
        Case "ingreso", "salida", "devolucion cliente", "devolucion proveedor"
        Case Else
            MsgBox "Selecciona un tipo de movimiento de la lista.", vbExclamation, "Tipo no valido"
            Exit Sub
    End Select
    If Not IsNumeric(wsForm.Range("B11").Value) Then
        MsgBox "Ingresa una cantidad numerica mayor que cero.", vbExclamation, "Cantidad no valida"
        Exit Sub
    End If
    cantidad = CDbl(wsForm.Range("B11").Value)
    If cantidad <= 0 Then
        MsgBox "La cantidad debe ser mayor que cero.", vbExclamation, "Cantidad no valida"
        Exit Sub
    End If
    If responsable = "" Then
        MsgBox "Ingresa quien entrega o recibe el producto.", vbExclamation, "Responsable requerido"
        Exit Sub
    End If

    Set encontrado = wsProd.Range("A5:A100").Find(What:=codigo, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
    If encontrado Is Nothing Then
        MsgBox "No encontre el codigo en Productos.", vbExclamation, "Producto no valido"
        Exit Sub
    End If
    existencia = 0
    If IsNumeric(wsProd.Cells(encontrado.Row, 6).Value) Then existencia = CDbl(wsProd.Cells(encontrado.Row, 6).Value)
    unidad = LCase$(Trim$(CStr(wsProd.Cells(encontrado.Row, 4).Value)))
    If (tipoNormalizado = "salida" Or tipoNormalizado = "devolucion proveedor") And cantidad > existencia Then
        MsgBox "La cantidad supera la existencia actual (" & Format(existencia, "#,##0.##") & ").", vbExclamation, "Existencia insuficiente"
        Exit Sub
    End If

    equivalencia = Empty
    If unidad = "m" And IsNumeric(wsProd.Cells(encontrado.Row, 9).Value) Then
        factorTubo = CDbl(wsProd.Cells(encontrado.Row, 9).Value)
        If factorTubo > 0 Then equivalencia = cantidad / factorTubo
    End If
    If documento = "" Then documento = "MOV-" & Format(Now, "yyyymmdd-hhnnss")

    filaDestino = 0
    For filaActual = 5 To 200
        If Len(Trim$(CStr(wsMov.Cells(filaActual, 2).Value))) = 0 Then
            filaDestino = filaActual
            Exit For
        End If
    Next filaActual
    If filaDestino = 0 Then
        MsgBox "El registro alcanzo la fila 200. Amplia Movimientos antes de continuar.", vbExclamation, "Registro lleno"
        Exit Sub
    End If

    With wsMov
        .Cells(filaDestino, 1).Value = fecha
        .Cells(filaDestino, 1).NumberFormatLocal = "dd-mmm-aaaa"
        .Cells(filaDestino, 2).Value = documento
        .Cells(filaDestino, 3).Value = codigo
        .Cells(filaDestino, 4).Value = tipo
        .Cells(filaDestino, 5).Value = responsable
        .Cells(filaDestino, 6).Value = cantidad
        .Cells(filaDestino, 6).NumberFormat = "#,##0.##"
        If IsEmpty(equivalencia) Then
            .Cells(filaDestino, 7).ClearContents
        Else
            .Cells(filaDestino, 7).Value = equivalencia
            .Cells(filaDestino, 7).NumberFormat = "#,##0.##"
        End If
        .Cells(filaDestino, 8).Value = observacion
        .Cells(filaDestino, 9).Value = tercero
    End With

    Application.CalculateFull
    filaFormula = encontrado.Row
    wsForm.Range("A17").Value = "Registrado: " & tipo & " de " & Format(cantidad, "#,##0.##") & " " & unidad & _
        " para " & codigo & ". Existencia actual: " & Format(wsProd.Cells(filaFormula, 6).Value, "#,##0.##")
    wsForm.Range("A17").Font.Color = RGB(22, 110, 80)
    wsForm.Range("B6:B7").ClearContents
    wsForm.Range("B11:B15").ClearContents
    wsForm.Range("B5").Value = Date
    wsForm.Range("B5").NumberFormatLocal = "dd-mmm-aaaa"
    MsgBox "Movimiento agregado y existencia actualizada.", vbInformation, "Inventario actualizado"
End Sub
