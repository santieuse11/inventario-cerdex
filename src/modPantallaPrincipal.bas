Attribute VB_Name = "modPantallaPrincipal"
Option Explicit

' Pantalla de inicio y accesos del Kardex.
' Las existencias siguen calculandose en Productos desde Movimientos.

Private Const NOMBRE_INICIO As String = "Inicio"
Private Const FILA_PRIMER_PRODUCTO As Long = 5
Private Const FILA_ULTIMO_PRODUCTO As Long = 100

Public Sub ConfigurarPantallaPrincipal()
    Dim wsInicio As Worksheet
    Dim pantallaAnterior As Boolean

    On Error GoTo ControlError
    pantallaAnterior = Application.ScreenUpdating
    Application.ScreenUpdating = False

    Set wsInicio = ObtenerOCrearInicio()
    ConstruirPantalla wsInicio

    AgregarBotonInicio ThisWorkbook.Worksheets("Registro"), "G4"
    AgregarBotonInicio ThisWorkbook.Worksheets("Panel"), "J4"
    AgregarBotonInicio ThisWorkbook.Worksheets("Movimientos"), "A3"
    AgregarBotonInicio ThisWorkbook.Worksheets("Productos"), "A3"
    AsignarAccionGuardar

    If wsInicio.Index > 1 Then wsInicio.Move Before:=ThisWorkbook.Worksheets(1)
    Application.Calculate
    AbrirInicio

Salida:
    Application.ScreenUpdating = pantallaAnterior
    Exit Sub

ControlError:
    Application.ScreenUpdating = pantallaAnterior
    MsgBox "No se pudo configurar la pantalla principal: " & Err.Description, _
           vbExclamation, "Inventario CERDEX"
End Sub

Public Sub AbrirInicio()
    On Error GoTo ControlError
    ThisWorkbook.Worksheets(NOMBRE_INICIO).Activate
    With ActiveWindow
        .DisplayGridlines = False
        .Zoom = 60
    End With
    ThisWorkbook.Worksheets(NOMBRE_INICIO).Range("A1").Select
    Application.Calculate
    Exit Sub

ControlError:
    On Error Resume Next
    ThisWorkbook.Worksheets("Panel").Activate
    MsgBox "No encuentro la pantalla Inicio. Ejecuta ConfigurarPantallaPrincipal.", _
           vbInformation, "Inventario CERDEX"
End Sub

Public Sub AbrirRegistroEntrada()
    AbrirRegistroConTipo "Ingreso"
End Sub

Public Sub AbrirRegistroSalida()
    AbrirRegistroConTipo "Salida"
End Sub

Public Sub AbrirRegistroDevolucion()
    Dim respuesta As VbMsgBoxResult
    Dim tipo As String

    respuesta = MsgBox("Elige el tipo de devolucion:" & vbCrLf & vbCrLf & _
                       "Si: devolucion de cliente" & vbCrLf & _
                       "No: devolucion a proveedor" & vbCrLf & _
                       "Cancelar: volver al inicio", _
                       vbYesNoCancel + vbQuestion, "Registrar devolucion")

    Select Case respuesta
        Case vbYes
            tipo = TextoDevolucionCliente()
        Case vbNo
            tipo = TextoDevolucionProveedor()
        Case Else
            Exit Sub
    End Select

    AbrirRegistroConTipo tipo
End Sub

Public Sub VerInventario()
    ActivarHoja "Panel"
End Sub

Public Sub VerProductosAgotados()
    Dim ws As Worksheet
    Dim cantidad As Long

    On Error GoTo ControlError
    Set ws = ThisWorkbook.Worksheets("Productos")
    cantidad = Application.WorksheetFunction.CountIfs( _
        ws.Range("A" & FILA_PRIMER_PRODUCTO & ":A" & FILA_ULTIMO_PRODUCTO), "<>", _
        ws.Range("F" & FILA_PRIMER_PRODUCTO & ":F" & FILA_ULTIMO_PRODUCTO), "<=0")

    If cantidad = 0 Then
        LimpiarFiltroProductos ws
        ws.Activate
        MsgBox "No hay referencias agotadas en el inventario actual.", _
               vbInformation, "Productos agotados"
        Exit Sub
    End If

    AplicarFiltroProductos ws, 6, "<=0"
    ws.Activate
    ws.Range("A4").Select
    Exit Sub

ControlError:
    MsgBox "No se pudo consultar los productos agotados: " & Err.Description, _
           vbExclamation, "Inventario CERDEX"
End Sub

Public Sub VerProductosPorAgotarse()
    Dim ws As Worksheet
    Dim minimosConfigurados As Long
    Dim porAgotarse As Long

    On Error GoTo ControlError
    Set ws = ThisWorkbook.Worksheets("Productos")
    minimosConfigurados = Application.WorksheetFunction.CountIfs( _
        ws.Range("A" & FILA_PRIMER_PRODUCTO & ":A" & FILA_ULTIMO_PRODUCTO), "<>", _
        ws.Range("G" & FILA_PRIMER_PRODUCTO & ":G" & FILA_ULTIMO_PRODUCTO), CriterioSinMinimo())
    porAgotarse = Application.WorksheetFunction.CountIf( _
        ws.Range("G" & FILA_PRIMER_PRODUCTO & ":G" & FILA_ULTIMO_PRODUCTO), "Por agotarse")

    If minimosConfigurados = 0 Then
        LimpiarFiltroProductos ws
        ws.Activate
        MsgBox "No se puede calcular que productos estan por agotarse: " & _
               "el catalogo no tiene niveles minimos configurados.", _
               vbInformation, "Minimos no configurados"
        Exit Sub
    End If

    If porAgotarse = 0 Then
        LimpiarFiltroProductos ws
        ws.Activate
        MsgBox "No hay referencias clasificadas como por agotarse.", _
               vbInformation, "Productos por agotarse"
        Exit Sub
    End If

    AplicarFiltroProductos ws, 7, "Por agotarse"
    ws.Activate
    ws.Range("A4").Select
    Exit Sub

ControlError:
    MsgBox "No se pudo consultar los productos por agotarse: " & Err.Description, _
           vbExclamation, "Inventario CERDEX"
End Sub

Public Sub VerHistorialMovimientos()
    ActivarHoja "Movimientos"
End Sub

Public Sub VerListadoProductos()
    Dim ws As Worksheet

    On Error GoTo ControlError
    Set ws = ThisWorkbook.Worksheets("Productos")
    LimpiarFiltroProductos ws
    ws.Activate
    ws.Range("A4").Select
    Exit Sub

ControlError:
    MsgBox "No se pudo abrir el listado de productos: " & Err.Description, _
           vbExclamation, "Inventario CERDEX"
End Sub

' Llama a la macro de registro existente sin duplicar su logica de inventario.
' Al completar un movimiento, vuelve a Inicio. Si hay una validacion fallida,
' permanece en Registro para corregir los datos.
Public Sub GuardarMovimientoYVolver()
    Static enProceso As Boolean
    Dim wsRegistro As Worksheet
    Dim estadoAnterior As String
    Dim estadoNuevo As String

    If enProceso Then Exit Sub
    enProceso = True
    On Error GoTo ControlError

    Set wsRegistro = ThisWorkbook.Worksheets("Registro")
    estadoAnterior = CStr(wsRegistro.Range("A17").Value)
    modInventario.AgregarMovimiento
    estadoNuevo = CStr(wsRegistro.Range("A17").Value)

    If Len(estadoNuevo) > 0 And estadoNuevo <> estadoAnterior Then
        If Left$(estadoNuevo, 11) = "Registrado:" Then AbrirInicio
    End If
    GoTo Salida

ControlError:
    MsgBox "No se pudo terminar el registro: " & Err.Description, _
           vbExclamation, "Inventario CERDEX"

Salida:
    enProceso = False
End Sub

Private Sub AbrirRegistroConTipo(ByVal tipo As String)
    Dim ws As Worksheet

    On Error GoTo ControlError
    Set ws = ThisWorkbook.Worksheets("Registro")

    If FormularioTieneDatos(ws) And _
       StrComp(Trim$(CStr(ws.Range("B6").Value)), tipo, vbTextCompare) <> 0 Then
        ws.Activate
        MsgBox "Hay datos pendientes en Registro. Finaliza ese movimiento antes " & _
               "de iniciar otro tipo. No se modifico ningun dato.", _
               vbExclamation, "Movimiento pendiente"
        Exit Sub
    End If

    If Not FormularioTieneDatos(ws) Then
        ws.Range("B5").Value = Date
        ws.Range("B5").NumberFormatLocal = "dd-mmm-aaaa"
    End If

    ws.Range("B6").Value = tipo
    ws.Activate
    ws.Range("B7").Select
    Exit Sub

ControlError:
    MsgBox "No se pudo abrir Registro: " & Err.Description, _
           vbExclamation, "Inventario CERDEX"
End Sub

Private Function FormularioTieneDatos(ByVal ws As Worksheet) As Boolean
    FormularioTieneDatos = _
        Len(Trim$(CStr(ws.Range("B7").Value))) > 0 Or _
        Len(Trim$(CStr(ws.Range("B11").Value))) > 0 Or _
        Len(Trim$(CStr(ws.Range("B13").Value))) > 0 Or _
        Len(Trim$(CStr(ws.Range("B14").Value))) > 0 Or _
        Len(Trim$(CStr(ws.Range("B15").Value))) > 0
End Function

Private Sub ActivarHoja(ByVal nombreHoja As String)
    On Error GoTo ControlError
    ThisWorkbook.Worksheets(nombreHoja).Activate
    Exit Sub

ControlError:
    MsgBox "No encuentro la hoja " & nombreHoja & ".", _
           vbExclamation, "Inventario CERDEX"
End Sub

Private Sub AplicarFiltroProductos(ByVal ws As Worksheet, _
                                   ByVal campo As Long, _
                                   ByVal criterio As String)
    If ws.FilterMode Then ws.ShowAllData
    If Not ws.AutoFilterMode Then ws.Range("A4:L100").AutoFilter
    ws.Range("A4:L100").AutoFilter Field:=campo, Criteria1:=criterio
End Sub

Private Sub LimpiarFiltroProductos(ByVal ws As Worksheet)
    On Error Resume Next
    If ws.FilterMode Then ws.ShowAllData
    On Error GoTo 0
End Sub

Private Function ObtenerOCrearInicio() As Worksheet
    On Error Resume Next
    Set ObtenerOCrearInicio = ThisWorkbook.Worksheets(NOMBRE_INICIO)
    On Error GoTo 0

    If ObtenerOCrearInicio Is Nothing Then
        Set ObtenerOCrearInicio = ThisWorkbook.Worksheets.Add( _
            Before:=ThisWorkbook.Worksheets(1))
        ObtenerOCrearInicio.Name = NOMBRE_INICIO
    End If
End Function

Private Sub ConstruirPantalla(ByVal ws As Worksheet)
    Dim bloque As Range
    Dim formulaNota As String

    Do While ws.Shapes.Count > 0
        ws.Shapes(1).Delete
    Loop
    ws.Range("A1:R27").UnMerge
    ws.Range("A1:R27").Clear

    With ws
        .Tab.Color = RGB(25, 135, 84)
        .Range("A1:Z27").Interior.Color = RGB(20, 34, 56)
        .Range("A1:R27").Font.Name = "Segoe UI"
        .Range("A1:R27").Font.Color = RGB(255, 255, 255)
        .Columns("A").ColumnWidth = 2
        .Columns("B:Q").ColumnWidth = 17.5
        .Columns("R").ColumnWidth = 2
        .Rows("1:27").RowHeight = 18
        .Rows("1:1").RowHeight = 8
        .Rows("2:2").RowHeight = 27
        .Rows("3:3").RowHeight = 18
        .Rows("4:4").RowHeight = 18
        .Rows("5:5").RowHeight = 8
        .Rows("6:6").RowHeight = 20
        .Rows("7:9").RowHeight = 17
        .Rows("10:10").RowHeight = 8
        .Rows("11:11").RowHeight = 20
        .Rows("12:14").RowHeight = 17
        .Rows("15:17").RowHeight = 17
        .Rows("18:18").RowHeight = 8
        .Rows("19:19").RowHeight = 20
        .Rows("20:22").RowHeight = 17
        .Rows("23:24").RowHeight = 14
        .Rows("25:25").RowHeight = 6
        .Rows("26:26").RowHeight = 14
    End With

    FusionarTexto ws, "B2:Q2", "INVENTARIO CERDEX", RGB(20, 34, 56), _
                  RGB(255, 255, 255), 22, True, xlLeft
    FusionarTexto ws, "B3:Q3", "Sistema de control de materiales e inventario", _
                  RGB(20, 34, 56), RGB(213, 224, 236), 11, False, xlLeft

    ws.Range("B4").Value = "Fecha"
    ws.Range("B4").Font.Bold = True
    ws.Range("B4").Font.Color = RGB(213, 224, 236)
    Set bloque = ws.Range("C4:E4")
    bloque.Merge
    bloque.Cells(1, 1).Formula = "=TODAY()"
    bloque.NumberFormat = "dd-mmm-yyyy"
    bloque.Font.Color = RGB(255, 255, 255)
    bloque.Font.Bold = True

    ws.Range("J4").Value = "Responsable"
    ws.Range("J4").Font.Bold = True
    ws.Range("J4").Font.Color = RGB(213, 224, 236)
    Set bloque = ws.Range("K4:Q4")
    bloque.Merge
    bloque.Cells(1, 1).Formula = "=IF(Registro!$B$12="""",""Sin asignar"",Registro!$B$12)"
    bloque.Font.Color = RGB(255, 255, 255)
    bloque.Font.Bold = True

    FusionarTexto ws, "B6:Q6", "OPERACIONES PRINCIPALES", RGB(36, 69, 99), _
                  RGB(255, 255, 255), 11, True, xlLeft
    FusionarTexto ws, "B11:Q11", "CONSULTAS", RGB(36, 69, 99), _
                  RGB(255, 255, 255), 11, True, xlLeft
    FusionarTexto ws, "B19:Q19", "INDICADORES DEL INVENTARIO", RGB(36, 69, 99), _
                  RGB(255, 255, 255), 11, True, xlLeft

    AgregarBotonMenu ws, "btnEntrada", "REGISTRAR ENTRADA", "B7:G9", _
                     "AbrirRegistroEntrada", RGB(25, 135, 84), RGB(255, 255, 255)
    AgregarBotonMenu ws, "btnSalida", "REGISTRAR SALIDA", "H7:L9", _
                     "AbrirRegistroSalida", RGB(220, 53, 69), RGB(255, 255, 255)
    AgregarBotonMenu ws, "btnDevolucion", _
                     "REGISTRAR DEVOLUCION", "M7:Q9", _
                     "AbrirRegistroDevolucion", RGB(255, 193, 7), RGB(20, 34, 56)

    AgregarBotonMenu ws, "btnInventario", "VER INVENTARIO", "B12:F14", _
                     "VerInventario", RGB(36, 69, 99), RGB(255, 255, 255)
    AgregarBotonMenu ws, "btnAgotados", "PRODUCTOS AGOTADOS", "G12:K14", _
                     "VerProductosAgotados", RGB(36, 69, 99), RGB(255, 255, 255)
    AgregarBotonMenu ws, "btnPorAgotarse", "PRODUCTOS POR AGOTARSE", "L12:Q14", _
                     "VerProductosPorAgotarse", RGB(36, 69, 99), RGB(255, 255, 255)
    AgregarBotonMenu ws, "btnHistorial", "HISTORIAL DE MOVIMIENTOS", "B15:I17", _
                     "VerHistorialMovimientos", RGB(36, 69, 99), RGB(255, 255, 255)
    AgregarBotonMenu ws, "btnListado", "LISTADO DE PRODUCTOS", "J15:Q17", _
                     "VerListadoProductos", RGB(36, 69, 99), RGB(255, 255, 255)

    CrearIndicador ws, "B20:E22", "REFERENCIAS REGISTRADAS", _
                   "=COUNTA(Productos!$A$5:$A$100)"
    CrearIndicador ws, "F20:I22", "REFERENCIAS CON EXISTENCIA", _
                   "=COUNTIFS(Productos!$A$5:$A$100,""<>"",Productos!$F$5:$F$100,"">0"")"
    CrearIndicador ws, "J20:M22", "REFERENCIAS AGOTADAS", _
                   "=COUNTIFS(Productos!$A$5:$A$100,""<>"",Productos!$F$5:$F$100,""<=0"")"
    CrearIndicador ws, "N20:Q22", "REFERENCIAS POR AGOTARSE", FormulaPorAgotarse()

    formulaNota = FormulaNotaEstado()
    FusionarTexto ws, "B23:Q24", formulaNota, _
        RGB(36, 69, 99), RGB(255, 255, 255), 10, False, xlLeft, True

    FusionarTexto ws, "B26:Q26", _
        "Las existencias se consultan desde el Kardex. Esta pantalla no permite editar saldos.", _
        RGB(20, 34, 56), RGB(183, 201, 220), 9, False, xlLeft

    ws.Activate
    With ActiveWindow
        .DisplayGridlines = False
        .Zoom = 60
    End With
End Sub

Private Sub CrearIndicador(ByVal ws As Worksheet, _
                           ByVal direccion As String, _
                           ByVal etiqueta As String, _
                           ByVal formulaValor As String)
    Dim tarjeta As Range
    Dim etiquetaRango As Range
    Dim valorRango As Range

    Set tarjeta = ws.Range(direccion)
    Set etiquetaRango = ws.Range(tarjeta.Cells(1, 1), tarjeta.Cells(1, tarjeta.Columns.Count))
    Set valorRango = ws.Range(tarjeta.Cells(2, 1), tarjeta.Cells(tarjeta.Rows.Count, tarjeta.Columns.Count))
    etiquetaRango.Merge
    valorRango.Merge

    tarjeta.Interior.Color = RGB(36, 69, 99)
    tarjeta.Borders.LineStyle = xlContinuous
    tarjeta.Borders.Color = RGB(80, 105, 132)

    With etiquetaRango
        .Value = etiqueta
        .Font.Name = "Segoe UI"
        .Font.Size = 9
        .Font.Bold = True
        .Font.Color = RGB(213, 224, 236)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With

    With valorRango
        .Cells(1, 1).Formula = formulaValor
        .Font.Name = "Segoe UI"
        .Font.Size = 20
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .NumberFormat = "#,##0"
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
End Sub

Private Sub FusionarTexto(ByVal ws As Worksheet, _
                          ByVal direccion As String, _
                          ByVal texto As String, _
                          ByVal colorFondo As Long, _
                          ByVal colorFuente As Long, _
                          ByVal tamanoFuente As Long, _
                          ByVal negrita As Boolean, _
                          ByVal alineacion As XlHAlign, _
                          Optional ByVal formula As Boolean = False)
    With ws.Range(direccion)
        .Merge
        If formula Then
            .Cells(1, 1).Formula = texto
        Else
            .Cells(1, 1).Value = texto
        End If
        .Interior.Color = colorFondo
        .Font.Name = "Segoe UI"
        .Font.Size = tamanoFuente
        .Font.Bold = negrita
        .Font.Color = colorFuente
        .HorizontalAlignment = alineacion
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With
End Sub

Private Sub AgregarBotonMenu(ByVal ws As Worksheet, _
                             ByVal nombre As String, _
                             ByVal etiqueta As String, _
                             ByVal rangoAncla As String, _
                             ByVal accion As String, _
                             ByVal colorFondo As Long, _
                             ByVal colorFuente As Long)
    Dim ancla As Range
    Dim boton As Shape

    Set ancla = ws.Range(rangoAncla)
    Set boton = ws.Shapes.AddShape(5, ancla.Left + 3, ancla.Top + 3, _
                                   ancla.Width - 6, ancla.Height - 6)
    With boton
        .Name = nombre
        .Fill.ForeColor.RGB = colorFondo
        .Line.Visible = 0
        .Placement = xlMoveAndSize
        .OnAction = AccionDeLibro(accion)
        With .TextFrame
            .Characters.Text = etiqueta
            .HorizontalAlignment = xlHAlignCenter
            .VerticalAlignment = xlVAlignCenter
            .MarginLeft = 8
            .MarginRight = 8
            .MarginTop = 4
            .MarginBottom = 4
            With .Characters.Font
                .Name = "Segoe UI"
                .Size = 14
                .Bold = True
                .Color = colorFuente
            End With
        End With
    End With
End Sub

Private Sub AgregarBotonInicio(ByVal ws As Worksheet, ByVal celda As String)
    Dim ancla As Range
    Dim boton As Shape
    Dim ancho As Single
    Dim alto As Single
    Dim desplazamientoSuperior As Single

    On Error Resume Next
    ws.Shapes("btnInicioKardex").Delete
    On Error GoTo 0

    Set ancla = ws.Range(celda)
    ancho = 86
    alto = 24
    desplazamientoSuperior = 2

    ' En hojas con tabla desde la fila 4, deja el acceso en la franja vacía de la fila 3.
    If ws.Name = "Movimientos" Or ws.Name = "Productos" Then
        ancho = 78
        alto = 17
        desplazamientoSuperior = 1
    End If

    Set boton = ws.Shapes.AddShape(5, ancla.Left + 2, ancla.Top + desplazamientoSuperior, ancho, alto)
    With boton
        .Name = "btnInicioKardex"
        .Fill.ForeColor.RGB = RGB(36, 69, 99)
        .Line.Visible = 0
        .Placement = xlMoveAndSize
        .OnAction = AccionDeLibro("AbrirInicio")
        With .TextFrame
            .Characters.Text = "INICIO"
            .HorizontalAlignment = xlHAlignCenter
            .VerticalAlignment = xlVAlignCenter
            With .Characters.Font
                .Name = "Segoe UI"
                .Size = 9
                .Bold = True
                .Color = RGB(255, 255, 255)
            End With
        End With
    End With
End Sub

Private Sub AsignarAccionGuardar()
    Dim ws As Worksheet
    Dim boton As Shape
    Dim texto As String

    Set ws = ThisWorkbook.Worksheets("Registro")
    For Each boton In ws.Shapes
        texto = ""
        On Error Resume Next
        texto = boton.TextFrame.Characters.Text
        On Error GoTo 0

        If InStr(1, texto, "AGREGAR MOVIMIENTO", vbTextCompare) > 0 Then
            boton.OnAction = AccionDeLibro("GuardarMovimientoYVolver")
            Exit Sub
        End If
    Next boton
End Sub

Private Function AccionDeLibro(ByVal nombreMacro As String) As String
    AccionDeLibro = "'" & ThisWorkbook.Name & "'!" & nombreMacro
End Function

Private Function CriterioSinMinimo() As String
    CriterioSinMinimo = "<>Sin m" & ChrW$(237) & "nimo"
End Function

Private Function FormulaPorAgotarse() As String
    FormulaPorAgotarse = _
        "=IF(COUNTIFS(Productos!$A$5:$A$100," & ComillasExcel("<>") & _
        ",Productos!$G$5:$G$100," & ComillasExcel(CriterioSinMinimo()) & _
        ")=0," & ComillasExcel("N/D") & _
        ",COUNTIFS(Productos!$A$5:$A$100," & ComillasExcel("<>") & _
        ",Productos!$G$5:$G$100," & ComillasExcel("Por agotarse") & "))"
End Function

Private Function FormulaNotaEstado() As String
    Dim textoNoDisponible As String
    Dim textoDisponible As String

    textoNoDisponible = "Por agotarse no disponible: el catalogo no tiene niveles m" & _
                        ChrW$(237) & "nimos configurados."
    textoDisponible = "Las alertas por agotarse se calculan con el estado del catalogo."
    FormulaNotaEstado = "=IF(COUNTIFS(Productos!$A$5:$A$100," & _
        ComillasExcel("<>") & ",Productos!$G$5:$G$100," & _
        ComillasExcel(CriterioSinMinimo()) & ")=0," & _
        ComillasExcel(textoNoDisponible) & "," & _
        ComillasExcel(textoDisponible) & ")"
End Function

Private Function ComillasExcel(ByVal texto As String) As String
    ComillasExcel = Chr$(34) & texto & Chr$(34)
End Function

Private Function TextoDevolucionCliente() As String
    TextoDevolucionCliente = "Devoluci" & ChrW$(243) & "n cliente"
End Function

Private Function TextoDevolucionProveedor() As String
    TextoDevolucionProveedor = "Devoluci" & ChrW$(243) & "n proveedor"
End Function
