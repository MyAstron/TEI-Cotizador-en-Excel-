VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Cotizacion 
   Caption         =   "..:: Cotizacion ::.. (Visual)"
   ClientHeight    =   9264.001
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   14304
   OleObjectBlob   =   "Cotizacion.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "Cotizacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Variables de Entorno
Dim ws As Worksheet
Dim wsa As Worksheet
Dim tbl As ListObject
Dim col As ListColumn
Dim cell As Range
Dim celda As Range
Dim precioIVA As Double
Dim precioISR As Double
Dim totalPrecio As Double
Dim valorPrecio As Double
Dim cotizacionTot As Double
Dim letrasCombo As MSForms.ComboBox

Private Sub btnClear_Click()
    lbDescipcion.Clear
    lbPrecios.Clear
    lbTotal.Clear
    lbHistorial.Clear
    btnEliminar.Enabled = False
    btnGenerar.Enabled = False
    cbSelect.Value = "- Seleccion -"
    txtCantidad.Text = "Cant."
End Sub

Private Sub btnEliminar_Click()
    Dim cotizacionSelect As Integer
    Dim itemCotizacion As String
    Dim itemPrecio As String
    cotizacionSelect = lbHistorial.ListIndex
    itemCotizacion = lbHistorial.Value
    itemPrecio = Mid(itemCotizacion, InStr(1, itemCotizacion, "($") + 2, Len(itemCotizacion) - InStr(1, itemCotizacion, "($") - 2)
    cotizacionTot = cotizacionTot - Val(itemPrecio)
    totalPrecio = 0
    totCotizacion
    lbHistorial.RemoveItem cotizacionSelect
    btnEliminar.Enabled = False
End Sub

Private Sub btnGenerar_Click()
    CotizacionPDF
    Unload InicioSesion
    Unload Cotizacion
End Sub

Private Sub btnGuardar_Click() ' Guardar datos en el historial local
    If cbSelect.Value = "- Seleccion -" Then
        MsgBox "Argumentos no validos", vbCritical, "Consola"
    ElseIf Not Val(txtCantidad.Text) > 0 Then
        MsgBox "Argumentos no validos", vbCritical, "Consola"
    ElseIf totalPrecio = 0 Then
        MsgBox "Argumentos no validos", vbCritical, "Consola"
    Else
        Prevista
        lbHistorial.AddItem "(x" & txtCantidad.Text & ") " & cbSelect.Text & " ($" & totalPrecio & ")"
        btnGenerar.Enabled = True
        totCotizacion
    End If
End Sub

Private Sub btnPreView_Click() ' prevista
    Prevista
End Sub

Private Sub chbCombo_Click()
    If chbProductos.Value Then
        chbProductos.Value = False
    Else
        chbCombo.Value = True
    End If
    cbSelect.Clear
    cbSelect.Value = "- Seleccion -"
    DetectarLetrasIniciales
End Sub

Private Sub chbProductos_Click()
    If chbCombo.Value Then
        chbCombo.Value = False
    Else
        chbProductos.Value = True
    End If
    cbSelect.Clear
    cbSelect.Value = "- Seleccion -"
    DetectarCodigos
End Sub

Private Sub CommandButton1_Click() ' boton inicial
        txtNombre.Text = InicioSesion.MiVariable(2)
        txtDireccion.Text = InicioSesion.MiVariable(3) & ", " & InicioSesion.MiVariable(4)
        txtMail.Text = InicioSesion.MiVariable(6)
        txtTel.Text = InicioSesion.MiVariable(5)
        txtID.Text = InicioSesion.MiVariable(1)
        lblAsesor.Caption = InicioSesion.MisDatos(1)
    DetectarLetrasIniciales
    Frame1.Visible = True
    Frame2.Visible = True
    CommandButton1.Visible = False
    btnClear.Visible = True
End Sub

Private Sub cbSelect_Change() ' al seleccionar un combo
    If txtCantidad.Text = "Cant." Then
        txtCantidad.Text = "1"
    ElseIf Not IsNumeric(txtCantidad.Text) Then
        txtCantidad.Text = "1"
    End If
    btnGuardar.Enabled = True
    btnPreView.Enabled = True
    Prevista
End Sub

Private Sub lbDescipcion_Click()
    Dim indexDesc As Integer
    indexDesc = lbDescipcion.ListIndex
    lbPrecios.Selected(indexDesc) = True
End Sub

Private Sub lbHistorial_Change() ' Al seleccionar un dato del historial
    If InStr(1, lbHistorial.Value, "Combo", vbTextCompare) > 0 Then
        chbCombo.Value = True
        chbProductos.Value = False
        DesglosarCadena
    Else
        chbCombo.Value = False
        chbProductos.Value = True
        DesglosarCadena
    End If
    btnEliminar.Enabled = True
End Sub

Private Sub lbPrecios_Click()
    Dim indexDesc As Integer
    indexDesc = lbPrecios.ListIndex
    lbDescipcion.Selected(indexDesc) = True
End Sub

Private Sub txtCantidad_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
' al editar el cuadro de texto 'cantidad'
    If txtCantidad.Text = "Cant." Then
        txtCantidad.Text = "1"
    End If
End Sub

Sub Prevista() ' funcion prevista
    lbDescipcion.Clear
    lbPrecios.Clear
    totalPrecio = 0
    valorPrecio = 0
    
    For Each row In tbl.ListRows
        For Each cell In row.Range.Cells
            Dim rowPos As Long
            Dim colPos As Long
            rowPos = cell.row - tbl.HeaderRowRange.row + 1
            colPos = cell.Column - tbl.HeaderRowRange.Column + 1
            If chbCombo.Value Then
                If colPos = 1 Then
                    If Left(cell.Value, 2) = Right(cbSelect.Text, 2) Then ' en el combo seleccionado buscar
                        lbDescipcion.AddItem ws.Cells(cell.row, col.Index + 1).Value ' Descripcion
                        lbPrecios.AddItem "$" & Val(ws.Cells(cell.row, col.Index + 2).Value) * Val(txtCantidad.Text) ' Precio
                        
                        valorPrecio = Val(ws.Cells(cell.row, col.Index + 2).Value) * Val(txtCantidad.Text) ' acumuladores (precio)
                        totalPrecio = totalPrecio + valorPrecio
                    End If
                End If
            Else
                If colPos = 2 Then
                    If cell.Value = cbSelect.Value Then ' en el combo seleccionado buscar
                        lbDescipcion.AddItem cell.Value ' Descripcion
                        lbPrecios.AddItem "$" & Val(ws.Cells(cell.row, col.Index + 1).Value) * Val(txtCantidad.Text) ' Precio
                        totalPrecio = Val(ws.Cells(cell.row, col.Index + 1).Value) * Val(txtCantidad.Text)
                    End If
                End If
            End If
        Next cell
    Next row
    ' ------------------- Decoracion -------------------
    lbDescipcion.AddItem "----"
    lbDescipcion.AddItem "IVA   - (12%)"
    lbDescipcion.AddItem "ISR   - (5%)"
    lbDescipcion.AddItem "Total"
        
    precioIVA = totalPrecio * 0.12
    precioISR = totalPrecio * 0.05
    totalPrecio = totalPrecio + precioIVA + precioISR
    lbPrecios.AddItem "----"
    lbPrecios.AddItem "$" & precioIVA
    lbPrecios.AddItem "$" & precioISR
    lbPrecios.AddItem "$" & totalPrecio
End Sub

Sub DetectarLetrasIniciales() ' Funcion para acomodar los combos segun el codigo de serie
    Set ws = ThisWorkbook.Sheets("Productos")
    Set tbl = ws.ListObjects("Catalogo")
    Set col = tbl.ListColumns("CODIGO")
    Set letrasCombo = cbSelect
    letrasCombo.Clear
    
    ' Inicializar una colección para almacenar temporalmente las letras iniciales
    Dim letras As New Collection
    For Each cell In col.DataBodyRange.Cells
        ' Obtener la letra inicial de cada celda
        Dim letra As String
        letra = Left(cell.Value, 2)
        
        ' Verificar si la letra inicial no está en la colección y agregarla si no está
        If Not Contiene(letras, letra) Then
            letras.Add letra
            letrasCombo.AddItem "Combo " & letra
        End If
    Next cell
End Sub

Sub DetectarCodigos() ' Funcion para acomodar los productos
    Set ws = ThisWorkbook.Sheets("Productos")
    Set tbl = ws.ListObjects("Catalogo")
    Set col = tbl.ListColumns("ARTICULO")
    Set letrasCombo = cbSelect
    letrasCombo.Clear
    
    ' Inicializar una colección para almacenar temporalmente las letras iniciales
    For Each cell In col.DataBodyRange.Cells
        letrasCombo.AddItem cell.Value
    Next cell
End Sub

Sub DesglosarCadena() ' re acomodar datos del historial a la previsualizacion
    Dim cadena As String
    Dim cantidad As String
    Dim nombreCombo As String
    
    If Not IsNull(lbHistorial.Value) Then
        cadena = lbHistorial.Value
        cantidad = Mid(cadena, 3, InStr(1, cadena, ")") - 3)
        nombreCombo = Mid(cadena, InStr(1, cadena, ")") + 2, InStr(1, cadena, "($") - InStr(1, cadena, ")") - 3)
        
        txtCantidad.Text = cantidad
        cbSelect.Value = nombreCombo
        Prevista
    End If
End Sub

Sub totCotizacion() ' Cambios sobre el total de la cotizacion (registro local)
    cotizacionTot = cotizacionTot + totalPrecio
    lbTotal.Clear
    lbTotal.AddItem "Cotizacion Total: $" & Format(cotizacionTot, "0.00")
End Sub

Function Contiene(col As Collection, valor As Variant) As Boolean ' colector para las letras iniciales de los numeros en serie
    Dim elem As Variant
    For Each elem In col
        If elem = valor Then
            Contiene = True
            Exit Function
        End If
    Next elem
    Contiene = False
End Function

Sub CotizacionPDF()
    Dim ruta As String
    Dim i As Integer
    Dim NumeroCotizacion As String
    Dim rng As Range
    Dim cell As Range
    Dim nombreHoja As String
    
    Set ws = ThisWorkbook.Sheets.Add
    ws.Visible = xlSheetHidden
    nombreHoja = ws.Name
    
    For i = 1 To Len(nombreHoja)
        If IsNumeric(Mid(nombreHoja, i, 1)) Then
            NumeroCotizacion = NumeroCotizacion & Mid(nombreHoja, i, 1)
        End If
    Next i
    ruta = ThisWorkbook.Path & "\Cotizaciones\" & "Cotizacion-" & NumeroCotizacion & "-" & Day(Date) & "." & Month(Date) & "." & Year(Date) & ".pdf"
    MsgBox "Al terminar la evaluacion se abrira automatico. " & ruta, vbInformation, "Consola"
    
    Set rng = ws.Range("D1:E1")
    rng.Merge
    With rng
        .Value = "COTIZACIÓN"
        .HorizontalAlignment = xlHAlignRight
        .VerticalAlignment = xlVAlignBottom
        .Font.Color = RGB(31, 70, 120)
        .Font.Size = 20
        .RowHeight = rng.RowHeight * 2
        .Columns.AutoFit
    End With
    
On Error GoTo ManejarError
    Dim IMG As Shape
    With ws.Range("A1")
        Set IMG = ws.Shapes.AddPicture(Filename:="https://encrypted-tbn0.gstatic.com/images?q=tbn:ANd9GcT41OXjDy-UdA42TTCttqmLlbkQ65wlVN10MGPAjJJujQ&s", _
                                LinkToFile:=False, _
                                SaveWithDocument:=True, _
                                Left:=ws.Range("A1").Left + ws.Range("A1").Width - 1, _
                                Top:=ws.Range("A1").Top, _
                                Width:=-1, _
                                Height:=-1)
        IMG.LockAspectRatio = msoTrue
        IMG.Height = .Height
    End With
ManejarError:
    Err.Clear
    
    Set rng = ws.Range("A9:B9")
    rng.Merge
    With rng
        .Value = "CLIENTE"
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(31, 70, 120)
        .Font.Size = 14
    End With
    
    Set rng = ws.Range("E3:E6")
    For Each cell In rng
        Select Case cell.Address
            Case "$E$3":
                cell.Value = Date
            Case "$E$4":
                cell.Value = NumeroCotizacion
            Case "$E$5":
                cell.Value = InicioSesion.MiVariable(1)
            Case "$E$6":
                cell.Value = DateAdd("m", 3, Date)
            Case Else:
                cell.Value = ""
        End Select
        With cell
            .HorizontalAlignment = xlHAlignCenter
            .Borders.LineStyle = xlContinuous
            .Borders.Color = RGB(0, 0, 0)
            .Borders.Weight = xlThin
            .Columns.AutoFit
        End With
    Next cell
    
    Set rng = ws.Range("A16:H16")
    For Each cell In rng
        Select Case cell.Address
            Case "$A$16":
                cell.Value = "CODIGO"
            Case "$B$16":
                cell.Value = "DESCRIPCIÓN"
            Case "$C$16":
                cell.Value = "CANT."
            Case "$D$16":
                cell.Value = "PRECIO"
            Case "$E$16":
                cell.Value = "SUB-TOTAL"
            Case "$F$16":
                cell.Value = "ISR"
            Case "$G$16":
                cell.Value = "IVA"
            Case "$H$16":
                cell.Value = "TOTAL"
            Case Else:
                cell.Value = ""
        End Select
        With cell
            .HorizontalAlignment = xlHAlignCenter
            .Font.Color = RGB(255, 255, 255)
            .Interior.Color = RGB(31, 70, 120)
            .Font.Size = 14
            .Columns.AutoFit
        End With
    Next cell
    
    Set rng = ws.Range("A2:A7")
    For Each cell In rng
        Select Case cell.Address
            Case "$A$2":
                cell.Value = InicioSesion.MisDatos(2)
            Case "$A$3":
                cell.Value = "Ciudad: " & InicioSesion.MisDatos(3)
            Case "$A$4":
                cell.Value = "Sitio Web: https://portal.ingenieria.usac.edu.gt/"
            Case "$A$5":
                cell.Value = "Teléfono: " & InicioSesion.MisDatos(4)
            Case "$A$6":
                cell.Value = "E-mail: " & InicioSesion.MisDatos(5)
            Case "$A$7":
                cell.Value = "Asesor de venta: " & InicioSesion.MisDatos(1)
                cell.Columns.AutoFit
            Case Else:
                cell.Value = ""
        End Select
        With cell
            .Font.Size = 14
        End With
    Next cell
    
    Set rng = ws.Range("A10:A14")
    For Each cell In rng
        Select Case cell.Address
            Case "$A$10":
                cell.Value = "Nombre: " & InicioSesion.MiVariable(2)
            Case "$A$11":
                cell.Value = InicioSesion.MiVariable(3)
            Case "$A$12":
                cell.Value = "Ciudad: " & InicioSesion.MiVariable(4)
            Case "$A$13":
                cell.Value = "Teléfono: " & InicioSesion.MiVariable(5)
            Case "$A$14":
                cell.Value = "E-mail: " & InicioSesion.MiVariable(6)
            Case Else:
                cell.Value = ""
        End Select
        With cell
            .Font.Size = 14
        End With
    Next cell
    
    Set rng = ws.Range("G4:G5")
    For Each cell In rng
        ws.Range("G4").Value = "12%"
        ws.Range("G5").Value = "5%"
        With cell
            .HorizontalAlignment = xlHAlignCenter
            .Font.Size = 14
            .Borders(xlEdgeBottom).Weight = xlThick
        End With
    Next cell
    
    With ws.Range("G3")
        .Font.Size = 14
        .Value = "IMPUESTOS"
        .HorizontalAlignment = xlHAlignCenter
        .Font.Bold = True
        .Columns.AutoFit
    End With
    
    Set rng = ws.Range("D3:D6")
    For Each cell In rng
        Select Case cell.Address
            Case "$D$3":
                cell.Value = "FECHA"
            Case "$D$4":
                cell.Value = "COTIZACIÓN #"
            Case "$D$5":
                cell.Value = "ID CLIENTE"
            Case "$D$6":
                cell.Value = "VALIDO HASTA"
            Case Else:
                cell.Value = ""
        End Select
        With cell
            .HorizontalAlignment = xlHAlignRight
            .Columns.AutoFit
            .Font.Size = 12
        End With
    Next cell
    
    Dim Rango As String
    Dim cadena As String
    Dim cantidad As String
    Dim nombreCombo As String
    Dim precioString As String
    Dim CODIGO As String
    Dim DESCRIPCION As String
    Dim precioNeto As Double
    Dim TotalBruto As Double
    Dim TotalImpuesto As Double
    Dim TotalNeto As Double
    
    For i = 0 To lbHistorial.ListCount - 1
        If Not IsNull(lbHistorial.List(i)) Then
            cadena = lbHistorial.List(i)
            cantidad = Mid(cadena, 3, InStr(1, cadena, ")") - 3)
            nombreCombo = Mid(cadena, InStr(1, cadena, ")") + 2, InStr(1, cadena, "($") - InStr(1, cadena, ")") - 3)
            precioString = Mid(cadena, InStr(InStr(InStr(InStr(InStr(cadena, "("), cadena, " "), cadena, " "), cadena, "("), cadena, "$") + 1)
        End If
        precioNeto = (Val(precioString) / (1 + 0.05 + 0.12)) / Val(cantidad)
        Rango = "$A$" & 17 + i & ":$H$" & 16 + Val(lbHistorial.ListCount)
        Set rng = ws.Range(Rango)
        CODIGO = ""
            DESCRIPCION = ""
        For Each cell In rng
            If cell.row Mod 2 = 0 Then
                cell.EntireRow.Interior.Color = RGB(220, 220, 220)
            Else
                cell.EntireRow.Interior.colorIndex = xlNone
            End If
            If InStr(1, nombreCombo, "Combo", vbTextCompare) > 0 Then
                Set wsa = ThisWorkbook.Sheets("Productos")
                Set tbl = wsa.ListObjects("Catalogo")
                Set col = tbl.ListColumns("CODIGO")
                
                Dim Articulos As New Collection
                For Each celda In col.DataBodyRange.Cells
                    Dim articulo As String
                    articulo = wsa.Cells(celda.row, celda.Column + 1).Value
                    If Left(celda.Value, 1) = Right(nombreCombo, 1) Then
                        If Not Contiene(Articulos, articulo) Then
                            Articulos.Add articulo
                            If Articulos.Count > 1 Then
                                DESCRIPCION = DESCRIPCION & " (-) " & articulo
                            Else
                                DESCRIPCION = DESCRIPCION & articulo
                            End If
                        End If
                    End If
                Next celda
                Set Articulos = New Collection
                CODIGO = nombreCombo
            Else
                DESCRIPCION = nombreCombo
                CODIGO = "Producto"
            End If
            If Len(DESCRIPCION) < 31 Then
                DESCRIPCION = DESCRIPCION
            Else
                DESCRIPCION = Left(DESCRIPCION, 25) & " [...]"
            End If
            Select Case cell.Column
                Case 1:
                    cell.Value = CODIGO
                    cell.HorizontalAlignment = xlHAlignCenter
                Case 2:
                    cell.Value = DESCRIPCION
                Case 3:
                    cell.Value = cantidad
                Case 4:
                    cell.Value = precioNeto
                    cell.NumberFormat = "#,##0.00 Q"
                Case 5:
                    cell.Value = precioNeto * Val(cantidad)
                    cell.NumberFormat = "#,##0.00 Q"
                    TotalBruto = TotalBruto + Val(precioString)
                Case 6:
                    cell.Value = precioNeto * 0.05
                    cell.NumberFormat = "#,##0.00 Q"
                    TotalImpuesto = TotalImpuesto + precioNeto * 0.05
                Case 7:
                    cell.Value = precioNeto * 0.12
                    cell.NumberFormat = "#,##0.00 Q"
                    TotalImpuesto = TotalImpuesto + precioNeto * 0.12
                Case 8:
                    cell.Value = Val(precioString)
                    cell.NumberFormat = "#,##0.00 Q"
                    TotalNeto = TotalNeto + Val(precioString)
                Case Else:
                    cell.Value = "N/A"
            End Select
            cell.Font.Size = 14
        Next cell
    Next i
    Dim LastRow As Long
    LastRow = rng.Rows(rng.Rows.Count).row + 1
    
    ' Agrega una fila nueva al final de la tabla
    With ws.Range("F" & LastRow)
        .Value = "SubTotal"
        .Font.Bold = True
        .Font.Size = 14
    End With
    With ws.Range("F" & LastRow + 1)
        .Value = "Impuesto Total"
        .Font.Bold = True
        .Font.Size = 14
    End With
    With ws.Range("F" & LastRow + 2)
        .Value = "Total"
        .Font.Bold = True
        .Font.Size = 14
        .Borders(xlEdgeTop).Weight = xlThin
        .Borders(xlEdgeBottom).Weight = xlThick
    End With
    
    With ws.Range("H" & LastRow)
        .Value = TotalBruto - TotalImpuesto
        .Font.Bold = True
        .NumberFormat = "#,##0.00 Q"
        .Font.Size = 14
    End With
    With ws.Range("H" & LastRow + 1)
        .Value = TotalImpuesto
        .Font.Bold = True
        .NumberFormat = "#,##0.00 Q"
        .Font.Size = 14
    End With
    With ws.Range("H" & LastRow + 2)
        .Value = TotalNeto
        .Font.Bold = True
        .NumberFormat = "#,##0.00 Q"
        .Font.Size = 14
        .Borders(xlEdgeTop).Weight = xlThin
        .Borders(xlEdgeBottom).Weight = xlThick
    End With
    With ws.Range("G" & LastRow + 2)
        .Borders(xlEdgeTop).Weight = xlThin
        .Borders(xlEdgeBottom).Weight = xlThick
    End With

    ws.Range("A1").Interior.Color = RGB(31, 70, 120)
    With ws.Range("B1")
        .Value = "COTIZACIONES USAC"
        .VerticalAlignment = xlVAlignCenter
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(31, 70, 120)
        .Font.Bold = True
        .Font.Size = 20
    End With
    ws.Range("A15").Columns.AutoFit

    Dim cola As Long
    
    LastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    
    ' Itera sobre cada columna y ajusta el ancho automáticamente
    For cola = 1 To lastCol
        ws.Columns(cola).AutoFit
    Next cola
    ' Agregar texto debajo de la línea de firma
    Dim ultimaFila As Long
    ultimaFila = ws.Cells(ws.Rows.Count, "A").End(xlUp).row
    
    ' Calcula la posición vertical de la firma debajo de la última fila de la tabla
    Dim firmaTop As Double
    firmaTop = ws.Cells(ultimaFila, 1).Top + ws.Cells(ultimaFila, 1).Height + 150
    
    Dim firmaLine As Shape
    Set firmaLine = ws.Shapes.AddLine(2, firmaTop, 300, firmaTop)
    With firmaLine.Line
        .ForeColor.RGB = RGB(31, 78, 120)
        .Weight = 1.5
    End With

    ' Ajustar posición de la línea de firma
    With ws.Range("A" & LastRow + 1)
        .Value = "TERMINOS Y CONDICIONES"
        .Font.Bold = True
        .Font.Size = 14
    End With
    With ws.Range("A" & LastRow + 3)
        .Value = "1. El pago será debitado antes de la entrega de bienes y servicios"
        .Font.Bold = True
        .Font.Size = 14
    End With
    With ws.Range("A" & LastRow + 4)
        .Value = "2. Enviar la cotización firmada al email indicado anteriormente"
        .Font.Bold = True
        .Font.Size = 14
    End With
    With ws.Range("A" & LastRow + 5)
        .Value = "La aceptación del cliente (firmar a continuación):"
        .Font.Bold = True
        .Font.Size = 14
    End With
    
    With ws.Range("A" & LastRow + 1 & ":C" & LastRow + 1)
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(31, 78, 120)
    End With
    With ws.Range("C" & LastRow + 1 & ":C" & LastRow + 5)
        .Borders(xlEdgeRight).Weight = xlThin
        .Borders(xlEdgeRight).Color = RGB(31, 78, 120)
    End With
    With ws.Range("A" & LastRow + 1 & ":A" & LastRow + 5)
        .Borders(xlEdgeLeft).Weight = xlThin
        .Borders(xlEdgeLeft).Color = RGB(31, 78, 120)
    End With
    With ws.Range("A" & LastRow + 5 & ":C" & LastRow + 5)
        .Borders(xlEdgeBottom).Weight = xlThin
        .Borders(xlEdgeBottom).Color = RGB(31, 78, 120)
    End With
    
    firmaLine.Left = ws.Range("A18").Left ' Alinea la línea con la celda B18
    
    Dim firmaText As Shape
    Set firmaText = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, firmaLine.Left, firmaLine.Top + firmaLine.Height + 5, 300, 20)
    With firmaText
        .Line.Visible = msoFalse ' Oculta el borde de la caja de texto
        .TextFrame.Characters.Text = "Firma."
        .TextFrame.Characters.Font.Bold = True
        .TextFrame.Characters.Font.Size = 16
        .TextFrame.HorizontalAlignment = xlHAlignCenter
    End With
    
    ' Centrar la línea de firma y el texto debajo de la línea
    firmaLine.Left = ws.Cells(18, 2).Left + (ws.Cells(18, 2).Width - firmaLine.Width) / 2
    firmaText.Left = firmaLine.Left + (firmaLine.Width - firmaText.Width) / 2

    With ws.PageSetup
        .Orientation = xlLandscape
        .PaperSize = xlPaperLegal
        .LeftMargin = Application.InchesToPoints(0.5)
        .RightMargin = Application.InchesToPoints(0)
        .TopMargin = Application.InchesToPoints(0.25)
        .BottomMargin = Application.InchesToPoints(0)
    End With
    ws.Visible = xlSheetVisible
    ws.ExportAsFixedFormat Type:=xlTypePDF, Filename:=ruta, Quality:=xlQualityStandard
    Application.DisplayAlerts = False
    ws.Delete
    Application.DisplayAlerts = True
    Shell "explorer.exe """ & ruta & """", vbNormalFocus
End Sub
