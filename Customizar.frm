VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Customizar 
   Caption         =   "..:: Nuevos Combos ::.."
   ClientHeight    =   7344
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   10740
   OleObjectBlob   =   "Customizar.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "Customizar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ws As Worksheet
Dim wsa As Worksheet
Dim tbl As ListObject
Dim col As ListColumn
Dim colum As ListColumn
Dim cell As Range
Dim celda As Range
Dim productos As New Collection

Private Sub btnAgregar_Click()
    If Not cbSelect.Text = "- Seleccionar -" Then
        If Not IsNull(cbSelect.Value) Then
            lbDatosCombo.AddItem cbSelect.Text
        End If
        btnGuardar.Enabled = True
        CargarSelect cbSelect.Text
    Else
        MsgBox "Debes seleccionar un producto.", vbExclamation, "Consola"
    End If
End Sub

Private Sub btnCargar_Click()
    OrdenarPorCodigo
    Set ws = ThisWorkbook.Sheets("Asesores de Venta")
    Set tbl = ws.ListObjects("Asesores")
    Dim Coincidencias As Double
    Dim VerificarEnExcel As Integer
    Coincidencias = False
    If Not tbl Is Nothing Then
        Set col = tbl.ListColumns(1)
        
        For Each cell In col.DataBodyRange.Cells
            If txtCodigo.Text = cell.Value Then
                If ws.Cells(cell.row, col.Index + 4).Value = txtTel.Text Then
                    frmCustom.Visible = True
                    MsgBox "Bienvenido/a " & ws.Cells(cell.row, col.Index + 1).Value & "."
                    Frame1.Caption = "Asesor: " & ws.Cells(cell.row, col.Index + 1).Value
                    Coincidencias = True
                End If
            End If
        Next cell
    Else
        MsgBox "La tabla de Clientes no se encontró.", vbCritical, "Consola"
    End If
    
    If Not Coincidencias Then
        frmCustom.Visible = False
        VerificarEnExcel = MsgBox("No coincideron los datos, ¿Quieres verificarlos en Excel?", vbYesNo, "Consola")
        If VerificarEnExcel = vbYes Then
            Sheets("Asesores de Venta").Activate
            Unload Customizar
        End If
    Else
        CargarSelect ""
    End If
    DetectarLetrasIniciales
End Sub

Private Sub btnEliminar_Click()
    BorrarFilasPorCodigo (Right(cbCombos.Value, 2))
        Customizar.Hide
        MsgBox "Todo listo para continuar"
        Unload Customizar
End Sub

Private Sub btnGuardar_Click()
    Set ws = ThisWorkbook.Sheets("Productos")
    Set tbl = ws.ListObjects("Catalogo")
    Set col = tbl.ListColumns("CODIGO")
    Dim ComoExistente As Boolean
    ComoExistente = False
    If lbDatosCombo.ListCount = 0 Then
        MsgBox "Deben Existir datos para el combo que vas a crear", vbExclamation, "Consola"
    ElseIf Left(txtCodigoCombo.Text, 2) = "XY" And Not chbExistente.Value Then
            MsgBox "Coloca un codigo diferente al generico.", vbExclamation, "Consola"
            ComboExistente = True
        Else
            Set letrasCombo = cbSelect
            letrasCombo.Clear
            
            ' Inicializar una colección para almacenar temporalmente las letras iniciales
            Dim letras As New Collection
            For Each cell In col.DataBodyRange.Cells
                ' Obtener la letra inicial de cada celda
                Dim letra As String
                letra = Left(cell.Value, 2)
                
                ' Verificar si la letra inicial no está en la colección y agregarla si no está
                If letra = Left(txtCodigoCombo.Text, 2) Then
                    ComboExistente = True
                End If
            Next cell
        End If
    If chbExistente.Value Then
        BorrarFilasPorCodigo Right(cbCombos.Value, 2)
    End If
    If ComboExistente Then
        MsgBox "El Combo ya fue creado con anterioridad, porfavor verifica los primeros caracteres en el codigo.", vbExclamation, "Consola"
    Else
        Dim nuevaFila As ListRow
        Set nuevaFila = tbl.ListRows.Add
        Set col = tbl.ListColumns("ARTICULO")
        
        For Each cell In col.DataBodyRange.Cells
            Dim rowPos As Long
            Dim colPos As Long
            rowPos = cell.row - tbl.HeaderRowRange.row + 1
            colPos = cell.Column - tbl.HeaderRowRange.Column + 1
            If colPos = 2 Then
                For i = 0 To lbDatosCombo.ListCount - 1
                    If cell.Value = lbDatosCombo.List(i) Then ' en el combo seleccionado buscar
                        nuevaFila.Range(i, 1).Value = txtCodigoCombo.Text
                        nuevaFila.Range(i, 2).Value = cell.Value
                        nuevaFila.Range(i, 3).Value = ws.Cells(cell.row, col.Index + 1).Value
                    End If
                Next i
            End If
        Next cell
        MsgBox "El nuevo combo `" & txtCodigoCombo.Text & "` ha sido creado con Exito.", vbInformation, "Consola"
        Customizar.Hide
        MsgBox "Todo listo para continuar"
        Unload Customizar
    End If
    CargarSelect ""
End Sub

Private Sub btnQuitar_Click()
    Dim cotizacionSelect As Integer
    cotizacionSelect = lbDatosCombo.ListIndex
    
    lbDatosCombo.RemoveItem cotizacionSelect
    CargarSelect ""
    btnQuitar.Enabled = False
End Sub

Private Sub cbCombos_Change()
    Set ws = ThisWorkbook.Sheets("Productos")
    Set tbl = ws.ListObjects("Catalogo")
    Set col = tbl.ListColumns("CODIGO")
    lbDatosCombo.Clear
    
    For Each cell In col.DataBodyRange.Cells
        If Left(cell.Value, 2) = Right(cbCombos.Value, 2) Then
            lbDatosCombo.AddItem ws.Cells(cell.row, col.Index + 1).Value
        End If
    Next cell
    txtCodigoCombo.Text = Right(cbCombos.Value, 2)
    CargarSelect ""
    btnGuardar.Enabled = True
End Sub

Private Sub chbExistente_Click()
    If chbNuevo.Value Then
        chbNuevo.Value = False
    Else
        chbExistente.Value = True
        cbCombos.Visible = True
        txtCodigoCombo.Visible = False
        btnEliminar.Enabled = True
    End If
End Sub

Private Sub chbNuevo_Click()
    If chbExistente.Value Then
        chbExistente.Value = False
    Else
        chbNuevo.Value = True
        cbCombos.Visible = False
        txtCodigoCombo.Visible = True
        btnGuardar.Enabled = False
        btnEliminar.Enabled = False
    End If
End Sub

Private Sub lbDatosCombo_Click()
    btnQuitar.Enabled = True
End Sub

Private Sub txtTel_Change()
    ValidarDatosDeTexto txtTel.Text, txtTel
End Sub

Private Sub txtCodigo_Change()
    ValidarDatosDeTexto txtCodigo.Text, txtCodigo
End Sub

Private Sub txtTel_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
' al editar el cuadro de texto 'cantidad'
    If txtTel.Text = "0000-0000" Then
        txtTel.Text = ""
    End If
End Sub

Private Sub txtCodigoCombo_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
' al editar el cuadro de texto 'cantidad'
    If txtCodigoCombo.Text = "XYZ-123" Then
        txtCodigoCombo.Text = ""
    End If
End Sub

Private Sub txtCodigoCombo_Change()
    Dim texto As String
        Dim i As Integer
        texto = txtCodigoCombo.Text
        For i = 1 To Len(texto)
            If Mid(texto, i, 1) Like "[a-zA-Z]" Then
                Mid(texto, i, 1) = UCase(Mid(texto, i, 1))
            End If
        Next i
        txtCodigoCombo.Text = texto
End Sub

Private Sub txtCodigo_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
' al editar el cuadro de texto 'cantidad'
    If txtCodigo.Text = "000" Then
        txtCodigo.Text = ""
    End If
End Sub

Sub ValidarDatosDeTexto(Dato As String, Cuadro As Object)
    If InStr(Dato, "-") = 5 Then
        Cuadro.Text = Cuadro.Text
    ElseIf Not IsNumeric(Dato) Then
        Cuadro.Text = ""
    End If
    If Val(Dato) <= 0 Then
        Cuadro.Text = ""
    End If
End Sub

Sub CargarSelect(Coincidencia As String)
    OrdenarPorCodigo
    Dim letrasCombo As ComboBox
    Dim item As Variant
    Set ws = ThisWorkbook.Sheets("Productos")
    Set tbl = ws.ListObjects("Catalogo")
    Set col = tbl.ListColumns("ARTICULO")
    Set letrasCombo = cbSelect
    letrasCombo.Clear
    
    For Each item In productos
        productos.Remove 1
    Next item
    
    ' Inicializar una colección para almacenar temporalmente las letras iniciales
    If Not lbDatosCombo.ListCount = 0 Then
        For i = 0 To lbDatosCombo.ListCount - 1
            productos.Add lbDatosCombo.List(i)
        Next i
    End If
    
    For Each cell In col.DataBodyRange.Cells
        ' Verificar si la letra inicial no está en la colección y agregarla si no está
        If Not Contiene(productos, cell.Value) Then
            cbSelect.AddItem cell.Value
        End If
    Next cell
    
    letrasCombo.Value = "- Seleccionar -"
End Sub

Sub DetectarLetrasIniciales() ' Funcion para acomodar los combos segun el codigo de serie
    Set ws = ThisWorkbook.Sheets("Productos")
    Set tbl = ws.ListObjects("Catalogo")
    Set col = tbl.ListColumns("CODIGO")
    Set letrasCombo = cbCombos
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

Sub OrdenarPorCodigo()
    Dim ws As Worksheet
    Dim rng As Range
    Dim LastRow As Long
    
    ' Definir la hoja de trabajo
    Set ws = ThisWorkbook.Sheets("Productos") ' Reemplaza "NombreDeTuHoja" con el nombre real de tu hoja
    
    ' Encontrar la última fila con datos en la columna A
    LastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).row
    
    ' Definir el rango que quieres ordenar
    Set rng = ws.Range("A1:C" & LastRow) ' Asumiendo que tus datos comienzan en la fila 2 y que la columna C contiene los precios
    
    ' Ordenar el rango por los caracteres iniciales del código
    With ws.Sort
        .SortFields.Clear
        .SortFields.Add Key:=rng.Columns(1), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SetRange rng
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub

Function BorrarFilasPorCodigo(ByVal codigoABorrar As String)
    Dim rng As Range
    Dim LastRow As Long
    Dim i As Long
    
    ' Definir la hoja de trabajo
    Set ws = ThisWorkbook.Sheets("Productos") ' Reemplaza "NombreDeTuHoja" con el nombre real de tu hoja
    
    ' Encontrar la última fila con datos en la columna A
    LastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).row
    
    ' Definir el rango que quieres recorrer
    Set rng = ws.Range("A1:A" & LastRow) ' Asumiendo que tus datos comienzan en la fila 2 y la columna A contiene los códigos
    
    ' Recorrer el rango y borrar las filas que coincidan con el código
    For i = rng.Rows.Count To 1 Step -1
        If Left(rng.Cells(i, 1).Value, 2) = codigoABorrar Then
            ws.Rows(i).EntireRow.Delete ' Sumamos 1 porque A2 corresponde a la fila 1 en la hoja de cálculo
        End If
    Next i
End Function
