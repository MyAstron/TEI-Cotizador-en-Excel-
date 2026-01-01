VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} InicioSesion 
   Caption         =   "..:: Inicio de Sesion ::.. (Registro de Datos)"
   ClientHeight    =   3072
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   8412.001
   OleObjectBlob   =   "InicioSesion.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "InicioSesion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Variables - Usos
Private USUARIO_INGRESADO() As String
Private ASESOR_VENTAS() As String
' Tablas
Dim ws As Worksheet
Dim tblClientes As ListObject
Dim tblAsesores As ListObject
Dim row As ListRow
Dim cell As Range
' BtnIngresar
Dim texto As String
Dim ClienteID As Double

Private Sub BtnIngresar_Click()
    Dim AsesorExistente As Boolean
    AsesorExistente = False
    ' buscar coincidencias en la tabla de "asesores"
    Set ws = ThisWorkbook.Sheets("Asesores de Venta")
    Set tblAsesores = ws.ListObjects("Asesores")
        
    ' Verificar si tblAsesores es válido antes de usarlo
    If Not tblAsesores Is Nothing Then
        Set col = tblAsesores.ListColumns(2)
        
        For Each cell In col.DataBodyRange.Cells
            If cbNombres.Value = cell.Value Then
                AsesorExistente = True
            End If
        Next cell
    Else
        MsgBox "La tabla de Clientes no se encontró.", vbCritical, "Consola"
    End If
    
    If Not AsesorExistente Then
        MsgBox "El Asesor de Ventas no es valido.", vbCritical, "Consola"
    Else
        InicioSesion.Hide
        Cotizacion.Show
    End If
End Sub

Private Sub btnNext_Click()
    texto = ""
    
    If Not IsNumeric(NumCliente.Text) Then ' Verificar que es un número
        MsgBox "No ingresaste un Dato Numerico", vbCritical, "Consola"
    Else
        ClienteID = Val(NumCliente.Text) ' Convertir el texto a un valor numérico
    End If
    
    If Len(Trim(NumCliente.Text)) = 0 Then ' Verificar caracteres existentes
        texto = "No ingresaste nada"
    ElseIf Len(texto) = 0 Then
        RecorerAsesores
        BuscarCliente
    End If
    frmCliente.Caption = "Cliente Registrado"
    lblCliente.Caption = "ID: " & NumCliente.Text
    NumCliente.Text = "Name: " & InicioSesion.MiVariable(2)
    frmAsesor.Visible = True
    frmCliente.Enabled = False
    BtnIngresar.Visible = True
    btnNext.Visible = False
End Sub

Sub BuscarCliente()
        ' buscar coincidencias en la tabla de "clientes"
        Set ws = ThisWorkbook.Sheets("Lista de Clientes")
        Set tblClientes = ws.ListObjects("Clientes")
        
        ' Verificar si tblClientes es válido antes de usarlo
        If Not tblClientes Is Nothing Then
            Set col = tblClientes.ListColumns(1)
            
            For Each cell In col.DataBodyRange.Cells
                ReDim Preserve USUARIO_INGRESADO(1 To 6)
                If ClienteID = cell.Value Then
                    texto = ws.Cells(cell.row, col.Index + 1).Value
                    For i = 1 To 6 ' Recorre las 6 columnas
                        USUARIO_INGRESADO(i) = ws.Cells(cell.row, col.Index + i - 1).Value
                    Next i
                    Exit For
                Else
                    For i = 1 To 6
                        USUARIO_INGRESADO(1) = "000"
                        USUARIO_INGRESADO(2) = "Usuario"
                        USUARIO_INGRESADO(i) = "-"
                    Next i
                End If
            Next cell
        Else
            MsgBox "La tabla de Clientes no se encontró.", vbCritical, "Consola"
        End If
    
        If texto = "" Then
            texto = "Usuario"
        End If
End Sub

Sub RecorerAsesores()
    ' buscar coincidencias en la tabla de "asesores"
    Set ws = ThisWorkbook.Sheets("Asesores de Venta")
    Set tblAsesores = ws.ListObjects("Asesores")
        
    ' Verificar si tblAsesores es válido antes de usarlo
    If Not tblAsesores Is Nothing Then
        Set col = tblAsesores.ListColumns(2)
            
        For Each cell In col.DataBodyRange.Cells
            cbNombres.AddItem cell.Value
        Next cell
    Else
        MsgBox "La tabla de Clientes no se encontró.", vbCritical, "Consola"
    End If
End Sub

Private Sub cbNombres_Change()
    ' buscar coincidencias en la tabla de "asesores"
    Set ws = ThisWorkbook.Sheets("Asesores de Venta")
    Set tblAsesores = ws.ListObjects("Asesores")
        
    ' Verificar si tblAsesores es válido antes de usarlo
    If Not tblAsesores Is Nothing Then
        Set col = tblAsesores.ListColumns(2)
        
        For Each cell In col.DataBodyRange.Cells
            If cbNombres.Value = cell.Value Then
                txtAsesor.Text = Left(cell.Value, 1) & ws.Cells(cell.row, col.Index - 1).Value
                ReDim Preserve ASESOR_VENTAS(1 To 6)
                For i = 1 To 6 ' Recorre las 6 columnas
                    ASESOR_VENTAS(i) = ws.Cells(cell.row, col.Index + i - 1).Value
                Next i
                Exit For
            End If
        Next cell
    Else
        MsgBox "La tabla de Clientes no se encontró.", vbCritical, "Consola"
    End If
End Sub

' ------- Otros Eventos -------
Private Sub NumCliente_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    NumCliente.Text = ""
End Sub

Private Sub txtAsesor_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    txtAsesor.Text = ""
End Sub

Public Property Get MiVariable(Index As Integer) As String
    MiVariable = USUARIO_INGRESADO(Index)
End Property

Public Property Let MiVariable(Index As Integer, ByVal Value As String)
    USUARIO_INGRESADO(Index) = Value
End Property

Public Property Get MisDatos(Index As Integer) As String
    MisDatos = ASESOR_VENTAS(Index)
End Property

Public Property Let MisDatos(Index As Integer, ByVal Value As String)
    ASESOR_VENTAS(Index) = Value
End Property

