' Versión refactorizada y mejorada del código original.
' Mejoras:
' - Código estructurado en funciones y subrutinas para mayor claridad.
' - Uso de constantes para "números mágicos" y cadenas de texto.
' - Búsqueda de festivos más eficiente usando un Dictionary.
' - Cálculo de la última fila más robusto.
' - Comentarios explicativos.
' - Desactivación de la actualización de pantalla para mejorar el rendimiento.

Option Explicit

' --- Constantes para mejorar la legibilidad y mantenibilidad ---
Private Const COL_SALDO As Long = 1
Private Const COL_NOMBRE As Long = 2 ' Asumiendo que los nombres están en la columna B
Private Const PRIMERA_COL_FECHA As Long = 3

Private Const CODIGO_DIA_LIBRE_1 As String = "L"
Private Const CODIGO_DIA_LIBRE_2 As String = "G"
Private Const CODIGO_SIN_SUPLENCIA As String = "9024"

Private Const SALDO_TURNO_NORMAL As Long = 9
Private Const SALDO_OTROS As Long = 7
Private Const SALDO_SIN_SUPLENCIA As Long = -7

'
' --- Subrutina Principal ---
'
Public Sub ProcesarCuadrante()
    On Error GoTo ErrorHandler

    ' --- Optimización de rendimiento y prevención de recursión ---
    Application.ScreenUpdating = False
    Application.EnableEvents = False ' Evita que la macro se dispare a sí misma por eventos de cambio de celda

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(1)

    ' Asegura que la columna de Saldo exista
    PrepararHoja ws

    ' Carga los festivos en una estructura de datos para búsqueda rápida
    Dim festivos As Object ' Scripting.Dictionary
    Set festivos = CargarFestivos()

    Dim lastRow As Long
    ' Se calcula la última fila basándose en la columna de nombres para mayor fiabilidad
    lastRow = ws.Cells(ws.Rows.Count, COL_NOMBRE).End(xlUp).Row

    Dim lastCol As Long
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    Dim reemplazosFecha As Long
    Dim reemplazos9024 As Long

    Dim i As Long
    For i = 2 To lastRow
        ProcesarFila ws, i, lastCol, festivos, reemplazosFecha, reemplazos9024
    Next i

    ' Mensaje de éxito antes de limpiar y salir
    MsgBox "Proceso completado." & vbCrLf & vbCrLf & _
           "Sustituciones por suplencia (fecha): " & reemplazosFecha & vbCrLf & _
           "Sustituciones por '" & CODIGO_SIN_SUPLENCIA & "': " & reemplazos9024, vbInformation, "Cálculo de Saldo Finalizado"

ExitSub:
    ' --- Restaurar configuración de la aplicación ---
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Sub

ErrorHandler:
    ' --- Manejo de errores ---
    MsgBox "Se ha producido un error en tiempo de ejecución: " & Err.Number & vbCrLf & vbCrLf & Err.Description, vbCritical, "Error en la Macro"
    Resume ExitSub
End Sub

'
' --- Funciones y Subrutinas de Ayuda ---
'

' Prepara la hoja de cálculo, insertando la columna "Saldo" si no existe.
Private Sub PrepararHoja(ByVal ws As Worksheet)
    If ws.Cells(1, COL_SALDO).Value <> "Saldo" Then
        ws.Columns(COL_SALDO).Insert Shift:=xlToRight
        ws.Cells(1, COL_SALDO).Value = "Saldo"
        ws.Cells(1, COL_SALDO).Font.Bold = True
    End If
End Sub

' Carga los días festivos desde un array a un Dictionary para una búsqueda eficiente.
' NOTA: Sería ideal cargar estos datos desde una hoja de configuración o un rango con nombre.
Private Function CargarFestivos() As Object ' As Scripting.Dictionary
    Set CargarFestivos = CreateObject("Scripting.Dictionary")

    ' Festivos para 2025 (hardcoded)
    Dim festivosArray As Variant
    festivosArray = Array("01/01/2025", "06/01/2025", "17/04/2025", "18/04/2025", _
                        "01/05/2025", "02/05/2025", "15/05/2025", "25/07/2025", _
                        "15/08/2025", "01/11/2025", "06/12/2025", "08/12/2025", _
                        "25/12/2025", "10/11/2025")

    Dim festivo As Variant
    For Each festivo In festivosArray
        If IsDate(festivo) Then
            ' Almacenamos la fecha sin la parte horaria para comparaciones fiables
            CargarFestivos(DateValue(CDate(festivo))) = True
        End If
    Next festivo
End Function

' Procesa una única fila de empleado.
Private Sub ProcesarFila(ByVal ws As Worksheet, ByVal rowIdx As Long, ByVal lastCol As Long, _
                         ByVal festivos As Object, ByRef reemplazosFecha As Long, ByRef reemplazos9024 As Long)

    Dim suplencias As Collection
    Set suplencias = ObtenerSuplencias(ws, rowIdx, lastCol)

    Dim saldo As Long
    saldo = 0

    Dim j As Long
    For j = PRIMERA_COL_FECHA To lastCol
        Dim celda As Range
        Set celda = ws.Cells(rowIdx, j)

        If Not IsDate(ws.Cells(1, j).Value) Then GoTo SiguienteCelda

        Dim fecha As Date
        fecha = DateValue(ws.Cells(1, j).Value) ' Usamos DateValue para ignorar la hora

        Dim valor As String
        valor = Trim(ws.Cells(rowIdx, j).Value)

        If (valor = CODIGO_DIA_LIBRE_1 Or valor = CODIGO_DIA_LIBRE_2) And EsDiaLaborable(fecha, festivos) Then
            ' Es un día libre que debe ser reemplazado
            If suplencias.Count > 0 Then
                ' Reemplazar con fecha de suplencia
                celda.Value = Format(suplencias(1), "dd-mmm")
                celda.Interior.Color = RGB(255, 165, 0) ' Naranja
                saldo = saldo + SALDO_TURNO_NORMAL
                reemplazosFecha = reemplazosFecha + 1
                suplencias.Remove 1
            Else
                ' Reemplazar con código de "sin suplencia"
                celda.Value = CODIGO_SIN_SUPLENCIA
                celda.Interior.Color = RGB(173, 216, 230) ' Azul claro
                saldo = saldo + SALDO_SIN_SUPLENCIA
                reemplazos9024 = reemplazos9024 + 1
            End If
        Else
            ' Es un día normal, calcular saldo según su código
            saldo = saldo + CalcularPuntosSaldo(valor)
        End If

SiguienteCelda:
    Next j

    ws.Cells(rowIdx, COL_SALDO).Value = saldo
End Sub

' Recopila los días de suplencia (trabajo en fin de semana) para un empleado.
Private Function ObtenerSuplencias(ByVal ws As Worksheet, ByVal rowIdx As Long, ByVal lastCol As Long) As Collection
    Set ObtenerSuplencias = New Collection

    Dim j As Long
    For j = PRIMERA_COL_FECHA To lastCol
        If IsDate(ws.Cells(1, j).Value) Then
            Dim fecha As Date
            fecha = ws.Cells(1, j).Value

            If Weekday(fecha, vbMonday) > 5 Then ' Sábado o Domingo
                Select Case ws.Cells(rowIdx, j).Value
                    Case "07", "15", "22", "B"
                        ObtenerSuplencias.Add fecha
                End Select
            End If
        End If
    Next j
End Function

' Calcula los puntos de saldo a sumar/restar según el código de la celda.
Private Function CalcularPuntosSaldo(ByVal valor As String) As Long
    Select Case valor
        Case "07", "15", "22"
            CalcularPuntosSaldo = SALDO_TURNO_NORMAL
        Case "MP", "FT", "V", "P", "B"
            CalcularPuntosSaldo = SALDO_OTROS
        Case CODIGO_SIN_SUPLENCIA
            CalcularPuntosSaldo = SALDO_SIN_SUPLENCIA
        Case Else
            CalcularPuntosSaldo = 0
    End Select
End Function

' Comprueba si una fecha es un día laborable (no fin de semana y no festivo).
Private Function EsDiaLaborable(ByVal fecha As Date, ByVal festivos As Object) As Boolean
    ' La fecha ya viene sin hora (DateValue)
    If Weekday(fecha, vbMonday) > 5 Then ' Es Sábado o Domingo
        EsDiaLaborable = False
        Exit Function
    End If

    If festivos.Exists(fecha) Then ' Es festivo
        EsDiaLaborable = False
        Exit Function
    End If

    EsDiaLaborable = True
End Function
