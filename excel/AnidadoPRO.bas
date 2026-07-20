Attribute VB_Name = "Mod_AnidadoPRO"
' =====================================================================
' Macro:  AnidadoPRO  -  Grupo KIMAK (Departamento de Compras)
'
' Calcula cuantos tableros hay que pedir por material a partir del CSV
' exportado del SIG, con nesting por franjas (guillotina) que respeta
' la direccion de la veta de cada pieza.
'
' Dos modos de uso con la misma macro:
'   1) Abrir un CSV del SIG y ejecutar AnidadoPRO.
'   2) Ejecutarla sin datos delante (libro vacio): abre el selector de
'      archivos de Windows, permite elegir varios CSV a la vez, los
'      combina en un libro nuevo y calcula el pedido completo.
'
' Debajo de los datos se generan dos tablas:
'   - MEDIDAS DE TABLERO: medida editable por material (3050x1220 por
'     defecto) y columna MISMO QUE (desplegable) para unificar
'     materiales duplicados por errores de escritura. Los posibles
'     duplicados se marcan en amarillo. Boton RECALCULAR incluido.
'   - TABLEROS A PEDIR: material, medida de tablero, calculo exacto y
'     unidades a pedir (redondeadas hacia arriba). Las piezas que no
'     caben en ningun tablero se avisan en rojo.
'
' Reglas de veta:
'   - Dir. Veta = "Largo", "Corto" o cualquier otro texto: la pieza no
'     gira; la medida de la columna Largo (donde va la veta) se alinea
'     con el largo del tablero. El SIG escribe "Corto" cuando la veta
'     va en el largo pero el largo es la medida pequena de la pieza.
'   - Dir. Veta = "Ancho": la pieza no gira; se coloca girada 90 grados
'     para que su Ancho quede alineado con el largo del tablero.
'   - Dir. Veta vacia o "0": sin veta; gira solo si Permite giro = 1.
'   - Si una fila trae veta y Permite giro = 1, la veta manda: no gira.
' =====================================================================
Option Explicit

' --- Parametros de corte (modificar si cambian las condiciones) ---
Private Const DEF_TAB_L As Double = 3050  ' Largo de tablero por defecto (mm)
Private Const DEF_TAB_A As Double = 1220  ' Ancho de tablero por defecto (mm)
Private Const MARGEN As Double = 12       ' Orilla en los 4 lados (mm)
Private Const SEP As Double = 17          ' Separacion piezas/franjas: 12 libre + 5 kerf (mm)

' --- Columnas del CSV del SIG (1=A, 2=B, 3=C...) ---
Private Const C_MAT As Long = 3     ' C: Mat. - recubr. - esp.
Private Const C_QTY As Long = 8     ' H: Cantidad
Private Const C_L As Long = 9       ' I: Largo (mm, direccion de la veta)
Private Const C_A As Long = 10      ' J: Ancho (mm)
Private Const C_GIRO As Long = 11   ' K: Permite giro (1 = si)
Private Const C_VETA As Long = 13   ' M: Dir. Veta ("Largo" / "Ancho" / vacio)

' --- Salida ---
Private Const TIT_MEDIDAS As String = "MEDIDAS DE TABLERO"
Private Const TIT_RESULT As String = "TABLEROS A PEDIR"
Private Const COL_OUT As Long = 3         ' Columna C: primera columna de las tablas
Private Const COL_LISTA As Long = 21      ' Columna U (oculta): lista para desplegables
Private Const BTN_NAME As String = "btnRecalcularTableros"


' =====================================================================
' Macro principal
' =====================================================================
Public Sub AnidadoPRO()
    Dim ws As Worksheet
    Set ws = ActiveSheet

    Dim dataLast As Long
    dataLast = UltimaFilaDatos(ws)

    ' --- Sin datos delante: modo importacion de varios CSV ---
    If dataLast < 2 Then
        If MsgBox("La hoja activa no tiene datos de piezas." & vbNewLine & vbNewLine & _
                  "Quieres seleccionar uno o varios CSV del SIG para importarlos " & _
                  "juntos y calcular los tableros del pedido completo?", _
                  vbYesNo + vbQuestion, "AnidadoPRO") <> vbYes Then Exit Sub
        Set ws = ImportarCSVs()
        If ws Is Nothing Then Exit Sub
        dataLast = UltimaFilaDatos(ws)
        If dataLast < 2 Then
            MsgBox "Los archivos seleccionados no contienen piezas.", _
                   vbExclamation, "AnidadoPRO"
            Exit Sub
        End If
    End If

    On Error GoTo Fallo
    Application.ScreenUpdating = False

    ' -----------------------------------------------------------------
    ' PASO 1: leer medidas y uniones de una ejecucion anterior
    ' (antes de limpiar la zona de resultados)
    ' -----------------------------------------------------------------
    Dim prevL As Object, prevA As Object, prevQ As Object
    Call LeerMedidasPrevias(ws, prevL, prevA, prevQ)

    ' -----------------------------------------------------------------
    ' PASO 2: materiales unicos en orden de aparicion
    ' -----------------------------------------------------------------
    Dim mats As Object
    Set mats = CreateObject("Scripting.Dictionary")
    mats.CompareMode = vbTextCompare

    Dim matNames() As String
    Dim matCount As Long
    matCount = 0

    Dim r As Long, mat As String
    For r = 2 To dataLast
        mat = Trim$(CStr(ws.Cells(r, C_MAT).Value))
        If mat <> "" Then
            If Not mats.Exists(mat) Then
                ReDim Preserve matNames(0 To matCount)
                matNames(matCount) = mat
                mats.Add mat, matCount
                matCount = matCount + 1
            End If
        End If
    Next r

    If matCount = 0 Then
        Application.ScreenUpdating = True
        MsgBox "No se encontraron materiales en la columna C.", vbExclamation
        Exit Sub
    End If

    ' -----------------------------------------------------------------
    ' PASO 3: medidas de tablero y uniones por material
    ' -----------------------------------------------------------------
    Dim tabL() As Double, tabA() As Double, mismoQue() As String
    ReDim tabL(0 To matCount - 1)
    ReDim tabA(0 To matCount - 1)
    ReDim mismoQue(0 To matCount - 1)

    Dim i As Long
    Dim medidasMalas As String
    For i = 0 To matCount - 1
        tabL(i) = DEF_TAB_L
        tabA(i) = DEF_TAB_A
        mismoQue(i) = ""
        If prevL.Exists(matNames(i)) Then
            If prevL(matNames(i)) > 0 Then tabL(i) = prevL(matNames(i))
            If prevA(matNames(i)) > 0 Then tabA(i) = prevA(matNames(i))
            mismoQue(i) = prevQ(matNames(i))
        End If
        ' Una union solo vale si apunta a un material existente y distinto
        If mismoQue(i) <> "" Then
            If Not mats.Exists(mismoQue(i)) Then
                mismoQue(i) = ""
            ElseIf mats(mismoQue(i)) = i Then
                mismoQue(i) = ""
            End If
        End If
        ' Medidas absurdas (menores que las orillas): volver al defecto
        If tabL(i) <= 2 * MARGEN + 10 Or tabA(i) <= 2 * MARGEN + 10 Then
            medidasMalas = medidasMalas & vbNewLine & "  - " & matNames(i)
            tabL(i) = DEF_TAB_L
            tabA(i) = DEF_TAB_A
        End If
    Next i

    If medidasMalas <> "" Then
        MsgBox "Estas medidas de tablero no eran validas y se han " & _
               "restablecido a " & DEF_TAB_L & " x " & DEF_TAB_A & ":" & _
               vbNewLine & medidasMalas, vbExclamation, "AnidadoPRO"
    End If

    ' Confirmar uniones entre espesores distintos (casi siempre error de dedo)
    Dim e1 As Double, e2 As Double
    For i = 0 To matCount - 1
        If mismoQue(i) <> "" Then
            e1 = ParseEspesor(matNames(i))
            e2 = ParseEspesor(mismoQue(i))
            If e1 > 0 And e2 > 0 And e1 <> e2 Then
                If MsgBox("Vas a unir dos materiales con ESPESOR DISTINTO:" & _
                          vbNewLine & vbNewLine & _
                          matNames(i) & "  (espesor " & e1 & ")" & vbNewLine & _
                          "con" & vbNewLine & _
                          mismoQue(i) & "  (espesor " & e2 & ")" & vbNewLine & vbNewLine & _
                          "Seguro que son el mismo material?", _
                          vbYesNo + vbExclamation, "Union de materiales") <> vbYes Then
                    mismoQue(i) = ""
                End If
            End If
        End If
    Next i

    ' Resolver cadenas de uniones (A -> B -> C acaba en C) y detectar circulos
    Dim destino() As Long
    ReDim destino(0 To matCount - 1)
    Dim ciclos As String
    Dim j As Long, pasos As Long
    For i = 0 To matCount - 1
        j = i
        pasos = 0
        Do While mismoQue(j) <> "" And pasos <= matCount
            j = mats(mismoQue(j))
            pasos = pasos + 1
        Loop
        If pasos > matCount Then
            destino(i) = i
            ciclos = ciclos & vbNewLine & "  - " & matNames(i)
        Else
            destino(i) = j
        End If
    Next i

    If ciclos <> "" Then
        MsgBox "Hay referencias circulares en la columna MISMO QUE " & _
               "(un material apunta a otro que apunta de vuelta al primero)." & _
               vbNewLine & "Se han ignorado las uniones de:" & vbNewLine & ciclos & _
               vbNewLine & vbNewLine & "Revisa los desplegables y pulsa RECALCULAR.", _
               vbExclamation, "AnidadoPRO"
    End If

    ' -----------------------------------------------------------------
    ' PASO 4: marcar posibles duplicados (nombres casi iguales)
    ' -----------------------------------------------------------------
    Dim sospechoso() As Boolean
    ReDim sospechoso(0 To matCount - 1)
    Dim ni As String, nj As String
    For i = 0 To matCount - 2
        For j = i + 1 To matCount - 1
            If mismoQue(i) = "" And mismoQue(j) = "" Then
                e1 = ParseEspesor(matNames(i))
                e2 = ParseEspesor(matNames(j))
                If e1 = e2 Then
                    ni = NormalizarNombre(matNames(i))
                    nj = NormalizarNombre(matNames(j))
                    If Levenshtein(ni, nj) <= 3 Then
                        sospechoso(i) = True
                        sospechoso(j) = True
                    End If
                End If
            End If
        Next j
    Next i

    ' -----------------------------------------------------------------
    ' PASO 5: nesting por grupo de materiales unificados
    ' -----------------------------------------------------------------
    Dim exactoG() As Double, pedirG() As Long
    ReDim exactoG(0 To matCount - 1)
    ReDim pedirG(0 To matCount - 1)
    Dim imposibles As New Collection

    Dim t As Long
    For t = 0 To matCount - 1
        If destino(t) = t Then
            Call CalcularGrupo(ws, dataLast, mats, destino, t, matNames(t), _
                               tabL(t), tabA(t), exactoG(t), pedirG(t), imposibles)
        End If
    Next t

    ' -----------------------------------------------------------------
    ' PASO 6: escribir tablas en la hoja
    ' -----------------------------------------------------------------
    Call LimpiarZonaSalida(ws, dataLast, matCount)

    Dim r0 As Long
    r0 = dataLast + 3

    ' ---- Tabla MEDIDAS DE TABLERO ----
    Call EstiloTitulo(ws.Range(ws.Cells(r0, COL_OUT), ws.Cells(r0, COL_OUT + 3)), TIT_MEDIDAS)
    Call EstiloCabecera(ws.Cells(r0 + 1, COL_OUT), "MATERIAL")
    Call EstiloCabecera(ws.Cells(r0 + 1, COL_OUT + 1), "LARGO TABLERO")
    Call EstiloCabecera(ws.Cells(r0 + 1, COL_OUT + 2), "ANCHO TABLERO")
    Call EstiloCabecera(ws.Cells(r0 + 1, COL_OUT + 3), "MISMO QUE")

    Dim fila As Long
    For i = 0 To matCount - 1
        fila = r0 + 2 + i
        With ws.Cells(fila, COL_OUT)
            .Value = matNames(i)
            If sospechoso(i) Then
                .Interior.Color = RGB(255, 235, 156)   ' amarillo: posible duplicado
            ElseIf i Mod 2 = 1 Then
                .Interior.Color = RGB(242, 246, 252)
            End If
        End With
        With ws.Cells(fila, COL_OUT + 1)
            .Value = tabL(i)
            .NumberFormat = "0"
            If i Mod 2 = 1 Then .Interior.Color = RGB(242, 246, 252)
        End With
        With ws.Cells(fila, COL_OUT + 2)
            .Value = tabA(i)
            .NumberFormat = "0"
            If i Mod 2 = 1 Then .Interior.Color = RGB(242, 246, 252)
        End With
        With ws.Cells(fila, COL_OUT + 3)
            .Value = mismoQue(i)
            If i Mod 2 = 1 Then .Interior.Color = RGB(242, 246, 252)
        End With
        ws.Cells(fila, COL_LISTA).Value = matNames(i)
    Next i

    ' Desplegable MISMO QUE alimentado por la lista oculta de la columna U
    Dim rngMismo As Range
    Set rngMismo = ws.Range(ws.Cells(r0 + 2, COL_OUT + 3), _
                            ws.Cells(r0 + 1 + matCount, COL_OUT + 3))
    With rngMismo.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
             Formula1:="=" & ws.Cells(r0 + 2, COL_LISTA).Address & ":" & _
                             ws.Cells(r0 + 1 + matCount, COL_LISTA).Address
        .IgnoreBlank = True
        .InCellDropdown = True
        .ErrorMessage = "Elige un material de la lista o deja la celda vacia."
    End With
    ws.Columns(COL_LISTA).Hidden = True

    Call Bordear(ws.Range(ws.Cells(r0 + 1, COL_OUT), ws.Cells(r0 + 1 + matCount, COL_OUT + 3)))

    ' Boton RECALCULAR junto al titulo
    Dim btn As Button
    Set btn = ws.Buttons.Add(ws.Cells(r0, COL_OUT + 4).Left + 5, _
                             ws.Cells(r0, COL_OUT + 4).Top - 2, 120, 24)
    btn.Name = BTN_NAME
    btn.Caption = "RECALCULAR"
    btn.Font.Bold = True
    btn.OnAction = "'" & ThisWorkbook.Name & "'!AnidadoPRO"

    ' ---- Tabla TABLEROS A PEDIR ----
    Dim rRes As Long
    rRes = r0 + matCount + 4
    Call EstiloTitulo(ws.Range(ws.Cells(rRes, COL_OUT), ws.Cells(rRes, COL_OUT + 3)), TIT_RESULT)
    Call EstiloCabecera(ws.Cells(rRes + 1, COL_OUT), "MATERIAL")
    Call EstiloCabecera(ws.Cells(rRes + 1, COL_OUT + 1), "MEDIDA TABLERO")
    Call EstiloCabecera(ws.Cells(rRes + 1, COL_OUT + 2), "EXACTO")
    Call EstiloCabecera(ws.Cells(rRes + 1, COL_OUT + 3), "A PEDIR")

    Dim totalPedir As Long
    totalPedir = 0
    fila = rRes + 1
    Dim nGrupos As Long
    nGrupos = 0

    For t = 0 To matCount - 1
        If destino(t) = t Then
            fila = fila + 1
            nGrupos = nGrupos + 1
            If nGrupos Mod 2 = 0 Then
                ws.Range(ws.Cells(fila, COL_OUT), ws.Cells(fila, COL_OUT + 2)) _
                  .Interior.Color = RGB(242, 246, 252)
            End If
            ws.Cells(fila, COL_OUT).Value = matNames(t)
            ws.Cells(fila, COL_OUT + 1).Value = _
                Format$(tabL(t), "0") & " x " & Format$(tabA(t), "0")
            With ws.Cells(fila, COL_OUT + 2)
                .Value = exactoG(t)
                .NumberFormat = "0.0"
            End With
            With ws.Cells(fila, COL_OUT + 3)
                .Value = pedirG(t)
                .NumberFormat = "0"
                .Font.Bold = True
                .Font.Size = 12
                .HorizontalAlignment = xlCenter
                .Interior.Color = RGB(226, 239, 218)   ' verde suave
            End With
            totalPedir = totalPedir + pedirG(t)
        End If
    Next t

    ' Fila de total
    fila = fila + 1
    With ws.Range(ws.Cells(fila, COL_OUT), ws.Cells(fila, COL_OUT + 3))
        .Interior.Color = RGB(31, 56, 100)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
    End With
    ws.Cells(fila, COL_OUT).Value = "TOTAL TABLEROS"
    With ws.Cells(fila, COL_OUT + 3)
        .Value = totalPedir
        .NumberFormat = "0"
        .Font.Size = 12
        .HorizontalAlignment = xlCenter
    End With

    Call Bordear(ws.Range(ws.Cells(rRes + 1, COL_OUT), ws.Cells(fila, COL_OUT + 3)))

    ' ---- Piezas que no caben en ningun tablero: aviso en rojo ----
    If imposibles.Count > 0 Then
        fila = fila + 2
        Dim k As Long
        For k = 1 To imposibles.Count
            With ws.Range(ws.Cells(fila, COL_OUT), ws.Cells(fila, COL_OUT + 3))
                .Interior.Color = RGB(192, 0, 0)
                .Font.Color = RGB(255, 255, 255)
                .Font.Bold = True
            End With
            ws.Cells(fila, COL_OUT).Value = imposibles(k)
            fila = fila + 1
        Next k
    End If

    ' Anchos minimos para que las tablas se lean bien
    If ws.Columns(COL_OUT).ColumnWidth < 48 Then ws.Columns(COL_OUT).ColumnWidth = 48
    If ws.Columns(COL_OUT + 1).ColumnWidth < 15 Then ws.Columns(COL_OUT + 1).ColumnWidth = 15
    If ws.Columns(COL_OUT + 2).ColumnWidth < 15 Then ws.Columns(COL_OUT + 2).ColumnWidth = 15
    If ws.Columns(COL_OUT + 3).ColumnWidth < 48 Then ws.Columns(COL_OUT + 3).ColumnWidth = 48

    Application.GoTo ws.Cells(r0, 1), True
    Application.ScreenUpdating = True
    Exit Sub

Fallo:
    Application.ScreenUpdating = True
    MsgBox "Error inesperado: " & Err.Description, vbCritical, "AnidadoPRO"
End Sub


' =====================================================================
' CalcularGrupo: recoge las piezas de todos los materiales unificados
' en el destino t y las anida. Devuelve exacto, a pedir e imposibles.
' =====================================================================
Private Sub CalcularGrupo(ws As Worksheet, dataLast As Long, mats As Object, _
                          destino() As Long, t As Long, nombreGrupo As String, _
                          tabLv As Double, tabAv As Double, _
                          ByRef exacto As Double, ByRef pedir As Long, _
                          ByRef imposibles As Collection)

    Dim utilL As Double, utilA As Double
    utilL = tabLv - 2 * MARGEN
    utilA = tabAv - 2 * MARGEN

    ' --- Primera pasada: contar piezas expandidas ---
    Dim nP As Long
    nP = 0
    Dim r As Long, mat As String, q As Long
    Dim rawL As Double, rawA As Double

    For r = 2 To dataLast
        mat = Trim$(CStr(ws.Cells(r, C_MAT).Value))
        If mat <> "" Then
            If mats.Exists(mat) Then
                If destino(mats(mat)) = t Then
                    rawL = ToNum(ws.Cells(r, C_L).Value)
                    rawA = ToNum(ws.Cells(r, C_A).Value)
                    If rawL > 0 And rawA > 0 Then
                        q = CLng(ToNum(ws.Cells(r, C_QTY).Value))
                        If q <= 0 Then q = 1
                        nP = nP + q
                    End If
                End If
            End If
        End If
    Next r

    exacto = 0
    pedir = 0
    If nP = 0 Then Exit Sub

    ' --- Segunda pasada: rellenar arrays de piezas ---
    Dim pL() As Double, pA() As Double
    Dim pVeta() As Long      ' 0 = sin veta, 1 = veta en Largo, 2 = veta en Ancho
    Dim pGiro() As Boolean
    ReDim pL(0 To nP - 1)
    ReDim pA(0 To nP - 1)
    ReDim pVeta(0 To nP - 1)
    ReDim pGiro(0 To nP - 1)

    Dim idx As Long, k As Long
    Dim vetaTxt As String, veta As Long
    Dim giro As Boolean
    idx = 0

    For r = 2 To dataLast
        mat = Trim$(CStr(ws.Cells(r, C_MAT).Value))
        If mat <> "" Then
            If mats.Exists(mat) Then
                If destino(mats(mat)) = t Then
                    rawL = ToNum(ws.Cells(r, C_L).Value)
                    rawA = ToNum(ws.Cells(r, C_A).Value)
                    If rawL > 0 And rawA > 0 Then
                        q = CLng(ToNum(ws.Cells(r, C_QTY).Value))
                        If q <= 0 Then q = 1

                        ' Dir. Veta: vacio o "0" = sin veta. "Ancho" = veta en la
                        ' columna Ancho. Cualquier otro texto ("Largo", "Corto"...)
                        ' significa veta en la medida de la columna Largo.
                        vetaTxt = LCase$(Trim$(CStr(ws.Cells(r, C_VETA).Value)))
                        If vetaTxt = "" Or vetaTxt = "0" Then
                            veta = 0
                        ElseIf InStr(vetaTxt, "anch") > 0 Then
                            veta = 2
                        Else
                            veta = 1
                        End If
                        ' La veta manda: solo gira si no hay veta y K = 1
                        giro = (veta = 0) And (ToNum(ws.Cells(r, C_GIRO).Value) = 1)

                        For k = 1 To q
                            pL(idx) = rawL
                            pA(idx) = rawA
                            pVeta(idx) = veta
                            pGiro(idx) = giro
                            idx = idx + 1
                        Next k
                    End If
                End If
            End If
        End If
    Next r

    Call AnidarPiezas(pL, pA, pVeta, pGiro, nP, utilL, utilA, _
                      nombreGrupo, tabLv, tabAv, exacto, pedir, imposibles)
End Sub


' =====================================================================
' AnidarPiezas: nesting por franjas (guillotina) respetando la veta
' =====================================================================
Private Sub AnidarPiezas(pL() As Double, pA() As Double, pVeta() As Long, _
                         pGiro() As Boolean, nP As Long, _
                         utilL As Double, utilA As Double, _
                         nombreGrupo As String, tabLv As Double, tabAv As Double, _
                         ByRef exacto As Double, ByRef pedir As Long, _
                         ByRef imposibles As Collection)

    Dim p As Long, o As Long
    Dim tmp As Double

    ' Piezas que pueden girar: normalizar (largo >= ancho) solo para ordenar
    For p = 0 To nP - 1
        If pGiro(p) And pA(p) > pL(p) Then
            tmp = pL(p): pL(p) = pA(p): pA(p) = tmp
        End If
    Next p

    ' Orientaciones permitidas de cada pieza: (ancho en X, alto en Y)
    ' El eje Y recorre el largo util del tablero (la veta del tablero)
    Dim nO() As Long, oW() As Double, oH() As Double
    ReDim nO(0 To nP - 1)
    ReDim oW(0 To nP - 1, 0 To 1)
    ReDim oH(0 To nP - 1, 0 To 1)

    For p = 0 To nP - 1
        Select Case pVeta(p)
            Case 1      ' veta en Largo: Largo alineado con el largo del tablero
                nO(p) = 1
                oW(p, 0) = pA(p): oH(p, 0) = pL(p)
            Case 2      ' veta en Ancho: pieza girada 90 grados de forma forzosa
                nO(p) = 1
                oW(p, 0) = pL(p): oH(p, 0) = pA(p)
            Case Else
                If pGiro(p) Then
                    nO(p) = 2
                    oW(p, 0) = pA(p): oH(p, 0) = pL(p)
                    oW(p, 1) = pL(p): oH(p, 1) = pA(p)
                Else
                    nO(p) = 1
                    oW(p, 0) = pA(p): oH(p, 0) = pL(p)
                End If
        End Select
    Next p

    ' Piezas imposibles: no caben en un tablero vacio en ninguna orientacion
    Dim pSkip() As Boolean
    ReDim pSkip(0 To nP - 1)
    For p = 0 To nP - 1
        pSkip(p) = True
        For o = 0 To nO(p) - 1
            If oW(p, o) <= utilA And oH(p, o) <= utilL Then
                pSkip(p) = False
                Exit For
            End If
        Next o
        If pSkip(p) Then
            imposibles.Add "NO CABE: pieza de " & Format$(pL(p), "0.##") & " x " & _
                           Format$(pA(p), "0.##") & " mm en tablero de " & _
                           Format$(tabLv, "0") & " x " & Format$(tabAv, "0") & _
                           "  (" & nombreGrupo & ")"
        End If
    Next p

    ' Ordenar de mayor a menor por alto de colocacion; en empate, por ancho
    Dim ii As Long, jj As Long
    Dim doSwap As Boolean
    Dim ki As Double, kj As Double
    Dim tL As Double, tA As Double, tV As Long
    Dim tG As Boolean, tS As Boolean, tN As Long
    Dim tW0 As Double, tW1 As Double, tH0 As Double, tH1 As Double

    For ii = 0 To nP - 2
        For jj = ii + 1 To nP - 1
            ki = oH(ii, 0): kj = oH(jj, 0)
            doSwap = False
            If ki < kj Then
                doSwap = True
            ElseIf ki = kj And oW(ii, 0) < oW(jj, 0) Then
                doSwap = True
            End If
            If doSwap Then
                tL = pL(ii): pL(ii) = pL(jj): pL(jj) = tL
                tA = pA(ii): pA(ii) = pA(jj): pA(jj) = tA
                tV = pVeta(ii): pVeta(ii) = pVeta(jj): pVeta(jj) = tV
                tG = pGiro(ii): pGiro(ii) = pGiro(jj): pGiro(jj) = tG
                tS = pSkip(ii): pSkip(ii) = pSkip(jj): pSkip(jj) = tS
                tN = nO(ii): nO(ii) = nO(jj): nO(jj) = tN
                tW0 = oW(ii, 0): oW(ii, 0) = oW(jj, 0): oW(jj, 0) = tW0
                tW1 = oW(ii, 1): oW(ii, 1) = oW(jj, 1): oW(jj, 1) = tW1
                tH0 = oH(ii, 0): oH(ii, 0) = oH(jj, 0): oH(jj, 0) = tH0
                tH1 = oH(ii, 1): oH(ii, 1) = oH(jj, 1): oH(jj, 1) = tH1
            End If
        Next jj
    Next ii

    ' --- Colocacion por franjas ---
    Dim tableros As Long      ' tableros completos cerrados
    Dim curX As Double        ' X ocupado en la franja actual
    Dim curY As Double        ' Y de inicio de la franja actual
    Dim shelfH As Double      ' alto de la franja actual
    Dim areaLast As Double    ' area de piezas en el tablero abierto
    tableros = 0: curX = 0: curY = 0: shelfH = 0: areaLast = 0

    Dim placed As Boolean
    Dim w As Double, h As Double, needX As Double, newH As Double, newY As Double

    For p = 0 To nP - 1
        If Not pSkip(p) Then
            placed = False

            ' Paso A: colocar en la franja actual
            For o = 0 To nO(p) - 1
                w = oW(p, o): h = oH(p, o)
                If curX = 0 Then needX = w Else needX = curX + SEP + w
                newH = shelfH: If h > newH Then newH = h
                If needX <= utilA And curY + newH <= utilL Then
                    curX = needX
                    shelfH = newH
                    areaLast = areaLast + pL(p) * pA(p)
                    placed = True
                    Exit For
                End If
            Next o

            ' Paso B: nueva franja en el mismo tablero
            If Not placed Then
                For o = 0 To nO(p) - 1
                    w = oW(p, o): h = oH(p, o)
                    newY = curY + shelfH + SEP
                    If newY + h <= utilL And w <= utilA Then
                        curY = newY
                        shelfH = h
                        curX = w
                        areaLast = areaLast + pL(p) * pA(p)
                        placed = True
                        Exit For
                    End If
                Next o
            End If

            ' Paso C: cerrar tablero y abrir uno nuevo
            If Not placed Then
                tableros = tableros + 1
                curX = 0: curY = 0: shelfH = 0: areaLast = 0
                For o = 0 To nO(p) - 1
                    w = oW(p, o): h = oH(p, o)
                    If w <= utilA And h <= utilL Then
                        curX = w
                        shelfH = h
                        areaLast = pL(p) * pA(p)
                        placed = True
                        Exit For
                    End If
                Next o
            End If
        End If
    Next p

    ' --- Resultado: exacto (1 decimal) y unidades a pedir ---
    Dim areaUtil As Double, dec As Double
    areaUtil = utilL * utilA
    dec = 0
    If areaLast > 0 Then
        dec = Int(areaLast / areaUtil * 10 + 0.5) / 10
        If dec < 0.1 Then dec = 0.1
        If dec > 1 Then dec = 1
    End If

    exacto = tableros + dec
    pedir = tableros + IIf(areaLast > 0, 1, 0)
End Sub


' =====================================================================
' ImportarCSVs: selecciona varios CSV del SIG y los combina en un
' libro nuevo (una sola cabecera). Devuelve la hoja combinada.
' =====================================================================
Private Function ImportarCSVs() As Worksheet
    Dim archivos As Variant
    archivos = Application.GetOpenFilename( _
        FileFilter:="Archivos CSV (*.csv), *.csv", _
        Title:="Selecciona los CSV del pedido (Ctrl+clic para varios)", _
        MultiSelect:=True)
    If Not IsArray(archivos) Then Exit Function    ' cancelado por el usuario

    Dim wb As Workbook
    Set wb = Workbooks.Add(xlWBATWorksheet)
    Dim ws As Worksheet
    Set ws = wb.Worksheets(1)
    On Error Resume Next
    ws.Name = "PEDIDO"
    On Error GoTo 0

    Dim outR As Long
    outR = 1
    Dim f As Long, ff As Integer
    Dim linea As String, campos As Variant, c As Long

    For f = LBound(archivos) To UBound(archivos)
        ff = FreeFile
        Open archivos(f) For Input As #ff
        Do While Not EOF(ff)
            Line Input #ff, linea
            linea = Replace$(linea, Chr$(13), "")
            If Trim$(Replace$(linea, ";", "")) <> "" Then
                campos = Split(linea, ";")
                If UBound(campos) + 1 >= C_A Then
                    If InStr(1, CStr(campos(0)), "Cod. Ensamblaje", vbTextCompare) > 0 Then
                        ' Cabecera: escribirla solo una vez
                        If outR = 1 Then
                            For c = 0 To UBound(campos)
                                ws.Cells(1, c + 1).Value = campos(c)
                            Next c
                            ws.Rows(1).Font.Bold = True
                            outR = 2
                        End If
                    Else
                        If outR = 1 Then outR = 2
                        For c = 0 To UBound(campos)
                            Select Case c + 1
                                Case C_QTY, C_L, C_A, C_GIRO
                                    ws.Cells(outR, c + 1).Value = ToNum(campos(c))
                                Case Else
                                    ws.Cells(outR, c + 1).Value = CStr(campos(c))
                            End Select
                        Next c
                        outR = outR + 1
                    End If
                End If
            End If
        Loop
        Close #ff
    Next f

    ws.Columns("A:S").AutoFit
    Set ImportarCSVs = ws
End Function


' =====================================================================
' LimpiarZonaSalida: borra tablas, formatos, validaciones y boton de
' una ejecucion anterior (todo lo que hay debajo de los datos)
' =====================================================================
Private Sub LimpiarZonaSalida(ws As Worksheet, dataLast As Long, matCount As Long)
    Dim hastaFila As Long
    hastaFila = dataLast + 2 * matCount + 500

    Dim zona As Range
    Set zona = ws.Rows(dataLast + 1 & ":" & hastaFila)
    On Error Resume Next
    zona.Validation.Delete
    ws.Buttons(BTN_NAME).Delete
    On Error GoTo 0
    zona.Clear
End Sub


' =====================================================================
' UltimaFilaDatos: ultima fila real de datos, ignorando las tablas de
' resultados de una ejecucion anterior
' =====================================================================
Private Function UltimaFilaDatos(ws As Worksheet) As Long
    Dim mRow As Long
    mRow = FilaMarcador(ws)

    Dim lastC As Long
    If mRow > 0 Then
        lastC = mRow - 1
        Do While lastC >= 2
            If Trim$(CStr(ws.Cells(lastC, C_MAT).Value)) <> "" Then Exit Do
            lastC = lastC - 1
        Loop
        If lastC < 2 Then lastC = 0
    Else
        lastC = ws.Cells(ws.Rows.Count, C_MAT).End(xlUp).Row
        If lastC < 2 Then lastC = 0
    End If
    UltimaFilaDatos = lastC
End Function


' FilaMarcador: fila del titulo MEDIDAS DE TABLERO (0 si no existe)
Private Function FilaMarcador(ws As Worksheet) As Long
    Dim f As Range
    Set f = ws.Columns(C_MAT).Find(What:=TIT_MEDIDAS, LookIn:=xlValues, _
                                   LookAt:=xlWhole, MatchCase:=False)
    If Not f Is Nothing Then FilaMarcador = f.Row
End Function


' =====================================================================
' LeerMedidasPrevias: recupera medidas y uniones de la tabla de una
' ejecucion anterior para no tener que reintroducirlas al recalcular
' =====================================================================
Private Sub LeerMedidasPrevias(ws As Worksheet, ByRef dL As Object, _
                               ByRef dA As Object, ByRef dQ As Object)
    Set dL = CreateObject("Scripting.Dictionary"): dL.CompareMode = vbTextCompare
    Set dA = CreateObject("Scripting.Dictionary"): dA.CompareMode = vbTextCompare
    Set dQ = CreateObject("Scripting.Dictionary"): dQ.CompareMode = vbTextCompare

    Dim mRow As Long
    mRow = FilaMarcador(ws)
    If mRow = 0 Then Exit Sub

    Dim r As Long, mat As String
    r = mRow + 2
    Do While Trim$(CStr(ws.Cells(r, COL_OUT).Value)) <> ""
        mat = Trim$(CStr(ws.Cells(r, COL_OUT).Value))
        If Not dL.Exists(mat) Then
            dL.Add mat, ToNum(ws.Cells(r, COL_OUT + 1).Value)
            dA.Add mat, ToNum(ws.Cells(r, COL_OUT + 2).Value)
            dQ.Add mat, Trim$(CStr(ws.Cells(r, COL_OUT + 3).Value))
        End If
        r = r + 1
    Loop
End Sub


' =====================================================================
' Utilidades
' =====================================================================

' ToNum: convierte a numero aceptando punto o coma decimal
Private Function ToNum(v As Variant) As Double
    If IsNumeric(v) Then
        ToNum = CDbl(v)
    Else
        ToNum = Val(Replace$(Trim$(CStr(v)), ",", "."))
    End If
End Function

' ParseEspesor: numero final del nombre del material (tras el ultimo guion)
Private Function ParseEspesor(mat As String) As Double
    Dim pos As Long
    pos = InStrRev(mat, "-")
    If pos > 0 And pos < Len(mat) Then
        ParseEspesor = Val(Replace$(Trim$(Mid$(mat, pos + 1)), ",", "."))
    End If
End Function

' NormalizarNombre: quita espacios y separadores para comparar nombres
Private Function NormalizarNombre(s As String) As String
    Dim t As String
    t = UCase$(Trim$(s))
    t = Replace$(t, " ", "")
    t = Replace$(t, "-", "")
    t = Replace$(t, "_", "")
    t = Replace$(t, ".", "")
    NormalizarNombre = t
End Function

' Levenshtein: numero de ediciones entre dos textos
Private Function Levenshtein(a As String, b As String) As Long
    Dim la As Long, lb As Long, i As Long, j As Long, costo As Long
    la = Len(a): lb = Len(b)
    If la = 0 Then Levenshtein = lb: Exit Function
    If lb = 0 Then Levenshtein = la: Exit Function

    Dim d() As Long
    ReDim d(0 To la, 0 To lb)
    For i = 0 To la: d(i, 0) = i: Next i
    For j = 0 To lb: d(0, j) = j: Next j

    Dim m As Long
    For i = 1 To la
        For j = 1 To lb
            If Mid$(a, i, 1) = Mid$(b, j, 1) Then costo = 0 Else costo = 1
            m = d(i - 1, j) + 1
            If d(i, j - 1) + 1 < m Then m = d(i, j - 1) + 1
            If d(i - 1, j - 1) + costo < m Then m = d(i - 1, j - 1) + costo
            d(i, j) = m
        Next j
    Next i
    Levenshtein = d(la, lb)
End Function

' Estilos
Private Sub EstiloTitulo(rng As Range, texto As String)
    rng.Interior.Color = RGB(31, 56, 100)
    rng.Font.Color = RGB(255, 255, 255)
    rng.Font.Bold = True
    rng.Cells(1, 1).Value = texto
End Sub

Private Sub EstiloCabecera(cel As Range, texto As String)
    With cel
        .Value = texto
        .Font.Bold = True
        .Interior.Color = RGB(46, 95, 163)
        .Font.Color = RGB(255, 255, 255)
    End With
End Sub

Private Sub Bordear(rng As Range)
    Dim i As Long
    For i = 7 To 10   ' xlEdgeLeft, xlEdgeTop, xlEdgeBottom, xlEdgeRight
        rng.Borders(i).LineStyle = xlContinuous
        rng.Borders(i).Color = RGB(150, 160, 175)
    Next i
    rng.Borders(xlInsideHorizontal).LineStyle = xlContinuous
    rng.Borders(xlInsideHorizontal).Color = RGB(200, 208, 220)
    If rng.Columns.Count > 1 Then
        rng.Borders(xlInsideVertical).LineStyle = xlContinuous
        rng.Borders(xlInsideVertical).Color = RGB(200, 208, 220)
    End If
End Sub
