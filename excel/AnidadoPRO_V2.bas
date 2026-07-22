Attribute VB_Name = "Mod_AnidadoPRO_V2"
' =====================================================================
' Macro:  AnidadoPRO_V2  -  Grupo KIMAK (Departamento de Compras)
' =====================================================================
'
' QUE HACE
' --------
' A partir del CSV de despiece exportado del SIG calcula, para un pedido
' de fabricacion:
'   1) Cuantos TABLEROS hay que pedir de cada material (nesting por
'      franjas / guillotina, respetando la direccion de la veta).
'   2) Cuantos METROS DE CANTO hay que pedir de cada material de canto
'      (a partir de la columna Canteado y las medidas de cada pieza).
'
' Novedad de la V2 respecto a la V1: el tercer cuadro "CANTOS A PEDIR".
'
' DOS MODOS DE USO (misma macro)
' ------------------------------
'   1) Abrir un CSV del SIG y ejecutar AnidadoPRO_V2.
'   2) Ejecutarla sin datos delante (libro vacio): abre el selector de
'      archivos de Windows, se eligen varios CSV a la vez, se combinan
'      en un libro nuevo y se calcula el pedido completo.
'
' TABLAS QUE GENERA (debajo de los datos)
' ---------------------------------------
'   - MEDIDAS DE TABLERO: medida editable por material (3050x1220 por
'     defecto) y columna MISMO QUE (desplegable) para unificar materiales
'     duplicados por errores de escritura. Posibles duplicados en amarillo.
'   - TABLEROS A PEDIR: material, medida, calculo exacto, unidades a pedir
'     (redondeadas hacia arriba) y M2 reales. Avisos rojo/naranja (abajo).
'     M2 reales = EXACTO x area del tablero completo (incluye piezas,
'     separaciones y orilla). El ultimo tablero parcial cuenta solo su
'     fraccion (0.1 -> 0.1 x area), no un tablero entero.
'   - CANTOS A PEDIR: material de canto, MISMO QUE (desplegable para
'     unir cantos iguales escritos distinto), metros exactos y metros a
'     pedir (con el extra de seguridad). Posibles duplicados en amarillo.
'   Todas se crean como tablas de Excel con filtros para ordenar.
'   Un unico boton RECALCULAR recalcula los tres cuadros a la vez.
'
' REGLAS DE VETA (nesting de tableros)
' ------------------------------------
'   - Dir. Veta = "Largo", "Corto" o cualquier otro texto: la pieza NO
'     gira; la medida de la columna Largo (donde va la veta) se alinea
'     con el largo del tablero. El SIG escribe "Corto" cuando la veta va
'     en el largo pero el largo es la medida pequena de la pieza.
'   - Dir. Veta = "Ancho": la pieza NO gira; se coloca girada 90 grados.
'   - Dir. Veta vacia o "0": sin veta; gira solo si Permite giro = 1.
'   - Si una fila trae veta y Permite giro = 1, la veta manda: no gira.
'
' PIEZAS GRANDES (nesting de tableros)
' ------------------------------------
'   - Pieza que no cabe en el area util (con orillas) pero si en el
'     tablero bruto (3050x1220, 3050x200, 100x1220...): se anida aparte
'     usando el tablero completo, cuenta en los tableros a pedir y se
'     avisa en NARANJA con su Cod. Pieza (va sin margen de orilla).
'   - Pieza que no cabe ni en el tablero bruto: aviso en ROJO, no cuenta.
'
' CALCULO DE CANTO (metros)
' -------------------------
'   Columna Canteado (ej: "1L", "2C 2L", "1C 2L", "0", vacio):
'     nL = numero de cantos en el lado LARGO; nC = en el lado CORTO.
'   Lado largo  = la medida MAYOR de la pieza  (max de Largo y Ancho).
'   Lado corto  = la medida MENOR de la pieza  (min de Largo y Ancho).
'   Metros por pieza (exacto) = nL*ladoLargo + nC*ladoCorto.
'   Extra de seguridad = EXTRA_CANTO_MM por CADA canto (borde).
'     -> a pedir = exacto + EXTRA_CANTO_MM * (nL + nC), por cada pieza.
'   El calculo no depende del nesting (es propio de la pieza) y se cuenta
'   aunque la pieza no quepa en tablero.
'
'   Material del canto: SIEMPRE el material del tablero sin el espesor
'     (ej: "AGLOMERADO-S/ORDEN-16" -> "AGLOMERADO-S/ORDEN"). Si la fila
'     tiene "Comentarios Canteado", ese texto se anade entre parentesis y
'     la celda se pinta de NARANJA para revisarla
'     (ej: "AGLOMERADO-S/ORDEN (CANTO PVC BLANCO)").
'   La columna Ingletado NO influye en el canto (se ignora).
'
' MATERIALES IGUALES ESCRITOS DISTINTO
' ------------------------------------
'   La macro considera automaticamente el mismo material los nombres que
'   solo difieren en guiones o espacios de mas: "MDF--16", "MDF---16" y
'   "MDF   --16" se agrupan como uno solo (ver ClaveMaterial). Para
'   diferencias mayores (una palabra escrita de otra forma) se marca el
'   posible duplicado en amarillo y se une a mano con MISMO QUE.
'
' =====================================================================
Option Explicit

' =====================================================================
'   PARAMETROS QUE PUEDES CAMBIAR
'   (cambia aqui los valores; el resto de la macro se ajusta solo)
' =====================================================================
Private Const DEF_TAB_L      As Double = 3050  ' Largo tablero estandar (mm)
Private Const DEF_TAB_A      As Double = 1220  ' Ancho tablero estandar (mm)
Private Const MARGEN         As Double = 12    ' Orilla en los 4 lados (mm)
Private Const SEP            As Double = 17    ' Separacion entre piezas y franjas (mm)
                                               '   (12 mm libres + 5 mm de kerf del disco)
Private Const EXTRA_CANTO_MM As Double = 20    ' Extra de canto por cada borde (mm)

' --- Columnas del CSV del SIG (1=A, 2=B, 3=C...) ---
Private Const C_MAT       As Long = 3     ' C: Mat. - recubr. - esp.
Private Const C_COD       As Long = 5     ' E: Cod. Pieza
Private Const C_QTY       As Long = 8     ' H: Cantidad
Private Const C_L         As Long = 9     ' I: Largo (mm, direccion de la veta)
Private Const C_A         As Long = 10    ' J: Ancho (mm)
Private Const C_GIRO      As Long = 11    ' K: Permite giro (1 = si)
Private Const C_VETA      As Long = 13    ' M: Dir. Veta ("Largo"/"Corto"/"Ancho"/vacio)
Private Const C_CANTO     As Long = 14    ' N: Canteado ("1L", "2C 2L"...)
Private Const C_COM_CANTO As Long = 15    ' O: Comentarios Canteado (material del canto)

' --- Salida ---
Private Const TIT_MEDIDAS As String = "MEDIDAS DE TABLERO"
Private Const TIT_RESULT  As String = "TABLEROS A PEDIR"
Private Const TIT_CANTOS  As String = "CANTOS A PEDIR"
Private Const COL_OUT     As Long = 3     ' Columna C: primera columna de las tablas
Private Const COL_LISTA   As Long = 21    ' Columna U (oculta): lista comun de desplegables
Private Const COL_SUG     As Long = 30    ' Columna AD+ (ocultas): listas por fila con los
                                          ' candidatos a duplicado colocados al principio
Private Const BTN_NAME    As String = "btnRecalcularV2"


' =====================================================================
' Macro principal
' =====================================================================
Public Sub AnidadoPRO_V2()
    Dim ws As Worksheet
    Set ws = ActiveSheet

    Dim dataLast As Long
    dataLast = UltimaFilaDatos(ws)

    ' --- Sin datos delante: modo importacion de varios CSV ---
    If dataLast < 2 Then
        If MsgBox("La hoja activa no tiene datos de piezas." & vbNewLine & vbNewLine & _
                  "Quieres seleccionar uno o varios CSV del SIG para importarlos " & _
                  "juntos y calcular el pedido completo?", _
                  vbYesNo + vbQuestion, "AnidadoPRO_V2") <> vbYes Then Exit Sub
        Set ws = ImportarCSVs()
        If ws Is Nothing Then Exit Sub
        dataLast = UltimaFilaDatos(ws)
        If dataLast < 2 Then
            MsgBox "Los archivos seleccionados no contienen piezas.", _
                   vbExclamation, "AnidadoPRO_V2"
            Exit Sub
        End If
    End If

    On Error GoTo Fallo
    Application.ScreenUpdating = False

    ' -----------------------------------------------------------------
    ' PASO 1: leer estado previo (antes de limpiar la zona de resultados)
    ' -----------------------------------------------------------------
    Dim prevL As Object, prevA As Object, prevQ As Object
    Call LeerMedidasPrevias(ws, prevL, prevA, prevQ)
    Dim prevCQ As Object
    Call LeerUnionesCanto(ws, prevCQ)

    ' -----------------------------------------------------------------
    ' PASO 2: materiales de tablero unicos, en orden de aparicion
    ' -----------------------------------------------------------------
    Dim mats As Object
    Set mats = CreateObject("Scripting.Dictionary")
    mats.CompareMode = vbTextCompare
    Dim matNames() As String, matCount As Long
    matCount = 0

    ' La clave de agrupacion (ClaveMaterial) ignora guiones y espacios de
    ' mas, asi "MDF--16" y "MDF---16" caen en el mismo grupo. Como nombre a
    ' mostrar se usa la primera grafia encontrada.
    Dim r As Long, mat As String
    For r = 2 To dataLast
        mat = Trim$(CStr(ws.Cells(r, C_MAT).Value))
        If mat <> "" Then
            If Not mats.Exists(ClaveMaterial(mat)) Then
                ReDim Preserve matNames(0 To matCount)
                matNames(matCount) = mat
                mats.Add ClaveMaterial(mat), matCount
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

    Dim i As Long, medidasMalas As String
    For i = 0 To matCount - 1
        tabL(i) = DEF_TAB_L
        tabA(i) = DEF_TAB_A
        mismoQue(i) = ""
        If prevL.Exists(ClaveMaterial(matNames(i))) Then
            If prevL(ClaveMaterial(matNames(i))) > 0 Then tabL(i) = prevL(ClaveMaterial(matNames(i)))
            If prevA(ClaveMaterial(matNames(i))) > 0 Then tabA(i) = prevA(ClaveMaterial(matNames(i)))
            mismoQue(i) = prevQ(ClaveMaterial(matNames(i)))
        End If
        If mismoQue(i) <> "" Then
            If Not mats.Exists(ClaveMaterial(mismoQue(i))) Then
                mismoQue(i) = ""
            ElseIf mats(ClaveMaterial(mismoQue(i))) = i Then
                mismoQue(i) = ""
            End If
        End If
        If tabL(i) <= 2 * MARGEN + 10 Or tabA(i) <= 2 * MARGEN + 10 Then
            medidasMalas = medidasMalas & vbNewLine & "  - " & matNames(i)
            tabL(i) = DEF_TAB_L
            tabA(i) = DEF_TAB_A
        End If
    Next i

    If medidasMalas <> "" Then
        MsgBox "Estas medidas de tablero no eran validas y se han " & _
               "restablecido a " & DEF_TAB_L & " x " & DEF_TAB_A & ":" & _
               vbNewLine & medidasMalas, vbExclamation, "AnidadoPRO_V2"
    End If

    ' Confirmar uniones entre espesores distintos (casi siempre error de dedo)
    Dim e1 As Double, e2 As Double
    For i = 0 To matCount - 1
        If mismoQue(i) <> "" Then
            e1 = ParseEspesor(matNames(i))
            e2 = ParseEspesor(mismoQue(i))
            If e1 > 0 And e2 > 0 And e1 <> e2 Then
                If MsgBox("Vas a unir dos materiales con ESPESOR DISTINTO:" & _
                          vbNewLine & vbNewLine & matNames(i) & "  (espesor " & e1 & ")" & _
                          vbNewLine & "con" & vbNewLine & mismoQue(i) & "  (espesor " & e2 & ")" & _
                          vbNewLine & vbNewLine & "Seguro que son el mismo material?", _
                          vbYesNo + vbExclamation, "Union de materiales") <> vbYes Then
                    mismoQue(i) = ""
                End If
            End If
        End If
    Next i

    ' Resolver cadenas de uniones y detectar circulos; marcar duplicados
    Dim destino() As Long, sospechoso() As Boolean, parejas() As String
    Dim ciclos As String
    Call ResolverUniones(matNames, matCount, mats, mismoQue, True, _
                         destino, sospechoso, parejas, ciclos)
    If ciclos <> "" Then AvisoCiclos ciclos

    ' -----------------------------------------------------------------
    ' PASO 4: nesting por grupo de materiales unificados
    ' -----------------------------------------------------------------
    Dim exactoG() As Double, pedirG() As Long
    ReDim exactoG(0 To matCount - 1)
    ReDim pedirG(0 To matCount - 1)
    Dim imposibles As New Collection
    Dim avisos As New Collection

    Dim t As Long
    For t = 0 To matCount - 1
        If destino(t) = t Then
            Call CalcularGrupo(ws, dataLast, mats, destino, t, matNames(t), _
                               tabL(t), tabA(t), exactoG(t), pedirG(t), imposibles, avisos)
        End If
    Next t

    ' -----------------------------------------------------------------
    ' PASO 5: materiales de canto, uniones y metros
    ' -----------------------------------------------------------------
    Dim cantos As Object
    Set cantos = CreateObject("Scripting.Dictionary")
    cantos.CompareMode = vbTextCompare
    Dim cantoNames() As String, cantoNaranja() As Boolean, cantoCount As Long
    cantoCount = 0

    Dim nLc As Long, nCc As Long, cmat As String
    For r = 2 To dataLast
        If Trim$(CStr(ws.Cells(r, C_MAT).Value)) <> "" Then
            Call ParseCanteado(CStr(ws.Cells(r, C_CANTO).Value), nLc, nCc)
            If nLc + nCc > 0 Then
                If ToNum(ws.Cells(r, C_L).Value) > 0 And ToNum(ws.Cells(r, C_A).Value) > 0 Then
                    cmat = CantoMaterial(ws, r)
                    If cmat <> "" And Not cantos.Exists(ClaveMaterial(cmat)) Then
                        ReDim Preserve cantoNames(0 To cantoCount)
                        ReDim Preserve cantoNaranja(0 To cantoCount)
                        cantoNames(cantoCount) = cmat
                        cantoNaranja(cantoCount) = CantoTieneComentario(ws, r)
                        cantos.Add ClaveMaterial(cmat), cantoCount
                        cantoCount = cantoCount + 1
                    End If
                End If
            End If
        End If
    Next r

    Dim mismoQueC() As String, destinoC() As Long
    Dim sospechosoC() As Boolean, parejasC() As String, ciclosC As String
    Dim exactMM() As Double, pedirMM() As Double
    Dim c As Long

    If cantoCount > 0 Then
        ReDim mismoQueC(0 To cantoCount - 1)
        For c = 0 To cantoCount - 1
            mismoQueC(c) = ""
            If prevCQ.Exists(ClaveMaterial(cantoNames(c))) Then _
                mismoQueC(c) = prevCQ(ClaveMaterial(cantoNames(c)))
            If mismoQueC(c) <> "" Then
                If Not cantos.Exists(ClaveMaterial(mismoQueC(c))) Then
                    mismoQueC(c) = ""
                ElseIf cantos(ClaveMaterial(mismoQueC(c))) = c Then
                    mismoQueC(c) = ""
                End If
            End If
        Next c

        Call ResolverUniones(cantoNames, cantoCount, cantos, mismoQueC, False, _
                             destinoC, sospechosoC, parejasC, ciclosC)
        If ciclosC <> "" Then AvisoCiclos ciclosC

        ' Acumular metros de canto en el destino de cada material de canto
        ReDim exactMM(0 To cantoCount - 1)
        ReDim pedirMM(0 To cantoCount - 1)
        Dim q As Long, rawL As Double, rawA As Double
        Dim ladoL As Double, ladoC As Double, tt As Double, cdst As Long
        For r = 2 To dataLast
            If Trim$(CStr(ws.Cells(r, C_MAT).Value)) <> "" Then
                Call ParseCanteado(CStr(ws.Cells(r, C_CANTO).Value), nLc, nCc)
                If nLc + nCc > 0 Then
                    rawL = ToNum(ws.Cells(r, C_L).Value)
                    rawA = ToNum(ws.Cells(r, C_A).Value)
                    If rawL > 0 And rawA > 0 Then
                        q = CLng(ToNum(ws.Cells(r, C_QTY).Value))
                        If q <= 0 Then q = 1
                        cmat = CantoMaterial(ws, r)
                        If cantos.Exists(ClaveMaterial(cmat)) Then
                            cdst = destinoC(cantos(ClaveMaterial(cmat)))
                            ladoL = rawL: ladoC = rawA
                            If ladoC > ladoL Then tt = ladoL: ladoL = ladoC: ladoC = tt
                            exactMM(cdst) = exactMM(cdst) + (nLc * ladoL + nCc * ladoC) * q
                            pedirMM(cdst) = pedirMM(cdst) + _
                                (nLc * ladoL + nCc * ladoC) * q + EXTRA_CANTO_MM * (nLc + nCc) * q
                        End If
                    End If
                End If
            End If
        Next r
    End If

    ' -----------------------------------------------------------------
    ' PASO 6: escribir las tablas en la hoja
    ' -----------------------------------------------------------------
    Call LimpiarZonaSalida(ws, dataLast, matCount + cantoCount)

    Dim r0 As Long
    r0 = dataLast + 3
    Call AsegurarBotonRecalcular(ws, r0)   ' primero: nunca deja la hoja sin boton

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
                .Interior.Color = RGB(255, 235, 156)
            ElseIf i Mod 2 = 1 Then
                .Interior.Color = RGB(242, 246, 252)
            End If
        End With
        With ws.Cells(fila, COL_OUT + 1)
            .Value = tabL(i): .NumberFormat = "0"
            If i Mod 2 = 1 Then .Interior.Color = RGB(242, 246, 252)
        End With
        With ws.Cells(fila, COL_OUT + 2)
            .Value = tabA(i): .NumberFormat = "0"
            If i Mod 2 = 1 Then .Interior.Color = RGB(242, 246, 252)
        End With
        With ws.Cells(fila, COL_OUT + 3)
            .Value = mismoQue(i)
            If i Mod 2 = 1 Then .Interior.Color = RGB(242, 246, 252)
        End With
        ws.Cells(fila, COL_LISTA).Value = matNames(i)
    Next i

    Dim listaComun As String
    listaComun = "=" & ws.Cells(r0 + 2, COL_LISTA).Address & ":" & _
                       ws.Cells(r0 + 1 + matCount, COL_LISTA).Address
    For i = 0 To matCount - 1
        Call PonerDesplegable(ws, r0 + 2 + i, COL_OUT + 3, matNames, matCount, i, _
                              parejas(i), listaComun)
    Next i

    Call Bordear(ws.Range(ws.Cells(r0 + 1, COL_OUT), ws.Cells(r0 + 1 + matCount, COL_OUT + 3)))
    Call ConvertirEnTabla(ws, r0 + 1, r0 + 1 + matCount, "TablaMedidasTableroV2")

    ' ---- Tabla TABLEROS A PEDIR ----
    Dim rRes As Long
    rRes = r0 + matCount + 4
    Call EstiloTitulo(ws.Range(ws.Cells(rRes, COL_OUT), ws.Cells(rRes, COL_OUT + 4)), TIT_RESULT)
    Call EstiloCabecera(ws.Cells(rRes + 1, COL_OUT), "MATERIAL")
    Call EstiloCabecera(ws.Cells(rRes + 1, COL_OUT + 1), "MEDIDA TABLERO")
    Call EstiloCabecera(ws.Cells(rRes + 1, COL_OUT + 2), "EXACTO")
    Call EstiloCabecera(ws.Cells(rRes + 1, COL_OUT + 3), "A PEDIR")
    Call EstiloCabecera(ws.Cells(rRes + 1, COL_OUT + 4), "M2 REALES")

    Dim totalPedir As Long, nGrupos As Long
    Dim totalM2 As Double, m2 As Double
    totalPedir = 0: nGrupos = 0: totalM2 = 0
    fila = rRes + 1
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
                .Value = exactoG(t): .NumberFormat = "0.0"
            End With
            With ws.Cells(fila, COL_OUT + 3)
                .Value = pedirG(t): .NumberFormat = "0"
                .Font.Bold = True: .Font.Size = 12
                .HorizontalAlignment = xlCenter
                .Interior.Color = RGB(226, 239, 218)
            End With
            ' M2 reales = tableros EXACTOS x area del tablero completo (incluye
            ' piezas, separaciones y orilla). El ultimo tablero parcial cuenta
            ' solo su fraccion (0.1 -> 0.1 x area), no un tablero entero.
            m2 = exactoG(t) * tabL(t) * tabA(t) / 1000000#
            With ws.Cells(fila, COL_OUT + 4)
                .Value = m2: .NumberFormat = "0.00"
                If nGrupos Mod 2 = 0 Then .Interior.Color = RGB(242, 246, 252)
            End With
            totalPedir = totalPedir + pedirG(t)
            totalM2 = totalM2 + m2
        End If
    Next t

    fila = fila + 1
    With ws.Range(ws.Cells(fila, COL_OUT), ws.Cells(fila, COL_OUT + 4))
        .Interior.Color = RGB(31, 56, 100)
        .Font.Color = RGB(255, 255, 255): .Font.Bold = True
    End With
    ws.Cells(fila, COL_OUT).Value = "TOTAL TABLEROS"
    With ws.Cells(fila, COL_OUT + 3)
        .Value = totalPedir: .NumberFormat = "0"
        .Font.Size = 12: .HorizontalAlignment = xlCenter
    End With
    With ws.Cells(fila, COL_OUT + 4)
        .Value = totalM2: .NumberFormat = "0.00"
        .Font.Size = 12: .HorizontalAlignment = xlCenter
    End With

    Call Bordear(ws.Range(ws.Cells(rRes + 1, COL_OUT), ws.Cells(fila, COL_OUT + 4)))
    If nGrupos > 0 Then Call ConvertirEnTabla(ws, rRes + 1, fila - 1, "TablaTablerosAPedirV2", COL_OUT + 4)

    ' ---- Avisos: rojo = no cabe; naranja = usa el tablero sin margen ----
    Dim k As Long
    If imposibles.Count > 0 Then
        fila = fila + 2
        For k = 1 To imposibles.Count
            With ws.Range(ws.Cells(fila, COL_OUT), ws.Cells(fila, COL_OUT + 3))
                .Interior.Color = RGB(192, 0, 0)
                .Font.Color = RGB(255, 255, 255): .Font.Bold = True
            End With
            ws.Cells(fila, COL_OUT).Value = imposibles(k)
            fila = fila + 1
        Next k
    End If
    If avisos.Count > 0 Then
        If imposibles.Count = 0 Then fila = fila + 2
        For k = 1 To avisos.Count
            With ws.Range(ws.Cells(fila, COL_OUT), ws.Cells(fila, COL_OUT + 3))
                .Interior.Color = RGB(237, 125, 49)
                .Font.Color = RGB(255, 255, 255): .Font.Bold = True
            End With
            ws.Cells(fila, COL_OUT).Value = avisos(k)
            fila = fila + 1
        Next k
    End If

    ' ---- Tabla CANTOS A PEDIR ----
    If cantoCount > 0 Then
        Dim rCanto As Long
        rCanto = fila + 3
        Call EstiloTitulo(ws.Range(ws.Cells(rCanto, COL_OUT), ws.Cells(rCanto, COL_OUT + 3)), TIT_CANTOS)
        Call EstiloCabecera(ws.Cells(rCanto + 1, COL_OUT), "MATERIAL CANTO")
        Call EstiloCabecera(ws.Cells(rCanto + 1, COL_OUT + 1), "METROS EXACTOS")
        Call EstiloCabecera(ws.Cells(rCanto + 1, COL_OUT + 2), "METROS A PEDIR")
        Call EstiloCabecera(ws.Cells(rCanto + 1, COL_OUT + 3), "MISMO QUE")

        Dim totalMetros As Double
        totalMetros = 0
        For c = 0 To cantoCount - 1
            fila = rCanto + 2 + c
            With ws.Cells(fila, COL_OUT)
                .Value = cantoNames(c)
                If cantoNaranja(c) Then
                    .Interior.Color = RGB(255, 192, 0)     ' naranja: tiene comentario, revisar
                ElseIf sospechosoC(c) Then
                    .Interior.Color = RGB(255, 235, 156)   ' amarillo: posible duplicado
                ElseIf c Mod 2 = 1 Then
                    .Interior.Color = RGB(242, 246, 252)
                End If
            End With
            If destinoC(c) = c Then
                With ws.Cells(fila, COL_OUT + 1)
                    .Value = exactMM(c) / 1000: .NumberFormat = "0.00"
                    If c Mod 2 = 1 Then .Interior.Color = RGB(242, 246, 252)
                End With
                With ws.Cells(fila, COL_OUT + 2)
                    .Value = pedirMM(c) / 1000: .NumberFormat = "0.00"
                    .Font.Bold = True: .Font.Size = 12
                    .HorizontalAlignment = xlCenter
                    .Interior.Color = RGB(226, 239, 218)
                End With
                totalMetros = totalMetros + pedirMM(c) / 1000
            Else
                If c Mod 2 = 1 Then
                    ws.Range(ws.Cells(fila, COL_OUT + 1), ws.Cells(fila, COL_OUT + 2)) _
                      .Interior.Color = RGB(242, 246, 252)
                End If
            End If
            With ws.Cells(fila, COL_OUT + 3)
                .Value = mismoQueC(c)
                If c Mod 2 = 1 Then .Interior.Color = RGB(242, 246, 252)
            End With
            ws.Cells(fila, COL_LISTA).Value = cantoNames(c)
        Next c

        Dim listaComunC As String
        listaComunC = "=" & ws.Cells(rCanto + 2, COL_LISTA).Address & ":" & _
                            ws.Cells(rCanto + 1 + cantoCount, COL_LISTA).Address
        For c = 0 To cantoCount - 1
            Call PonerDesplegable(ws, rCanto + 2 + c, COL_OUT + 3, cantoNames, cantoCount, c, _
                                  parejasC(c), listaComunC)
        Next c

        ' Fila de total (metros a pedir)
        fila = rCanto + 2 + cantoCount
        With ws.Range(ws.Cells(fila, COL_OUT), ws.Cells(fila, COL_OUT + 3))
            .Interior.Color = RGB(31, 56, 100)
            .Font.Color = RGB(255, 255, 255): .Font.Bold = True
        End With
        ws.Cells(fila, COL_OUT).Value = "TOTAL METROS"
        With ws.Cells(fila, COL_OUT + 2)
            .Value = totalMetros: .NumberFormat = "0.00"
            .Font.Size = 12: .HorizontalAlignment = xlCenter
        End With

        Call Bordear(ws.Range(ws.Cells(rCanto + 1, COL_OUT), ws.Cells(fila, COL_OUT + 3)))
        Call ConvertirEnTabla(ws, rCanto + 1, rCanto + 1 + cantoCount, "TablaCantosAPedirV2")
    End If

    ' ---- Ocultar columnas auxiliares de los desplegables ----
    Dim maxN As Long
    maxN = matCount
    If cantoCount > maxN Then maxN = cantoCount
    ws.Columns(COL_LISTA).Hidden = True
    If maxN > 0 Then
        ws.Range(ws.Columns(COL_SUG), ws.Columns(COL_SUG + maxN - 1)).Hidden = True
    End If

    ' Anchos minimos para que se lea bien
    If ws.Columns(COL_OUT).ColumnWidth < 48 Then ws.Columns(COL_OUT).ColumnWidth = 48
    If ws.Columns(COL_OUT + 1).ColumnWidth < 15 Then ws.Columns(COL_OUT + 1).ColumnWidth = 15
    If ws.Columns(COL_OUT + 2).ColumnWidth < 15 Then ws.Columns(COL_OUT + 2).ColumnWidth = 15
    If ws.Columns(COL_OUT + 3).ColumnWidth < 48 Then ws.Columns(COL_OUT + 3).ColumnWidth = 48
    If ws.Columns(COL_OUT + 4).ColumnWidth < 12 Then ws.Columns(COL_OUT + 4).ColumnWidth = 12

    Application.GoTo ws.Cells(r0, 1), True
    Application.ScreenUpdating = True
    Exit Sub

Fallo:
    Application.ScreenUpdating = True
    MsgBox "Error inesperado: " & Err.Description, vbCritical, "AnidadoPRO_V2"
End Sub


' =====================================================================
' ResolverUniones: a partir de mismoQue() (ya validado por el llamante)
' resuelve las cadenas A->B->C (destino final), detecta circulos y marca
' los posibles duplicados (nombres casi iguales) con sus candidatos.
'   conEspesor = True  -> solo marca duplicados con el mismo espesor
'                         (materiales de tablero; -16 y -19 no se mezclan)
'   conEspesor = False -> marca por parecido de nombre (materiales de canto)
' =====================================================================
Private Sub ResolverUniones(nombres() As String, n As Long, dict As Object, _
                            mismoQ() As String, conEspesor As Boolean, _
                            ByRef destino() As Long, ByRef sospechoso() As Boolean, _
                            ByRef parejas() As String, ByRef ciclos As String)
    ReDim destino(0 To n - 1)
    ReDim sospechoso(0 To n - 1)
    ReDim parejas(0 To n - 1)
    ciclos = ""

    ' Cadenas de uniones y deteccion de circulos
    Dim i As Long, j As Long, pasos As Long
    For i = 0 To n - 1
        j = i: pasos = 0
        Do While mismoQ(j) <> "" And pasos <= n
            j = dict(ClaveMaterial(mismoQ(j)))
            pasos = pasos + 1
        Loop
        If pasos > n Then
            destino(i) = i
            ciclos = ciclos & vbNewLine & "  - " & nombres(i)
        Else
            destino(i) = j
        End If
    Next i

    ' Posibles duplicados (solo entre los que aun no estan unidos)
    Dim e1 As Double, e2 As Double
    For i = 0 To n - 2
        For j = i + 1 To n - 1
            If mismoQ(i) = "" And mismoQ(j) = "" Then
                If conEspesor Then
                    e1 = ParseEspesor(nombres(i)): e2 = ParseEspesor(nombres(j))
                    If e1 <> e2 Then GoTo SiguienteJ
                End If
                If SonParecidos(nombres(i), nombres(j)) Then
                    sospechoso(i) = True: sospechoso(j) = True
                    parejas(i) = parejas(i) & j & ";"
                    parejas(j) = parejas(j) & i & ";"
                End If
            End If
SiguienteJ:
        Next j
    Next i
End Sub


' PonerDesplegable: crea el desplegable MISMO QUE en una celda. Si la fila
' tiene candidatos a duplicado, arma una lista propia (columnas ocultas
' COL_SUG+) con esos candidatos al principio; si no, usa la lista comun.
Private Sub PonerDesplegable(ws As Worksheet, fila As Long, colCelda As Long, _
                             nombres() As String, n As Long, idxSelf As Long, _
                             parejasStr As String, listaComun As String)
    Dim formulaLista As String
    If parejasStr <> "" Then
        Dim nSug As Long, jj As Long, kk As Long, yaEsta As Boolean
        Dim idxPar() As String
        nSug = 0
        idxPar = Split(parejasStr, ";")
        For jj = 0 To UBound(idxPar) - 1     ' ultimo elemento vacio por el ; final
            ws.Cells(fila, COL_SUG + nSug).Value = nombres(CLng(idxPar(jj)))
            nSug = nSug + 1
        Next jj
        For jj = 0 To n - 1
            If jj <> idxSelf Then
                yaEsta = False
                For kk = 0 To UBound(idxPar) - 1
                    If CLng(idxPar(kk)) = jj Then yaEsta = True: Exit For
                Next kk
                If Not yaEsta Then
                    ws.Cells(fila, COL_SUG + nSug).Value = nombres(jj)
                    nSug = nSug + 1
                End If
            End If
        Next jj
        formulaLista = "=" & ws.Range(ws.Cells(fila, COL_SUG), _
                                      ws.Cells(fila, COL_SUG + nSug - 1)).Address
    Else
        formulaLista = listaComun
    End If

    With ws.Cells(fila, colCelda).Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:=formulaLista
        .IgnoreBlank = True
        .InCellDropdown = True
        .ErrorMessage = "Elige un valor de la lista o deja la celda vacia."
    End With
End Sub


' ConvertirEnTabla: convierte un rango en tabla de Excel con filtros.
' colFin es la ultima columna; si se omite, COL_OUT+3 (4 columnas).
Private Sub ConvertirEnTabla(ws As Worksheet, filaCab As Long, filaFin As Long, _
                             nombre As String, Optional colFin As Long = -1)
    If colFin < 0 Then colFin = COL_OUT + 3
    On Error Resume Next
    Dim lo As ListObject
    Set lo = ws.ListObjects.Add(xlSrcRange, _
        ws.Range(ws.Cells(filaCab, COL_OUT), ws.Cells(filaFin, colFin)), , xlYes)
    lo.Name = nombre
    lo.TableStyle = ""
    On Error GoTo 0
End Sub


' AvisoCiclos: mensaje cuando hay referencias circulares en MISMO QUE
Private Sub AvisoCiclos(ciclos As String)
    MsgBox "Hay referencias circulares en la columna MISMO QUE " & _
           "(un elemento apunta a otro que apunta de vuelta al primero)." & _
           vbNewLine & "Se han ignorado las uniones de:" & vbNewLine & ciclos & _
           vbNewLine & vbNewLine & "Revisa los desplegables y pulsa RECALCULAR.", _
           vbExclamation, "AnidadoPRO_V2"
End Sub


' =====================================================================
' CalcularGrupo: recoge las piezas de todos los materiales unificados
' en el destino t y las anida. Devuelve exacto, a pedir e imposibles.
' =====================================================================
Private Sub CalcularGrupo(ws As Worksheet, dataLast As Long, mats As Object, _
                          destino() As Long, t As Long, nombreGrupo As String, _
                          tabLv As Double, tabAv As Double, _
                          ByRef exacto As Double, ByRef pedir As Long, _
                          ByRef imposibles As Collection, ByRef avisos As Collection)

    Dim utilL As Double, utilA As Double
    utilL = tabLv - 2 * MARGEN
    utilA = tabAv - 2 * MARGEN

    Dim nP As Long
    nP = 0
    Dim r As Long, mat As String, q As Long
    Dim rawL As Double, rawA As Double

    For r = 2 To dataLast
        mat = Trim$(CStr(ws.Cells(r, C_MAT).Value))
        If mat <> "" Then
            If mats.Exists(ClaveMaterial(mat)) Then
                If destino(mats(ClaveMaterial(mat))) = t Then
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

    exacto = 0: pedir = 0
    If nP = 0 Then Exit Sub

    Dim pL() As Double, pA() As Double, pVeta() As Long, pGiro() As Boolean, pCod() As String
    ReDim pL(0 To nP - 1)
    ReDim pA(0 To nP - 1)
    ReDim pVeta(0 To nP - 1)
    ReDim pGiro(0 To nP - 1)
    ReDim pCod(0 To nP - 1)

    Dim idx As Long, k As Long
    Dim vetaTxt As String, veta As Long
    Dim giro As Boolean
    idx = 0

    For r = 2 To dataLast
        mat = Trim$(CStr(ws.Cells(r, C_MAT).Value))
        If mat <> "" Then
            If mats.Exists(ClaveMaterial(mat)) Then
                If destino(mats(ClaveMaterial(mat))) = t Then
                    rawL = ToNum(ws.Cells(r, C_L).Value)
                    rawA = ToNum(ws.Cells(r, C_A).Value)
                    If rawL > 0 And rawA > 0 Then
                        q = CLng(ToNum(ws.Cells(r, C_QTY).Value))
                        If q <= 0 Then q = 1

                        vetaTxt = LCase$(Trim$(CStr(ws.Cells(r, C_VETA).Value)))
                        If vetaTxt = "" Or vetaTxt = "0" Then
                            veta = 0
                        ElseIf InStr(vetaTxt, "anch") > 0 Then
                            veta = 2
                        Else
                            veta = 1
                        End If
                        giro = (veta = 0) And (ToNum(ws.Cells(r, C_GIRO).Value) = 1)

                        For k = 1 To q
                            pL(idx) = rawL
                            pA(idx) = rawA
                            pVeta(idx) = veta
                            pGiro(idx) = giro
                            pCod(idx) = Trim$(CStr(ws.Cells(r, C_COD).Value))
                            idx = idx + 1
                        Next k
                    End If
                End If
            End If
        End If
    Next r

    Call AnidarPiezas(pL, pA, pVeta, pGiro, pCod, nP, utilL, utilA, _
                      nombreGrupo, tabLv, tabAv, exacto, pedir, imposibles, avisos)
End Sub


' =====================================================================
' AnidarPiezas: nesting por franjas (guillotina) respetando la veta
' =====================================================================
Private Sub AnidarPiezas(pL() As Double, pA() As Double, pVeta() As Long, _
                         pGiro() As Boolean, pCod() As String, nP As Long, _
                         utilL As Double, utilA As Double, _
                         nombreGrupo As String, tabLv As Double, tabAv As Double, _
                         ByRef exacto As Double, ByRef pedir As Long, _
                         ByRef imposibles As Collection, ByRef avisos As Collection)

    Dim p As Long, o As Long, tmp As Double

    For p = 0 To nP - 1
        If pGiro(p) And pA(p) > pL(p) Then
            tmp = pL(p): pL(p) = pA(p): pA(p) = tmp
        End If
    Next p

    Dim nO() As Long, oW() As Double, oH() As Double
    ReDim nO(0 To nP - 1)
    ReDim oW(0 To nP - 1, 0 To 1)
    ReDim oH(0 To nP - 1, 0 To 1)

    For p = 0 To nP - 1
        Select Case pVeta(p)
            Case 1
                nO(p) = 1
                oW(p, 0) = pA(p): oH(p, 0) = pL(p)
            Case 2
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

    ' Clasificar: 0 = normal (con orillas), 1 = sin margen (tablero completo),
    ' 2 = imposible (no cabe ni en el bruto). Avisos agrupados por Cod. Pieza.
    Dim pClase() As Long
    ReDim pClase(0 To nP - 1)
    Dim impDict As Object, smDict As Object
    Set impDict = CreateObject("Scripting.Dictionary")
    Set smDict = CreateObject("Scripting.Dictionary")
    Dim clave As Variant

    For p = 0 To nP - 1
        pClase(p) = 2
        For o = 0 To nO(p) - 1
            If oW(p, o) <= utilA And oH(p, o) <= utilL Then
                pClase(p) = 0
                Exit For
            ElseIf oW(p, o) <= tabAv And oH(p, o) <= tabLv Then
                pClase(p) = 1
            End If
        Next o
        If pClase(p) > 0 Then
            clave = pCod(p) & "|" & FmtMM(pL(p)) & " x " & FmtMM(pA(p))
            If pClase(p) = 2 Then
                If impDict.Exists(clave) Then impDict(clave) = impDict(clave) + 1 _
                                         Else impDict.Add clave, 1
            Else
                If smDict.Exists(clave) Then smDict(clave) = smDict(clave) + 1 _
                                        Else smDict.Add clave, 1
            End If
        End If
    Next p

    Dim partes() As String
    For Each clave In impDict.Keys
        partes = Split(CStr(clave), "|")
        imposibles.Add "NO CABE: " & impDict(clave) & " ud" & IIf(impDict(clave) > 1, "s", "") & _
                       " de " & partes(1) & " mm - Cod. Pieza " & partes(0) & _
                       "  (" & nombreGrupo & ", tablero " & Format$(tabLv, "0") & _
                       " x " & Format$(tabAv, "0") & ")"
    Next clave
    For Each clave In smDict.Keys
        partes = Split(CStr(clave), "|")
        avisos.Add "SIN MARGEN: " & smDict(clave) & " ud" & IIf(smDict(clave) > 1, "s", "") & _
                   " de " & partes(1) & " mm - Cod. Pieza " & partes(0) & _
                   "  (" & nombreGrupo & "): usa el tablero completo de " & _
                   Format$(tabLv, "0") & " x " & Format$(tabAv, "0") & _
                   " sin margen de orilla - revisar antes de cortar"
    Next clave

    ' Ordenar de mayor a menor por alto de colocacion; en empate, por ancho
    Dim ii As Long, jj As Long, doSwap As Boolean
    Dim ki As Double, kj As Double
    Dim tL As Double, tA As Double, tV As Long, tG As Boolean, tS As Long, tN As Long
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
                tS = pClase(ii): pClase(ii) = pClase(jj): pClase(jj) = tS
                tN = nO(ii): nO(ii) = nO(jj): nO(jj) = tN
                tW0 = oW(ii, 0): oW(ii, 0) = oW(jj, 0): oW(jj, 0) = tW0
                tW1 = oW(ii, 1): oW(ii, 1) = oW(jj, 1): oW(jj, 1) = tW1
                tH0 = oH(ii, 0): oH(ii, 0) = oH(jj, 0): oH(jj, 0) = tH0
                tH1 = oH(ii, 1): oH(ii, 1) = oH(jj, 1): oH(jj, 1) = tH1
            End If
        Next jj
    Next ii

    Dim tab1 As Long, area1 As Double
    Call ColocarPasada(pL, pA, nO, oW, oH, pClase, nP, 0, utilL, utilA, tab1, area1)
    Dim tab2 As Long, area2 As Double
    Call ColocarPasada(pL, pA, nO, oW, oH, pClase, nP, 1, tabLv, tabAv, tab2, area2)

    exacto = tab1 + Decimal1(area1, utilL * utilA) + tab2 + Decimal1(area2, tabLv * tabAv)
    pedir = tab1 + IIf(area1 > 0, 1, 0) + tab2 + IIf(area2 > 0, 1, 0)
End Sub


' ColocarPasada: nesting por franjas de las piezas de una clase concreta
Private Sub ColocarPasada(pL() As Double, pA() As Double, nO() As Long, _
                          oW() As Double, oH() As Double, pClase() As Long, _
                          nP As Long, clase As Long, limL As Double, limA As Double, _
                          ByRef tableros As Long, ByRef areaLast As Double)
    Dim p As Long, o As Long
    Dim curX As Double, curY As Double, shelfH As Double
    tableros = 0: curX = 0: curY = 0: shelfH = 0: areaLast = 0

    Dim placed As Boolean
    Dim w As Double, h As Double, needX As Double, newH As Double, newY As Double

    For p = 0 To nP - 1
        If pClase(p) = clase Then
            placed = False
            For o = 0 To nO(p) - 1
                w = oW(p, o): h = oH(p, o)
                If curX = 0 Then needX = w Else needX = curX + SEP + w
                newH = shelfH: If h > newH Then newH = h
                If needX <= limA And curY + newH <= limL Then
                    curX = needX: shelfH = newH
                    areaLast = areaLast + pL(p) * pA(p)
                    placed = True: Exit For
                End If
            Next o
            If Not placed Then
                For o = 0 To nO(p) - 1
                    w = oW(p, o): h = oH(p, o)
                    newY = curY + shelfH + SEP
                    If newY + h <= limL And w <= limA Then
                        curY = newY: shelfH = h: curX = w
                        areaLast = areaLast + pL(p) * pA(p)
                        placed = True: Exit For
                    End If
                Next o
            End If
            If Not placed Then
                tableros = tableros + 1
                curX = 0: curY = 0: shelfH = 0: areaLast = 0
                For o = 0 To nO(p) - 1
                    w = oW(p, o): h = oH(p, o)
                    If w <= limA And h <= limL Then
                        curX = w: shelfH = h
                        areaLast = pL(p) * pA(p)
                        placed = True: Exit For
                    End If
                Next o
            End If
        End If
    Next p
End Sub


' Decimal1: fraccion de tablero ocupada, redondeada a 1 decimal (min 0.1)
Private Function Decimal1(areaPiezas As Double, areaTablero As Double) As Double
    Dim dec As Double
    dec = 0
    If areaPiezas > 0 And areaTablero > 0 Then
        dec = Int(areaPiezas / areaTablero * 10 + 0.5) / 10
        If dec < 0.1 Then dec = 0.1
        If dec > 1 Then dec = 1
    End If
    Decimal1 = dec
End Function


' =====================================================================
' ImportarCSVs: selecciona varios CSV del SIG y los combina en un libro
' nuevo (una sola cabecera). Devuelve la hoja combinada.
' =====================================================================
Private Function ImportarCSVs() As Worksheet
    Dim archivos As Variant
    archivos = Application.GetOpenFilename( _
        FileFilter:="Archivos CSV (*.csv), *.csv", _
        Title:="Selecciona los CSV del pedido (Ctrl+clic para varios)", _
        MultiSelect:=True)
    If Not IsArray(archivos) Then Exit Function

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
' LimpiarZonaSalida: borra tablas, formatos y validaciones de una
' ejecucion anterior (todo lo que hay debajo de los datos)
' =====================================================================
Private Sub LimpiarZonaSalida(ws As Worksheet, dataLast As Long, nElems As Long)
    Dim hastaFila As Long
    hastaFila = dataLast + 3 * nElems + 800

    Dim zona As Range
    Set zona = ws.Rows(dataLast + 1 & ":" & hastaFila)
    On Error Resume Next
    Dim i As Long
    For i = ws.ListObjects.Count To 1 Step -1
        If ws.ListObjects(i).Range.Row > dataLast Then ws.ListObjects(i).Delete
    Next i
    zona.Validation.Delete
    ws.Buttons(BTN_NAME).Delete
    On Error GoTo 0
    ' El boton-forma NO se borra: se reutiliza en AsegurarBotonRecalcular.
    zona.Clear
End Sub


' =====================================================================
' AsegurarBotonRecalcular: crea el boton amarillo si no existe, o lo
' reposiciona si ya estaba. Nunca deja la hoja sin boton.
' =====================================================================
Private Sub AsegurarBotonRecalcular(ws As Worksheet, r0 As Long)
    Dim btn As Shape
    Dim posL As Double, posT As Double
    posL = ws.Cells(r0, COL_OUT + 4).Left + 5
    posT = ws.Cells(r0, COL_OUT + 4).Top - 2

    Dim s As Shape, encontrado As Boolean, k As Long
    encontrado = False
    For k = ws.Shapes.Count To 1 Step -1
        Set s = ws.Shapes(k)
        If s.Name = BTN_NAME Then
            If encontrado Then s.Delete Else Set btn = s: encontrado = True
        End If
    Next k

    If btn Is Nothing Then
        Set btn = ws.Shapes.AddShape(msoShapeRoundedRectangle, posL, posT, 130, 26)
        btn.Name = BTN_NAME
        btn.Fill.ForeColor.RGB = RGB(255, 204, 0)
        btn.Line.ForeColor.RGB = RGB(175, 140, 0)
        With btn.TextFrame2
            .TextRange.Text = "RECALCULAR"
            .TextRange.Font.Bold = msoTrue
            .TextRange.Font.Size = 11
            .TextRange.Font.Fill.ForeColor.RGB = RGB(0, 0, 0)
            .TextRange.ParagraphFormat.Alignment = msoAlignCenter
            .VerticalAnchor = msoAnchorMiddle
        End With
    Else
        btn.Left = posL
        btn.Top = posT
    End If

    btn.Placement = xlFreeFloating
    btn.OnAction = "'" & ThisWorkbook.Name & "'!AnidadoPRO_V2"
End Sub


' =====================================================================
' UltimaFilaDatos: ultima fila real de datos, ignorando las tablas de
' resultados de una ejecucion anterior
' =====================================================================
Private Function UltimaFilaDatos(ws As Worksheet) As Long
    Dim mRow As Long
    mRow = FilaMarcador(ws, TIT_MEDIDAS)

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


' FilaMarcador: fila donde esta el titulo indicado (0 si no existe)
Private Function FilaMarcador(ws As Worksheet, titulo As String) As Long
    Dim f As Range
    Set f = ws.Columns(C_MAT).Find(What:=titulo, LookIn:=xlValues, _
                                   LookAt:=xlWhole, MatchCase:=False)
    If Not f Is Nothing Then FilaMarcador = f.Row
End Function


' LeerMedidasPrevias: recupera medidas de tablero y uniones de la tabla
' MEDIDAS de una ejecucion anterior (para no reintroducirlas al recalcular)
Private Sub LeerMedidasPrevias(ws As Worksheet, ByRef dL As Object, _
                               ByRef dA As Object, ByRef dQ As Object)
    Set dL = CreateObject("Scripting.Dictionary"): dL.CompareMode = vbTextCompare
    Set dA = CreateObject("Scripting.Dictionary"): dA.CompareMode = vbTextCompare
    Set dQ = CreateObject("Scripting.Dictionary"): dQ.CompareMode = vbTextCompare

    Dim mRow As Long
    mRow = FilaMarcador(ws, TIT_MEDIDAS)
    If mRow = 0 Then Exit Sub

    Dim r As Long, clave As String
    r = mRow + 2
    Do While Trim$(CStr(ws.Cells(r, COL_OUT).Value)) <> ""
        clave = ClaveMaterial(Trim$(CStr(ws.Cells(r, COL_OUT).Value)))
        If Not dL.Exists(clave) Then
            dL.Add clave, ToNum(ws.Cells(r, COL_OUT + 1).Value)
            dA.Add clave, ToNum(ws.Cells(r, COL_OUT + 2).Value)
            dQ.Add clave, Trim$(CStr(ws.Cells(r, COL_OUT + 3).Value))
        End If
        r = r + 1
    Loop
End Sub


' LeerUnionesCanto: recupera las uniones MISMO QUE de la tabla CANTOS de
' una ejecucion anterior (columna MISMO QUE en COL_OUT+3)
Private Sub LeerUnionesCanto(ws As Worksheet, ByRef dQ As Object)
    Set dQ = CreateObject("Scripting.Dictionary"): dQ.CompareMode = vbTextCompare

    Dim mRow As Long
    mRow = FilaMarcador(ws, TIT_CANTOS)
    If mRow = 0 Then Exit Sub

    Dim r As Long, cm As String
    r = mRow + 2
    Do While Trim$(CStr(ws.Cells(r, COL_OUT).Value)) <> "" And _
             Trim$(CStr(ws.Cells(r, COL_OUT).Value)) <> "TOTAL METROS"
        cm = ClaveMaterial(Trim$(CStr(ws.Cells(r, COL_OUT).Value)))
        If Not dQ.Exists(cm) Then dQ.Add cm, Trim$(CStr(ws.Cells(r, COL_OUT + 3).Value))
        r = r + 1
    Loop
End Sub


' =====================================================================
' Utilidades
' =====================================================================

' ParseCanteado: cuenta cantos del lado largo (nL) y del corto (nC) a
' partir del texto de la columna Canteado ("1L", "2C 2L", "1C 2L", "0"...)
Private Sub ParseCanteado(txt As String, ByRef nL As Long, ByRef nC As Long)
    nL = 0: nC = 0
    Dim s As String
    s = UCase$(Trim$(txt))
    Dim i As Long, ch As String, num As String, cuenta As Long
    num = ""
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch >= "0" And ch <= "9" Then
            num = num & ch
        ElseIf ch = "L" Then
            If num = "" Then cuenta = 1 Else cuenta = CLng(num)
            nL = nL + cuenta: num = ""
        ElseIf ch = "C" Then
            If num = "" Then cuenta = 1 Else cuenta = CLng(num)
            nC = nC + cuenta: num = ""
        Else
            num = ""
        End If
    Next i
End Sub


' CantoMaterial: material del canto de una fila. SIEMPRE el material del
' tablero sin el espesor; si "Comentarios Canteado" tiene texto, se anade
' entre parentesis (esa fila se pintara de naranja para revisar).
'   ""                 -> "AGLOMERADO-S/ORDEN"
'   "CANTO PVC BLANCO" -> "AGLOMERADO-S/ORDEN (CANTO PVC BLANCO)"
Private Function CantoMaterial(ws As Worksheet, r As Long) As String
    Dim base As String, com As String
    base = MaterialSinEspesor(Trim$(CStr(ws.Cells(r, C_MAT).Value)))
    com = Trim$(CStr(ws.Cells(r, C_COM_CANTO).Value))
    If com <> "" Then
        CantoMaterial = base & " (" & com & ")"
    Else
        CantoMaterial = base
    End If
End Function

' CantoTieneComentario: True si la fila trae texto en Comentarios Canteado
Private Function CantoTieneComentario(ws As Worksheet, r As Long) As Boolean
    CantoTieneComentario = (Trim$(CStr(ws.Cells(r, C_COM_CANTO).Value)) <> "")
End Function


' MaterialSinEspesor: quita el espesor final del nombre del material
' ("AGLOMERADO-S/ORDEN-16" -> "AGLOMERADO-S/ORDEN")
Private Function MaterialSinEspesor(mat As String) As String
    Dim pos As Long, tail As String, base As String
    pos = InStrRev(mat, "-")
    If pos > 0 And pos < Len(mat) Then
        tail = Trim$(Mid$(mat, pos + 1))
        If IsNumeric(Replace$(tail, ",", ".")) Then
            base = RTrim$(Left$(mat, pos - 1))
            Do While Len(base) > 0 And Right$(base, 1) = "-"
                base = Left$(base, Len(base) - 1)
            Loop
            If base <> "" Then MaterialSinEspesor = base Else MaterialSinEspesor = mat
            Exit Function
        End If
    End If
    MaterialSinEspesor = mat
End Function


' SonParecidos: True si dos nombres son casi iguales (<=3 ediciones) o uno
' es prefijo del otro (p.ej. "CANTO PVC BLANCO" y su variante con nota)
Private Function SonParecidos(a As String, b As String) As Boolean
    Dim na As String, nb As String
    na = NormalizarNombre(a): nb = NormalizarNombre(b)
    If na = "" Or nb = "" Then Exit Function
    If Levenshtein(na, nb) <= 3 Then SonParecidos = True: Exit Function

    Dim corto As String, largo As String
    If Len(na) <= Len(nb) Then corto = na: largo = nb Else corto = nb: largo = na
    If Len(corto) >= 6 Then
        If Left$(largo, Len(corto)) = corto Then SonParecidos = True
    End If
End Function


' FmtMM: formatea una medida sin dejar separador decimal colgando
Private Function FmtMM(x As Double) As String
    If x = Int(x) Then FmtMM = Format$(x, "0") Else FmtMM = Format$(x, "0.##")
End Function

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

' ClaveMaterial: clave de agrupacion que ignora guiones y espacios de mas.
' Colapsa cualquier serie de espacios o guiones y quita los espacios pegados
' a un guion, de modo que "MDF--16", "MDF---16" y "MDF   --16" dan la misma
' clave ("MDF-16"). Solo toca separadores: nombres realmente distintos siguen
' siendo distintos.
Private Function ClaveMaterial(s As String) As String
    Dim t As String
    t = UCase$(Trim$(s))
    ' 1) espacios multiples -> uno
    Do While InStr(t, "  ") > 0
        t = Replace$(t, "  ", " ")
    Loop
    ' 2) quitar espacios pegados a un guion ("MDF -16" / "MDF- 16" -> "MDF-16")
    Do While InStr(t, " -") > 0 Or InStr(t, "- ") > 0
        t = Replace$(t, " -", "-")
        t = Replace$(t, "- ", "-")
    Loop
    ' 3) guiones multiples -> uno
    Do While InStr(t, "--") > 0
        t = Replace$(t, "--", "-")
    Loop
    ClaveMaterial = Trim$(t)
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
    For i = 7 To 10
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
