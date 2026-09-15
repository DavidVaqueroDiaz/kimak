#Requires -Version 5.1
<#
.SYNOPSIS
    Copia UNA sola copia de cada pieza CNC listada en un indice JSON.

.DESCRIPTION
    Lee un indice JSON (el que genera Generar-Indice.ps1 o equivalente), detecta
    el codigo de pieza dentro del nombre de cada archivo (p.ej. 1002275-001-001_A),
    agrupa los archivos que son la misma pieza y copia a la carpeta destino
    solo el mas reciente de cada grupo.

    Reglas aplicadas:
      - El codigo completo INCLUYE la revision (_A, _B, _D): 881048-001-001_A y
        881048-001-001_D son piezas distintas y se copian las dos.
      - "Cara A" y "Cara B" se tratan por separado (opcional): de cada cara se
        copia su archivo mas reciente, para no perder el programa de la 2a cara.
      - Los archivos sin codigo de pieza (nestings, tableros: "DM 25mm.",
        "PLACAS PVC...") NO se copian; salen listados en el informe CSV.
      - Por defecto solo se procesan los .ard. Si anades otras extensiones, cada
        una se trata por separado (.cix / .hop son maquinas distintas).

.EXAMPLE
    .\Copiar-Piezas.ps1
    Abre la ventana con el cuadro de texto de la ruta destino.

.EXAMPLE
    .\Copiar-Piezas.ps1 -SinInterfaz -Indice "C:\indices\M2.json" -Destino "D:\PIEZAS" -Simular
    Ejecuta sin ventana (para tareas programadas) y solo simula.
#>
[CmdletBinding()]
param(
    [string]   $Indice,
    [string]   $Destino,
    [string[]] $Extensiones,
    [string]   $FiltroRuta,
    [switch]   $NoSepararCaras,
    [switch]   $FechasDeDisco,
    [switch]   $UsarFechaCreacion,
    [switch]   $Sobrescribir,
    [switch]   $Simular,
    [switch]   $SinInterfaz
)

$ErrorActionPreference = 'Stop'

# ============================================================================
#  CONFIGURACION
# ============================================================================

$OPC = [System.Text.RegularExpressions.RegexOptions]::IgnoreCase -bor `
       [System.Text.RegularExpressions.RegexOptions]::Compiled

# Codigo de pieza: 6-9 digitos - 3/4 digitos - 2/4 digitos [ _ o - + letra de revision ]
#   1002275-001-001_A   902275-001-001_A   000875483-001-001_A   826928-001-0036_A
#   943625-001-001-B    865618-001-001 (sin revision)
# Se busca en CUALQUIER parte del nombre: "PIEZA 10_943251-001-001_A MDF 16mm" tambien vale.
$RX_CODIGO = [regex]::new(
    '(?<![0-9])(?<num>\d{6,9})-(?<g2>\d{3,4})-(?<g3>\d{2,4})(?:[_\-](?<rev>[A-Za-z]))?(?![0-9A-Za-z])',
    [System.Text.RegularExpressions.RegexOptions]::Compiled)

# Cara A / Cara B. El (?![a-z]) evita cazar "cara buena" y "cara mala".
$RX_CARA_B = [regex]::new('cara\s*[_\-\.]?\s*b(?![a-z])', $OPC)
$RX_CARA_A = [regex]::new('cara\s*[_\-\.]?\s*a(?![a-z])', $OPC)

$EXT_POR_DEFECTO = '.ard'
$CARPETA_CONFIG  = [Environment]::GetFolderPath('ApplicationData')
if (-not $CARPETA_CONFIG) { $CARPETA_CONFIG = [System.IO.Path]::GetTempPath() }
$ARCHIVO_CONFIG  = Join-Path (Join-Path $CARPETA_CONFIG 'KIMAK') 'copiar-piezas.json'

$script:Cancelar = $false

# ============================================================================
#  MOTOR (sin interfaz: se puede usar desde linea de comandos)
# ============================================================================

function ConvertTo-RutaLarga {
    # Prefijo \\?\ para poder trabajar con rutas de mas de 260 caracteres.
    # En el indice hay rutas de hasta 293 caracteres: sin esto, Copy-Item falla.
    param([string]$Ruta)
    if ([string]::IsNullOrWhiteSpace($Ruta)) { return $Ruta }
    if ($Ruta.StartsWith('\\?\'))            { return $Ruta }
    if ($Ruta.StartsWith('\\'))              { return '\\?\UNC\' + $Ruta.Substring(2) }
    if ($Ruta -match '^[A-Za-z]:\\')         { return '\\?\' + $Ruta }
    return $Ruta
}

function Get-Campo {
    # Lee un campo tanto si el registro es PSCustomObject (ConvertFrom-Json)
    # como si es Dictionary (JavaScriptSerializer, PowerShell 5.1 con JSON grande).
    param($Objeto, [string]$Nombre)
    if ($null -eq $Objeto) { return $null }
    if ($Objeto -is [System.Collections.IDictionary]) {
        if ($Objeto.Contains($Nombre)) { return $Objeto[$Nombre] }
        return $null
    }
    $p = $Objeto.PSObject.Properties[$Nombre]
    if ($p) { return $p.Value }
    return $null
}

function Read-IndicePiezas {
    param([string]$Ruta)
    $rutaL = ConvertTo-RutaLarga $Ruta
    if (-not [System.IO.File]::Exists($rutaL)) {
        throw "No existe el indice: $Ruta"
    }
    $texto = [System.IO.File]::ReadAllText($rutaL, [System.Text.Encoding]::UTF8)
    if ($texto.Length -gt 0 -and $texto[0] -eq [char]0xFEFF) { $texto = $texto.Substring(1) }
    try {
        return @(ConvertFrom-Json -InputObject $texto)
    } catch {
        # PowerShell 5.1 con archivos JSON grandes: segundo intento con el
        # serializador de .NET, que no tiene el limite de tamano.
        Add-Type -AssemblyName System.Web.Extensions
        $js = New-Object System.Web.Script.Serialization.JavaScriptSerializer
        $js.MaxJsonLength  = [int]::MaxValue
        $js.RecursionLimit = 200
        return @($js.DeserializeObject($texto))
    }
}

function ConvertTo-Fecha {
    param([string]$Texto)
    if ([string]::IsNullOrWhiteSpace($Texto)) { return [datetime]::MinValue }
    $f = [datetime]::MinValue
    $estilos = [System.Globalization.DateTimeStyles]::RoundtripKind
    if ([datetime]::TryParse($Texto, [System.Globalization.CultureInfo]::InvariantCulture, $estilos, [ref]$f)) {
        return $f
    }
    if ([datetime]::TryParse($Texto, [ref]$f)) { return $f }
    return [datetime]::MinValue
}

function Get-CodigoPieza {
    # Devuelve el codigo normalizado (818463-001-001_A) o $null si no hay.
    param([string]$Nombre)
    $m = $RX_CODIGO.Match($Nombre)
    if (-not $m.Success) { return $null }
    $num = $m.Groups['num'].Value.TrimStart('0'); if ($num.Length -lt 6) { $num = $num.PadLeft(6, '0') }
    $g2  = $m.Groups['g2'].Value.TrimStart('0');  if ($g2.Length  -lt 3) { $g2  = $g2.PadLeft(3, '0') }
    $g3  = $m.Groups['g3'].Value.TrimStart('0');  if ($g3.Length  -lt 3) { $g3  = $g3.PadLeft(3, '0') }
    $cod = "$num-$g2-$g3"
    if ($m.Groups['rev'].Success) { $cod += '_' + $m.Groups['rev'].Value.ToUpperInvariant() }
    return $cod
}

function Get-CaraPieza {
    param([string]$Nombre)
    if ($RX_CARA_B.IsMatch($Nombre)) { return 'B' }
    if ($RX_CARA_A.IsMatch($Nombre)) { return 'A' }
    return ''
}

function Invoke-CopiaPiezas {
    param(
        [Parameter(Mandatory = $true)][string] $Indice,
        [Parameter(Mandatory = $true)][string] $Destino,
        [string[]]    $Extensiones   = @('.ard'),
        [string]      $FiltroRuta    = '',
        [bool]        $SepararCaras  = $true,
        [bool]        $FechasDeDisco = $false,
        [bool]        $UsarFechaCreacion = $false,
        [bool]        $Sobrescribir  = $false,
        [bool]        $Simular       = $false,
        [scriptblock] $OnLog         = $null,
        [scriptblock] $OnProgreso    = $null
    )

    function Escribe([string]$t) { if ($OnLog) { & $OnLog $t } else { Write-Host $t } }
    function Avanza([int]$v, [int]$max, [string]$t) { if ($OnProgreso) { & $OnProgreso $v $max $t } }

    $script:Cancelar = $false
    $crono = [System.Diagnostics.Stopwatch]::StartNew()

    # ---- 1. Leer indice -----------------------------------------------------
    Escribe "Leyendo indice: $Indice"
    $registros = Read-IndicePiezas $Indice
    $total = $registros.Count
    Escribe ("Registros en el indice: {0:N0}" -f $total)
    if ($total -eq 0) { throw 'El indice no contiene registros.' }

    # ---- 2. Preparar filtros ------------------------------------------------
    $extSet = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($e in $Extensiones) {
        $x = "$e".Trim()
        if ($x) { if (-not $x.StartsWith('.')) { $x = '.' + $x }; [void]$extSet.Add($x) }
    }
    if ($extSet.Count -eq 0) { throw 'No has indicado ninguna extension.' }

    $destinoAbs = [System.IO.Path]::GetFullPath($Destino.Trim().TrimEnd('\'))
    $destinoL   = ConvertTo-RutaLarga $destinoAbs
    if (-not $Simular) { [void][System.IO.Directory]::CreateDirectory($destinoL) }

    # ---- 3. Clasificar ------------------------------------------------------
    $grupos    = New-Object 'System.Collections.Generic.Dictionary[string,System.Collections.Generic.List[object]]' ([System.StringComparer]::OrdinalIgnoreCase)
    $sinCodigo = New-Object 'System.Collections.Generic.List[object]'
    $nExtFuera = 0
    $nRutaFuera = 0
    $i = 0

    Escribe 'Clasificando archivos...'
    foreach ($r in $registros) {
        $i++
        if (($i % 2000) -eq 0) { Avanza $i $total "Clasificando $i / $total"; if ($script:Cancelar) { return $null } }

        $nombre = [string](Get-Campo $r 'nombre')
        if ([string]::IsNullOrWhiteSpace($nombre)) { continue }

        $ruta = [string](Get-Campo $r 'ruta_completa')
        if ([string]::IsNullOrWhiteSpace($ruta)) {
            $carpeta = [string](Get-Campo $r 'ruta')
            if ($carpeta) { $ruta = (Join-Path $carpeta $nombre) } else { continue }
        }

        $ext = [string](Get-Campo $r 'extension')
        if ([string]::IsNullOrWhiteSpace($ext)) { $ext = [System.IO.Path]::GetExtension($nombre) }
        if (-not $extSet.Contains($ext)) { $nExtFuera++; continue }

        if ($FiltroRuta -and ($ruta.IndexOf($FiltroRuta, [System.StringComparison]::OrdinalIgnoreCase) -lt 0)) {
            $nRutaFuera++; continue
        }

        $sinExt = [string](Get-Campo $r 'nombre_sin_ext')
        if ([string]::IsNullOrWhiteSpace($sinExt)) { $sinExt = [System.IO.Path]::GetFileNameWithoutExtension($nombre) }

        $fechaTxt = [string](Get-Campo $r 'fecha_modificacion')
        $fecha    = ConvertTo-Fecha $fechaTxt
        if ($UsarFechaCreacion) {
            $fechaCre = [string](Get-Campo $r 'fecha_creacion')   # solo si el indice la trae
            if ($fechaCre) { $fc = ConvertTo-Fecha $fechaCre; if ($fc -gt $fecha) { $fecha = $fc } }
        }

        $item = [pscustomobject]@{
            Nombre      = $nombre
            Ruta        = $ruta
            Ext         = $ext
            FechaIndice = $fecha
            FechaUsada  = $fecha
            Codigo      = Get-CodigoPieza $sinExt
            Cara        = ''
            Estado      = ''
            Motivo      = ''
            Destino     = ''
        }

        if (-not $item.Codigo) { $sinCodigo.Add($item); continue }
        if ($SepararCaras) { $item.Cara = Get-CaraPieza $sinExt }

        $clave = '{0}|{1}|{2}' -f $item.Codigo, $item.Cara, $ext
        $lista = $null
        if (-not $grupos.TryGetValue($clave, [ref]$lista)) {
            $lista = New-Object 'System.Collections.Generic.List[object]'
            $grupos[$clave] = $lista
        }
        $lista.Add($item)
    }

    $nConCodigo = 0; foreach ($g in $grupos.Values) { $nConCodigo += $g.Count }
    Escribe ("  Con codigo de pieza : {0:N0} archivos" -f $nConCodigo)
    Escribe ("  Piezas distintas    : {0:N0}" -f $grupos.Count)
    Escribe ("  Sin codigo (ignorados): {0:N0}" -f $sinCodigo.Count)
    if ($nExtFuera)  { Escribe ("  Fuera por extension : {0:N0}" -f $nExtFuera) }
    if ($nRutaFuera) { Escribe ("  Fuera por filtro de ruta: {0:N0}" -f $nRutaFuera) }

    # ---- 4. Fechas reales de disco (solo donde hay que desempatar) ----------
    if ($FechasDeDisco) {
        $aRevisar = @($grupos.Values | Where-Object { $_.Count -gt 1 })
        $n = 0; $tot = $aRevisar.Count
        Escribe ("Leyendo fechas reales en disco de {0:N0} grupos duplicados..." -f $tot)
        foreach ($g in $aRevisar) {
            $n++
            if (($n % 50) -eq 0) { Avanza $n $tot "Fechas en disco $n / $tot"; if ($script:Cancelar) { return $null } }
            foreach ($it in $g) {
                try {
                    $fi = New-Object System.IO.FileInfo((ConvertTo-RutaLarga $it.Ruta))
                    if ($fi.Exists) {
                        $f = $fi.LastWriteTime
                        if ($UsarFechaCreacion -and $fi.CreationTime -gt $f) { $f = $fi.CreationTime }
                        $it.FechaUsada = $f
                    }
                } catch { }
            }
        }
    }

    # ---- 5. Elegir ganador y copiar -----------------------------------------
    $usados = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    $nCopiados = 0; $nOmitidos = 0; $nErrores = 0; $nNoEncontrados = 0; $nDescartados = 0
    $informe = New-Object 'System.Collections.Generic.List[object]'
    $n = 0; $tot = $grupos.Count

    Escribe ''
    Escribe $(if ($Simular) { '*** SIMULACION: no se copia nada ***' } else { "Copiando a: $destinoAbs" })

    foreach ($clave in $grupos.Keys) {
        $n++
        if (($n % 100) -eq 0) {
            Avanza $n $tot ("Copiando {0:N0} / {1:N0}" -f $n, $tot)
            if ($script:Cancelar) { Escribe 'CANCELADO por el usuario.'; break }
        }

        # Mas reciente primero; a igualdad de fecha, la ruta mas corta (suele ser la original)
        $lista = $grupos[$clave]
        if ($lista.Count -eq 1) {
            $orden = @($lista[0])
        } else {
            $orden = @($lista | Sort-Object -Property @{Expression = 'FechaUsada'; Descending = $true},
                                                      @{Expression = { $_.Ruta.Length }; Descending = $false},
                                                      @{Expression = 'Ruta'; Descending = $false})
        }

        $ganador = $null
        foreach ($cand in $orden) {
            if ([System.IO.File]::Exists((ConvertTo-RutaLarga $cand.Ruta))) { $ganador = $cand; break }
            $cand.Estado = 'NO ENCONTRADO'
            $cand.Motivo = 'El archivo ya no existe en esa ruta'
            $nNoEncontrados++
        }

        foreach ($cand in $orden) {
            if ($cand -eq $ganador) { continue }
            if ($cand.Estado -eq 'NO ENCONTRADO') { $informe.Add($cand); continue }
            $cand.Estado = 'DESCARTADO'
            $cand.Motivo = if ($ganador) { 'Duplicado: hay otro mas reciente (' + $ganador.FechaUsada.ToString('yyyy-MM-dd HH:mm') + ')' } else { 'Duplicado' }
            $nDescartados++
            $informe.Add($cand)
        }

        if (-not $ganador) { continue }

        # Nombre en destino, evitando pisar un archivo distinto con el mismo nombre
        $baseNom = [System.IO.Path]::GetFileNameWithoutExtension($ganador.Nombre)
        $extNom  = [System.IO.Path]::GetExtension($ganador.Nombre)
        $nomFinal = $ganador.Nombre
        $k = 1
        while ($usados.Contains($nomFinal)) { $k++; $nomFinal = "{0}_{1}{2}" -f $baseNom, $k, $extNom }
        [void]$usados.Add($nomFinal)

        $rutaDest  = Join-Path $destinoAbs $nomFinal
        $rutaDestL = ConvertTo-RutaLarga $rutaDest
        $ganador.Destino = $rutaDest

        try {
            if ($Simular) {
                $ganador.Estado = 'SIMULADO'
                $ganador.Motivo = 'Se copiaria'
                $nCopiados++
            } elseif ([System.IO.File]::Exists($rutaDestL) -and -not $Sobrescribir) {
                $ganador.Estado = 'OMITIDO'
                $ganador.Motivo = 'Ya existe en destino (marca Sobrescribir para reemplazarlo)'
                $nOmitidos++
            } else {
                [System.IO.File]::Copy((ConvertTo-RutaLarga $ganador.Ruta), $rutaDestL, $true)
                $ganador.Estado = 'COPIADO'
                $nCopiados++
            }
        } catch {
            $ganador.Estado = 'ERROR'
            $ganador.Motivo = $_.Exception.Message
            $nErrores++
            Escribe ("  ERROR con {0}: {1}" -f $ganador.Nombre, $_.Exception.Message)
        }
        $informe.Add($ganador)
    }

    foreach ($it in $sinCodigo) {
        $it.Estado = 'SIN CODIGO'
        $it.Motivo = 'No se detecto codigo de pieza en el nombre (nesting / tablero)'
        $informe.Add($it)
    }

    # ---- 6. Informe CSV -----------------------------------------------------
    $sello    = (Get-Date).ToString('yyyyMMdd_HHmmss')
    $rutaCsv  = Join-Path $destinoAbs ("_informe_piezas_$sello.csv")
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine('Estado;Codigo;Cara;Extension;FechaIndice;FechaUsada;Nombre;RutaOrigen;ArchivoDestino;Motivo')
    foreach ($it in $informe) {
        $fi = if ($it.FechaIndice -eq [datetime]::MinValue) { '' } else { $it.FechaIndice.ToString('yyyy-MM-dd HH:mm:ss') }
        $fu = if ($it.FechaUsada  -eq [datetime]::MinValue) { '' } else { $it.FechaUsada.ToString('yyyy-MM-dd HH:mm:ss') }
        $campos = @($it.Estado, $it.Codigo, $it.Cara, $it.Ext, $fi, $fu, $it.Nombre, $it.Ruta, $it.Destino, $it.Motivo)
        $linea = ($campos | ForEach-Object { '"' + ("$_" -replace '"', '""') + '"' }) -join ';'
        [void]$sb.AppendLine($linea)
    }
    try {
        [void][System.IO.Directory]::CreateDirectory($destinoL)
        [System.IO.File]::WriteAllText((ConvertTo-RutaLarga $rutaCsv), $sb.ToString(), (New-Object System.Text.UTF8Encoding($true)))
    } catch { Escribe ("No se pudo escribir el informe: {0}" -f $_.Exception.Message); $rutaCsv = '' }

    $crono.Stop()
    Escribe ''
    Escribe '================ RESUMEN ================'
    Escribe ("Piezas distintas encontradas : {0:N0}" -f $grupos.Count)
    Escribe ("Copiadas                     : {0:N0}" -f $nCopiados)
    Escribe ("Omitidas (ya estaban)        : {0:N0}" -f $nOmitidos)
    Escribe ("Descartadas por duplicadas   : {0:N0}" -f $nDescartados)
    Escribe ("No encontradas en disco      : {0:N0}" -f $nNoEncontrados)
    Escribe ("Errores                      : {0:N0}" -f $nErrores)
    Escribe ("Sin codigo (no copiadas)     : {0:N0}" -f $sinCodigo.Count)
    Escribe ("Tiempo                       : {0:mm\:ss}" -f $crono.Elapsed)
    if ($rutaCsv) { Escribe ("Informe                      : {0}" -f $rutaCsv) }
    Escribe '========================================='

    return [pscustomobject]@{
        Piezas       = $grupos.Count
        Copiadas     = $nCopiados
        Omitidas     = $nOmitidos
        Descartadas  = $nDescartados
        NoEncontradas= $nNoEncontrados
        Errores      = $nErrores
        SinCodigo    = $sinCodigo.Count
        Informe      = $rutaCsv
        Destino      = $destinoAbs
        Cancelado    = $script:Cancelar
    }
}

# ============================================================================
#  INTERFAZ GRAFICA
# ============================================================================

function Get-Preferencias {
    try {
        if ([System.IO.File]::Exists($ARCHIVO_CONFIG)) {
            return (Get-Content -LiteralPath $ARCHIVO_CONFIG -Raw -Encoding UTF8 | ConvertFrom-Json)
        }
    } catch { }
    return $null
}

function Save-Preferencias {
    param($Datos)
    try {
        $carpeta = Split-Path -Parent $ARCHIVO_CONFIG
        if (-not (Test-Path -LiteralPath $carpeta)) { [void](New-Item -ItemType Directory -Path $carpeta -Force) }
        ($Datos | ConvertTo-Json) | Set-Content -LiteralPath $ARCHIVO_CONFIG -Encoding UTF8
    } catch { }
}

function Show-Ventana {
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing
    [System.Windows.Forms.Application]::EnableVisualStyles()

    $pref = Get-Preferencias

    $frm                 = New-Object System.Windows.Forms.Form
    $frm.Text            = 'KIMAK - Copiar piezas desde indice JSON'
    $frm.Size            = New-Object System.Drawing.Size(880, 660)
    $frm.MinimumSize     = New-Object System.Drawing.Size(760, 560)
    $frm.StartPosition   = 'CenterScreen'
    $frm.Font            = New-Object System.Drawing.Font('Segoe UI', 9)

    function NuevaEtiqueta([string]$txt, [int]$x, [int]$y) {
        $l = New-Object System.Windows.Forms.Label
        $l.Text = $txt; $l.Location = New-Object System.Drawing.Point($x, $y)
        $l.AutoSize = $true
        return $l
    }

    # --- Indice JSON ---
    $frm.Controls.Add((NuevaEtiqueta '1) Indice JSON de la carpeta a procesar:' 14 14))
    $txtIndice           = New-Object System.Windows.Forms.TextBox
    $txtIndice.Location  = New-Object System.Drawing.Point(14, 34)
    $txtIndice.Size      = New-Object System.Drawing.Size(730, 24)
    $txtIndice.Anchor    = 'Top,Left,Right'
    if ($pref) { $txtIndice.Text = [string]$pref.Indice }
    $btnIndice           = New-Object System.Windows.Forms.Button
    $btnIndice.Text      = 'Examinar...'
    $btnIndice.Location  = New-Object System.Drawing.Point(752, 33)
    $btnIndice.Size      = New-Object System.Drawing.Size(100, 26)
    $btnIndice.Anchor    = 'Top,Right'
    $frm.Controls.AddRange(@($txtIndice, $btnIndice))

    # --- Carpeta destino (el cuadro de texto que pediste) ---
    $frm.Controls.Add((NuevaEtiqueta '2) Carpeta donde guardar las copias de las piezas:' 14 68))
    $txtDestino          = New-Object System.Windows.Forms.TextBox
    $txtDestino.Location = New-Object System.Drawing.Point(14, 88)
    $txtDestino.Size     = New-Object System.Drawing.Size(730, 24)
    $txtDestino.Anchor   = 'Top,Left,Right'
    if ($pref) { $txtDestino.Text = [string]$pref.Destino }
    $btnDestino          = New-Object System.Windows.Forms.Button
    $btnDestino.Text     = 'Examinar...'
    $btnDestino.Location = New-Object System.Drawing.Point(752, 87)
    $btnDestino.Size     = New-Object System.Drawing.Size(100, 26)
    $btnDestino.Anchor   = 'Top,Right'
    $frm.Controls.AddRange(@($txtDestino, $btnDestino))

    # --- Opciones ---
    $gb          = New-Object System.Windows.Forms.GroupBox
    $gb.Text     = '3) Opciones'
    $gb.Location = New-Object System.Drawing.Point(14, 122)
    $gb.Size     = New-Object System.Drawing.Size(838, 170)
    $gb.Anchor   = 'Top,Left,Right'

    $gb.Controls.Add((NuevaEtiqueta 'Extensiones (separadas por ; ):' 14 26))
    $txtExt          = New-Object System.Windows.Forms.TextBox
    $txtExt.Location = New-Object System.Drawing.Point(200, 23)
    $txtExt.Size     = New-Object System.Drawing.Size(170, 24)
    $txtExt.Text     = if ($pref -and $pref.Extensiones) { [string]$pref.Extensiones } else { $EXT_POR_DEFECTO }
    $gb.Controls.Add($txtExt)

    $gb.Controls.Add((NuevaEtiqueta 'Filtro de ruta (opcional):' 400 26))
    $txtFiltro          = New-Object System.Windows.Forms.TextBox
    $txtFiltro.Location = New-Object System.Drawing.Point(560, 23)
    $txtFiltro.Size     = New-Object System.Drawing.Size(260, 24)
    $txtFiltro.Anchor   = 'Top,Left,Right'
    if ($pref) { $txtFiltro.Text = [string]$pref.FiltroRuta }
    $gb.Controls.Add($txtFiltro)

    $chkCaras          = New-Object System.Windows.Forms.CheckBox
    $chkCaras.Text     = 'Separar Cara A / Cara B (copia el mas reciente de cada cara)'
    $chkCaras.Location = New-Object System.Drawing.Point(16, 56)
    $chkCaras.Size     = New-Object System.Drawing.Size(420, 22)
    $chkCaras.Checked  = if ($pref -and $null -ne $pref.SepararCaras) { [bool]$pref.SepararCaras } else { $true }

    $chkDisco          = New-Object System.Windows.Forms.CheckBox
    $chkDisco.Text     = 'Comprobar fechas reales en disco al desempatar (mas lento)'
    $chkDisco.Location = New-Object System.Drawing.Point(16, 82)
    $chkDisco.Size     = New-Object System.Drawing.Size(420, 22)
    $chkDisco.Checked  = if ($pref -and $null -ne $pref.FechasDeDisco) { [bool]$pref.FechasDeDisco } else { $false }

    $chkSobre          = New-Object System.Windows.Forms.CheckBox
    $chkSobre.Text     = 'Sobrescribir si el archivo ya esta en la carpeta destino'
    $chkSobre.Location = New-Object System.Drawing.Point(450, 56)
    $chkSobre.Size     = New-Object System.Drawing.Size(370, 22)
    $chkSobre.Checked  = if ($pref -and $null -ne $pref.Sobrescribir) { [bool]$pref.Sobrescribir } else { $false }

    $chkCreacion          = New-Object System.Windows.Forms.CheckBox
    $chkCreacion.Text     = 'Tener en cuenta tambien la fecha de creacion (la mas reciente de las dos)'
    $chkCreacion.Location = New-Object System.Drawing.Point(16, 108)
    $chkCreacion.Size     = New-Object System.Drawing.Size(430, 22)
    $chkCreacion.Checked  = if ($pref -and $null -ne $pref.UsarFechaCreacion) { [bool]$pref.UsarFechaCreacion } else { $false }

    $chkSimular          = New-Object System.Windows.Forms.CheckBox
    $chkSimular.Text     = 'SIMULAR: no copia nada, solo genera el informe CSV'
    $chkSimular.Location = New-Object System.Drawing.Point(450, 82)
    $chkSimular.Size     = New-Object System.Drawing.Size(370, 22)
    $chkSimular.Checked  = $false

    $lblAyuda      = NuevaEtiqueta 'Se copia un archivo por pieza: el de fecha mas cercana a hoy. La revision (_A, _B) cuenta como pieza distinta.' 16 132
    $lblAyuda.ForeColor = [System.Drawing.Color]::DimGray
    $gb.Controls.AddRange(@($chkCaras, $chkDisco, $chkSobre, $chkSimular, $chkCreacion, $lblAyuda))
    $frm.Controls.Add($gb)

    # --- Botones ---
    $btnProcesar          = New-Object System.Windows.Forms.Button
    $btnProcesar.Text     = 'PROCESAR'
    $btnProcesar.Location = New-Object System.Drawing.Point(14, 304)
    $btnProcesar.Size     = New-Object System.Drawing.Size(140, 34)
    $btnProcesar.Font     = New-Object System.Drawing.Font('Segoe UI', 10, [System.Drawing.FontStyle]::Bold)

    $btnCancelar          = New-Object System.Windows.Forms.Button
    $btnCancelar.Text     = 'Cancelar'
    $btnCancelar.Location = New-Object System.Drawing.Point(164, 304)
    $btnCancelar.Size     = New-Object System.Drawing.Size(100, 34)
    $btnCancelar.Enabled  = $false

    $btnAbrir             = New-Object System.Windows.Forms.Button
    $btnAbrir.Text        = 'Abrir carpeta destino'
    $btnAbrir.Location    = New-Object System.Drawing.Point(274, 304)
    $btnAbrir.Size        = New-Object System.Drawing.Size(150, 34)

    $pb          = New-Object System.Windows.Forms.ProgressBar
    $pb.Location = New-Object System.Drawing.Point(440, 310)
    $pb.Size     = New-Object System.Drawing.Size(412, 22)
    $pb.Anchor   = 'Top,Left,Right'
    $frm.Controls.AddRange(@($btnProcesar, $btnCancelar, $btnAbrir, $pb))

    $txtLog              = New-Object System.Windows.Forms.TextBox
    $txtLog.Location     = New-Object System.Drawing.Point(14, 348)
    $txtLog.Size         = New-Object System.Drawing.Size(838, 248)
    $txtLog.Multiline    = $true
    $txtLog.ScrollBars   = 'Vertical'
    $txtLog.ReadOnly     = $true
    $txtLog.BackColor    = [System.Drawing.Color]::White
    $txtLog.Font         = New-Object System.Drawing.Font('Consolas', 9)
    $txtLog.Anchor       = 'Top,Left,Right,Bottom'
    $frm.Controls.Add($txtLog)

    # --- Eventos ---
    $btnIndice.Add_Click({
        $d = New-Object System.Windows.Forms.OpenFileDialog
        $d.Filter = 'Indice JSON (*.json)|*.json|Todos (*.*)|*.*'
        if ($txtIndice.Text) { try { $d.InitialDirectory = Split-Path -Parent $txtIndice.Text } catch { } }
        if ($d.ShowDialog() -eq 'OK') { $txtIndice.Text = $d.FileName }
    })

    $btnDestino.Add_Click({
        $d = New-Object System.Windows.Forms.FolderBrowserDialog
        $d.Description = 'Elige la carpeta donde guardar las copias de las piezas'
        if ($txtDestino.Text) { $d.SelectedPath = $txtDestino.Text }
        if ($d.ShowDialog() -eq 'OK') { $txtDestino.Text = $d.SelectedPath }
    })

    $btnAbrir.Add_Click({
        if ($txtDestino.Text -and (Test-Path -LiteralPath $txtDestino.Text)) {
            Start-Process explorer.exe $txtDestino.Text
        } else {
            [void][System.Windows.Forms.MessageBox]::Show('Esa carpeta todavia no existe.', 'KIMAK')
        }
    })

    $btnCancelar.Add_Click({ $script:Cancelar = $true })

    $btnProcesar.Add_Click({
        if (-not $txtIndice.Text.Trim())  { [void][System.Windows.Forms.MessageBox]::Show('Indica el archivo JSON del indice.', 'Falta el indice'); return }
        if (-not $txtDestino.Text.Trim()) { [void][System.Windows.Forms.MessageBox]::Show('Indica la carpeta destino.', 'Falta el destino'); return }

        $txtLog.Clear()
        $btnProcesar.Enabled = $false
        $btnCancelar.Enabled = $true
        $pb.Value = 0
        $frm.Cursor = [System.Windows.Forms.Cursors]::WaitCursor

        $log = {
            param($t)
            $txtLog.AppendText("$t`r`n")
            [System.Windows.Forms.Application]::DoEvents()
        }
        $prog = {
            param($v, $max, $t)
            if ($max -gt 0) {
                $p = [int](100 * $v / $max); if ($p -gt 100) { $p = 100 }
                $pb.Value = $p
            }
            [System.Windows.Forms.Application]::DoEvents()
        }

        try {
            Save-Preferencias ([pscustomobject]@{
                Indice        = $txtIndice.Text
                Destino       = $txtDestino.Text
                Extensiones   = $txtExt.Text
                FiltroRuta    = $txtFiltro.Text
                SepararCaras  = $chkCaras.Checked
                FechasDeDisco = $chkDisco.Checked
                Sobrescribir  = $chkSobre.Checked
                UsarFechaCreacion = $chkCreacion.Checked
            })

            $res = Invoke-CopiaPiezas -Indice $txtIndice.Text.Trim() `
                                      -Destino $txtDestino.Text.Trim() `
                                      -Extensiones ($txtExt.Text -split '[;,]') `
                                      -FiltroRuta $txtFiltro.Text.Trim() `
                                      -SepararCaras $chkCaras.Checked `
                                      -FechasDeDisco $chkDisco.Checked `
                                      -UsarFechaCreacion $chkCreacion.Checked `
                                      -Sobrescribir $chkSobre.Checked `
                                      -Simular $chkSimular.Checked `
                                      -OnLog $log -OnProgreso $prog
            $pb.Value = 100
            if ($res) {
                $msg = "Piezas distintas: {0:N0}`nCopiadas: {1:N0}`nOmitidas (ya estaban): {2:N0}`nDescartadas por duplicadas: {3:N0}`nErrores: {4:N0}`n`nInforme:`n{5}" -f `
                       $res.Piezas, $res.Copiadas, $res.Omitidas, $res.Descartadas, $res.Errores, $res.Informe
                [void][System.Windows.Forms.MessageBox]::Show($msg, 'Proceso terminado', 'OK', 'Information')
            }
        } catch {
            & $log ("ERROR: " + $_.Exception.Message)
            [void][System.Windows.Forms.MessageBox]::Show($_.Exception.Message, 'Error', 'OK', 'Error')
        } finally {
            $frm.Cursor = [System.Windows.Forms.Cursors]::Default
            $btnProcesar.Enabled = $true
            $btnCancelar.Enabled = $false
        }
    })

    [void]$frm.ShowDialog()
}

# ============================================================================
#  ARRANQUE
# ============================================================================

if ($SinInterfaz) {
    if (-not $Indice -or -not $Destino) { throw 'Con -SinInterfaz hay que indicar -Indice y -Destino.' }
    $ext = if ($Extensiones) { $Extensiones } else { $EXT_POR_DEFECTO -split ';' }
    $r = Invoke-CopiaPiezas -Indice $Indice -Destino $Destino -Extensiones $ext `
                            -FiltroRuta $FiltroRuta `
                            -SepararCaras (-not $NoSepararCaras) `
                            -FechasDeDisco ([bool]$FechasDeDisco) `
                            -UsarFechaCreacion ([bool]$UsarFechaCreacion) `
                            -Sobrescribir ([bool]$Sobrescribir) `
                            -Simular ([bool]$Simular)
    $r | Format-List | Out-String | Write-Host
} else {
    Show-Ventana
}
