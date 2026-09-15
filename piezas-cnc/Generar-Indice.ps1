#Requires -Version 5.1
<#
.SYNOPSIS
    Genera el indice JSON de una carpeta (el que luego usa Copiar-Piezas.ps1).

.DESCRIPTION
    Recorre una carpeta y todas sus subcarpetas y escribe un JSON con un
    registro por archivo, con el mismo formato que el indice de M2:

      nombre, nombre_sin_ext, extension, servidor, ruta, ruta_completa,
      fecha_modificacion  (+ fecha_creacion y tamano, que son informativos)

    Soporta rutas de red largas (mas de 260 caracteres) y no se detiene si
    alguna carpeta esta protegida: la anota y sigue.

.EXAMPLE
    .\Generar-Indice.ps1
    Abre la ventana para elegir carpeta y archivo de salida.

.EXAMPLE
    .\Generar-Indice.ps1 -SinInterfaz -Carpeta "\\servidor10\Intercambio_General\CNC M2\ZARA" -Salida "C:\indices\zara.json"
#>
[CmdletBinding()]
param(
    [string]   $Carpeta,
    [string]   $Salida,
    [string[]] $Extensiones,
    [string]   $Servidor = 'M2',
    [switch]   $SoloEstaCarpeta,
    [switch]   $SinInterfaz
)

$ErrorActionPreference = 'Stop'

$EXT_POR_DEFECTO = '.ard'
$CARPETA_CONFIG  = [Environment]::GetFolderPath('ApplicationData')
if (-not $CARPETA_CONFIG) { $CARPETA_CONFIG = [System.IO.Path]::GetTempPath() }
$ARCHIVO_CONFIG  = Join-Path (Join-Path $CARPETA_CONFIG 'KIMAK') 'generar-indice.json'

$script:Cancelar = $false

# ============================================================================
#  MOTOR
# ============================================================================

function ConvertTo-RutaLarga {
    param([string]$Ruta)
    if ([string]::IsNullOrWhiteSpace($Ruta)) { return $Ruta }
    if ($Ruta.StartsWith('\\?\'))            { return $Ruta }
    if ($Ruta.StartsWith('\\'))              { return '\\?\UNC\' + $Ruta.Substring(2) }
    if ($Ruta -match '^[A-Za-z]:\\')         { return '\\?\' + $Ruta }
    return $Ruta
}

function ConvertFrom-RutaLarga {
    param([string]$Ruta)
    if ([string]::IsNullOrWhiteSpace($Ruta)) { return $Ruta }
    if ($Ruta.StartsWith('\\?\UNC\'))        { return '\\' + $Ruta.Substring(8) }
    if ($Ruta.StartsWith('\\?\'))            { return $Ruta.Substring(4) }
    return $Ruta
}

function ConvertTo-TextoJson {
    param([string]$Texto)
    if ([string]::IsNullOrEmpty($Texto)) { return '""' }
    $t = $Texto.Replace('\', '\\').Replace('"', '\"').Replace("`r", '\r').Replace("`n", '\n').Replace("`t", '\t')
    return '"' + $t + '"'
}

function Invoke-GenerarIndice {
    param(
        [Parameter(Mandatory = $true)][string] $Carpeta,
        [Parameter(Mandatory = $true)][string] $Salida,
        [string[]]    $Extensiones     = @('.ard'),
        [string]      $Servidor        = 'M2',
        [bool]        $IncluirSubcarpetas = $true,
        [scriptblock] $OnLog           = $null,
        [scriptblock] $OnProgreso      = $null
    )

    function Escribe([string]$t) { if ($OnLog) { & $OnLog $t } else { Write-Host $t } }

    $script:Cancelar = $false
    $crono = [System.Diagnostics.Stopwatch]::StartNew()

    $raiz = $Carpeta.Trim().TrimEnd('\')
    $raizL = ConvertTo-RutaLarga $raiz
    if (-not [System.IO.Directory]::Exists($raizL)) { throw "No existe la carpeta: $raiz" }

    $extSet = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    $todas = $false
    foreach ($e in $Extensiones) {
        $x = "$e".Trim()
        if ($x -eq '*' -or $x -eq '.*') { $todas = $true; continue }
        if ($x) { if (-not $x.StartsWith('.')) { $x = '.' + $x }; [void]$extSet.Add($x) }
    }
    if (-not $todas -and $extSet.Count -eq 0) { throw 'No has indicado ninguna extension.' }

    Escribe "Recorriendo: $raiz"
    if ($todas) { Escribe 'Extensiones: TODAS' } else { Escribe ('Extensiones: ' + ($extSet -join ' ')) }

    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append('[')

    $pendientes = New-Object 'System.Collections.Generic.Stack[string]'
    $pendientes.Push($raizL)
    $nArchivos = 0; $nCarpetas = 0; $nBloqueadas = 0; $primero = $true

    while ($pendientes.Count -gt 0) {
        if ($script:Cancelar) { Escribe 'CANCELADO por el usuario.'; break }
        $dir = $pendientes.Pop()
        $nCarpetas++
        if (($nCarpetas % 25) -eq 0) {
            if ($OnProgreso) { & $OnProgreso $nCarpetas $nArchivos (ConvertFrom-RutaLarga $dir) }
        }

        if ($IncluirSubcarpetas) {
            try { foreach ($sub in [System.IO.Directory]::EnumerateDirectories($dir)) { $pendientes.Push($sub) } }
            catch { $nBloqueadas++; Escribe ("  (sin acceso) " + (ConvertFrom-RutaLarga $dir)) }
        }

        try {
            foreach ($f in [System.IO.Directory]::EnumerateFiles($dir)) {
                $ext = [System.IO.Path]::GetExtension($f)
                if (-not $todas -and -not $extSet.Contains($ext)) { continue }
                try {
                    $fi = New-Object System.IO.FileInfo($f)
                    $rutaCompleta = ConvertFrom-RutaLarga $f
                    $rutaCarpeta  = ConvertFrom-RutaLarga $dir
                    if (-not $primero) { [void]$sb.Append(',') }
                    $primero = $false
                    [void]$sb.Append('{"nombre":').Append((ConvertTo-TextoJson $fi.Name))
                    [void]$sb.Append(',"nombre_sin_ext":').Append((ConvertTo-TextoJson ([System.IO.Path]::GetFileNameWithoutExtension($fi.Name))))
                    [void]$sb.Append(',"extension":').Append((ConvertTo-TextoJson $ext))
                    [void]$sb.Append(',"servidor":').Append((ConvertTo-TextoJson $Servidor))
                    [void]$sb.Append(',"ruta":').Append((ConvertTo-TextoJson $rutaCarpeta))
                    [void]$sb.Append(',"ruta_completa":').Append((ConvertTo-TextoJson $rutaCompleta))
                    [void]$sb.Append(',"fecha_modificacion":').Append((ConvertTo-TextoJson $fi.LastWriteTime.ToString('o')))
                    [void]$sb.Append(',"fecha_creacion":').Append((ConvertTo-TextoJson $fi.CreationTime.ToString('o')))
                    [void]$sb.Append(',"tamano":').Append($fi.Length)
                    [void]$sb.Append('}')
                    $nArchivos++
                } catch { }
            }
        } catch { $nBloqueadas++ }
    }

    [void]$sb.Append(']')

    $salidaAbs = [System.IO.Path]::GetFullPath($Salida.Trim())
    $carpetaSalida = [System.IO.Path]::GetDirectoryName($salidaAbs)
    if ($carpetaSalida) { [void][System.IO.Directory]::CreateDirectory((ConvertTo-RutaLarga $carpetaSalida)) }
    [System.IO.File]::WriteAllText((ConvertTo-RutaLarga $salidaAbs), $sb.ToString(), (New-Object System.Text.UTF8Encoding($true)))

    $crono.Stop()
    Escribe ''
    Escribe '================ RESUMEN ================'
    Escribe ("Carpetas recorridas : {0:N0}" -f $nCarpetas)
    Escribe ("Archivos indexados  : {0:N0}" -f $nArchivos)
    if ($nBloqueadas) { Escribe ("Carpetas sin acceso : {0:N0}" -f $nBloqueadas) }
    Escribe ("Tiempo              : {0:mm\:ss}" -f $crono.Elapsed)
    Escribe ("Indice              : {0}" -f $salidaAbs)
    Escribe '========================================='

    return [pscustomobject]@{
        Carpetas  = $nCarpetas
        Archivos  = $nArchivos
        SinAcceso = $nBloqueadas
        Indice    = $salidaAbs
        Cancelado = $script:Cancelar
    }
}

# ============================================================================
#  INTERFAZ GRAFICA
# ============================================================================

function Show-Ventana {
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing
    [System.Windows.Forms.Application]::EnableVisualStyles()

    $pref = $null
    try { if ([System.IO.File]::Exists($ARCHIVO_CONFIG)) { $pref = Get-Content -LiteralPath $ARCHIVO_CONFIG -Raw -Encoding UTF8 | ConvertFrom-Json } } catch { }

    $frm               = New-Object System.Windows.Forms.Form
    $frm.Text          = 'KIMAK - Generar indice JSON de una carpeta'
    $frm.Size          = New-Object System.Drawing.Size(880, 560)
    $frm.MinimumSize   = New-Object System.Drawing.Size(760, 480)
    $frm.StartPosition = 'CenterScreen'
    $frm.Font          = New-Object System.Drawing.Font('Segoe UI', 9)

    function NuevaEtiqueta([string]$txt, [int]$x, [int]$y) {
        $l = New-Object System.Windows.Forms.Label
        $l.Text = $txt; $l.Location = New-Object System.Drawing.Point($x, $y); $l.AutoSize = $true
        return $l
    }

    $frm.Controls.Add((NuevaEtiqueta '1) Carpeta a indexar (vale una ruta de red \\servidor10\...):' 14 14))
    $txtCarpeta          = New-Object System.Windows.Forms.TextBox
    $txtCarpeta.Location = New-Object System.Drawing.Point(14, 34)
    $txtCarpeta.Size     = New-Object System.Drawing.Size(730, 24)
    $txtCarpeta.Anchor   = 'Top,Left,Right'
    if ($pref) { $txtCarpeta.Text = [string]$pref.Carpeta }
    $btnCarpeta          = New-Object System.Windows.Forms.Button
    $btnCarpeta.Text     = 'Examinar...'
    $btnCarpeta.Location = New-Object System.Drawing.Point(752, 33)
    $btnCarpeta.Size     = New-Object System.Drawing.Size(100, 26)
    $btnCarpeta.Anchor   = 'Top,Right'
    $frm.Controls.AddRange(@($txtCarpeta, $btnCarpeta))

    $frm.Controls.Add((NuevaEtiqueta '2) Archivo JSON de salida:' 14 68))
    $txtSalida          = New-Object System.Windows.Forms.TextBox
    $txtSalida.Location = New-Object System.Drawing.Point(14, 88)
    $txtSalida.Size     = New-Object System.Drawing.Size(730, 24)
    $txtSalida.Anchor   = 'Top,Left,Right'
    if ($pref -and $pref.Salida) {
        $txtSalida.Text = [string]$pref.Salida
    } else {
        $txtSalida.Text = Join-Path ([Environment]::GetFolderPath('MyDocuments')) ('indice_' + (Get-Date).ToString('yyyyMMdd') + '.json')
    }
    $btnSalida          = New-Object System.Windows.Forms.Button
    $btnSalida.Text     = 'Examinar...'
    $btnSalida.Location = New-Object System.Drawing.Point(752, 87)
    $btnSalida.Size     = New-Object System.Drawing.Size(100, 26)
    $btnSalida.Anchor   = 'Top,Right'
    $frm.Controls.AddRange(@($txtSalida, $btnSalida))

    $gb          = New-Object System.Windows.Forms.GroupBox
    $gb.Text     = '3) Opciones'
    $gb.Location = New-Object System.Drawing.Point(14, 122)
    $gb.Size     = New-Object System.Drawing.Size(838, 90)
    $gb.Anchor   = 'Top,Left,Right'

    $gb.Controls.Add((NuevaEtiqueta 'Extensiones (separadas por ; , o * para todas):' 14 26))
    $txtExt          = New-Object System.Windows.Forms.TextBox
    $txtExt.Location = New-Object System.Drawing.Point(290, 23)
    $txtExt.Size     = New-Object System.Drawing.Size(180, 24)
    $txtExt.Text     = if ($pref -and $pref.Extensiones) { [string]$pref.Extensiones } else { $EXT_POR_DEFECTO }

    $gb.Controls.Add((NuevaEtiqueta 'Servidor:' 500 26))
    $txtServidor          = New-Object System.Windows.Forms.TextBox
    $txtServidor.Location = New-Object System.Drawing.Point(570, 23)
    $txtServidor.Size     = New-Object System.Drawing.Size(120, 24)
    $txtServidor.Text     = if ($pref -and $pref.Servidor) { [string]$pref.Servidor } else { 'M2' }

    $chkSub          = New-Object System.Windows.Forms.CheckBox
    $chkSub.Text     = 'Incluir subcarpetas'
    $chkSub.Location = New-Object System.Drawing.Point(16, 56)
    $chkSub.Size     = New-Object System.Drawing.Size(200, 22)
    $chkSub.Checked  = $true
    $gb.Controls.AddRange(@($txtExt, $txtServidor, $chkSub))
    $frm.Controls.Add($gb)

    $btnGenerar          = New-Object System.Windows.Forms.Button
    $btnGenerar.Text     = 'GENERAR INDICE'
    $btnGenerar.Location = New-Object System.Drawing.Point(14, 224)
    $btnGenerar.Size     = New-Object System.Drawing.Size(160, 34)
    $btnGenerar.Font     = New-Object System.Drawing.Font('Segoe UI', 10, [System.Drawing.FontStyle]::Bold)

    $btnCancelar          = New-Object System.Windows.Forms.Button
    $btnCancelar.Text     = 'Cancelar'
    $btnCancelar.Location = New-Object System.Drawing.Point(184, 224)
    $btnCancelar.Size     = New-Object System.Drawing.Size(100, 34)
    $btnCancelar.Enabled  = $false

    $lblEstado          = New-Object System.Windows.Forms.Label
    $lblEstado.Location = New-Object System.Drawing.Point(300, 234)
    $lblEstado.Size     = New-Object System.Drawing.Size(552, 20)
    $lblEstado.Anchor   = 'Top,Left,Right'
    $lblEstado.ForeColor = [System.Drawing.Color]::DimGray
    $frm.Controls.AddRange(@($btnGenerar, $btnCancelar, $lblEstado))

    $txtLog            = New-Object System.Windows.Forms.TextBox
    $txtLog.Location   = New-Object System.Drawing.Point(14, 268)
    $txtLog.Size       = New-Object System.Drawing.Size(838, 230)
    $txtLog.Multiline  = $true
    $txtLog.ScrollBars = 'Vertical'
    $txtLog.ReadOnly   = $true
    $txtLog.BackColor  = [System.Drawing.Color]::White
    $txtLog.Font       = New-Object System.Drawing.Font('Consolas', 9)
    $txtLog.Anchor     = 'Top,Left,Right,Bottom'
    $frm.Controls.Add($txtLog)

    $btnCarpeta.Add_Click({
        $d = New-Object System.Windows.Forms.FolderBrowserDialog
        $d.Description = 'Elige la carpeta a indexar'
        if ($txtCarpeta.Text) { $d.SelectedPath = $txtCarpeta.Text }
        if ($d.ShowDialog() -eq 'OK') { $txtCarpeta.Text = $d.SelectedPath }
    })
    $btnSalida.Add_Click({
        $d = New-Object System.Windows.Forms.SaveFileDialog
        $d.Filter = 'Indice JSON (*.json)|*.json'
        $d.FileName = [System.IO.Path]::GetFileName($txtSalida.Text)
        if ($d.ShowDialog() -eq 'OK') { $txtSalida.Text = $d.FileName }
    })
    $btnCancelar.Add_Click({ $script:Cancelar = $true })

    $btnGenerar.Add_Click({
        if (-not $txtCarpeta.Text.Trim()) { [void][System.Windows.Forms.MessageBox]::Show('Indica la carpeta a indexar.', 'Falta la carpeta'); return }
        if (-not $txtSalida.Text.Trim())  { [void][System.Windows.Forms.MessageBox]::Show('Indica el archivo JSON de salida.', 'Falta la salida'); return }
        $txtLog.Clear()
        $btnGenerar.Enabled = $false; $btnCancelar.Enabled = $true
        $frm.Cursor = [System.Windows.Forms.Cursors]::WaitCursor

        $log  = { param($t) $txtLog.AppendText("$t`r`n"); [System.Windows.Forms.Application]::DoEvents() }
        $prog = { param($c, $a, $d) $lblEstado.Text = ("{0:N0} carpetas | {1:N0} archivos | {2}" -f $c, $a, $d); [System.Windows.Forms.Application]::DoEvents() }

        try {
            try {
                $cfg = Join-Path $CARPETA_CONFIG 'KIMAK'
                if (-not (Test-Path -LiteralPath $cfg)) { [void](New-Item -ItemType Directory -Path $cfg -Force) }
                ([pscustomobject]@{ Carpeta = $txtCarpeta.Text; Salida = $txtSalida.Text; Extensiones = $txtExt.Text; Servidor = $txtServidor.Text } |
                    ConvertTo-Json) | Set-Content -LiteralPath $ARCHIVO_CONFIG -Encoding UTF8
            } catch { }

            $res = Invoke-GenerarIndice -Carpeta $txtCarpeta.Text.Trim() -Salida $txtSalida.Text.Trim() `
                                        -Extensiones ($txtExt.Text -split '[;,]') -Servidor $txtServidor.Text.Trim() `
                                        -IncluirSubcarpetas $chkSub.Checked -OnLog $log -OnProgreso $prog
            if ($res) {
                [void][System.Windows.Forms.MessageBox]::Show(
                    ("Archivos indexados: {0:N0}`n`nIndice:`n{1}" -f $res.Archivos, $res.Indice),
                    'Indice generado', 'OK', 'Information')
            }
        } catch {
            & $log ('ERROR: ' + $_.Exception.Message)
            [void][System.Windows.Forms.MessageBox]::Show($_.Exception.Message, 'Error', 'OK', 'Error')
        } finally {
            $frm.Cursor = [System.Windows.Forms.Cursors]::Default
            $btnGenerar.Enabled = $true; $btnCancelar.Enabled = $false
        }
    })

    [void]$frm.ShowDialog()
}

# ============================================================================
#  ARRANQUE
# ============================================================================

if ($SinInterfaz) {
    if (-not $Carpeta -or -not $Salida) { throw 'Con -SinInterfaz hay que indicar -Carpeta y -Salida.' }
    $ext = if ($Extensiones) { $Extensiones } else { $EXT_POR_DEFECTO -split ';' }
    $r = Invoke-GenerarIndice -Carpeta $Carpeta -Salida $Salida -Extensiones $ext -Servidor $Servidor `
                              -IncluirSubcarpetas (-not $SoloEstaCarpeta)
    $r | Format-List | Out-String | Write-Host
} else {
    Show-Ventana
}
