Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Controleer of IntuneWinAppUtil.exe aanwezig is, anders downloaden
$IntuneToolPath = "C:\tools\IntuneWinAppUtil.exe"
if (-Not (Test-Path $IntuneToolPath)) {
    Write-Host "🔍 IntuneWinAppUtil.exe niet gevonden. Downloaden van GitHub..."

    $toolsDir = "C:\tools"
    New-Item -ItemType Directory -Path $toolsDir -Force | Out-Null

    try {
        $releaseInfo = Invoke-RestMethod -Uri "https://api.github.com/repos/microsoft/Microsoft-Win32-Content-Prep-Tool/releases/latest"
        $zipAsset = $releaseInfo.assets | Where-Object { $_.name -like '*.zip' } | Select-Object -First 1

        if ($zipAsset -ne $null) {
            $zipPath = "$env:TEMP\IntuneWinAppUtil.zip"
            Invoke-WebRequest -Uri $zipAsset.browser_download_url -OutFile $zipPath
            Expand-Archive -Path $zipPath -DestinationPath $toolsDir -Force
            Remove-Item $zipPath
            Write-Host "✅ IntuneWinAppUtil.exe succesvol gedownload en uitgepakt naar $toolsDir"
        } else {
            Write-Error "❌ Geen ZIP-bestand gevonden in de laatste GitHub-release."
            exit 1
        }
    } catch {
        Write-Error "❌ Fout bij downloaden van IntuneWinAppUtil.exe: $_"
        exit 1
    }
} else {
    Write-Host "✅ IntuneWinAppUtil.exe al aanwezig in $IntuneToolPath"
}

# GUI voor printerselectie
$form = New-Object System.Windows.Forms.Form
$form.Text = "Selecteer printers"
$form.Size = New-Object System.Drawing.Size(400,400)
$form.StartPosition = "CenterScreen"

$listbox = New-Object System.Windows.Forms.CheckedListBox
$listbox.Size = New-Object System.Drawing.Size(360,280)
$listbox.Location = New-Object System.Drawing.Point(10,10)
$listbox.CheckOnClick = $true

$printers = Get-Printer | Select-Object -ExpandProperty Name
$printers | ForEach-Object { [void]$listbox.Items.Add($_) }
$form.Controls.Add($listbox)

$okButton = New-Object System.Windows.Forms.Button
$okButton.Text = "OK"
$okButton.Location = New-Object System.Drawing.Point(150,310)
$okButton.Add_Click({ $form.Close() })
$form.Controls.Add($okButton)

$form.ShowDialog() | Out-Null
$selectedPrinters = $listbox.CheckedItems

# Padstructuur
$BaseOutput = "C:\temp\Printers"
$Version = "1.0.0"

foreach ($printer in $selectedPrinters) {
    try {
        Write-Host "`nVerwerken van printer: $printer"
        $printerDir = Join-Path $BaseOutput $printer
        $outputDir = Join-Path $BaseOutput "output\$printer"
        New-Item -ItemType Directory -Path $printerDir -Force | Out-Null
        New-Item -ItemType Directory -Path $outputDir -Force | Out-Null

        try {
            $printerObj = Get-Printer -Name $printer
            $driverName = $printerObj.DriverName
            $portName = $printerObj.PortName
            $portObj = Get-PrinterPort -Name $portName
            $printerIP = $portObj.PrinterHostAddress
            $portType = if ($portObj.SNMPEnabled -eq $false -and $portObj.Protocol -eq 1) { "RAW" } elseif ($portObj.Protocol -eq 2) { "LPR" } else { "Onbekend" }
            $lprQueue = if ($portType -eq "LPR") { $portObj.LprQueue } else { "" }
            $infPath = (Get-PrinterDriver -Name $driverName).InfPath
            $infFile = Split-Path $infPath -Leaf
            $driverFolder = Split-Path $infPath -Parent
        } catch {
            Write-Error "❌ Fout bij ophalen van driverinformatie voor printer ${printer}: $_"
            continue
        }

        try {
            Copy-Item -Path "$driverFolder\*" -Destination $printerDir -Recurse -Force
        } catch {
            Write-Error "❌ Fout bij kopiëren van driverbestanden voor ${printer}: $_"
            continue
        }

        try {
            $configFile = Join-Path $printerDir $printer
            Start-Process -FilePath "RUNDLL32" -ArgumentList "PRINTUI.DLL,PrintUIEntry /Ss /n `"$printer`" /a `"$configFile`" d g" -Wait
        } catch {
            Write-Error "❌ Fout bij exporteren van configuratiebestand voor ${printer}: $_"
            continue
        }

        try {
            Invoke-WebRequest -Uri "https://raw.githubusercontent.com/NietThomas/PackagePrinter/main/Install.ps1" -OutFile (Join-Path $printerDir "Install.ps1") -ErrorAction Stop
            Invoke-WebRequest -Uri "https://raw.githubusercontent.com/NietThomas/PackagePrinter/main/Uninstall.ps1" -OutFile (Join-Path $printerDir "Uninstall.ps1") -ErrorAction Stop
        } catch {
            Write-Error "❌ Fout bij downloaden van install/uninstall scripts voor ${printer}: $_"
            continue
        }

        try {
            Start-Process -FilePath $IntuneToolPath -ArgumentList "-c `"$printerDir`" -s Install.ps1 -o `"$outputDir`"" -Wait
            Write-Host " ===================================================================================================================================="
            Write-Host "✅ Printer '$printer' succesvol verpakt in $outputDir"

            # Samenvatting
            Write-Host ""
            Write-Host "📦 Samenvatting voor '${printer}':"
            Write-Host "   ➤ IP-adres: $printerIP"
            Write-Host "   ➤ INF-bestand: $infFile"
            Write-Host "   ➤ Drivernaam: $driverName"
            Write-Host "   ➤ Poorttype: $portType"
            if ($portType -eq "LPR") {
                Write-Host "   ➤ LPR Queue: $lprQueue"
            }
            Write-Host ""

            $extraLpr = ""
            if ($portType -eq "LPR") {
                $extraLpr = '-LprQueueName "' + $lprQueue + '" '
            }

            $installCmd = 'C:\Windows\sysnative\WindowsPowerShell\v1.0\powershell.exe -executionpolicy bypass -file Install.ps1 -PortType "' + $portType + '" ' +
                          $extraLpr +
                          '-PortName "' + $portName + '" -PrinterIP "' + $printerIP + '" -PrinterName "' + $printer + '" -DriverName "' + $driverName + '" -INFFile "' + $infFile + '" -Version ' + $Version

            $uninstallCmd = 'C:\Windows\sysnative\WindowsPowerShell\v1.0\powershell.exe -executionpolicy bypass -file Uninstall.ps1 -PrinterName "' + $printer + '"'

            Write-Host "🛠️ Installatiecommando:"
            Write-Host "   $installCmd"
            Write-Host ""
            Write-Host "🗑️ Uninstall commando:"
            Write-Host "   $uninstallCmd"
            Write-Host ""
            Write-Host "🔍 Detectieregel 1:"
            Write-Host "   Key Path:   HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Print\Printers\$printer"
            Write-Host "   Name:       Name"
            Write-Host "   Operator:   Equals"
            Write-Host "   Value:      $printer"
            Write-Host ""
            Write-Host "🔍 Detectieregel 2:"
            Write-Host "   Key Path:   HKEY_LOCAL_MACHINE\SOFTWARE\IntunePrinters\$printer"
            Write-Host "   Name:       Version"
            Write-Host "   Operator:   Equals"
            Write-Host "   Value:      $Version"
            Write-Host " ===================================================================================================================================="

        } catch {
            Write-Error "❌ Fout bij packagen van printer ${printer}: $_"
            continue
        }

    } catch {
        Write-Error "❌ Onverwachte fout bij verwerken van printer ${printer}: $_"
    }
}
