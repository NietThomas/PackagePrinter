<#
.Synopsis
Created on:   31/12/2021
Modified on:  7/11/2024
Created by:   Ben Whitmore
Modified by:  Thomas Vnh
Filename:     Install-Printer.ps1

Simple script to install a network printer from an INF file. The INF and required CAB files hould be in the same directory as the script if creating a Win32app
Also with configuration

#### Win32 app Commands ####

Install:
RAW Printer: %WINDIR%\sysnative\WindowsPowerShell\v1.0\powershell.exe -executionpolicy bypass -file Install.ps1 -Porttype "RAW" -PortName "IP_10.10.1.1" -PrinterIP "10.1.1.1" -PrinterName "Canon Printer Upstairs" -DriverName "Canon Generic Plus UFR II" -INFFile "CNLB0MA64.inf" -Version 1.0.0
LPR Printer: %WINDIR%\sysnative\WindowsPowerShell\v1.0\powershell.exe -executionpolicy bypass -file Install.ps1 -Porttype "LPR" -LprQueueName "LPR" -PortName "IP_10.10.1.1" -PrinterIP "10.1.1.1" -PrinterName "Canon Printer Upstairs" -DriverName "Canon Generic Plus UFR II" -INFFile "CNLB0MA64.inf" -Version 1.0.0

Uninstall:
%WINDIR%\sysnative\WindowsPowerShell\v1.0\powershell.exe -executionpolicy bypass -file Uninstall.ps1 -PrinterName "Canon-Garage-Lade1-A4"

Detection:
HKEY_LOCAL_MACHINE\Software\Microsoft\Windows NT\CurrentVersion\Print\Printers\Canon Printer Upstairs
Name = "Canon Printer Upstairs"

HKEY_LOCAL_MACHINE\SOFTWARE\IntunePrinters\Canon Printer Upstairs
Version = 1.0.0

.Example
RAW Printer: %WINDIR%\sysnative\WindowsPowerShell\v1.0\powershell.exe -executionpolicy bypass -file Install.ps1 -Porttype "RAW" -PortName "IP_10.10.1.1" -PrinterIP "10.1.1.1" -PrinterName "Canon Printer Upstairs" -DriverName "Canon Generic Plus UFR II" -INFFile "CNLB0MA64.inf" -Version 1.0.0
LPR Printer: %WINDIR%\sysnative\WindowsPowerShell\v1.0\powershell.exe -executionpolicy bypass -file Install.ps1 -Porttype "LPR" -LprQueueName "LPR" -PortName "IP_10.10.1.1" -PrinterIP "10.1.1.1" -PrinterName "Canon Printer Upstairs" -DriverName "Canon Generic Plus UFR II" -INFFile "CNLB0MA64.inf" -Version 1.0.0
#>

[CmdletBinding()]
Param (
    [Parameter(Mandatory = $True)]
    [String]$PortName,
    [Parameter(Mandatory = $True)]
    [String]$Porttype,
    [Parameter(Mandatory = $True)]
    [String]$PrinterIP,
    [Parameter(Mandatory = $True)]
    [String]$PrinterName,
    [Parameter(Mandatory = $True)]
    [String]$DriverName,
    [Parameter(Mandatory = $True)]
    [String]$INFFile,
    [Parameter(Mandatory = $True)]
    [String]$Version,
    [Parameter(Mandatory = $False)]
    [String]$LprQueueName
)

$registryPath = "HKLM:\SOFTWARE\IntunePrinters\" + $printerName
$registryName = "Version"

#Reset Error catching variable
$Throwbad = $Null

#Run script in 64bit PowerShell to enumerate correct path for pnputil
If ($ENV:PROCESSOR_ARCHITEW6432 -eq "AMD64") {
    Try {
        &"$ENV:WINDIR\SysNative\WindowsPowershell\v1.0\PowerShell.exe" -File $PSCOMMANDPATH -PortName $PortName -PrinterIP $PrinterIP -DriverName $DriverName -PrinterName $PrinterName -INFFile $INFFile -Version $Version
    }
    Catch {
        Write-Error "Failed to start $PSCOMMANDPATH"
        Write-Warning "$($_.Exception.Message)"
        $Throwbad = $True
    }
}

function Write-LogEntry {
    param (
        [parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$Value,
        [parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [string]$FileName = "$($PrinterName).log",
        [switch]$Stamp
    )

    #Build Log File appending System Date/Time to output
    $LogFile = Join-Path -Path $env:SystemRoot -ChildPath $("Temp\$FileName")
    $Time = -join @((Get-Date -Format "HH:mm:ss.fff"), " ", (Get-WmiObject -Class Win32_TimeZone | Select-Object -ExpandProperty Bias))
    $Date = (Get-Date -Format "MM-dd-yyyy")

    If ($Stamp) {
        $LogText = "<$($Value)> <time=""$($Time)"" date=""$($Date)"">"
    }
    else {
        $LogText = "$($Value)"   
    }
	
    Try {
        Out-File -InputObject $LogText -Append -NoClobber -Encoding Default -FilePath $LogFile -ErrorAction Stop
    }
    Catch [System.Exception] {
        Write-Warning -Message "Unable to add log entry to $LogFile.log file. Error message at line $($_.InvocationInfo.ScriptLineNumber): $($_.Exception.Message)"
    }
}

Write-LogEntry -Value "##################################"
Write-LogEntry -Stamp -Value "Installation started"
Write-LogEntry -Value "##################################"
Write-LogEntry -Value "Install Printer using the following values..."
Write-LogEntry -Value "Port Name: $PortName"
Write-LogEntry -Value "Printer IP: $PrinterIP"
Write-LogEntry -Value "Printer Name: $PrinterName"
Write-LogEntry -Value "Driver Name: $DriverName"
Write-LogEntry -Value "INF File: $INFFile"
Write-LogEntry -Value "Config File: $Printername"
Write-LogEntry -Value "Version: $Version"

$INFARGS = @(
    "/add-driver"
    "$INFFile"
)

If (-not $ThrowBad) {

    Try {

        #Stage driver to driver store
        Write-LogEntry -Stamp -Value "Staging Driver to Windows Driver Store using INF ""$($INFFile)"""
        Write-LogEntry -Stamp -Value "Running command: Start-Process pnputil.exe -ArgumentList $($INFARGS) -wait -passthru"
        Start-Process pnputil.exe -ArgumentList $INFARGS -wait -passthru

    }
    Catch {
        Write-Warning "Error staging driver to Driver Store"
        Write-Warning "$($_.Exception.Message)"
        Write-LogEntry -Stamp -Value "Error staging driver to Driver Store"
        Write-LogEntry -Stamp -Value "$($_.Exception)"
        $ThrowBad = $True
    }
}

If (-not $ThrowBad) {
    Try {
    
        #Install driver
        $DriverExist = Get-PrinterDriver -Name $DriverName -ErrorAction SilentlyContinue
        if (-not $DriverExist) {
            Write-LogEntry -Stamp -Value "Adding Printer Driver ""$($DriverName)"""
            Add-PrinterDriver -Name $DriverName -Confirm:$false
        }
        else {
            Write-LogEntry -Stamp -Value "Print Driver ""$($DriverName)"" already exists. Skipping driver installation."
        }
    }
    Catch {
        Write-Warning "Error installing Printer Driver"
        Write-Warning "$($_.Exception.Message)"
        Write-LogEntry -Stamp -Value "Error installing Printer Driver"
        Write-LogEntry -Stamp -Value "$($_.Exception)"
        $ThrowBad = $True
    }
}

If (-not $ThrowBad) {
    Try {
        #Create Printer Port
        $PortExist = Get-Printerport -Name $PortName -ErrorAction SilentlyContinue
        if ($PortExist) {
            if ($PortExist.Protocol -ne $PortType) {
                Write-LogEntry -Stamp -Value "Removing Port ""$($PortName)"" because protocol does not match"
                Remove-PrinterPort -Name $PortName -Confirm:$false
                $PortExist = $null
            }
        }
        if (-not $PortExist) {
            Write-LogEntry -Stamp -Value "Adding Port ""$($PortName)"""
            if ($PortType -eq "RAW") {
                Add-PrinterPort -Name $PortName -PrinterHostAddress $PrinterIP -Confirm:$false
            }
            elseif ($PortType -eq "LPR") {
                Add-PrinterPort -Name $PortName -LprHostAddress $PrinterIP -Confirm:$false -LprByteCounting -LprQueueName $LprQueueName
            }
        }
    }
    Catch {
        Write-Warning "Error creating Printer Port"
        Write-Warning "$($_.Exception.Message)"
        Write-LogEntry -Stamp -Value "Error creating Printer Port"
        Write-LogEntry -Stamp -Value "$($_.Exception)"
        $ThrowBad = $True
    }
}

If (-not $ThrowBad) {
    Try {

        #Add Printer
        $PrinterExist = Get-Printer -Name $PrinterName -ErrorAction SilentlyContinue
        if (-not $PrinterExist) {
            Write-LogEntry -Stamp -Value "Adding Printer ""$($PrinterName)"""
            Add-Printer -Name $PrinterName -DriverName $DriverName -PortName $PortName -Confirm:$false
        }
        else {
            Write-LogEntry -Stamp -Value "Printer ""$($PrinterName)"" already exists. Removing old printer..."
            Remove-Printer -Name $PrinterName -Confirm:$false
            Write-LogEntry -Stamp -Value "Adding Printer ""$($PrinterName)"""
            Add-Printer -Name $PrinterName -DriverName $DriverName -PortName $PortName -Confirm:$false
        }

        $PrinterExist2 = Get-Printer -Name $PrinterName -ErrorAction SilentlyContinue
        if ($PrinterExist2) {
            Write-LogEntry -Stamp -Value "Printer ""$($PrinterName)"" added successfully"
        }
        else {
            Write-Warning "Error creating Printer"
            Write-LogEntry -Stamp -Value "Printer ""$($PrinterName)"" error creating printer"
            $ThrowBad = $True
        }
    }
    Catch {
        Write-Warning "Error creating Printer"
        Write-Warning "$($_.Exception.Message)"
        Write-LogEntry -Stamp -Value "Error creating Printer"
        Write-LogEntry -Stamp -Value "$($_.Exception)"
        $ThrowBad = $True
    }
}


If (-not $ThrowBad) {
    Try {

        #Add config
        RUNDLL32 PRINTUI.DLL,PrintUIEntry /Sr /n $Printername /a $PrinterName d g r
    }
    Catch {
        Write-Warning "Error adding configuration Printer"
        Write-Warning "$($_.Exception.Message)"
        Write-LogEntry -Stamp -Value "Error adding configuration Printer"
        Write-LogEntry -Stamp -Value "$($_.Exception)"
        $ThrowBad = $True
    }
}

If (-not $ThrowBad) {
    Try {

        #Add Version number
        if (-not (Test-Path $registryPath)) {
        New-Item -Path $registryPath -Force
    }

        #Stel de registerwaarde in
        Set-ItemProperty -Path $registryPath -Name $registryName -Value $Version
    }
    Catch {
        Write-Warning "Error adding regkey"
        Write-Warning "$($_.Exception.Message)"
        Write-LogEntry -Stamp -Value "Error adding regkey"
        Write-LogEntry -Stamp -Value "$($_.Exception)"
        $ThrowBad = $True
    }
}

If ($ThrowBad) {
    Write-Error "An error was thrown during installation. Installation failed. Refer to the log file in %temp% for details"
    Write-LogEntry -Stamp -Value "Installation Failed"
}
