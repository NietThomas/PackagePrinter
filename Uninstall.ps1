<#
.Synopsis
Created on:   31/12/2021
Created by:   Ben Whitmore
Modified on:  17/03/2025
Modified by:  Thomas Vanhooren
Filename:     Uninstall.ps1

powershell.exe -executionpolicy bypass -file .\Uninstall.ps1 -PrinterName "Canon Printer Upstairs"

.Example
.\Uninstall.ps1 -PrinterName "Canon Printer Upstairs"
#>

[CmdletBinding()]
Param (
    [Parameter(Mandatory = $True)]
    [String]$PrinterName
)
# Uninstalls the printer
Try {
    # Remove Printer
    $PrinterExist = Get-Printer -Name $PrinterName -ErrorAction SilentlyContinue
    $registryPath = "HKLM:\SOFTWARE\IntunePrinters\" + $PrinterName
    if ($PrinterExist) {
        Remove-Printer -Name $PrinterName -Confirm:$false
        Get-Item $registryPath | Remove-Item -Force -Verbose
    } else {
        Write-Warning "Printer '$PrinterName' does not exist."
    }
}
Catch {
    Write-Warning "Error removing Printer: $($_.Exception.Message)"
}
