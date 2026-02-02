[CmdletBinding(
    SupportsShouldProcess = $true,  # Allows -WhatIf and -Confirm parameters
    ConfirmImpact = 'Medium'
)]
param( 
    [Parameter( Mandatory = $false,  HelpMessage = 'Specify the OutputFileName.'  )]
    [string]$OutputFileName     = 'output.xlsx',
    [switch]$o
)

begin {
 
    Write-Verbose "Starting script execution."
    $scriptDir = Split-Path -Path $MyInvocation.MyCommand.Definition -Parent
    $SavePath = Join-Path -Path $scriptDir -ChildPath "data\$OutputFileName"
    $logPath = Join-Path -Path $scriptDir -ChildPath 'script.log'
  
    function Write-Log {
        param(
            [string]$Entry
        )
        "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss.fff') $Entry" | Out-File -FilePath $logPath -Append
    } 
} 
process {
    Clear-Host
    try { 
        $config = Get-Content -Raw -Path "$PSScriptRoot\config.json" | ConvertFrom-Json  
        $SourceSheetPath = $config.SourceSheetPath -replace '~', $PSScriptRoot.Replace('\', '\\')
        $LookupSheetPath = $config.LookupSheetPath -replace '~', $PSScriptRoot.Replace('\', '\\')
        $SavePath = $config.SavePath -replace '~', $PSScriptRoot.Replace('\', '\\')
        # 2. Import the CSV natively
        $distinctCodes = Import-Csv -Path $SourceSheetPath | 
                        Select-Object -ExpandProperty $config.SourceColumnName -Unique
        Write-Log "Distinct Source codes extracted: $($distinctCodes -join ', ')."
         
        $lookupData = Import-Csv -Path $LookupSheetPath -ErrorAction Stop
        
        $foundRows = foreach ($code in $distinctCodes) { 
            $lookupData | Where-Object { $_.$($config.LookupColumnName) -eq $code }
        }
        $distinctFoundCodes = $foundRows | Select-Object -ExpandProperty $config.LookupColumnName -Unique
       
        $distinctCodesNotFoundInLookup = $distinctCodes | Where-Object { $_ -notin $distinctFoundCodes }

        if ($foundRows) {
            $foundRows | Export-Csv -Path $SavePath -NoTypeInformation -Force
            Write-Host "Success! $($foundRows.Count) rows saved to $SavePath" -ForegroundColor Green
            $foundRows | Format-Table -AutoSize 
            
            Write-Host "Distinct Found codes: $($distinctFoundCodes -join ', ')."
            Write-Host "Distinct Codes not found in Lookup: $($distinctCodesNotFoundInLookup -join ', ')." -BackgroundColor Red
             Write-Host ""
            if ($o) {
                Start-Process -FilePath $SavePath
            }
        }
        else {
            Write-Warning "No matches were found in the lookup file."
        } 
    }
    catch {
        Write-Error "An error occurred: $($_.Exception.Message)" 
    }
}

end {
    # Cleanup code goes here (e.g., closing connections)
    Write-Verbose "Script execution finished."
}
#endregion Main Script Logic