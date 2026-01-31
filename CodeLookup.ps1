[CmdletBinding(
    SupportsShouldProcess = $true,  # Allows -WhatIf and -Confirm parameters
    ConfirmImpact = 'Medium'
)]
param( 
    [Parameter( Mandatory = $false,  HelpMessage = 'Specify the OutputFileName.'  )]
    [string]$OutputFileName     = 'output.xlsx'
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
 
    try { 
        $config = Get-Content -Raw -Path "$PSScriptRoot\config.json" | ConvertFrom-Json  
 
        # 2. Import the CSV natively
        $distinctCodes = Import-Csv -Path $config.SourceSheetPath | 
                        Select-Object -ExpandProperty $config.SourceColumnName -Unique
        Write-Log "Distinct codes extracted: $($distinctCodes -join ', ')."
         
        $lookupData = Import-Csv -Path $config.LookupSheetPath -ErrorAction Stop
   
        $foundRows = foreach ($code in $distinctCodes) { 
            $lookupData | Where-Object { $_.Code -eq $code }
        }
 
        if ($foundRows) {
            $foundRows | Export-Csv -Path $config.SavePath -NoTypeInformation -Force
            Write-Host "Success! $($foundRows.Count) rows saved to $($config.SavePath)" -ForegroundColor Green
            $foundRows | Format-Table -AutoSize
            Start-Process -FilePath $config.SavePath
        }
        else {
            Write-Warning "No matches were found in the lookup file."
        } 
    }
    catch {
        Write-Error "An error occurred: $($_.Exception.Message)"
        # Exit with a non-zero code to indicate failure if needed
        # exit 1
    }
}

end {
    # Cleanup code goes here (e.g., closing connections)
    Write-Verbose "Script execution finished."
}
#endregion Main Script Logic