if (-NOT ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator))
{
    # Get the current script path
    $scriptPath = $myinvocation.mycommand.definition

    # Start a new PowerShell process with elevated privileges
    Start-Process powershell -Verb runAs -ArgumentList "& '$scriptPath'"
    Break  # Exit the current script
}

#Install-Module -Name ImportExcel
# Read user input for the Excel name
$excelName = Read-Host "Enter the Excel name"

# Construct the full path to the Excel file in the same folder
$excelFile = Join-Path $PSScriptRoot "$excelName.xlsx"

# Now you can use $excelFile as the desired Excel file path
# For example, you can import data from the Excel file:
$excelData = Import-Excel -Path $excelFile


foreach ($row in $excelData) {
    if ($row.'Requested For') {
        
        $EID = $row.'Requested For'
        Write-Host "Requested For: $EID"
        $adUser = Get-ADUser -Filter "NAME -eq '$EID'" -Properties a-companyCode -SearchBase "OU=People,DC=dir,DC=svc,DC=accenture,DC=com"
        
        $CC = $adUser.'a-companyCode'
        $row.CC = $CC
        Write-Host "Con el EID: $EID se extrajo el CC $CC "
    }
}


try {

    Export-Excel -InputObject $excelData -Path $excelFile -WorksheetName "Page 1" -AutoSize -AutoFilter
    Write-Host "Excel file saved successfully."

} catch {
    Write-Host "Error saving the Excel file: $_"
}

