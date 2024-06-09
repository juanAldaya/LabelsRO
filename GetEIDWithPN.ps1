if (-NOT ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator))
{
    # Get the current script path
    $scriptPath = $myinvocation.mycommand.definition

    # Start a new PowerShell process with elevated privileges
    Start-Process powershell -Verb runAs -ArgumentList "& '$scriptPath'"
    Break  # Exit the current script
}

Install-Module -Name ImportExcel
# Read user input for the Excel name
$excelName = Read-Host "Enter the Excel name"

# Concatenate ".xlsx" to the user-provided name
$excelFile = $excelName + ".xlsx"

# Now you can use $excelFile as the desired Excel file name
# For example, you can import data from the Excel file:
$excelData = Import-Excel -Path $excelFile


foreach ($row in $excelData) {
    if ($row.'Employee Number') {

        $EID = $row.'Requested For'
        $adUser = Get-ADUser -Filter "NAME -eq '$EID'" -Properties a-companyCode -SearchBase "OU=People,DC=dir,DC=svc,DC=accenture,DC=com"
        
        $CC = $adUser.Name
        $row.'CC' = $CC
        #Write-Host "Con el EID: $PN se extrajo el EID $EnterpriseID "
    }
}


try {

    Export-Excel -InputObject $excelData -Path "C:\Induccion25-03-2024.xlsx" -AutoSize -AutoFilter
    Write-Host "Excel file saved successfully."

} catch {
    Write-Host "Error saving the Excel file: $_"
}

