Clear-Host

# Abrimos la conexi?n con azureAD en caso de no estar conectado
if($azureConnection.Account -eq $null){
    $azureConnection = Connect-AzureAD
}    

# Ficheros requeridos con Funciones auxiliares 
$ScriptDirectory = Split-Path -Path $MyInvocation.MyCommand.Definition -Parent
try {
    . ("$ScriptDirectory\menu.ps1")    
    . ("$ScriptDirectory\commonFunctions.ps1")  
    . ("$ScriptDirectory\RandomPasswords.ps1")         
}
catch {
    Write-Host "Error cargando ficheros auxiliares PowerShell Scripts" -ForegroundColor Red
}
#endregion

# lanzamos la aplicaci?n llamando al men? principal
showMainMenu
