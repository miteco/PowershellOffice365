Function NewCSV {
    Write-Host "Ha seleccionado la carga masiva desde CSV"
    Write-Host ""
    Read-Host -Prompt 'Pulse Intro para continuar'
}

Function UpdateCSV {
    Write-Host "Ha seleccionado actualización masiva desde CSV"
    Write-Host ""
    Read-Host -Prompt 'Pulse Intro para continuar'
}

Function SearchUserByDNI{
    Write-Host "Ha seleccionado actualización masiva desde CSV"
    Write-Host ""
    Read-Host -Prompt 'Pulse Intro para continuar'
}

Function SearchUserByDistinguisedName{
    Write-Host "Ha seleccionado actualización masiva desde CSV"
    Write-Host ""
    Read-Host -Prompt 'Pulse Intro para continuar'
}
Function UpdateUserByDNI{
    Write-Host "Ha seleccionado actualización masiva desde CSV"
    Write-Host ""
    Read-Host -Prompt 'Pulse Intro para continuar'
}

Function UpdateUserByDistinguisedName{
    Write-Host "Ha seleccionado actualización masiva desde CSV"
    Write-Host ""
    Read-Host -Prompt 'Pulse Intro para continuar'
}


Function main {
    Clear-Host

    Do {
        Write-Host "Menu de identidades"
        Write-Host "==================="
        Write-Host ""
        Write-Host "1. Buscar usuario por DNI"
        Write-Host "2. Buscar usuario por distinguisedName"
        Write-Host "3. Editar usuario por DNI"
        Write-Host "4. Editar usuario por distinguisedName"
        Write-Host "5. Altas desde CSV"
        Write-Host "6. Editar desde CSV"
        Write-Host ""
        Write-Host "E. Salir"
        Write-Host ""
        $opt=Read-Host -Prompt 'Introduzca opcion'

        Clear-Host
        if($opt -eq 1){          
            SearchUserByDNI
        } elseif ($opt -eq 2) {
            SearchUserByDistinguisedName
        } elseif ($opt -eq 3) {
            UpdateUserByDNI
        } elseif ($opt -eq 4) {
            UpdateUserByDistinguisedName
        } elseif ($opt -eq 5) {
            NewCSV
        } elseif ($opt -eq 6) {
            UpdateCSV
        } elseif ($opt -eq "E") {
            Write-Host "Adios!!"                
        }

    } While($opt -ne "E") 

    
    
    Exit-PSHostProcess

}

main
