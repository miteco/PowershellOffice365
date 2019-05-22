<#
    Autor:  Luis Felipe Vicente Díez
    Fecha:  22 de Mayo de 2019
    Desc:   Generación de un menú para la GUI de aplicación
#>

Function Invoke-Menu {
    [cmdletbinding()]
    Param(
    [Parameter(Position=0,Mandatory=$True,HelpMessage="Enter your menu text")]
    [ValidateNotNullOrEmpty()]
    [string]$Menu,
    [Parameter(Position=1)]
    [ValidateNotNullOrEmpty()]
    [string]$Title = "Menu",
    [Alias("cls")]
    [switch]$ClearScreen
    )
     
    #clear the screen if requested
    if ($ClearScreen) { 
     Clear-Host 
    }
     
    #build the menu prompt
    $menuPrompt = $title
    #add a return
    $menuprompt+="`n"
    #add an underline
    $menuprompt+="-"*$title.Length
    #add another return
    $menuprompt+="`n"
    #add the menu
    $menuPrompt+=$menu
     
    Read-Host -Prompt $menuprompt
     
    } #end function

$menu=@"

1. Buscar usuario por DNI
2. Buscar usuario por distinguisedName
3. Editar DNI de usuario 
4. Editar usuario por distinguisedName
5. Altas desde CSV
6. Editar desde CSV
7. Exportar usuarios a CSV

R. Conectar con AzureAD
Q. Salir

Seleccione el numero de Tarea o Q para Salir
"@

function showMainMenu {
    
    Do {
        Clear-Host
        #use a Switch construct to take action depending on what menu choice
        #is selected.
        Switch (Invoke-Menu -menu $menu -title "Tareas de Azure AD" -clear ) {
        "1" {   Start-Sleep -seconds 1                 
                Clear-Host
                Write-Host "Buscar Usuario por DNI" -ForegroundColor Yellow
                FindUserByDNI
            } 
        "2" {   Start-Sleep -seconds 1
                Clear-Host
                Write-Host "Buscar Usuario por DistinguisedName" -ForegroundColor Green                
                FindUserByDistinguisedName
            }
        "3" {   Start-Sleep -seconds 1
                Clear-Host
                Write-Host "Actualizar DNI de un Usuario por DistinguisedName" -ForegroundColor  Magenta
                UpdateDniByUser
            }
        "4" {   Start-Sleep -seconds 1
                Clear-Host
                Write-Host "Actualizar Usuario DistinguisedName OPCIÓN NO DISPONIBLE" -ForegroundColor  Red
                Start-Sleep -seconds 1
            }
        "5" {   Start-Sleep -seconds 1
                Clear-Host
                Write-Host "Altas desde CSV" -ForegroundColor  White
                NewUsersCSV
            }
        "6" {   Start-Sleep -seconds 1
                Clear-Host
                Write-Host "Modificaciones desde CSV" -ForegroundColor  Blue
                EditUsersCSV
            }
        "7" {   Start-Sleep -seconds 1
                Clear-Host
                Write-Host "Exportar Usuarios a CSV" -ForegroundColor  Cyan                
                exportUsersCSV
            }
        "R" {   Clear-Host
                Write-Host "Conectando con AzureAD" -ForegroundColor Cyan
                Start-Sleep -seconds 1
                Connect-AzureAD
            }
        "Q" {   Clear-Host
                Write-Host "¡Adios!" -ForegroundColor Cyan     
                Return       
                #Exit-PSHostProcess
            }
        Default {   Write-Warning "Opción invalida. Intentelo de nuevo."
                    Start-Sleep -milliseconds 750}
        } #switch
    } While ($True)

}