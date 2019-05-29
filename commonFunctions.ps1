<#
    Autor:  Luis Felipe Vicente Díez
    Fecha:  22 de Mayo de 2019
    Desc:   Colección de Funciones para el funcionamiento de la aplicación de Gestión de Azure AD
#>

function ExportUsersCSV {
    # Desc: Exporta a un fichero CSV la información de los usuarios dados de alta en Azure AD
    
    $file=$env:userprofile + "\Documents\azuread_accounts"+$(get-date -f yyMMddhhmmss)+".csv"

    # Encadena 3 comandos: Get-AzureADUser, Select y export csv
    # Los campos a exportar al csv son valiables en función de las necesidades planteadas
    Get-AzureADUser -All $true | Select UserPrincipalName,@{name="EmployeeId";E={$_.ExtensionProperty["employeeId"]}},`
        DisplayName,GivenName,Surname,MailNickName,Department,JobTitle,TelephoneNumber,PhysicalDeliveryOfficeName,StreetAddress,PostalCode,City,State,Country, `
        Mail,UserType,CreationType,AccountEnabled,`
        @{name='Licensed';expression={if($_.AssignedLicenses){$TRUE}else{$False}}},`
        @{name='Plan';expression={if($_.AssignedPlans){$TRUE}else{$False}}},`
        PreferredLanguage, UsageLocation | export-csv $file -NoTypeInformation -Encoding UTF8 -Delimiter ";" -Verbose        

    Write-Host $file
    Write-Host ""

    Read-Host -Prompt 'Pulse Intro para continuar' 
}

Function FindUserByDNI($employeeId){
    
    # Desc: La función recuepra la información para el DNI buscado. 
    #       El DNI puede llegar como parametro o en caso contrario será la propia función la que los solicite
    # 

    if ($employeeId -eq "" -or $employeeId -eq $null){    
        $employeeId = Read-Host -Prompt "Introduzca el DNI que esta buscando "
    }

    if($employeeId -ne $null -and $employeeId -ne ""){
        
        # Puede que haya más de un usuario con el mismo DNI (no debería). Recuepramos la info como colección de usuarios y la recorremos.
        $userList = (Get-AzureADUser -All $True | Where-Object { $_.extensionProperty.employeeId -eq $EmployeeId } | Select *)        

        if ($userList -ne $null){
            ForEach ($user in $userList){
                Write-Host ""
                Write-Host $user.UserPrincipalName
                Write-Host $user.DisplayName
                Write-Host ""
            }
        }else{            
            Write-Host ""
            Write-Warning "No se ha encontrado el DNI "
            Write-Host ""    
        }
    } else {
        Write-Host ""
        Write-Host "No ha introducido un DNI valido"
        Write-Host ""
        }
        
    Read-Host -Prompt 'Pulse Intro para continuar'
}

Function FindUserByDistinguisedName($dn){
    # Desc: La función recupera la información para el idnetificador único de usario buscado. 
    #       El Identificado puede llegar como parametro o en caso contrario será la propia función la que los solicite
    #

    if($dn -eq "" -or $dn -eq $null ) {
        $dn = Read-Host -Prompt "Introduzca el Identificador que esta buscando "
    }
    
    try {
        $user = (Get-AzureADUser -ObjectId $dn )
    
            if($dn -ne $null -and $dn -ne ""){        
                # De momento lo pintamos como Json aunque puede cambiar en un futuro cuando se concreten los requerimientos
                Write-Host $user.ToJson()        
            } else {
                Write-Warning "No ha introducido un Identificador valido"
            }
            
        } catch{
            Write-Warning "No se ha encontrado el Identificador especificado "
        }   

    Write-Host "" 
    Read-Host -Prompt 'Pulse Intro para continuar'
}
Function UpdateDniByUser{
    # Desc: La función recupera la información para el DNI buscado. 
    #       El DNI puede llegar como parametro o en caso contrario será la propia función la que los solicite
    #
    try {
        $userPrincipalName = Read-Host -Prompt "Introduzca el Usuario que esta buscando "
        Get-AzureADUser -ObjectId $userPrincipalName

        $employeeId = Read-Host -Prompt "Introduzca el DNI que quiere asignar "

        if($employeeId -ne $null -and $employeeId -ne ""){
            
            $employeeId = FormatDNI($employeeId)

            if (ExistsEmployeeId($employeeId) -eq $true){
                Write-Warning -Message "El DNI ya existe asignado a un usuario"
                FindUserByDNI($employeeId)                
            }else{
                $extension = New-Object "System.Collections.Generic.Dictionary``2[System.String,System.String]"
                $extension.Add("employeeId", $employeeId)

                Set-AzureADUser -ObjectId $userPrincipalName -ExtensionProperty $extension
                Write-Host "Actualización completada" -ForegroundColor Green
                Start-Sleep -seconds 1
            }

        }
                
    } catch [Microsoft.Open.AzureAD16.Client.ApiException] {
        Write-Warning -Message "Cuenta no encontrada "
    } catch {
        Write-Error -Message $_.Exception
        throw $_.Exception
    }
}

Function UpdateUserByDistinguisedName{
    Write-Host "Realiza esta operación desde la web de office 365"
    Write-Host ""
    Read-Host -Prompt 'Pulse Intro para continuar'
}

function ComprobarCSV {
    param(
    [string]$CsvFilePath
    )
<#  
    Autor:  Luis Felipe Vicente Díez
    Desc:   Batería de comprobaciones sobre los datos a cargar a partir del csv de carga 
            Si las verificaciones nos son satisfactorias se muestra por pantalla y se genera un fichero con todas las incongruencias detectadas para su resolución 
#>

    # Preparamos el fichero de resultado de las comprobaciones    
    $txtResultsTest =$env:userprofile + "\Documents\comprobaciones365_"+$(get-date -f yyMMddhhmmss)+".csv"

    # Fichero con los usuarios de alta nuevos    
    $NewUsers = import-csv -Path $CsvFilePath -Encoding UTF8 -Delimiter ";"

    $i=0    
    $Mensaje=""

    # Recorremos la colección de nuevos usuarios
    Foreach ($NewUser in $NewUsers) { 
        
        # Mostramos una barra de progreso de la tarea para que el usuario tenga información del estado del proceso
        $i++        
        Write-Progress -Activity 'Comprobando usuarios' -CurrentOperation $NewUser.UserPrincipalName -PercentComplete (($i / $NewUsers.count) * 100) 
        Start-Sleep -Milliseconds 500
        
        $UserPrincipalName = $NewUser.UserPrincipalName
        $EmployeeId = $NewUser.EmployeeId
        $DisplayName = $NewUser.DisplayName
        $GivenName = $NewUser.GivenName
        $Surname = $NewUser.Surname
        $MailNickName = $NewUser.MailNickName
        $Department  = $NewUser.Department
        $JobTitle = $NewUser.JobTitle
        $TelephoneNumber  = $NewUser.TelephoneNumber
        $PhysicalDeliveryOfficeName = $NewUser.PhysicalDeliveryOfficeName
        $StreetAddress = $NewUser.StreetAddress
        $PostalCode = $NewUser.PostalCode
        $City = $NewUser.City
        $State = $NewUser.State
        $Country = $NewUser.Country
        $StandardPack = $NewUser.StandardPack
        $FlowFree =  $NewUser.FlowFree

        $EmployeeId = FormatDNI($EmployeeId)

        # Validamos que los campos requeridos tengan contenido
        if ($UserPrincipalName -eq "" -or $UserPrincipalName -eq $null -or $DisplayName -eq "" -or $DisplayName -eq $null -or $GivenName -eq "" -or $GivenName -eq $null -or
            $Surname -eq "" -or $Surname -eq $null -or $MailNickName -eq "" -or $MailNickName -eq $null -or $Department -eq "" -or $Department -eq $null -or
            $City -eq "" -or $City -eq $null -or $State -eq "" -or $State -eq $null -or $Country -eq "" -or $Country -eq $null ){            
            $Mensaje="El fichero no es valido"
            OutMensaje -NumLinea 0 -Mensaje $Mensaje  -txtResultsPath $txtResultsTest
            break
            
        } else {
            # Comprobar si ya existe el usuario
            if (ExistsUserPrincipalName($UserPrincipalName) -eq $true){            
                $Mensaje="UserPrincipalName ya existe "+$UserPrincipalName
                OutMensaje -NumLinea $i -Mensaje $Mensaje  -txtResultsPath $txtResultsTest
            } else {
                # Comprobamos con una expresión regular si el correo tiene un formato valido abcd@acbd.es 
                if ($UserPrincipalName -notmatch "^[a-zA-Z0-9.!#$%&'*+/=?^_`{|}~-]+@[a-zA-Z0-9-]+(?:\.[a-zA-Z0-9-]+)*$"){
                    $Mensaje="UserPrincipalName no es un correo normalizado " + $UserPrincipalName
                    OutMensaje -NumLinea $i -Mensaje $Mensaje  -txtResultsPath $txtResultsTest
                }
            }
            
            # Comprobaciones EmployeeId
            if ($EmployeeId -ne "" -and $EmployeeId -ne $null){             
                if (ExistsEmployeeId($EmployeeId) -eq $true){
                    $Mensaje="EmployeeId (DNI) ya existe "+$EmployeeId
                    OutMensaje -NumLinea $i -Mensaje $Mensaje  -txtResultsPath $txtResultsTest
                } else {
                    # if ($EmployeeId -notmatch "/[0-9]/"){} # Comprobar DNI

                }
            } else {
                $Mensaje="EmployeeId (DNI) esta vacío"
                OutMensaje -NumLinea $i -Mensaje $Mensaje  -txtResultsPath $txtResultsTest
            }    

            # Comprobaciones DisplayName
            if ($DisplayName -eq "" -or $DisplayName -eq $null) {
                $Mensaje="El DisplayName esta vacío"
                OutMensaje -NumLinea $i -Mensaje $Mensaje  -txtResultsPath $txtResultsTest
            }
            
            # Comprobaciones GivenName
            if ($GivenName -eq "" -or $GivenName -eq $null) {
                $Mensaje="El GivenName esta vacío"
                OutMensaje -NumLinea $i -Mensaje $Mensaje  -txtResultsPath $txtResultsTest
            }

            # Comprobaciones Surname
            if ($Surname -eq "" -or $Surname -eq $null) {
                $Mensaje="El Surname esta vacío"
                OutMensaje -NumLinea $i -Mensaje $Mensaje  -txtResultsPath $txtResultsTest
            }

            # Comprobaciones MailNickName
            if ($MailNickName -eq "" -or $MailNickName -eq $null) {
                $Mensaje="MailNickName esta vacío"
                OutMensaje -NumLinea $i -Mensaje $Mensaje  -txtResultsPath $txtResultsTest
            }

            # Comprobaciones Department
            if ($Department -eq "" -or $Department -eq $null) {
                $Mensaje="Department esta vacío"
                OutMensaje -NumLinea $i -Mensaje $Mensaje  -txtResultsPath $txtResultsTest
            }

            # Comprobaciones JobTitle
            if ($JobTitle -eq "" -or $JobTitle -eq $null) {
                $Mensaje="El JobTitle esta vacío"
                OutMensaje -NumLinea $i -Mensaje $Mensaje  -txtResultsPath $txtResultsTest
            }
        <#
            # Comprobaciones TelephoneNumber
            if ($TelephoneNumber -eq "" -or $TelephoneNumber -eq $null) {
                $Mensaje="TelephoneNumber esta vacío"
                OutMensaje -NumLinea $i -Mensaje $Mensaje  -txtResultsPath $txtResultsTest
            }

            # Comprobaciones PhysicalDeliveryOfficeName
            if ($PhysicalDeliveryOfficeName -eq "" -or $PhysicalDeliveryOfficeName -eq $null) {
                $Mensaje="PhysicalDeliveryOfficeName esta vacío"
                OutMensaje -NumLinea $i -Mensaje $Mensaje  -txtResultsPath $txtResultsTest
            }

            # Comprobaciones StreetAddress
            if ($StreetAddress -eq "" -or $StreetAddress -eq $null) {
                $Mensaje="StreetAddress esta vacío"
                OutMensaje -NumLinea $i -Mensaje $Mensaje  -txtResultsPath $txtResultsTest
            }
        #>
            # Comprobaciones PostalCode
            if ($PostalCode -eq "" -or $PostalCode -eq $null) {
                #$Mensaje="PostalCode esta vacío"
                #OutMensaje -NumLinea $i -Mensaje $Mensaje  -txtResultsPath $txtResultsTest
            }else{
                # Comprobamos que el Código Postal tenga valores númericos
                if($PostalCode -notmatch "^\d+$"){
                    $Mensaje="PostalCode no tiene un valor numerico"
                    OutMensaje -NumLinea $i -Mensaje $Mensaje  -txtResultsPath $txtResultsTest
                }
            }

            # Comprobaciones City
            if ($City -eq "" -or $City -eq $null) {
                $Mensaje="City esta vacío"
                OutMensaje -NumLinea $i -Mensaje $Mensaje  -txtResultsPath $txtResultsTest
            }

            # Comprobaciones State
            if ($State -eq "" -or $State -eq $null) {
                $Mensaje="State esta vacío"
                OutMensaje -NumLinea $i -Mensaje $Mensaje  -txtResultsPath $txtResultsTest
            }

            # Comprobaciones Country
            if ($Country -eq "" -or $Country -eq $null) {
                $Mensaje="Country esta vacío"
                OutMensaje -NumLinea $i -Mensaje $Mensaje  -txtResultsPath $txtResultsTest
            }         

            if (!($StandardPack -eq "true" -or $StandardPack -eq "false")){
                $Mensaje="StandardPack incorrecto"
                OutMensaje -NumLinea $i -Mensaje $Mensaje  -txtResultsPath $txtResultsTest
            }
            if (!($FlowFree -eq "true" -or $FlowFree -eq "false")){
                $Mensaje="FlowFree incorrecto"
                OutMensaje -NumLinea $i -Mensaje $Mensaje  -txtResultsPath $txtResultsTest
            }
        }
    }

    if ($Mensaje -eq ""){
        return $true
    }else{
        return $false
    }
}

function OutMensaje {
    param([int]$NumLinea, [string]$Mensaje, [string]$txtResultsPath) 
    # Rellena el fichero de mensajes de error y los muestra por pantalla
    Write-Host "Línea "$NumLinea". "$Mensaje -ForegroundColor Red
    "Línea "+$NumLinea+". "+$Mensaje | Out-File -FilePath $txtResultsPath -Append

}

Function ExistsUserPrincipalName($UserPrincipalName) {
    # Devuelve un valor booleano en caso de existir el UserPrincipalName
    try{
        if (Get-AzureADUser -ObjectId $UserPrincipalName){
            return $True
        } else {
            return $False
        }
    }catch{
        return $False
    }
}

Function ExistsEmployeeId($EmployeeId) {
    # Devuelve un valor booleano en caso de existir el DNI-EmployeeId
    try{
        if (Get-AzureADUser -All $True | Where-Object { $_.extensionProperty.employeeId -eq $EmployeeId } | Select-Object *){
            return $True
        } else {
            return $False
        }
    }catch{
        return $False
    }
 
}

function CargaUsuariosCSV {
    param(
    [string]$CsvFilePath
    )
    
    $TimeStamp=$(get-date -f yyMMddhhmmss)
    $CsvFilePass ="D:\365Pass"+$TimeStamp+".csv"  

    # Fichero con los usuarios de alta nuevos    
    $NewUsers = import-csv -Path $CsvFilePath -Encoding UTF8 -Delimiter ";"
    
    $i=0

    # Recorremos la colección de nuevos usuarios
    Foreach ($NewUser in $NewUsers) { 

        $i++        
        Write-Progress -Activity 'Comprobando usuarios' -CurrentOperation $NewUser.UserPrincipalName -PercentComplete (($i / $NewUsers.count) * 100)
        Start-Sleep -Milliseconds 500

        $UserPrincipalName = $NewUser.UserPrincipalName.Trim()
        $EmployeeId = $NewUser.EmployeeId.Trim()
        $DisplayName = $NewUser.DisplayName.Trim()
        $GivenName = $NewUser.GivenName.Trim()
        $Surname = $NewUser.Surname.Trim()
        $MailNickName = $NewUser.MailNickName.Trim()
        $Department  = $NewUser.Department.Trim()
        $JobTitle = $NewUser.JobTitle.Trim()
        $TelephoneNumber  = $NewUser.TelephoneNumber
        $PhysicalDeliveryOfficeName = $NewUser.PhysicalDeliveryOfficeName
        $StreetAddress = $NewUser.StreetAddress
        $PostalCode = $NewUser.PostalCode
        $City = $NewUser.City.Trim()
        $State = $NewUser.State.Trim()
        $Country = $NewUser.Country.Trim()
        $StandardPack = $NewUser.StandardPack
        $FlowFree =  $NewUser.FlowFree
            
        $EmployeeId = FormatDNI($EmployeeId)
                    
        # Preparamos una password aleatoria
        $PasswordProfile = New-Object -TypeName Microsoft.Open.AzureAD.Model.PasswordProfile
        $PasswordProfile.Password = Scramble-String($password)
        $PasswordProfile.ForceChangePasswordNextLogin = $true
        
        # El campo es requerido
        if ($PhysicalDeliveryOfficeName -eq ""){$PhysicalDeliveryOfficeName = $null}
        if ($PostalCode -eq ""){$PostalCode = $null}
        if ($StreetAddress -eq ""){$StreetAddress = $null}
        if ($TelephoneNumber -eq ""){$TelephoneNumber = $null}
        
        # Creamos el usuario con los datos del csv
        New-AzureADUser  -AccountEnabled $true  -City $City -Country $Country -Department $Department -DisplayName $DisplayName -GivenName $GivenName -PasswordProfile $PasswordProfile `
            -JobTitle $JobTitle -PhysicalDeliveryOfficeName $PhysicalDeliveryOfficeName -PostalCode $PostalCode -PreferredLanguage "es-ES"  `
            -State $State -StreetAddress $StreetAddress -Surname $Surname -UserPrincipalName $UserPrincipalName -UserType "Member" -MailNickName $MailNickName -UsageLocation "ES" -TelephoneNumber $TelephoneNumber 
            
        
            # Agregamos las passwords al fichero de salida
        $UserPrincipalName+";"+$PasswordProfile.Password| Out-File -FilePath $CsvFilePass -Append

        # Cargamos el DNI del usuario en el objeto creado en la ExtensionProperty EmployeeId
        $user = (Get-AzureADUser -ObjectId $UserPrincipalName)
        $ObjectId = $user.ObjectId
                        
        $extension = New-Object "System.Collections.Generic.Dictionary``2[System.String,System.String]"
        $extension.Add("employeeId", $EmployeeId)

        Set-AzureADUser -ObjectId $ObjectId  -ExtensionProperty $extension

        # Licenciamos el usuario para que tenga disponibles los elementos del paquete Office.

        # Asignamos las licencias de Office365.
        # Es necesario recuperar el Id(SkuId) de la licencia a partir del literal de la misma para poder aignarsela al usuario
        if ($StandardPack -eq "true"){
            $planName="STANDARDPACK"
            $License = New-Object -TypeName Microsoft.Open.AzureAD.Model.AssignedLicense
            $License.SkuId = (Get-AzureADSubscribedSku | Where-Object -Property SkuPartNumber -Value $planName -EQ).SkuID
            $LicensesToAssign = New-Object -TypeName Microsoft.Open.AzureAD.Model.AssignedLicenses
            $LicensesToAssign.AddLicenses = $License
            Set-AzureADUserLicense -ObjectId $ObjectId -AssignedLicenses $LicensesToAssign
        }

        if ($FlowFree -eq "true"){
            $planName="FLOW_FREE"
            $License = New-Object -TypeName Microsoft.Open.AzureAD.Model.AssignedLicense
            $License.SkuId = (Get-AzureADSubscribedSku | Where-Object -Property SkuPartNumber -Value $planName -EQ).SkuID
            $LicensesToAssign = New-Object -TypeName Microsoft.Open.AzureAD.Model.AssignedLicenses
            $LicensesToAssign.AddLicenses = $License

            Set-AzureADUserLicense -ObjectId $ObjectId -AssignedLicenses $LicensesToAssign    
        }
    }

}

function  NewUsersCSV {

    Clear-Host

    # Fichero con los usuarios de alta nuevos
    $CsvFilePath = (Read-Host -Prompt 'Introduzca la ruta del fichero que desea cargar')

    if([System.IO.File]::Exists($CsvFilePath)){

        if (ComprobarCSV -CsvFilePath $CsvFilePath){    
            CargaUsuariosCSV -CsvFilePath $CsvFilePath
        } else {
            Write-Host "No se ha podido realizar la carga. Compruebe los errores." -ForegroundColor Yellow
        }
    }else{
        Write-Host "No se ha encontrado el fichero expecificado " $FileCSV -ForegroundColor Yellow
    }

    Read-Host -Prompt 'Pulse Intro para continuar'

}

function FormatDNI($EmployeeId) {
    # Rellemamos con 0 por la izquierda    
    if ($EmployeeId -ne ""){                        
        $dni = ("00000000"+$EmployeeId)
        $length = $dni.length
        $dni = $dni.Substring($length-9, 9)
        $EmployeeId = $dni
    }
    return $EmployeeId
}