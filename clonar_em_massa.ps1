try {
    #using
    if (!(Get-Module -Name 'ActiveDirectory')) {
        Install-Module -Name 'ActiveDirectory' -ErrorAction Stop
    }
    
    if (!(Test-Path 'C:\Program Files\WindowsPowerShell\Modules\ImportExcel')) {
        Install-Module -Name 'ImportExcel' -ErrorAction Stop
    }
    
    #main
    
    $endereco = Read-Host "Informe o caminho completo da Planilha"
    
    if (!(Test-Path -LiteralPath $endereco)) {
        throw "Endereco da planilha inexistente"
    }
    
    $template = Read-Host "Digite o login de rede do usuario espelho"

    if (($null -eq (get-aduser $template))) {
        throw "Usuario template nulo ou inexistente, por favor, validar no Active Directory"
    }

    $usuarios = Import-Excel -Path $endereco -ErrorAction Stop
    
    $autenticacao = Get-Credential -ErrorAction Stop
    
    if ($null -eq ($autenticacao.Password -and $autenticacao.UserName)){
        throw "Autenticacao nula"
    }


    $Office = Get-ADUser $template -Properties "Office" | 
                Select-Object -ExpandProperty Office
    $Department = Get-ADUser $template -Properties "Department" | 
                Select-Object -ExpandProperty Department
    $Description = Get-ADUser $template -Properties "Description" | 
                Select-Object -ExpandProperty Description
    $Company = Get-ADUser $template -Properties "Company" | 
                Select-Object -ExpandProperty Company
    $ScriptPath = Get-ADUser $template -Properties "ScriptPath" | 
                Select-Object -ExpandProperty ScriptPath
    

    Write-Host "#Propriedades do usuario Template que serao Clonadas#" -ForegroundColor Green

    Write-Host "#Escritorio: $Office"
    Write-Host "#Departamento: $Department"
    Write-Host "#Descricao: $Description"
    Write-Host "#Empresa: $Company"
    Write-Host "#Script de Logon: $ScriptPath `n"

    $ddNameEspelho = Get-ADUser $template -Properties "DistinguishedName", "DisplayName" | 
                        Select-Object -Property "DistinguishedName", "DisplayName"
    $ouEspelho = $ddNameEspelho.DistinguishedName -replace "CN=$($ddNameEspelho.DisplayName),", ''

    Write-Host "#Unidade Organizacional: $ouEspelho `n"

    $gruposTemplate = Get-ADUser -Identity $template -Properties memberof | 
                        Select-Object -ExpandProperty memberof

    Write-Host "####################################################" -ForegroundColor Green

    foreach ($usuarioAfetado in $usuarios.login) {

       try {

            if ($null -eq $usuarioAfetado) {
                throw "Usuario afetado inexistente"
            }

            function Export-BkpUsuarioAfetado {
                param (
                    [Parameter(Mandatory)]
                    [String]$loginDeRede
                )
        
                $data = ((Get-Date -UFormat "%d-%m-%Y_%T").Replace(":",""))
        
                $log = "$loginDeRede-bkp-$data.txt"
        
                [String]$dados = (Get-ADUser -Identity $loginDeRede -Properties * |
                                    Select-Object -Property *)
        
                $dados.Split(";") *>> $log
        
                Write-Host "`nGrupos:`n" *>> $log 
                $grupos = (Get-ADUser -Identity $loginDeRede -Properties memberof | 
                            Select-Object -ExpandProperty memberof)
                
                    try {
                        foreach ($item in $grupos.Split(",")) {
                            if ($item.Contains("CN=")) {
                                $formatado = "$($item.Replace("CN=",''));"
                                Write-Host "$($formatado.Replace("\",''))" *>> $log
                            }
                        }
                    } catch {
                        Write-Host "Usuario nao tem grupos para fazer bkp"
                    }
                
            }

            
            Export-BkpUsuarioAfetado($usuarioAfetado)

            
            Write-Host "Atualizando os dados do usuario $usuarioAfetado" -ForegroundColor Yellow
            Write-Host "Atualizando Escritorio para: $Office"
            Set-ADUser $usuarioAfetado -Office $Office
            Write-Host "Atualizando Departamento para: $Department"
            Set-ADUser $usuarioAfetado -Department $Department
            Write-Host "Atualizando Escricao para: $Description"
            Set-ADUser $usuarioAfetado -Description $Description
            Write-Host "Atualizando Empresa para: $Company"
            Set-ADUser $usuarioAfetado -Company $Company
            Write-Host "Atualizando Script para: $ScriptPath"
            Set-ADUser $usuarioAfetado -ScriptPath $ScriptPath
            Write-Host "Atualizando OU para: $ouEspelho"
            Move-ADObject -Identity (Get-ADUser $usuarioAfetado |
                                        Select-Object -Property "DistinguishedName").DistinguishedName -TargetPath $ouEspelho
            Write-Host "********************************************" -ForegroundColor Yellow 

            #remove o usuário usuarioAfetado de todos os seus grupos atuais

            $gruposAfetado = (Get-ADUser -Identity $usuarioAfetado -Properties memberof | 
                                Select-Object -ExpandProperty memberof)

            foreach ($grupoAntigo in $gruposAfetado) {

                try {
                    Remove-ADGroupMember -Identity $grupoAntigo -Members $usuarioAfetado -Credential $autenticacao -Confirm:$false
                } catch {
                    Write-Warning "Nao foi possivel remover o usuario $usuarioAfetado do grupoNovo $grupoAntigo"
                    $_.Exception.Message

                } 

            }

            Write-Host "::Removendo os grupos antigos de(a) $usuarioAfetado"

            foreach ($grupoNovo in $gruposTemplate) {

                try {
                    Add-ADGroupMember -Identity $grupoNovo -Members $usuarioAfetado -Credential $autenticacao
                }
                catch {
                    write-warning "Nao foi possivel adicionar o usuario $usuarioAfetado no $grupoNovo"
                    $_.Exception.Message
                }

            }

            Write-Host "::Adicionando os novos grupos em $usuarioAfetado`n"
        }
        catch {
            Write-Host 'Ocorreu uma excecao'
            Write-Host $_
            Write-Host $_.ScriptStackTrace
        }

    }
}
catch {
    Write-Host 'Ocorreu uma excecao durante a execucao'
    Write-Host $_
    Write-Host '********************************'
    Write-Host $_.ScriptStackTrace
}

#Autor: Antônio Marcelo
#contato marcelo.bezerra@aec.com.br