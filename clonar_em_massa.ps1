try {
    #dependencias
    if (!(Get-Module -Name 'ActiveDirectory')) {
        Install-Module -Name 'ActiveDirectory' -ErrorAction Stop
    }
    
    if (!(Test-Path 'C:\Program Files\WindowsPowerShell\Modules\ImportExcel')) {
        Install-Module -Name 'ImportExcel' -ErrorAction Stop
    }
    
    #processamento
    [String]$endereco = Read-Host "Informe o caminho completo da Planilha"
    [String]$template = Read-Host "Digite o login de rede do usuario espelho"

    if (($null -eq (get-aduser $template))) {
        throw "varivel do usuario espelho recebeu um valor nulo"
    } elseif ($null -eq ($autenticacao.Password -and $autenticacao.UserName)){
        throw "autenticacao vazia"
    } elseif (!(Test-Path -LiteralPath $endereco)) {
        throw "endereco da planilha inexistente"
    }

    $usuarios = Import-Excel -Path $endereco -ErrorAction Stop
    $autenticacao = Get-Credential

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
    

    Write-Host "Escritorio: $Office" -ForegroundColor Green
    Write-Host "Departamento: $Department" -ForegroundColor Green
    Write-Host "Descricao: $Description" -ForegroundColor Green
    Write-Host "Empresa: $Company" -ForegroundColor Green
    Write-Host "Script de Logon: $ScriptPath `n" -ForegroundColor Green

    $ddNameEspelho = Get-ADUser $template -Properties "DistinguishedName", "DisplayName" | 
                        Select-Object -Property "DistinguishedName", "DisplayName"
    $ouEspelho = $ddNameEspelho.DistinguishedName -replace "CN=$($ddNameEspelho.DisplayName),", ''

    Write-Host "Unidade Organizacional: $ouEspelho `n" -ForegroundColor Green

    $gruposTemplate = Get-ADUser -Identity $template -Properties memberof | 
                        Select-Object -ExpandProperty memberof


    foreach ($usuarioAfetado in $usuarios.login) {

       try {

            if ($null -eq $usuarioAfetado) {
                throw "Usuario Inexistente"
            }

            function Export-Bkp {
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

            
                Export-Bkp($usuarioAfetado)

            
            Write-Host "atualizando os dados do usuario $usuarioAfetado" -ForegroundColor Yellow
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
                    write-warning "Nao foi possivel remover o usuario $usuarioAfetado do grupoNovo $grupoAntigo"
                    $_.Exception.Message

                } 

            }
            Write-Host "Removendo usuario usuarioAfetado dos seus antigos grupos..."

            foreach ($grupoNovo in $gruposTemplate) {

                try {
                    Add-ADGroupMember -Identity $grupoNovo -Members $usuarioAfetado -Credential $autenticacao
                }
                catch {
                    write-warning "Nao foi possivel adicionar o usuario $usuarioAfetado no $grupoNovo"
                    $_.Exception.Message
                }

            }

            Write-Host "Adicionando usuario usuarioAfetado nos grupos do template...`n"
        }
        catch {
            Write-Host "Ocorreu uma excecao"
            Write-Host $_
            Write-Host $_.ScriptStackTrace
        }

    }
}
catch {
    Write-Host 'Ocorreu uma excecao durante a execucao'
    Write-Host $_
    Write-Host "********************************"
    Write-Host $_.ScriptStackTrace
}

#Autor: Antônio Marcelo
#contato marcelo.bezerra@aec.com.br