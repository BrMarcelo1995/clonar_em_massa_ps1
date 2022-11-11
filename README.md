# clonar_em_massa_ps1


<h3> Descrição do Projeto </h3>

Script feito para clonar configurações de perfil de usuário em grande quantidade,
bastando inserir os logins afetados na coluna login da planilha antes da execução e ao executar passar os parâmetro do caminho da planilha e o login template
que servirá como espelho para os afetados.



<h3>Configuração do ambiente</h3>

1º Abra o powershell como administrador e conceda permissão para execução de script remoto com o comando "Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser"

2º ainda com o powershell como administrador execute o script para instalação das dependências declaradas no script, caso ocorrer problema para importar o módulo do ActiveDirectory recomendo realizar a instalação manual abrindo o painel de configurações do windows, buscar por recursos opcionais e instalar o pacote de ferramenta abaixo:
![image](https://user-images.githubusercontent.com/32343597/201237736-f5850b4a-f812-4f02-9035-79f093a4c377.png)

Após isso o ambiente está configurado e pronto para execução do script

<h3> Campos Clonados do Usuário Template/Espelho </h3>

<ol>
<li>Escritorio</li>
<li>Departamento</li>
<li>Descrição</li>
<li>Empresa</li>
<li>Script de Logon</li>
<li>Unidade Organizacional</li>
</ol>
