



# Outlook365

Envie, leia, responda e-mails e gerencie sua caixa de correio do Outlook.

*Read this in other languages: [English](Manual_Outlook365.md), [Português](Manual_Outlook365.pr.md), [Español](Manual_Outlook365.es.md)*

![banner](imgs/Banner_Outlook365.png)
## Como instalar este módulo

Para instalar o módulo no Rocketbot Studio, pode ser feito de duas formas:
1. Manual: __Baixe__ o arquivo .zip e descompacte-o na pasta módulos. O nome da pasta deve ser o mesmo do módulo e dentro dela devem ter os seguintes arquivos e pastas: \__init__.py, package.json, docs, example e libs. Se você tiver o aplicativo aberto, atualize seu navegador para poder usar o novo módulo.
2. Automático: Ao entrar no Rocketbot Studio na margem direita você encontrará a seção **Addons**, selecione **Install Mods**, procure o módulo desejado e aperte instalar.

## Como utilizar este módulo

Para criar uma nova palavra-passe de aplicação para uma aplicação ou dispositivo, siga os passos abaixo:
1. Vá para a página de noções básicas de segurança (https://account.microsoft.com/security?refd=support.microsoft.com) e faça login na sua conta Microsoft.
2. Selecione Mais opções de segurança.
3. Em Senhas de aplicativo, selecione Criar uma nova senha de aplicativo. Uma nova senha de aplicativo é gerada e aparece na tela.
4. Digite esta senha de aplicativo onde você digitaria sua senha normal da conta Microsoft no aplicativo.


## Descrição do comando

### Configuração do Servidor

Configuração do Servidor
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|Usuário||user@example.com|
|Timeout||99|
|Senha||******|
|Não é conexão IMAP|Se esta caixa estiver marcada, evita a conexão IMAP.||
|Atribuir resultado a uma variável||Variable|

### Enviar Email

Envie um email, antes você deve configurar o servidor
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|Para||to@mail.com, to2@mail.com|
|Cc||cc@mail.com, cc2@mail.com|
|Bcc||bcc@mail.com, bcc2@mail.com|
|Assunto||Nuevo mail|
|Corpo||Esto es una prueba|
|Arquivo Anexo||C:\User\Desktop\test.txt|
|Pasta (Vários arquivos)||C:\User\Desktop\Files|
|Corpo do email contém HTML|Marque esta caixa se o corpo do email contiver HTML.||

### Lista todos os email

Lista todos os email, você pode especificar um filtro
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|Filtro||SUBJECT "COMPRA*"|
|Pasta||345|
|Atribuir a variável||Variable|

### Lista emails não lidos

Lista emails não lidos
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|Filtro||SUBJECT "COMPRA*"|
|Pasta||inbox|
|Atribuir a variável||Variable|

### Ler email por ID

Ler email por ID
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|ID do email||345|
|Pasta||inbox|
|Atribuir a variável||Variable|
|Caminho para baixar anexos||C:\User\Desktop|
|Corpo do email em HTML|Se esta caixa for marcada, trará a versão HTML do corpo do email.||

### Criar Pasta

Cria uma pasta
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|Nome da Pasta||Ingrese nombre de la carpeta|

### Mover email para pasta

Move email para pasta
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|ID do email||Ingrese ID del email|
|Nome da pasta para enviar||test|
|Nome da pasta de origem||test|
|Atribuir resultado para variável||Variable|

### Responder email por ID

Responder email por ID
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|ID do email||355|
|Pasta do email para responder||inbox|
|Cc||cc@mail.com, cc2@mail.com|
|Cópia oculta||bcc@mail.com, bcc2@mail.com|
|Corpo||Body|
|Arquivo Anexo||C:\User\Desktop\test.txt|
|Pasta (Vários arquivos)||C:\User\Desktop\Files|

### Encaminhar email por ID

Encaminhar email por ID
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|ID do email||355|
|Email||test@email.com|
|Cc||cc@mail.com, cc2@mail.com|
|Bcc||bcc@mail.com, bcc2@mail.com|

### Listar Pastas

Retorna todas as pastas
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|Atribuir resultado para variável||Variable|

### Marcar email como não lido

Marcar email como não lido
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|Nome da Pasta||inbox|
|ID do email||Ingrese ID del email|

### Baixar anexos por ID

Baixe anexos por ID e salve-os na pasta especificada
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|ID do email||345|
|Pasta||inbox|
|Caminho para baixar anexos||C:\User\Desktop|

### Fechar Conexão

Fecha a conexão do servidor
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
