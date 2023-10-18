



# Outlook365
  
Send, read, reply to emails and manage your Outlook mailbox.  

*Read this in other languages: [English](Manual_Outlook365.md), [Português](Manual_Outlook365.pr.md), [Español](Manual_Outlook365.es.md)*
  
![banner](imgs/Banner_Outlook365.png)
## How to install this module
  
To install the module in Rocketbot Studio, it can be done in two ways:
1. Manual: __Download__ the .zip file and unzip it in the modules folder. The folder name must be the same as the module and inside it must have the following files and folders: \__init__.py, package.json, docs, example and libs. If you have the application open, refresh your browser to be able to use the new module.
2. Automatic: When entering Rocketbot Studio on the right margin you will find the **Addons** section, select **Install Mods**, search for the desired module and press install.  


## Description of the commands

### Server Configuration
  
Server Configuration
|Parameters|Description|example|
| --- | --- | --- |
|Usuario||user@example.com|
|Timeout||99|
|Password||******|
|Not IMAP connection|If this box is marked, it avoids IMAP connection.||
|Assign result to a Variable||Variable|

### Send Email
  
Send email, before you must configurate the server
|Parameters|Description|example|
| --- | --- | --- |
|To||to@mail.com, to2@mail.com|
|Cc||cc@mail.com, cc2@mail.com|
|Bcc||bcc@mail.com, bcc2@mail.com|
|Subject||Nuevo mail|
|Body||Esto es una prueba|
|Attached File||C:\User\Desktop\test.txt|
|Folder (Multiple files)||C:\User\Desktop\Files|

### List all email
  
List all email, you can specify a filter
|Parameters|Description|example|
| --- | --- | --- |
|Filter||SUBJECT "COMPRA*"|
|Folder||345|
|Asign to var||Variable|

### List unread emails
  
List all unread email, you can specify a filter
|Parameters|Description|example|
| --- | --- | --- |
|Filter||SUBJECT "COMPRA*"|
|Folder||inbox|
|Asign to var||Variable|

### Read email for ID
  
Read email for ID
|Parameters|Description|example|
| --- | --- | --- |
|Email ID||345|
|Folder||inbox|
|Asign to var||Variable|
|Path for download attachment||C:\User\Desktop|
|Email HTML body|If this box is marked, will bring the HTML version of email body.||

### Create Folder
  
Create Folder
|Parameters|Description|example|
| --- | --- | --- |
|Folder Name||Ingrese nombre de la carpeta|

### Move email to folder
  
Move email to folder
|Parameters|Description|example|
| --- | --- | --- |
|Email ID||Ingrese ID del email|
|Folder name to send||test|
|Source folder name||test|
|Asign result to var||Variable|

### Reply email for ID
  
Reply email for ID
|Parameters|Description|example|
| --- | --- | --- |
|Email ID||355|
|Email Folder to reply||inbox|
|Cc||cc@mail.com, cc2@mail.com|
|Bcc||bcc@mail.com, bcc2@mail.com|
|Body||Body|
|Attached File||C:\User\Desktop\test.txt|
|Folder (Multiple files)||C:\User\Desktop\Files|

### Forward email for ID
  
Forward email for ID
|Parameters|Description|example|
| --- | --- | --- |
|Email ID||355|
|Email||test@email.com|
|Cc||cc@mail.com, cc2@mail.com|
|Bcc||bcc@mail.com, bcc2@mail.com|

### List Folders
  
List all Folders
|Parameters|Description|example|
| --- | --- | --- |
|Asign result to var||Variable|

### Mark email as unread
  
Mark email as unread
|Parameters|Description|example|
| --- | --- | --- |
|Folder Name||inbox|
|Email ID||Ingrese ID del email|

### Download attachments by ID
  
Download attachments by ID and save them in the specified folder
|Parameters|Description|example|
| --- | --- | --- |
|Email ID||345|
|Folder||inbox|
|Path for download attachment||C:\User\Desktop|

### Close Server
  
Close server connection
|Parameters|Description|example|
| --- | --- | --- |
