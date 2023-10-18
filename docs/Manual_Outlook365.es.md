



# Outlook365
  
Envia, lee, responde correos y gestiona tu casilla de Outlook.  

*Read this in other languages: [English](Manual_Outlook365.md), [Português](Manual_Outlook365.pr.md), [Español](Manual_Outlook365.es.md)*
  
![banner](imgs/Banner_Outlook365.png)
## Como instalar este módulo
  
Para instalar el módulo en Rocketbot Studio, se puede hacer de dos formas:
1. Manual: __Descargar__ el archivo .zip y descomprimirlo en la carpeta modules. El nombre de la carpeta debe ser el mismo al del módulo y dentro debe tener los siguientes archivos y carpetas: \__init__.py, package.json, docs, example y libs. Si tiene abierta la aplicación, refresca el navegador para poder utilizar el nuevo modulo.
2. Automática: Al ingresar a Rocketbot Studio sobre el margen derecho encontrara la sección de **Addons**, seleccionar **Install Mods**, buscar el modulo deseado y presionar install.  


## Descripción de los comandos

### Configurar Servidor
  
Configurar Servidor
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|User||user@example.com|
|Timeout||99|
|Contraseña||******|
|Sin conexión IMAP|Si se marca esta casilla, evita la conexión IMAP.||
|Asignar resultado a variable||Variable|

### Enviar Email
  
Envia un email, previamente debe configurar el servidor
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Para||to@mail.com, to2@mail.com|
|Cc||cc@mail.com, cc2@mail.com|
|Bcc||bcc@mail.com, bcc2@mail.com|
|Asunto||Nuevo mail|
|Mensaje||Esto es una prueba|
|Archivo Adjunto||C:\User\Desktop\test.txt|
|Carpeta (Varios archivos)||C:\User\Desktop\Files|

### Lista todos los email
  
Lista todos los email, se puede especificar un filtro
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Filtro||SUBJECT "COMPRA*"|
|Carpeta||345|
|Asignar a variable||Variable|

### Lista emails no leídos
  
Lista emails no leídos
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Filtro||SUBJECT "COMPRA*"|
|Carpeta||inbox|
|Asignar a variable||Variable|

### Leer email por ID
  
Leer email por ID
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|ID del email||345|
|Carpeta||inbox|
|Asignar a variable||Variable|
|Ruta para descargar adjuntos||C:\User\Desktop|
|Cuerpo de email en HTML|Si se marca esta casilla, devolvera el cuerpo del email en versión HTML.||

### Crear Carpeta
  
Crea una carpeta
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Nombre Carpeta||Ingrese nombre de la carpeta|

### Mover email a carpeta
  
Mueve email a carpeta
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|ID del email||Ingrese ID del email|
|Carpeta de destino||test|
|Nombre de la carpeta de origen||test|
|Asignar resultado a variable||Variable|

### Responder email por ID
  
Responder email por ID
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|ID Email||355|
|Carpeta del mail a responder||inbox|
|Cc||cc@mail.com, cc2@mail.com|
|Copia oculta||bcc@mail.com, bcc2@mail.com|
|Mensaje||Body|
|Archivo Adjunto||C:\User\Desktop\test.txt|
|Carpeta (Varios archivos)||C:\User\Desktop\Files|

### Reenviar email por ID
  
Reenviar email por ID
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|ID Email||355|
|Email||test@email.com|
|Cc||cc@mail.com, cc2@mail.com|
|Copia oculta||bcc@mail.com, bcc2@mail.com|

### Listar Carpetas
  
Devuelve todas las carpetas
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Asignar resultado a variable||Variable|

### Marcar email como no leído
  
Marcar email como no leído
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Nombre Carpeta||inbox|
|ID del email||Ingrese ID del email|

### Descargar adjuntos por ID
  
Descarga adjuntos por ID y los guardar en la carpeta especificada
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|ID del email||345|
|Carpeta||inbox|
|Ruta para descargar adjuntos||C:\User\Desktop|

### Cerrar Conexión
  
Cierra la conexión del servidor
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
