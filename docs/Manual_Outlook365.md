



# Outlook365
  
Módulo para realizar acciones en Outlook Office 365  
  
![banner](/docs/imgs/Banner_Outlook365.png)
## Como instalar este módulo
  
__Descarga__ e __instala__ el contenido en la carpeta 'modules' en la ruta de rocketbot.  



## Descripción de los comandos

### Configurar Servidor
  
Configurar Servidor
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|User||user@example.com|
|Timeout||99|
|Contraseña||******|
|Asignar resultado a variable||Variable|

### Enviar Email
  
Envia un email, previamente debe configurar el servidor
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Para||to@mail.com, to2@mail.com|
|Copia||cc@mail.com, cc2@mail.com|
|Copia oculta||bcc@mail.com, bcc2@mail.com|
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
|Mensaje||Esto es una prueba|
|Archivo Adjunto||C:\User\Desktop\test.txt|

### Reenviar email por ID
  
Reenviar email por ID
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|ID Email||355|
|Email||test@email.com|

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
