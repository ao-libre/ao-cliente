# AO-LIBRE 800X600

Este fork simplemente trae de nuevo el viejo frm para jugar a 800x600

TODO:
-Poner los botones de desafios y amigos

# IMPORTANTE

Para que este fork funcione tenes que ir al server y cambiar las constantes del mod Areas para adaptarse a la nueva resolucion del render
Cambiar:

Public Const XWindow            As Byte = 23

Public Const YWindow            As Byte = 19


Por:

Public Const XWindow            As Byte = 17

Public Const YWindow            As Byte = 13

# Cliente Argentum Online Libre
Importante, no bajar el codigo con el boton Download as a ZIP de github por que lo descarga mal, muchos archivos por el encoding quedan corruptos.

Tenes que bajar el codigo con un cliente de git, con el cliente original de la linea de comandos seria:
```
git clone https://www.github.com/ao-libre/ao-cliente
```


![AO Logo](https://ao-libre.github.io/ao-website/assets/images/logo.png)

## Wiki Desarrollo Argentum Online
[Manual para entender el codigo de Argentum Online](http://es.dao.wikia.com/wiki/Wiki_Desarrollo_Argentum_Online).

## Diagrama Arquitectura Aplicaciones AO-LIBRE
https://www.reddit.com/r/argentumonlineoficial/comments/f402p9/argentum_online_libre_diagrama_arquitectura/

## Logs publicos de nuestro Server
AO es un juego open-source y por ello abrimos nuestros logs del server al publico para que puedan ver que errores hay en el servidor y poder ayudar a repararlos 

- http://argentumonline.org/logs-desarrollo.html
- http://argentumonline.org/logs-gms.html
- http://argentumonline.org/logs-errores.html
- http://argentumonline.org/logs-statistics.html

## F.A.Q:

#### Error - Al abrir el proyecto en Visual Basic 6 no puede cargar todas las dependencias:
Este es un error comun que les suele pasar a varias personas, esto es debido que el EOL del archivo esta corrupto.
Visual Basic 6 lee el .vbp en CLRF, hay varias formas de solucionarlo:

Opcion a:
Con Notepad++ cambiar el EOL del archivo a CLRF

Opcion b:
Abrir un editor de texto y reemplazar todos los `'\n'` por `'\r\n'`

## Autoupdates:

El programa al iniciar comparara la version del programa que se encuentra en `INIT/Config.ini` en el parámetro [version](https://github.com/ao-libre/ao-cliente/blob/master/INIT/Config.ini) con la ultima version que se encuentra en el [Endpoint Github Releases](https://api.github.com/repos/ao-libre/ao-cliente/releases/latest). En caso de ser diferente, se ejecuta nuestro programa `ao-autoupdate` para poder hacer el update.

Para mas información sobre este proceso:

[Funcion para comparar versiones](https://github.com/ao-libre/ao-cliente/blob/master/CODIGO/Formularios/frmCargando.frm#L121)

[Codigo fuente ao-autoupdate](https://github.com/ao-libre/ao-autoupdate)

## Revisar/Probar Pull Requests:
En caso que se quiera probar un PULL REQUEST hay que estar en el branch `master` y luego hacer un pull del Pull Request de la siguiente manera: `git pull origin pull/135/head` donde 135 es el numero de Pull Request

## Como hacer un release?
Aqui se deja explicado como hacer un release para cualquiera de las aplicaciones de Argentum Online Libre 
https://github.com/ao-libre/ao-cliente/wiki/How-to-create-and-publish-Releases%3F

## Documentacion oficial Visual Basic 6
While the Visual Basic 6.0 IDE is no longer supported, Microsoft's goal is that Visual Basic 6.0 applications continue to run on supported Windows versions. The resources available from this page should help you as you maintain existing applications, and as you migrate your functionality to .NET.

https://docs.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/visual-basic-6.0-documentation

--------------------------

We start our branch from this version / old code:
* http://www.gs-zone.org/temas/cliente-y-servidor-13-3-dx8-v1.95611/




