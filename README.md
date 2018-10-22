# Cliente Argentum Online Libre
![AO Logo](https://ao-libre.github.io/ao-website/assets/images/logo.png)

## Autoupdates:

El programa al iniciar comparara la actual version del programa que se encuentra en `INIT/Version.ini` en el parámetro [version](https://github.com/ao-libre/ao-cliente/blob/92ec2e263f33b0e762b1ddef4875bbf220f634c4/INIT/Version.ini#L2) con la ultima version que se encuentra en en la ultima version lanzada mediante el [Endpoint Github Releases](https://api.github.com/repos/ao-libre/ao-cliente/releases/latest). En caso de ser diferente, se ejecuta nuestro programa `ao-autoupdate` para poder hacer el update.

Para mas información sobre este proceso:

[Funcion para comparar versiones](https://github.com/ao-libre/ao-cliente/blob/92ec2e263f33b0e762b1ddef4875bbf220f634c4/CODIGO/frmCargando.frm#L121)

[Codigo fuente ao-autoupdate](https://github.com/ao-libre/ao-autoupdate)


Se tomo como base esta version:
* http://www.gs-zone.org/temas/cliente-y-servidor-13-3-dx8-v1.95611/


