# finanzasTECHO-moduloFC-OB
Summary: 
- Modulo de creación de nóminas de reembolsos para el toolkit Office Banking.

Requirements
- Para la ejecución del aplicativo es necesario tener el client_secret.json. Se debe descargar del Google API Console con la cuenta del área.
- Validar la configuración del archivo .ini dependiendo del ambiente de instalación.

Funcionalidades
- Generación local de archivos de reembolsos para pago de nómina en OfficeBanking desde los archivos de flujo de caja y de personas
- Configurable para permitir la generación de semana actual o de cualquier semana en el año

Changelog
- V160418: Añade mensaje de ejecución y de errores en MATCH entre flujo de caja y personas
- V090418: Cambio de extensión a xls, cambio de nombramientod de Sheet1 a Hoja1.
- V040418: Versión release 1 para generación de reembolsos desde GSheets. Toma como base los archivos de Flujo de caja y el de personas en el TDrive de Finanzas.
- <V040418: Versiones anteriores: Deprecated