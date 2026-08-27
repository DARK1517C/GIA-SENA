# Retiro de firma PDF

La firma de documentos PDF fue retirada del alcance funcional de GIA.

El módulo conserva los campos históricos de firma en el modelo/base de datos para no realizar una migración destructiva sobre evidencias existentes. Sin embargo:

- no se muestra la opción Firmar en el visor;
- no se muestra “Firmar PDF y aprobar”;
- los endpoints de firma/documento firmado quedan deshabilitados;
- la aprobación de una evidencia ya no depende de una firma;
- el visor propio continúa para el PDF original con descarga, impresión, paginación y zoom.
