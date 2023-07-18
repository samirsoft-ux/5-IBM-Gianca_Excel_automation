# IBM Cloud generar reporte Ferreyros 🗒️

## 📃 Introducción
Automatización de la generación del reporte para Ferreyros utilizando python y flask.
Este repositorio se encuentra conectado con Render(cuenta de google) para el despliegue de la aplicación de forma automática.

## 📑 Índice  
1. [Descargar y subir el snapshot de VMware](#1-Descargar-y-subir-el-snapshot-de-VMware)
2. [Descargar, ordenar e ingresar los datos del Client Details](#2-Descargar-,-ordenar-e-ingresar-los-datos-del-Client-Details)
<br />

## 1° Descargar y subir el snapshot de VMware
```Este es el archivo input que va a servir para que se genere el reporte y se consigue a través del vCloud Director console.```

1. Ingresa a la sección de ***Lista de recursos*** y dentro ubicar el apartado ***Compute***.

2. Ingresar al servicio de ***VMware Solutions*** del cual se va a generar el reporte.

3. Dentro seleccionar el botón "vCloud Director console" e ingresar al portal con sus credenciales.

4. Dentro ingresar al Virtual Data Center del cual se va a generar el reporte y seleccionar el botón "EXPORT VMS" (dejar por defecto todos los parámetros de descarga del snapshot).

5. Abrir el archivo descargado y seleccionar el botón "Save As...".

6. Guardar el archivo con el nombre "principal" y en un formato ".xlsx".

7. Dirigirse al generador de reportes y seleccionar el botón "Choose File" y subir el archivo que se acaba de guardar con el nombre "principal.xlsx".


<p align="center">
   <img src=https://github.com/samirsoft-ux/Gianca_FFS/blob/main/Images/first.gif>
</p>

   **Notas**
   * En caso no te aparezca el botón "EXPORT VMS" cambiar la vista en la que se muestran las VM's.
<p align="center">
   <img src=https://github.com/samirsoft-ux/Gianca_FFS/blob/main/Images/Nota1.png>
</p>

## 2° Descargar, ordenar e ingresar los datos del Client Details
```Este es el archivo de los montos de facturación del cliente que se encuentran desde el portal Client Details.```


<p align="center">
   <img src=>
</p>

   **Notas**
   * LOREM IPSUM   