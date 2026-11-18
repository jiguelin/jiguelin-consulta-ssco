# Consulta SSCO

Aplicación de Streamlit para verificar si el RUC de un proveedor de compras figura en la relación de Sujetos sin Capacidad Operativa publicada por SUNAT.

## Uso

1. En **Padrón SSCO**, selecciona **Descargar de SUNAT** para consultar el listado oficial. También puedes elegir **Cargar Excel del padrón** y seleccionar una copia del archivo oficial.
2. En **Cruzar compras**, carga uno o varios archivos y pulsa **Cruzar con el padrón**. En **Buscar por RUC**, puedes consultar directamente un proveedor sin cargar compras.
3. Revisa las coincidencias y las filas sin RUC consultable. **Descargar reporte Excel** genera un archivo con el cruce completo, las coincidencias, el detalle de la carga y la versión del padrón.

| Formato | Identificador que se cruza |
| --- | --- |
| Registro contable `COMPRAS.xlsx` | Documento de identidad del proveedor en el bloque **Información del proveedor** |
| `resultado_*.xlsx` de NewContaSis | Documento de proveedor en la columna H de **Comprobantes** |
| Propuesta de compras SIRE en ZIP, TXT o CSV | **Nro Doc Identidad**, con tipo de documento RUC |
| Excel con encabezados | Columna **RUC proveedor** o un encabezado equivalente |

Se admiten comprobantes de años anteriores. El cruce compara cada RUC con la versión del padrón que aparece en el reporte. Para revisar con una versión anterior, carga el Excel oficial de esa fecha si lo conservas. Las coincidencias identifican proveedores que figuran en ese listado para su revisión contable.

Cada archivo puede tener hasta 15 MB. Una consulta admite hasta 45 MB y 40.000 filas. Se omiten copias idénticas del mismo archivo. Cada origen reconocido conserva sus propias filas.

[Padrón oficial de SUNAT](https://www.sunat.gob.pe/padronesnotificaciones/sujeSinCapacidadOperativa.html)

## Ejecutar localmente

Se usa Python 3.12.

```sh
python -m pip install -r requirements.txt
streamlit run streamlit_app.py
```

## Desplegar en Streamlit Community Cloud

En **Create app**, selecciona el repositorio `jiguelin/jiguelin-consulta-ssco`, la rama `main` y el archivo `streamlit_app.py`. En **Settings → Sharing**, selecciona **Only specific people can view this app** para usarla con registros contables del equipo.

La aplicación procesa los archivos en la sesión de Streamlit. GitHub aloja el código fuente.

[Despliegue en Streamlit](https://docs.streamlit.io/deploy/streamlit-community-cloud/deploy-your-app/deploy) · [Privacidad de la app](https://docs.streamlit.io/deploy/streamlit-community-cloud/share-your-app#make-your-app-public-or-private)
