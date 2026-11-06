# FIAS · Centro de decisiones del Directorio — v0.3

Portal documental con **interfaz nueva** (no derivada visualmente de los proyectos Donaciones / Movilizaciones) y autenticación Microsoft 365 / Microsoft Graph basada en el patrón técnico del portal Movilizaciones FIAS.

## Instalación

```bash
npm install
npm run dev
npm run build
```

### Antes de iniciar sesión

1. Registrar en Microsoft Entra ID la dirección de la aplicación como **Single-page application (SPA) redirect URI**, exactamente igual a su URL pública (o `http://localhost:5173/` para pruebas).
2. Confirmar que la app identificada en `src/sharepoint.ts` está autorizada en el tenant para usar permisos delegados `User.Read`, `Files.Read.All` y `Sites.Read.All`, con consentimiento administrativo si el tenant lo requiere. Los identificadores tenant/client se tomaron del ejemplo Movilizaciones; se pueden sobrescribir mediante `VITE_ENTRA_TENANT_ID` y `VITE_ENTRA_CLIENT_ID`.
3. El usuario debe tener permiso de lectura sobre la matriz compartida y los PDF del repositorio.
4. Publicar en alojamiento HTTPS (salvo localhost), con redirección SPA a `index.html`.

## Flujo de consulta en vivo

- MSAL 4 gestiona inicio de sesión Microsoft 365, sesión en sessionStorage, renovación de token y cierre.
- Graph resuelve el enlace compartido mediante `/shares/u!{base64url}/driveItem`.
- Desde el driveItem se obtiene la URL temporal de descarga y se lee el XLSX en memoria, sin subirlo a terceros.
- Se indexan las hojas, asuntos y resoluciones; se heredan los enlaces de la misma acta cuando aparecen en filas posteriores.
- Al pulsar **Sincronizar** se vuelve a consultar SharePoint. El portal **no se sincroniza automáticamente en segundo plano**.
- Si Graph, las políticas del tenant o CORS de la URL temporal impiden descargar la matriz, se informa el error: puede usarse la importación local de contingencia.

## Privacidad y alcance

- La distribución NO incluye una copia de los registros de la matriz Excel; para verlos se requiere inicio de sesión o importación manual.
- No se usan proxies públicos, workers anónimos, claves secretas ni tokens persistidos por el portal fuera de la caché MSAL de sesión.
- Los permisos de apertura de cada PDF siguen definidos por SharePoint.
- No se extrae texto de PDF, no hay OCR ni IA semántica en esta versión; la búsqueda es textual sobre la matriz, y los resultados no son resúmenes generados por IA.
- Los indicadores cuentan registros de asuntos/resoluciones por fila y actas agrupadas, no infieren cuáles decisiones están aprobadas.

## Estructura

- `src/sharepoint.ts`: autenticación, lectura Graph, descarga segura del Excel.
- `src/data.ts`: índice en memoria, normalización de enlaces, años y etiquetas.
- `src/main.tsx`: interfaz documental Enterprise, filtros, vistas, gráficos, panel detalle.
- `src/style.css`: diseño responsive claro / oscuro.

## Diagnóstico de accesos

- **AADSTS50011**: la URL de redirección SPA no coincide con la registrada en Entra.
- **403**: el usuario, el enlace compartido o el registro de aplicación no dispone del permiso necesario.
- **Consent required**: aprobar los permisos delegados de Graph o configurar la app según las políticas del tenant.
- **CORS en descarga**: revisar directivas de descarga de SharePoint. Un proxy corporativo autenticado o API backend administrada puede ser necesario; no usar un proxy público no autenticado para documentos institucionales.
