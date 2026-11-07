# Echoes Market Tracker

Automatiza la consulta diaria de la API Echoes y mantiene un Excel histórico en OneDrive personal.

## Qué hace
- Consulta `https://echoes.mobi/api/v1/prices` con los 34 `type_ids` incluidos.
- Guarda `prices_as_of`, `type_id`, nombre, `avg`, `avg_no_outliers`, `base_price`, versión y fecha de captura.
- Evita duplicados por combinación `prices_as_of + type_id`.
- Actualiza `Historical_Data`, `Latest_Prices`, `Price_Analysis`, `Dashboard` y `Run_Info`.
- Programa el trabajo para las 08:00 de Alemania (CET/CEST). GitHub Actions puede retrasar los disparos programados; no es un servicio de ejecución al segundo.

## Configuración inicial

### 1. Crear repositorio
1. Crea un repositorio privado en GitHub, por ejemplo `echoes-market-tracker`.
2. Sube todos los archivos de esta carpeta manteniendo la estructura.

### 2. Conectar OneDrive personal mediante rclone
Necesitas un ordenador solo para realizar esta autorización inicial:
1. Instala rclone desde https://rclone.org/downloads/
2. Abre una terminal y ejecuta `rclone config`.
3. Elige `n` (nuevo remote), pon el nombre `onedrive`, selecciona el tipo `onedrive`.
4. Completa el inicio de sesión en el navegador con tu cuenta Microsoft personal y concede acceso a OneDrive.
5. Comprueba que funciona con `rclone lsd onedrive:`.
6. Localiza el archivo de configuración con `rclone config file`.
7. Codifica el archivo completo en Base64:
   - macOS/Linux: `base64 -w 0 ~/.config/rclone/rclone.conf`
   - Windows PowerShell: `[Convert]::ToBase64String([IO.File]::ReadAllBytes("$env:APPDATA\rclone\rclone.conf"))`
   En Windows, comprueba la ruta real que muestra `rclone config file`.

**Seguridad:** el contenido Base64 equivale a una credencial. No lo pegues en el repositorio ni en el chat. Guárdalo únicamente como secreto de GitHub.

### 3. Añadir secreto y variable en GitHub
En el repositorio: Settings → Secrets and variables → Actions.

Secret:
- `RCLONE_CONFIG_B64`: el texto Base64 del archivo rclone.conf.

Variable:
- `RCLONE_REMOTE`: `onedrive` (debe coincidir con el nombre elegido en rclone).

### 4. Primera ejecución
1. En GitHub abre Actions → `Actualizar histórico de precios Echoes`.
2. Pulsa `Run workflow`.
3. La ejecución manual no espera a las 08:00.
4. Comprueba que aparece `Echoes_Market_History.xlsx` en la raíz de tu OneDrive.

## Notas importantes
- El Excel es el almacén histórico inicial. Para historiales muy grandes, convendrá migrar los datos a una base de datos y dejar Excel como informe.
- La clave de deduplicación es `prices_as_of + type_id`. Si Echoes publica varios snapshots en el mismo día, se conservarán todos si sus timestamps difieren.
- La lista de 34 IDs procede de la URL facilitada. La API puede devolver menos elementos; el script registra exactamente lo que devuelve y no inventa los ausentes.
- Si la API cambia el formato, la tarea fallará en vez de sobrescribir el histórico con datos incompletos.
- GitHub Actions no garantiza ejecución exacta al minuto; puede haber retrasos o ejecuciones omitidas en periodos de alta carga.
- Revisa las condiciones de uso y límites de la API Echoes antes de automatizar consultas periódicas.
