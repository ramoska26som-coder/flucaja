# Flujo de Caja — Instrucciones de configuración

Herramienta de control de efectivo en caja, con almacenamiento en Google Sheets y acceso desde cualquier dispositivo a través de GitHub Pages.

---

## Estructura del proyecto

```
flujo-caja-web/
├── index.html    ← La aplicación web (sube a GitHub Pages)
└── Code.gs       ← Backend de Google Apps Script (copia al editor)
```

---

## Paso 1 — Crear la hoja de cálculo en Google Sheets

1. Abre [Google Sheets](https://sheets.google.com) y crea una hoja nueva.
2. Dale un nombre descriptivo, por ejemplo: **Flujo de Caja**.
3. **No necesitas crear ninguna hoja ni columna a mano**; el script las crea automáticamente la primera vez que se ejecuta.

---

## Paso 2 — Crear el proyecto de Apps Script

1. Dentro de la hoja de cálculo, ve al menú **Extensiones → Apps Script**.
2. Se abrirá el editor de código. Borra todo el contenido que aparece por defecto en `Code.gs`.
3. Copia todo el contenido del archivo **`Code.gs`** de este repositorio y pégalo en el editor.
4. Guarda el proyecto con **Ctrl+S** (o el ícono de disco). Puedes nombrar el proyecto como quieras, ej. *Flujo de Caja Backend*.

---

## Paso 3 — Desplegar como Web App

1. En el editor de Apps Script, haz clic en el botón **Implementar** (arriba a la derecha) → **Nueva implementación**.
2. Haz clic en el ícono de engranaje junto a "Seleccionar tipo" y elige **Aplicación web**.
3. Configura los campos así:

   | Campo | Valor |
   |---|---|
   | Descripción | Flujo de Caja v1 |
   | Ejecutar como | **Yo** (tu cuenta de Google) |
   | Quién puede acceder | **Cualquier usuario** |

4. Haz clic en **Implementar**.
5. La primera vez, Google te pedirá que **autorices el acceso**. Acepta todos los permisos — el script solo necesita acceder a tu Google Sheet.
6. Copia la **URL de la aplicación web** que aparece. Tiene este formato:
   ```
   https://script.google.com/macros/s/AKfycb.../exec
   ```
   **Guarda esta URL**, la necesitarás en el Paso 5.

> **Nota:** Cada vez que modifiques el código de `Code.gs`, deberás crear una nueva implementación para que los cambios entren en efecto.

---

## Paso 4 — Publicar el frontend en GitHub Pages

1. Crea un repositorio en GitHub (puede ser público o privado).
2. Sube el archivo **`index.html`** a la raíz del repositorio.
3. Ve a **Settings → Pages** en tu repositorio.
4. En **Source**, selecciona la rama `main` y carpeta `/ (root)`.
5. Haz clic en **Save**. GitHub te dará una URL pública como:
   ```
   https://tu-usuario.github.io/nombre-del-repo/
   ```
6. Espera 1-2 minutos y abre esa URL en tu navegador o celular.

---

## Paso 5 — Conectar la app con el backend

1. Abre la URL de GitHub Pages.
2. Verás una pantalla de configuración pidiendo la URL del Web App.
3. Pega la URL que copiaste en el Paso 3 y haz clic en **Conectar y cargar datos**.
4. La app se conectará a Google Sheets, creará las hojas necesarias y mostrará el panel principal.

La URL del servidor se guarda en el navegador — solo tendrás que configurarla una vez por dispositivo.

---

## Uso diario

- **Agregar movimiento:** botón **+** abajo a la derecha.
- **Recargar datos:** botón **↻ Recargar** en la barra superior o en el pie de página (útil si otro usuario registró movimientos desde otra PC o celular).
- **Comparar saldos:** los campos de Excel y Sistema se comparan contra el Saldo ajustado (saldo del flujo + efectivo pendiente).
- **Ver inventario de billetes:** panel debajo de la comparación, muestra cuántos billetes hay por denominación según los movimientos registrados.

---

## Acceso desde múltiples dispositivos

Cualquier persona con acceso a la URL de GitHub Pages puede ver y registrar movimientos, y los datos se sincronizan en tiempo real con el Google Sheet. Para recargar los datos de otros usuarios, usa el botón **↻ Recargar**.

---

## Notas de seguridad

- El Google Sheet y el Apps Script están bajo tu cuenta de Google.
- La configuración de **"Cualquier usuario"** permite que cualquier persona con la URL del Web App pueda leer y escribir datos. Si necesitas restringir el acceso, comparte la URL del Web App solo con las personas autorizadas.
- No hay autenticación de usuarios en la app web — cualquiera que tenga la URL de GitHub Pages puede usarla.

---

## Solución de problemas

**La app muestra "No se pudo conectar"**
- Verifica que la URL del Web App sea correcta (debe empezar con `https://script.google.com/macros/s/`).
- Confirma que desplegaste la app con acceso de **"Cualquier usuario"**.
- Intenta recargar la página y reconectar.

**Cambié el código de Code.gs pero no funciona**
- Debes crear una **nueva implementación** cada vez que modifiques el script. Una implementación existente no se actualiza automáticamente.

**Aparece una pantalla de inicio de sesión de Google al abrir la app**
- Esto ocurre si configuraste el acceso como "Cualquier usuario con cuenta de Google" en lugar de "Cualquier usuario". Cambia esta configuración y vuelve a implementar.
