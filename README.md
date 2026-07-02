# Panel de Traslados — actualización automática (Mercado Público)

Tubería que mantiene el dashboard `dashboard.html` al día con las nuevas órdenes de compra
de un **grupo de proveedores** (los del archivo Excel original), consultando la API pública
de ChileCompra.

## Idea general

1. **Siembra** el histórico completo desde el Excel a una base local SQLite (`mp_traslados.db`).
   Así no hay que rehacer 20 años vía API.
2. Cada día, un proceso **incremental** consulta la API solo *hacia adelante* (desde la última
   fecha almacenada), trae las OC nuevas de cada proveedor del grupo y las guarda.
3. Se **reconstruye** `dashboard.html` (autocontenido) a partir de la base.

El navegador no puede llamar la API directamente (CORS + el ticket no debe exponerse en el HTML),
por eso la actualización ocurre en este proceso — no dentro de la página.

## Requisitos

- Python 3.10+
- Un **ticket** de la API: se solicita gratis en ChileCompra (menú de la API en
  `api.mercadopublico.cl`). Llega por correo. Nunca lo pongas en el código; se pasa por la
  variable de entorno `MP_TICKET`.

```bash
pip install -r requirements.txt
export MP_TICKET="TU-TICKET-AQUI"
```

## Puesta en marcha (una sola vez)

```bash
# 1) Cargar el histórico desde el Excel
python pipeline.py seed --excel /ruta/Mercado_Publico_de_Traslados_20250819_v1.xlsx --snapshot 2025-08-19

# 2) Resolver el código de proveedor de cada RUT (se cachea en providers.csv)
python pipeline.py resolve

# 3) (Recomendado) Verificar el esquema real de una OC antes de la primera actualización
python pipeline.py probe 1057490-2417-SE23
```

`probe` imprime el JSON crudo de una OC. Compara sus claves con las que usa
`mp_api.extraer_campos()`. Los campos más propensos a diferir son el **monto total**
(`Total` / `MontoTotalOC`…), el bloque **Proveedor**/**Comprador** y **TipoMoneda**. Si alguno
tiene otro nombre, ajústalo en esa única función.

## Uso diario

```bash
python pipeline.py run                    # = resolve + update + build
python pipeline.py run --sugerencias 30   # además busca proveedores candidatos (últimos 30 días)
```

o por pasos:

```bash
python pipeline.py update           # trae OC nuevas (desde la última fecha en base)
python pipeline.py candidates --dias 60   # descubre proveedores de traslados fuera de la nómina
python pipeline.py build            # regenera dashboard.html
```

Abre `dashboard.html` en el navegador. Es un archivo único, sin dependencias.

## Tipo de cambio UF/UTM

Las OC en CLP usan factor 1. Las OC en **UF (CLF)** o **UTM** se convierten a CLP con el valor del
indicador **en la fecha de la orden**, obtenido de `mindicador.cl` (sin ticket) y cacheado en la
tabla `fx` para no repetir consultas. El panel muestra una barra con el último valor UF/UTM aplicado
y cuántas OC del histórico están en esas monedas. El histórico ya trae su conversión desde la
columna `ConversionRate` del Excel; `mindicador.cl` solo se usa para las OC nuevas.

## Proveedores sugeridos (agregar desde el propio panel)

`python pipeline.py candidates` escanea la lista diaria nacional de OC, se queda con las que en su
nombre indican traslados/ambulancia/paciente, y acumula los proveedores que **no** están en tu
nómina (con su RUT, N° de OC y monto). Esos candidatos quedan embebidos en el dashboard.

En el panel **"Proveedores sugeridos para la nómina"** marcas los que te interesan y pulsas
**"Descargar nómina actualizada"**: se genera un `providers.csv` con tu nómina actual + los
seleccionados. Reemplaza el archivo en el repositorio y, en la próxima corrida, `resolve` obtiene
sus códigos y `update` empieza a seguirlos. (El navegador no puede escribir en tu disco ni en el
repositorio; por eso la incorporación es vía este archivo descargable, que es el mismo que consume
la tubería.)

> Nota: el escaneo recorre la lista nacional día por día y hace una consulta de detalle por cada OC
> con nombre de traslados. Es más pesado que la actualización normal; conviene correrlo con una
> ventana acotada (`--dias`) y no todos los días.

## Publicar en GitHub (automático + URL pública con Pages)

Puedes hacerlo **todo desde la web de GitHub**, sin terminal. Hay dos rutas:

### Ruta A — Solo publicar el panel ahora (sin ticket, sin automatización)
Para dejar el dashboard en línea en 3 minutos, con los datos actuales:
1. En GitHub: **New repository** (no marques "Add a README"). Créalo.
2. **Add file → Upload files**, arrastra `dashboard.html`. Confirma el commit.
3. Ábrelo en el repo → botón **✏️ (Edit)** → cambia el nombre del archivo a `index.html` →
   **Commit changes**. (Pages sirve `index.html` como página principal.)
4. **Settings → Pages → Source: Deploy from a branch → Branch: `main` / carpeta `/ (root)` → Save.**
5. Espera ~1 min: la URL aparece arriba en esa misma página (`https://TU-USUARIO.github.io/TU-REPO/`).

Los datos quedan congelados a la instantánea actual. Para que se actualicen solos, usa la Ruta B.

### Ruta B — Todo automático, configurado desde la web
El workflow siembra el histórico la primera vez, actualiza a diario y publica en Pages.
1. **New repository** (sin inicializar). Créalo.
2. **Add file → Upload files** y arrastra **toda la carpeta** `mp_pipeline` (incluye `pipeline.py`,
   `mp_api.py`, `requirements.txt`, `providers.csv`, `dashboard_template.html`, la carpeta
   `.github/`, y **el Excel** `Mercado_Publico_de_Traslados_20250819_v1.xlsx`). Confirma el commit.
   - *Consejo:* GitHub conserva las subcarpetas al arrastrar; verifica que quede
     `.github/workflows/update.yml` en el repo.
3. **Settings → Secrets and variables → Actions → New repository secret.** Nombre `MP_TICKET`,
   valor tu ticket. **Add secret**.
4. **Settings → Pages → Source: GitHub Actions** (no "Deploy from a branch"; el workflow despliega).
5. **Actions →** habilita los workflows si te lo pide → **"Actualizar y publicar panel Traslados"
   → Run workflow.** Al terminar (check verde), la URL queda en el job **publicar** y en
   **Settings → Pages**.

De ahí en adelante corre solo con el `cron` diario. En esta ruta no necesitas Git ni Python en tu
equipo: la máquina de GitHub Actions hace la siembra, la actualización y el despliegue.

> **Primera corrida (importante):** al no existir aún la base, el primer `run` se pone al día desde
> la instantánea (ago-2025) recorriendo día por día, lo que puede tardar y, si el rango es muy
> grande, acercarse al límite de 6 h por job. Si ocurre, vuelve a lanzar el workflow: cada corrida
> retoma desde la última fecha ya guardada. Alternativa más rápida: hacer ese primer catch-up en tu
> equipo (sección siguiente) y subir el `mp_traslados.db` ya al día.

### Configuración local (opcional, para el primer catch-up rápido)
Si prefieres poner la base al día en tu equipo antes de subirla (evita la corrida larga en Actions):

```bash
pip install -r requirements.txt
export MP_TICKET="TU-TICKET"
python pipeline.py seed --excel /ruta/Mercado_Publico_de_Traslados_20250819_v1.xlsx --snapshot 2025-08-19
python pipeline.py resolve
python pipeline.py update        # catch-up desde ago-2025 hasta hoy
python pipeline.py build
```

Luego sube el proyecto (con `mp_traslados.db` ya al día) por la Ruta B; el workflow detectará la
base existente y solo hará incrementos diarios.

### Programación
El horario está en el `cron` del workflow (`30 11 * * *` ≈ 08:30 en Chile). GitHub usa UTC;
ajusta según horario de verano/invierno si necesitas exactitud.

## Ejecutar sin GitHub
Cualquier programador de tareas sirve: exporta `MP_TICKET` y ejecuta `python pipeline.py run`
(cron en Linux/macOS, Programador de tareas en Windows). Abre el `dashboard.html` resultante.

## El grupo de proveedores

Está en `providers.csv` (`alias, rut, codigo_proveedor`), sembrado desde la hoja **RUT** del
Excel (75 proveedores). Para **agregar o quitar** proveedores, edita ese archivo: añade una fila
con su RUT (el `codigo_proveedor` se completa solo al correr `resolve`).

## Notas

- **Deduplicación:** la base usa el código de OC como clave, por lo que las OC repetidas del
  Excel se consolidan (11.853 OC únicas vs. 11.864 filas del Excel).
- **Monedas:** las OC en CLP usan `conversion_rate = 1`. Las OC en **UF/UTM** se convierten con el
  valor del indicador a la fecha (`mindicador.cl`), cacheado en la tabla `fx`.
- **Región:** se asigna con el mapa comprador→región derivado del Excel. Un comprador nuevo que
  no esté en el mapa queda como `Sin Region` hasta que lo agregues a `region_map`.
- **Carga de la puesta al día:** ponerse al día desde el snapshot (ago-2025) hasta hoy recorre
  día por día; es una corrida larga pero única y reanudable (guarda avance por fecha). El ritmo
  diario posterior es liviano.

## Archivos

| Archivo | Rol |
|---|---|
| `pipeline.py` | CLI: `seed`, `resolve`, `update`, `candidates`, `build`, `probe`, `run` |
| `mp_api.py` | Cliente de la API + normalización + tipo de cambio (`mindicador.cl`) + detección de traslados |
| `providers.csv` | Grupo de proveedores (alias, RUT, código) |
| `dashboard_template.html` | Plantilla del panel con el marcador `__DATA_JSON__` |
| `dashboard.html` | Panel generado (salida) |
| `mp_traslados.db` | Base local SQLite (histórico + incremental) |
| `.github/workflows/update.yml` | Actualización diaria + publicación en GitHub Pages |
| `.gitignore` | Excluye `site/` y `__pycache__/` |
