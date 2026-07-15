# 📁 Directorio Oficial de Scripts del Proyecto Xplora / Trademart

Este directorio centraliza todos los scripts operativos, de procesamiento (Python) y de inserción de datos (SQL) del proyecto, garantizando una arquitectura ordenada y evitando la dispersión de archivos.

---

## 📂 1. Scripts SQL por Tabla (`sql/`)

Para evitar confusiones al generar, inspeccionar o ejecutar inserciones de prueba y poblar la base de datos, se ha generado **un script SQL dedicado por cada tabla del sistema**, numerado estrictamente según la jerarquía de dependencias relacionales (`Foreign Keys`).

Cada uno de estos archivos incluye al final la instrucción de sincronización de secuencias de PostgreSQL (`SELECT setval(...)`), lo cual es **crítico** para que tras insertar IDs numéricos (`1, 2, ...`), los futuros guardados desde la API de Spring Boot/JPA sigan generando IDs correlativos sin errores de clave duplicada (`Duplicate Key Violation`).

### Orden exacto de ejecución (o dependencias):
1. **`00_limpieza.sql`**: Limpia todas las tablas relacionales (`pm_act`, `reporte`, `cronogramas`, `planning`, `pm`, `pdv`, `producto`, etc.) respetando las restricciones de llave foránea y preservando el usuario administrador (`usuario_id = 1`).
2. **`01_roles.sql`**: Inserta los roles del sistema (`ADMIN`, `GERENTE`, `SUPERVISOR`, `MERCADERISTA`) manejando conflictos por ID.
3. **`02_equipo_comercial.sql`**: Inserta los equipos comerciales o zonas territoriales.
4. **`03_usuario.sql`**: Inserta la base de usuarios (Supervisores y Mercaderistas).
5. **`04_usuario_rol.sql`**: Asigna roles a cada uno de los usuarios generados.
6. **`05_pdv.sql`**: Crea 500 Puntos de Venta (Mercados / Bodegas / Supermercados) en diversos distritos con estado capitalizado (`Activo`, `Inactivo`).
7. **`06_pm.sql`**: Crea los Puestos de Mercado (`P-01`, `P-02`, etc.) asociados a cada PDV.
8. **`07_producto.sql`**: Inserta el catálogo base de SKUs (`producto_id` de 1 a 20) necesarios para las promociones.
9. **`08_act.sql`**: Crea las actividades promocionales (`act`) vinculadas a los SKUs.
10. **`09_pm_act.sql`**: Vincula (N:M) cada puesto de mercado (`pm`) con las promociones vigentes (`act`).
11. **`10_planning.sql`**: Inserta exactamente **1 planning por cada mercaderista**, asignándole puestos y días de la semana.
12. **`11_cronogramas.sql`**: Inserta los cronogramas en el formato `datos_json` exacto requerido (`[{"fecha":"YYYY-MM-DD", "Ciudad":"", "MERCADO":"...", "DEX":"", "nroPuesto":"...", "encargado":"..."}]`).
13. **`12_reporte.sql`**: Inserta 3,000 reportes (Storechecks) distribuidos a lo largo del año 2026 con estados capitalizados (`Completado`, `Pendiente`, `En proceso`, `Observado`) y con relación verificada a `pm_id`.

> **Nota:** Si deseas ejecutar todo en un solo paso, puedes usar el archivo **`99_master_seed.sql`**, el cual consolida todos los scripts anteriores en un único archivo de ejecución y auto-ajusta las secuencias al finalizar.

---

## 🐍 2. Scripts de Procesamiento de Storecheck (`headless_storecheck.py`)

* **`headless_storecheck.py`**: Es el procesador headless en Python llamado directamente por el backend (`StorecheckProcessorController.java`). Toma el Excel base cargado por el usuario, filtra por promoción y PDVs seleccionados, y genera el archivo Excel de reporte (`output.xlsx`) con las pestañas y formato visual requerido.
* **`generate_split_scripts.py`**: Script generador programático de toda la suite SQL segmentada ubicada en `sql/`. Puedes ejecutar `python generate_split_scripts.py` en cualquier momento si deseas regenerar o actualizar los volúmenes de datos de prueba.
