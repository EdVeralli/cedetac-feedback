# CEDETAC - Feedback de usuarios (Boti)

Programa que baja de AWS Athena las respuestas de la encuesta de feedback (CXF) de los usuarios que pasaron por el flujo **CEDETAC** del chatbot Boti, las limpia y genera el Excel de indicadores con el mismo formato y las mismas fórmulas que el Excel testigo `CEDETAC - Feedback últimos 3 meses 2026`.

Validado el 02/10/2026 contra el testigo (01/07 al 30/09/2026): 1.752 filas idénticas y los 4 indicadores iguales.

---

## 1. Estructura

```
pedido cedetac/
├── cedetac_feedback.py        Programa principal (consola, argparse)
├── config_fechas.txt          Período a procesar
├── requirements.txt
├── sql/
│   └── cedetac_feedback.sql   Query de Athena con {fecha_inicio} / {fecha_fin}
├── plantilla/
│   └── CEDETAC_plantilla.xlsx Excel testigo sin datos (formato + hojas)
└── output/                    (se crea solo, no se sube al repo)
    ├── cedetac_bajada_<desde>_al_<hasta>.csv    bajada cruda de Athena
    ├── cedetac_limpio_<desde>_al_<hasta>.csv    bajada limpia
    ├── CEDETAC - Feedback <dd-mm-aaaa> al <dd-mm-aaaa>.xlsx
    └── logs/cedetac_<aaaammdd_hhmmss>.log
```

## 2. Requisitos (una sola vez)

- Python 3.10+ (Anaconda) con las librerías de `requirements.txt`:

  ```powershell
  cd "C:\GCBA\pedido cedetac"
  pip install -r requirements.txt
  ```

- AWS CLI v2 + `aws-azure-login` configurado en el profile `default` con el rol **PIBADataScientist** (misma configuración que Metricas_Boti_Mensual).
- VPN GCBA si la red lo requiere.

## 3. Configurar el período

Editar `config_fechas.txt`. Prioridad: **rango > últimos meses > mes**.

| Modo | Líneas | Resultado |
|---|---|---|
| Últimos N meses cerrados (por defecto) | `ULTIMOS_MESES=3` | Corrido en octubre 2026 → 01/07/2026 al 30/09/2026 |
| Mes completo | `MES=9` y `AÑO=2026` | 01/09/2026 al 30/09/2026 |
| Rango | `FECHA_INICIO=2026-07-01` y `FECHA_FIN=2026-09-30` | Ese rango exacto |

También se puede pisar el archivo por parámetro: `--desde 2026-07-01 --hasta 2026-09-30`.

## 4. Ejecución

```powershell
cd "C:\GCBA\pedido cedetac"
python cedetac_feedback.py
```

Al arrancar, el programa **pide el login** y muestra el comando:

```powershell
aws-azure-login --profile default --mode=gui
```

1. Ejecutarlo en **otra terminal** y completar el login en el navegador (rol PIBADataScientist).
2. Volver a la terminal del programa y presionar **ENTER**.
3. El programa verifica las credenciales (`sts get-caller-identity`). Si no son válidas vuelve a pedir el login (`q` para salir). Solo con credenciales válidas ejecuta el resto.

Si el token vence durante la query (dura 1 hora), vuelve a pedir el login y reintenta (hasta 3 veces).

### Otras formas

```powershell
cd "C:\GCBA\pedido cedetac"
# Rango puntual sin tocar config_fechas.txt
python cedetac_feedback.py --desde 2026-07-01 --hasta 2026-09-30

# Reprocesar un CSV ya bajado (no usa Athena ni login)
python cedetac_feedback.py --csv output\cedetac_bajada_2026-07-01_al_2026-09-30.csv --desde 2026-07-01 --hasta 2026-09-30

# Ayuda
python cedetac_feedback.py --help
```

## 5. Origen de los datos (query)

Base `caba-piba-consume-zone-db`, workgroup `Production-caba-piba-athena-boti-group`, región `us-east-1`.

1. **sesionescedetac**: sesiones de `boti_message_metrics_2` en el período que dispararon alguna regla `%SA01CAT01%` (flujo CEDETAC).
2. **sesiones**: de esas sesiones, la respuesta Sí/No a la pregunta de efectividad (`CXF01CUX01..04 Sí/No …`), excluyendo una lista de usuarios de prueba.
3. **feedback / tabla_feedback**: de `boti_user_vars_metrics_2` toma `sugerencianps`/`sugerenciacxf` (sugerencia), `ultimotemanps` y `ultimotema`.
4. **esfuerzo_satisfaccion**: regla de esfuerzo (Muy fácil … Muy difícil) y de satisfacción (Muy conforme … No sé) de cada sesión. Si no respondió: `Sin Datos`.

Los tres filtros de fecha de la query usan el período configurado (`{fecha_inicio}` / `{fecha_fin}`).

## 6. Limpieza

No se elimina ninguna fila por su contenido (`Sin Datos` y `No sé` se conservan).

1. Se excluyen filas con `fecha` fuera del período (queda logueado cuántas).
2. Se eliminan columnas: `session_id`, `Usuario`, `ultimotemanps`, `ultimotema`, `link_conversacion`.
3. En efectividad, esfuerzo y satisfacción se quita el código inicial (`CXF01CUX04 `, `CXF01CUX01 `, …) y el sufijo del flujo (`CATs`, `Integraciones`, `Estáticos`, `Pushes`). Ej: `CXF01CUX04 Muy fácil CATs` → `Muy fácil`.
4. Sugerencia: se colapsan espacios repetidos y se quitan espacios al inicio/fin.
5. Columnas finales: `fecha`, `Efectividad`, `Esfuerzo`, `Satisfaccion`, `Sugerencia del usuario` (última), ordenadas por fecha ascendente.

## 7. Indicadores (hoja "Indicadores generales")

Fórmulas idénticas al Excel testigo, sobre la hoja `Feedback CEDETAC`:

| Indicador | Fórmula |
|---|---|
| Total_Usuarios | `=CONTARA(A:A)` |
| TasaEfectividad | `=CONTAR.SI(B:B;"Sí") / CONTARA(B:B)` |
| Esfuerzo | `=(Muy difícil×5 + Difícil×4 + Más o menos×3 + Fácil×2 + Muy fácil×1) / CONTARA(C:C)` |
| Satisfacción % | `=(CONTAR.SI(D:D;"Muy conforme") + CONTAR.SI(D:D;"Conforme ")) / CONTARA(D:D)` |

### Comportamiento conocido de las fórmulas del testigo

Se mantienen tal cual para que los números sean comparables con lo reportado:

- `CONTARA` incluye la fila de encabezado: Total_Usuarios da +1 y todos los denominadores llevan +1.
- Esfuerzo y Satisfacción dividen por todas las filas, incluyendo `Sin Datos` (y `No sé` en Satisfacción).
- Satisfacción busca `"Conforme "` con espacio final, que no existe en los datos: solo suma `Muy conforme`.

El programa muestra en consola y en el log los mismos 4 valores calculados en Python, como control.

## 8. Problemas frecuentes

| Síntoma | Solución |
|---|---|
| Pide el login de nuevo después de ENTER | El login no terminó o eligió otro rol. Repetir `aws-azure-login --profile default --mode=gui`. |
| `ExpiredToken` | El programa vuelve a pedir el login solo; si se repite, renovar y relanzar. |
| `PermissionError` al guardar el Excel | El archivo está abierto en Excel: cerrarlo y volver a correr. |
| `La bajada no trajo filas` | Revisar el período en `config_fechas.txt`. |
