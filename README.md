# MPS Plásticos – Layconsa
Dashboard Dash para seguimiento diario de Plan vs. Real en inyección y soplado.

## Estructura de archivos
```
mps_plasticos/
├── app.py            ← dashboard principal
├── db.py             ← módulo SQLite (plan + avance real)
├── requirements.txt
├── plan_produccion.xlsx   ← TU EXCEL (columnas exactas abajo)
└── mps_plasticos.db       ← se crea automático al correr app.py
```

## Columnas requeridas en plan_produccion.xlsx
| Columna                  | Ejemplo         |
|--------------------------|-----------------|
| BELNR_ID                 | 3860            |
| ItemCode                 | 2310090439      |
| Descripción del artículo | Barril Plumón…  |
| Sec                      | 1               |
| Fecha Inicio             | 06/04/2026      |
| Cantidad Base            | 98,743          |
| Maquina                  | INY05           |

## Instalación
```bash
pip install -r requirements.txt
```

## Uso diario
1. Copiá tu Excel exportado del MRP como `plan_produccion.xlsx` en esta carpeta
2. Corré: `python app.py`
3. Abrí el navegador en: http://localhost:8050

El plan se recarga automáticamente desde el Excel cada vez que iniciás la app.
El avance real queda guardado en `mps_plasticos.db` (persiste entre sesiones).

## Registrar avance real
Botón **"＋ Registrar avance"** → seleccionás la OP, ingresás la cantidad producida y guardás.
Si registrás dos veces la misma OP en el mismo día, el sistema actualiza el último valor.

## Semáforo de colores
| Color    | Condición            |
|----------|----------------------|
| 🟢 Verde | ≥ 100% del plan      |
| 🟡 Amarillo | 75–99% del plan   |
| 🔴 Rojo  | < 75% o sin avance   |
| ⬜ Gris  | Sin producción (0)   |
