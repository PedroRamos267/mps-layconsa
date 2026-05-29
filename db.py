"""
db.py — MPS Plásticos Layconsa

Calendarios según Horas Programadas:
  24hs: Lun 07:00 → Sáb 07:00  (para fin de semana 48hs)
  12hs: Lun 07:00 → Vie 19:00  (cada día 07:00→19:00, para 19:00→07:00 y fin de semana)

Velocidad = Cantidad Base / Horas          (und/hora)
Cap. diaria = Velocidad × Horas Programadas
Duración (hs) = Planificado / Velocidad
"""

import sqlite3
import pandas as pd
from datetime import date, datetime, timedelta

HORA_INICIO = 7   # 07:00
HORA_FIN_12 = 19  # 19:00 para máquinas de 1 turno

# Puestos de trabajo paralelos — las OPs arrancan todas en la misma fecha
# sin esperar la cola secuencial
MAQUINAS_PARALELAS = {"AMAN02"}


# ── Helpers de calendario ─────────────────────────────────────────────────────

def siguiente_momento_habil(dt: datetime, horas_prog: int) -> datetime:
    """Avanza dt al próximo momento productivo respetando feriados y fines de semana."""
    MAX_ITER = 60
    for _ in range(MAX_ITER):
        d_date = dt.date()
        # Feriado o fin de semana no hábil → saltar al siguiente día a las 07:00
        if d_date in FERIADOS:
            dt = datetime(dt.year, dt.month, dt.day, HORA_INICIO) + timedelta(days=1)
            continue
        if horas_prog == 24:
            # Sábado 07:00+ → saltar al lunes
            if dt.weekday() == 5 and dt.hour >= HORA_INICIO:
                lun = datetime(dt.year, dt.month, dt.day, HORA_INICIO) + timedelta(days=2)
                dt = lun
                continue
            if dt.weekday() == 6:
                lun = datetime(dt.year, dt.month, dt.day, HORA_INICIO) + timedelta(days=1)
                dt = lun
                continue
        else:
            # 12hs: corte a las 19:00
            if dt.weekday() < 5 and dt.hour >= HORA_FIN_12:
                dt = datetime(dt.year, dt.month, dt.day, HORA_INICIO) + timedelta(days=1)
                continue
            if dt.weekday() >= 5:
                days_fwd = (7 - dt.weekday()) % 7 or 7
                dt = datetime(dt.year, dt.month, dt.day, HORA_INICIO) + timedelta(days=days_fwd)
                continue
            if dt.hour < HORA_INICIO:
                dt = datetime(dt.year, dt.month, dt.day, HORA_INICIO)
        break
    return dt


def sumar_horas_habiles(inicio: datetime, horas: float, horas_prog: int) -> datetime:
    """Avanza 'horas' productivas desde 'inicio' respetando el calendario."""
    actual = siguiente_momento_habil(inicio, horas_prog)
    restante = horas

    while restante > 0:
        if horas_prog == 24:
            # Corte: sábado 07:00
            dias_hasta_sab = (5 - actual.weekday()) % 7
            if dias_hasta_sab == 0 and actual.hour >= HORA_INICIO:
                actual = siguiente_momento_habil(actual, horas_prog)
                continue
            corte = datetime(actual.year, actual.month, actual.day,
                             HORA_INICIO) + timedelta(days=dias_hasta_sab)
            horas_disp = (corte - actual).total_seconds() / 3600
            if restante <= horas_disp:
                actual = actual + timedelta(hours=restante)
                restante = 0
            else:
                restante -= horas_disp
                # Avanzar al lunes 07:00 (saltando feriados)
                lun = corte + timedelta(days=2)
                actual = siguiente_momento_habil(lun, horas_prog)
        else:
            # Corte: 19:00 del mismo día
            corte_dia = datetime(actual.year, actual.month, actual.day, HORA_FIN_12)
            horas_disp = (corte_dia - actual).total_seconds() / 3600
            if restante <= horas_disp:
                actual = actual + timedelta(hours=restante)
                restante = 0
            else:
                restante -= horas_disp
                # Siguiente día hábil a las 07:00
                sig = datetime(actual.year, actual.month, actual.day,
                               HORA_INICIO) + timedelta(days=1)
                actual = siguiente_momento_habil(sig, horas_prog)
    return actual


# Feriados — agregar fechas manualmente en formato date(YYYY, M, D)
FERIADOS = {
    date(2026, 4, 2),   # Jueves Santo
    date(2026, 4, 3),   # Viernes Santo
}
MANTENIMIENTOS = {}  # {maquina: {date, ...}} — se carga desde Excel

def es_habil(d: date) -> bool:
    return d.weekday() < 5 and d not in FERIADOS


def limpiar_cantidad(serie: pd.Series) -> pd.Series:
    return (
        serie.astype(str)
        .str.replace(",", "", regex=False)
        .str.strip()
        .replace("nan", "0")
        .astype(float)
    )


# ── Init DB ───────────────────────────────────────────────────────────────────

def init_db(db_path: str):
    con = sqlite3.connect(db_path)
    con.executescript("""
        CREATE TABLE IF NOT EXISTS ordenes_plan (
            BELNR_ID          INTEGER PRIMARY KEY,
            ItemCode          TEXT,
            Descripcion       TEXT,
            Sec               INTEGER,
            Maquina           TEXT,
            Proceso           TEXT,
            Linea             TEXT,
            Subcomponente     TEXT,
            Horas_Prog        INTEGER,
            Fecha_Inicio      TEXT,
            Hora_Inicio       TEXT,
            Fecha_Fin         TEXT,
            Hora_Fin          TEXT,
            Velocidad         REAL,
            Cap_Diaria        REAL,
            Planificado       REAL,
            Duracion_Horas    REAL,
            Estado_OP         TEXT DEFAULT "ABIERTO",
            Avance_SAP        REAL DEFAULT 0,
            Mes               INTEGER DEFAULT 0,
            Fecha_Conta       TEXT,
            Fecha_Gantt_Desde TEXT
        );

        CREATE TABLE IF NOT EXISTS avance_real (
            id            INTEGER PRIMARY KEY AUTOINCREMENT,
            BELNR_ID      INTEGER,
            ItemCode      TEXT,
            Descripcion   TEXT,
            Fecha         TEXT,
            Cantidad_Real REAL,
            UNIQUE(BELNR_ID, Fecha)
        );
    """)
    con.commit()
    con.close()


# ── Carga desde Excel ─────────────────────────────────────────────────────────

FECHAS_EXCEL = "fechas_calculadas.xlsx"

def cargar_plan_desde_excel(db_path: str, excel_path: str):
    import os
    try:
        df = pd.read_excel(excel_path, sheet_name="Ordenes", dtype={"ItemCode": str})
    except Exception:
        df = pd.read_excel(excel_path, sheet_name=0, dtype={"ItemCode": str})

    df.columns = df.columns.str.strip()
    df = df.rename(columns={
        "Descripción del artículo": "Descripcion",
        "Descripción":              "Descripcion",
    })
    df = df.dropna(subset=["BELNR_ID"])
    df["BELNR_ID"] = df["BELNR_ID"].astype(int)

    # Limpiar filas vacías
    df = df.dropna(subset=["Sec", "Planificado"])
    df["Sec"] = df["Sec"].astype(int)

    # Si existe fechas_calculadas.xlsx, leer fechas de ahí
    _usar_fechas = False
    if os.path.exists(FECHAS_EXCEL):
        try:
            df_fc = pd.read_excel(FECHAS_EXCEL, sheet_name="Fechas", dtype={"ItemCode": str}, header=1)
            df_fc.columns = df_fc.columns.str.strip()
            df_fc = df_fc.rename(columns={
                "O. Prod.":      "BELNR_ID",
                "Hora Ini.":     "Hora_Inicio",
                "Fecha Gantt":   "Fecha_Gantt_Desde",
                "Fecha Inicio":  "Fecha_Inicio",
                "Fecha Fin":     "Fecha_Fin",
                "Hora Fin":      "Hora_Fin",
                "Cap. diaria":   "Cap_Diaria",
                "Dur.(hs)":      "Duracion_Horas",
                "Estado":        "Estado_OP",
                "Avance SAP":    "Avance_SAP",
                "Fecha Conta":   "Fecha_Conta",
                "Hs. Prog.":     "Horas_Prog",
            })
            df_fc["BELNR_ID"] = df_fc["BELNR_ID"].astype(int)
            cols_fc = [c for c in ["BELNR_ID","Fecha_Inicio","Hora_Inicio","Fecha_Fin","Hora_Fin",
                                    "Cap_Diaria","Duracion_Horas","Estado_OP","Avance_SAP",
                                    "Fecha_Conta","Fecha_Gantt_Desde"] if c in df_fc.columns]
            df = df.merge(df_fc[cols_fc], on="BELNR_ID", how="left")
            _usar_fechas = True
            print(f"  → Fechas leídas desde {FECHAS_EXCEL}")
        except Exception as e:
            print(f"  ⚠️  {FECHAS_EXCEL} no disponible: {e}")

    df["Cantidad_Base"]  = limpiar_cantidad(df["Cantidad Base"])
    df["Planificado"]    = limpiar_cantidad(df["Planificado"])
    df["Horas"]          = limpiar_cantidad(df["Horas"])
    df["Horas_Prog"]     = limpiar_cantidad(df["Horas Programadas"]).astype(int)
    df["Proceso"]        = df["Proceso"].fillna("Sin Proceso").str.strip()
    # Agrupar Inyección y Soplado bajo Plásticos
    plasticos = ["Inyección", "Inyeccion", "inyección", "inyeccion",
                 "Soplado", "soplado", "INYECCIÓN", "SOPLADO"]
    df["Proceso"] = df["Proceso"].apply(lambda p: "Plásticos" if p in plasticos else p)
    df["Linea"]          = df["Línea"].fillna("").str.strip() if "Línea" in df.columns else ""
    df["Subcomponente"]  = df["Subcomponente"].fillna("Sin clasificar").str.strip() if "Subcomponente" in df.columns else "Sin clasificar"
    df["Fecha_Inicio_Base"] = pd.to_datetime(df["Fecha Inicio"], dayfirst=True).dt.normalize()

    registros = []
    for (mes, maquina), grupo in df.groupby(["Mes", "Maquina"], sort=True):
        grupo = grupo.sort_values("Sec").reset_index(drop=True)
        horas_prog = int(grupo.loc[0, "Horas_Prog"])
        es_paralela = maquina in MAQUINAS_PARALELAS

        primera_fecha = grupo.loc[0, "Fecha_Inicio_Base"]
        cursor = datetime(primera_fecha.year, primera_fecha.month,
                          primera_fecha.day, HORA_INICIO)
        cursor = siguiente_momento_habil(cursor, horas_prog)

        for _, row in grupo.iterrows():
            hp = int(row["Horas_Prog"])
            horas = float(row["Horas"]) if float(row.get("Horas", 0)) > 0 else 1
            velocidad      = float(row["Cantidad_Base"]) / horas
            cap_diaria     = velocidad * hp
            duracion_horas = float(row["Planificado"]) / velocidad if velocidad > 0 else 0

            if _usar_fechas and pd.notna(row.get("Fecha_Inicio")):
                # Usar fechas del Excel congelado
                dt_inicio = datetime.fromisoformat(str(row["Fecha_Inicio"])[:10] + " " + str(row.get("Hora_Inicio","07:00")))
                dt_fin    = datetime.fromisoformat(str(row["Fecha_Fin"])[:10]    + " " + str(row.get("Hora_Fin","07:00")))
                cap_diaria     = float(row.get("Cap_Diaria", cap_diaria))
                duracion_horas = float(row.get("Duracion_Horas", duracion_horas))
            elif es_paralela:
                f_base = row["Fecha_Inicio_Base"]
                dt_inicio = siguiente_momento_habil(
                    datetime(f_base.year, f_base.month, f_base.day, HORA_INICIO), hp)
                dt_fin = sumar_horas_habiles(dt_inicio, duracion_horas, hp)
            else:
                dt_inicio = cursor
                dt_fin    = sumar_horas_habiles(cursor, duracion_horas, hp)
                cursor    = dt_fin

            registros.append({
                "BELNR_ID":          int(row["BELNR_ID"]),
                "ItemCode":          str(row["ItemCode"]).strip(),
                "Descripcion":       row["Descripcion"],
                "Sec":               int(row["Sec"]),
                "Maquina":           maquina,
                "Proceso":           row["Proceso"],
                "Linea":             row.get("Linea", row.get("Línea","")),
                "Subcomponente":     row.get("Subcomponente",""),
                "Horas_Prog":        hp,
                "Fecha_Inicio":      dt_inicio.strftime("%Y-%m-%d"),
                "Hora_Inicio":       dt_inicio.strftime("%H:%M"),
                "Fecha_Fin":         dt_fin.strftime("%Y-%m-%d"),
                "Hora_Fin":          dt_fin.strftime("%H:%M"),
                "Velocidad":         round(velocidad, 4),
                "Cap_Diaria":        round(cap_diaria, 0),
                "Planificado":       float(row["Planificado"]),
                "Duracion_Horas":    round(duracion_horas, 2),
                "Estado_OP":         str(row.get("Estado_OP", row.get("Estado","ABIERTO"))).strip().upper(),
                "Avance_SAP":        float(limpiar_cantidad(pd.Series([row.get("Avance_SAP",0)])).iloc[0]),
                "Mes":               int(mes),
                "Fecha_Conta":       str(pd.to_datetime(row["Fecha_Conta"], dayfirst=True).date())
                                     if pd.notna(row.get("Fecha_Conta")) and str(row.get("Fecha_Conta","")).strip()
                                     else dt_inicio.strftime("%Y-%m-%d"),
                "Fecha_Gantt_Desde": str(pd.to_datetime(row["Fecha_Gantt_Desde"], dayfirst=True).date())
                                     if pd.notna(row.get("Fecha_Gantt_Desde")) and str(row.get("Fecha_Gantt_Desde","")).strip()
                                     else dt_inicio.strftime("%Y-%m-%d"),
            })

    out = pd.DataFrame(registros)
    con = sqlite3.connect(db_path)
    out.to_sql("ordenes_plan", con, if_exists="replace", index=False)
    con.close()
    print(f"  → {len(out)} órdenes cargadas")


def cargar_ingresos_desde_excel(db_path: str, excel_path: str):
    """Deshabilitado — el avance real se registra solo desde el dashboard."""
    pass


# ── Queries ───────────────────────────────────────────────────────────────────

def get_ordenes_activas_en_dia(db_path: str, fecha: date) -> pd.DataFrame:
    fecha_str = str(fecha)
    con = sqlite3.connect(db_path)
    query = """
        SELECT
            p.BELNR_ID, p.ItemCode, p.Descripcion, p.Sec,
            p.Maquina, p.Proceso, p.Linea, p.Horas_Prog,
            p.Fecha_Inicio, p.Hora_Inicio, p.Fecha_Fin, p.Hora_Fin,
            p.Cap_Diaria, p.Planificado, p.Duracion_Horas,
            COALESCE(acum.total_real, 0)               AS Acumulado_Real,
            COALESCE(dia.Cantidad_Real, 0)             AS Real_Dia,
            p.Planificado - COALESCE(acum.total_real, 0) AS Pendiente
        FROM ordenes_plan p
        LEFT JOIN (
            SELECT BELNR_ID, SUM(Cantidad_Real) AS total_real
            FROM avance_real WHERE Fecha <= :fecha
            GROUP BY BELNR_ID
        ) acum ON p.BELNR_ID = acum.BELNR_ID
        LEFT JOIN avance_real dia
            ON p.BELNR_ID = dia.BELNR_ID AND dia.Fecha = :fecha
        WHERE (
            -- Todas las ABIERTAS del mes (para poder registrar avance)
            (
              COALESCE(p.Estado_OP, 'ABIERTO') != 'CERRADO'
              AND (
                CAST(COALESCE(p.Mes, 0) AS INTEGER) = 0
                OR CAST(COALESCE(p.Mes, 0) AS INTEGER) = CAST(strftime('%m', :fecha) AS INTEGER)
                OR CAST(strftime('%m', p.Fecha_Fin)    AS INTEGER) = CAST(strftime('%m', :fecha) AS INTEGER)
                OR CAST(strftime('%m', p.Fecha_Inicio) AS INTEGER) = CAST(strftime('%m', :fecha) AS INTEGER)
              )
            )
            OR (
              -- CERRADAS del mes para historial
              COALESCE(p.Estado_OP, 'ABIERTO') = 'CERRADO'
              AND (
                CAST(COALESCE(p.Mes, 0) AS INTEGER) = CAST(strftime('%m', :fecha) AS INTEGER)
                OR CAST(strftime('%m', p.Fecha_Inicio) AS INTEGER) = CAST(strftime('%m', :fecha) AS INTEGER)
              )
            )
          )
        ORDER BY p.Proceso, p.Maquina, p.Sec
    """
    df = pd.read_sql(query, con, params={"fecha": fecha_str})
    con.close()

    from datetime import timedelta as _td

    def calcular_plan_dia(row, fecha):
        """
        Plan teórico del día.
        Si hay avance real hasta ayer → usa ese para calcular el restante.
        Si no → usa plan teórico acumulado por días hábiles.
        """
        f_inicio  = date.fromisoformat(str(row["Fecha_Inicio"])[:10])
        cap       = row["Cap_Diaria"]
        planif    = row["Planificado"]
        acum_real = row["Acumulado_Real"] - row["Real_Dia"]

        if acum_real > 0:
            restante = max(planif - acum_real, 0)
            return min(cap, restante)

        dias_habiles_prev = 0
        d = f_inicio
        while d < fecha:
            if es_habil(d):
                dias_habiles_prev += 1
            d += _td(days=1)

        ya_planificado = min(cap * dias_habiles_prev, planif)
        restante = max(planif - ya_planificado, 0)
        return min(cap, restante)

    # Calcular Plan_Dia fila por fila con loop explícito (evita problemas de apply)
    plan_dia_vals = []
    for i in range(len(df)):
        row = df.iloc[i]
        try:
            val = float(calcular_plan_dia(row, fecha))
        except Exception:
            val = 0.0
        plan_dia_vals.append(val)
    df = df.copy()
    df["Plan_Dia"] = plan_dia_vals

    # ── Solapamiento de OPs en la misma máquina ──────────────────────────────
    df = df.sort_values(["Maquina", "Fecha_Inicio", "Hora_Inicio"]).reset_index(drop=True)
    maq_usado = {}

    plan_ajustado = []
    for i in range(len(df)):
        row       = df.iloc[i]
        maq       = row["Maquina"]
        cap       = float(row["Cap_Diaria"])
        hp        = int(row["Horas_Prog"])
        velocidad = cap / hp if hp > 0 else 0

        usado      = maq_usado.get(maq, 0.0)
        disponible = max(hp - usado, 0.0)

        if velocidad > 0 and disponible > 0:
            plan_orig = float(row["Plan_Dia"])
            horas_op  = min(plan_orig / velocidad, disponible)
            plan_aj   = round(horas_op * velocidad)
            maq_usado[maq] = usado + horas_op
        else:
            plan_aj = 0

        plan_ajustado.append(plan_aj)

    df["Plan_Dia"] = plan_ajustado
    return df


def get_todas_ordenes(db_path: str) -> pd.DataFrame:
    con = sqlite3.connect(db_path)
    df = pd.read_sql("SELECT * FROM ordenes_plan ORDER BY Proceso, Maquina, Sec", con)
    con.close()
    return df


def get_todos_ingresos(db_path: str) -> pd.DataFrame:
    con = sqlite3.connect(db_path)
    df = pd.read_sql("SELECT * FROM avance_real ORDER BY Fecha, BELNR_ID", con)
    con.close()
    return df


def get_primera_fecha(db_path: str) -> str:
    try:
        con = sqlite3.connect(db_path)
        row = con.execute("SELECT MIN(Fecha_Inicio) FROM ordenes_plan").fetchone()
        con.close()
        if row and row[0]:
            return row[0]
    except Exception:
        pass
    return str(date.today())


# ── MRP ───────────────────────────────────────────────────────────────────────

ALMACENES_STOCK = ["ALMA002", "ALMA089"]
ALMACENES_PROD  = ["ALMA070", "ALMA071", "ALMA072"]

def cargar_bom_stock(db_path: str, excel_path: str):
    """Carga BOM y Stock desde Excel a SQLite."""
    # BOM — renombrar por posición para evitar problema de columnas duplicadas
    try:
        bom = pd.read_excel(excel_path, sheet_name="BOM", header=None)
    except Exception:
        bom = pd.read_excel(excel_path, sheet_name=1, header=None)

    # Asignar nombres fijos por posición de columna
    bom.columns = ["Codigo_Padre","Posicion","Desc_Padre","Und_Padre",
                   "Qty_Padre","Proceso_BOM","Codigo_Comp","Desc_Comp",
                   "Und_Comp","Qty_Comp"]
    bom = bom.iloc[1:].reset_index(drop=True)  # quitar fila de encabezado

    bom["Codigo_Padre"] = bom["Codigo_Padre"].astype(str).str.strip()
    bom["Codigo_Comp"]  = bom["Codigo_Comp"].astype(str).str.strip()
    bom["Qty_Padre"]    = pd.to_numeric(bom["Qty_Padre"], errors="coerce").fillna(1)
    bom["Qty_Comp"]     = pd.to_numeric(bom["Qty_Comp"],  errors="coerce").fillna(0)
    bom["Factor"]       = bom["Qty_Comp"] / bom["Qty_Padre"]
    bom = bom.dropna(subset=["Codigo_Padre","Codigo_Comp"])

    con = sqlite3.connect(db_path)
    bom[["Codigo_Padre","Codigo_Comp","Desc_Padre","Desc_Comp",
         "Qty_Padre","Qty_Comp","Factor","Proceso_BOM"]].to_sql(
        "bom", con, if_exists="replace", index=False)
    con.close()
    print(f"  → {len(bom)} líneas de BOM cargadas")

    # Stock
    try:
        stk = pd.read_excel(excel_path, sheet_name="Stock")
    except Exception:
        try:
            stk = pd.read_excel(excel_path, sheet_name=2)
        except Exception:
            print("  ⚠️  No se encontró hoja Stock")
            return
    stk.columns = stk.columns.str.strip()
    stk = stk.rename(columns={
        "Número de artículo":   "ItemCode",
        "Descripción del artículo": "Descripcion",
        "Código de almacén":    "Almacen",
        "En stock":             "En_Stock",
    })
    stk["ItemCode"] = stk["ItemCode"].astype(str).str.strip()
    stk["En_Stock"] = pd.to_numeric(
        stk["En_Stock"].astype(str).str.replace(",",""), errors="coerce").fillna(0)

    con = sqlite3.connect(db_path)
    stk[["ItemCode","Descripcion","Almacen","En_Stock"]].to_sql(
        "stock", con, if_exists="replace", index=False)
    con.close()
    print(f"  → {len(stk)} registros de stock cargados")

    # BD_Spec - Tipo_Material2, Subcomponente, Lineas
    try:
        spec_raw = pd.read_excel(excel_path, sheet_name="BD_Spec")
        spec_raw.columns = spec_raw.columns.str.strip()
        # Mapeo directo por nombre exacto del Excel
        rename_map = {
            "Código_Material":  "ItemCode",
            "Tipo_Material2":   "Tipo_Material2",
            "Subcomponente":    "Subcomponente",
            "Líneas":           "Lineas",   # con tilde
            "Lineas":           "Lineas",   # sin tilde
        }
        spec_raw = spec_raw.rename(columns=rename_map)
        for c in ["ItemCode","Tipo_Material2","Subcomponente","Lineas"]:
            if c not in spec_raw.columns:
                spec_raw[c] = "Sin clasificar"
        spec = spec_raw[["ItemCode","Tipo_Material2","Subcomponente","Lineas"]].copy()
        spec["ItemCode"]      = spec["ItemCode"].astype(str).str.strip()
        spec["Tipo_Material2"]= spec["Tipo_Material2"].fillna("Sin clasificar").str.strip()
        spec["Subcomponente"] = spec["Subcomponente"].fillna("Sin clasificar").str.strip()
        spec["Lineas"]        = spec["Lineas"].fillna("Sin clasificar").str.strip()
        con = sqlite3.connect(db_path)
        spec.to_sql("bd_spec", con, if_exists="replace", index=False)
        con.close()
        print(f"  -> {len(spec)} registros de BD_Spec cargados")
    except Exception as e:
        print(f"  Warning:  BD_Spec no cargado: {e}")


def _explotar_bom_op(bom_df: pd.DataFrame, codigo: str, cantidad: float,
                     stk_semi: dict, nivel: int = 1) -> list:
    """
    Explosión recursiva BOM con la lógica correcta:

    SEMITERMINADOS (231xxx) — procesados en recursión:
      - Consulta stock disponible en stk_semi (compartido entre PTs, se descuenta en memoria)
      - Si stock cubre toda la necesidad → marca USAR_STOCK, no explota hijos
      - Si no alcanza → explota solo la diferencia (necesidad - stock)
      - El stock se reinicia por PT raíz (ver get_mrp_dia)

    COMPRADOS (no 231xxx) — nivel hoja:
      - Se acumulan como bruto, el descuento de stock se hace consolidado al final
    """
    if nivel > 10 or cantidad <= 0:
        return []

    hijos = bom_df[bom_df["Codigo_Padre"] == codigo]
    if hijos.empty:
        return []

    resultado = []
    for _, row in hijos.iterrows():
        comp    = str(row["Codigo_Comp"]).strip()
        req     = cantidad * row["Factor"]
        es_semi = comp.startswith("231")

        if es_semi:
            # Consultar stock disponible de este semiterminado
            disp = stk_semi.get(comp, 0)
            neto = max(req - disp, 0)
            # Descontar stock en memoria
            stk_semi[comp] = max(disp - req, 0)

            resultado.append({
                "Codigo_Comp": comp,
                "Desc_Comp":   str(row["Desc_Comp"]),
                "Req_Bruto":   req,
                "Req_Neto":    neto,
                "Stock_Comp":  min(disp, req),  # cuánto del stock se usa
                "Nivel":       nivel,
                "Es_Semi":     True,
                "Usar_Stock":  neto == 0,
            })

            # Solo explota hacia abajo si hay neto pendiente
            if neto > 0:
                sub = _explotar_bom_op(bom_df, comp, neto, stk_semi, nivel + 1)
                resultado.extend(sub)
        else:
            # Comprado — acumular bruto, descuento se hace al final consolidado
            resultado.append({
                "Codigo_Comp": comp,
                "Desc_Comp":   str(row["Desc_Comp"]),
                "Req_Bruto":   req,
                "Req_Neto":    req,   # se recalcula al consolidar
                "Stock_Comp":  0,
                "Nivel":       nivel,
                "Es_Semi":     False,
                "Usar_Stock":  False,
            })

    return resultado


def get_mrp_dia(db_path: str, fecha: date) -> pd.DataFrame:
    """
    Explota el BOM multinivel del plan del día.
    - Semiterminados (231xxx): cuánto producir
    - Materias primas (otros): cuánto comprar
    Solo explota OPs que están en el plan — no agrega materiales de productos no planificados.
    """
    df_plan = get_ordenes_activas_en_dia(db_path, fecha)
    if df_plan.empty:
        return pd.DataFrame()

    con = sqlite3.connect(db_path)
    bom     = pd.read_sql("SELECT * FROM bom", con)
    stk_raw = pd.read_sql("SELECT * FROM stock", con)
    try:
        spec = pd.read_sql("SELECT * FROM bd_spec", con)
    except Exception:
        spec = pd.DataFrame(columns=["ItemCode","Tipo_Material2"])
    con.close()

    bom["Codigo_Padre"] = bom["Codigo_Padre"].astype(str).str.strip()
    bom["Codigo_Comp"]  = bom["Codigo_Comp"].astype(str).str.strip()
    spec["ItemCode"]    = spec["ItemCode"].astype(str).str.strip()

    # Stock total por componente para pasar a la explosión
    def stock_alm(codigos):
        return (stk_raw[stk_raw["Almacen"].isin(codigos)]
                .groupby("ItemCode")["En_Stock"].sum().reset_index())

    s002  = stock_alm(["ALMA002"]).rename(columns={"En_Stock":"Stock_ALMA002"})
    s089  = stock_alm(["ALMA089"]).rename(columns={"En_Stock":"Stock_ALMA089"})
    sprod = stock_alm(ALMACENES_PROD).rename(columns={"En_Stock":"Stock_Prod"})

    stk_total_df = stk_raw[stk_raw["Almacen"].isin(ALMACENES_STOCK + ALMACENES_PROD)]
    stk_map = stk_total_df.groupby("ItemCode")["En_Stock"].sum().to_dict()

    # Explotar BOM desde cada OP usando nueva lógica semi/comprado
    todos_semi    = []  # semiterminados (con descuento en recursión)
    brutos_comp   = []  # comprados (bruto acumulado, descuento al final)

    # Stock de semiterminados compartido entre todas las OPs del día
    # Se descuenta progresivamente — si OP-A consume stock, OP-B lo ve reducido
    stk_semi = {k: v for k, v in stk_map.items() if k.startswith("231")}

    for _, op in df_plan.iterrows():
        item     = str(op["ItemCode"]).strip()
        plan_dia = op["Plan_Dia"]
        reqs = _explotar_bom_op(bom, item, plan_dia, stk_semi)
        for r in reqs:
            if r["Es_Semi"]:
                todos_semi.append(r)
            else:
                brutos_comp.append(r)

    if not todos_semi and not brutos_comp:
        return pd.DataFrame()

    # ── Procesar semiterminados ──────────────────────────────────────────────
    semi_rows = []
    if todos_semi:
        df_semi = pd.DataFrame(todos_semi)
        semi_agg = (df_semi.groupby(["Codigo_Comp","Desc_Comp"], as_index=False)
                    .agg(Req_Bruto=("Req_Bruto","sum"),
                         Req_Neto=("Req_Neto","sum"),
                         Stock_Comp=("Stock_Comp","sum"),
                         Nivel=("Nivel","min")))
        semi_agg["Es_Semi"] = True
        semi_rows = semi_agg.to_dict("records")

    # ── Procesar comprados — descuento consolidado ───────────────────────────
    comp_rows = []
    if brutos_comp:
        df_comp = pd.DataFrame(brutos_comp)
        comp_agg = (df_comp.groupby(["Codigo_Comp","Desc_Comp"], as_index=False)
                    .agg(Req_Bruto=("Req_Bruto","sum"), Nivel=("Nivel","min")))
        comp_agg["Es_Semi"] = False
        # Descuento de stock consolidado
        comp_agg["Stock_Total_Comp"] = comp_agg["Codigo_Comp"].map(
            lambda c: stk_map.get(c, 0))
        comp_agg["Req_Neto"]   = (comp_agg["Req_Bruto"] - comp_agg["Stock_Total_Comp"]).clip(lower=0)
        comp_agg["Stock_Comp"] = comp_agg["Stock_Total_Comp"]
        comp_rows = comp_agg.drop(columns=["Stock_Total_Comp"]).to_dict("records")

    req = pd.DataFrame(semi_rows + comp_rows)

    # Stock desglosado por almacén
    for s, col in [(s002,"Stock_ALMA002"),(s089,"Stock_ALMA089"),(sprod,"Stock_Prod")]:
        req = req.merge(s, left_on="Codigo_Comp", right_on="ItemCode", how="left").drop(
            columns=["ItemCode"], errors="ignore")
        req[col] = req[col].fillna(0)

    req["Stock_Total"] = req["Stock_ALMA002"] + req["Stock_ALMA089"] + req["Stock_Prod"]
    req["Req_Neto"]    = (req["Req_Bruto"] - req["Stock_Total"]).clip(lower=0)
    req["Diferencia"]  = req["Req_Neto"]   # alias claro para la UI
    req["Alerta"]      = req["Req_Neto"] > 0

    # Tipo_Material2 desde BD_Spec
    req = req.merge(spec, left_on="Codigo_Comp", right_on="ItemCode", how="left").drop(
        columns=["ItemCode"], errors="ignore")
    req["Tipo_Material2"] = req["Tipo_Material2"].fillna(
        req["Es_Semi"].apply(lambda x: "Semiterminado" if x else "Sin clasificar"))

    # Etiqueta de tipo para ordenar
    req["Tipo_Item"] = req["Es_Semi"].apply(
        lambda x: "1_Semiterminado a producir" if x else "2_Materia prima a comprar")

    cols = ["Tipo_Item","Tipo_Material2","Codigo_Comp","Desc_Comp","Nivel",
            "Req_Bruto","Stock_ALMA002","Stock_ALMA089","Stock_Prod",
            "Stock_Total","Req_Neto","Alerta"]
    return req[cols].sort_values(["Tipo_Item","Tipo_Material2","Codigo_Comp"]).reset_index(drop=True)


def get_mrp_proyectado(db_path: str) -> pd.DataFrame:
    """
    Calcula el stock proyectado día a día para cada componente
    a lo largo de todo el horizonte del plan.

    Retorna una fila por componente con:
      Tipo_Material2, Codigo_Comp, Desc_Comp,
      Stock_Total (inicial),
      Consumo_Total (todo el plan),
      Consumo_Diario_Prom,
      Dias_Cobertura,
      Fecha_Quiebre (primer día en que stock proyectado <= 0, o None),
      Semaforo ('rojo'|'amarillo'|'verde'),
      + columnas por cada día hábil del plan con stock proyectado
    """
    from datetime import timedelta

    con = sqlite3.connect(db_path)
    ordenes   = pd.read_sql("SELECT * FROM ordenes_plan", con)
    bom       = pd.read_sql("SELECT * FROM bom", con)
    stk_raw   = pd.read_sql("SELECT * FROM stock", con)
    try:
        spec  = pd.read_sql("SELECT * FROM bd_spec", con)
    except Exception:
        spec  = pd.DataFrame(columns=["ItemCode","Tipo_Material2"])
    con.close()

    if ordenes.empty or bom.empty:
        return pd.DataFrame()

    # Rango de días hábiles del plan completo
    fecha_min = date.fromisoformat(ordenes["Fecha_Inicio"].min()[:10])
    fecha_max = date.fromisoformat(ordenes["Fecha_Fin"].max()[:10])
    dias = []
    d = fecha_min
    while d <= fecha_max:
        if es_habil(d):
            dias.append(d)
        d += timedelta(days=1)

    # Stock inicial por componente
    def stock_alm(codigos):
        return (stk_raw[stk_raw["Almacen"].isin(codigos)]
                .groupby("ItemCode")["En_Stock"].sum())

    stk_total = (stock_alm(ALMACENES_STOCK + ALMACENES_PROD)
                 .reset_index()
                 .rename(columns={"En_Stock": "Stock_Total"}))
    stk_total["ItemCode"] = stk_total["ItemCode"].astype(str).str.strip()

    # Consumo diario TEÓRICO por componente — sin considerar avance real
    # Usa solo el plan (Cap_Diaria y Planificado) para proyectar el consumo futuro
    consumo_por_dia = {}   # {fecha: {codigo_comp: cantidad}}

    # Pre-calcular plan teórico acumulado por OP y día
    from datetime import timedelta as _td2

    # Calcular acumulado real por OP para usar Pendiente
    con_av = sqlite3.connect(db_path)
    try:
        av_df = pd.read_sql(
            "SELECT BELNR_ID, SUM(Cantidad_Real) as acum FROM avance_real GROUP BY BELNR_ID",
            con_av)
        acum_real_map = dict(zip(av_df["BELNR_ID"].astype(int), av_df["acum"]))
    except Exception:
        acum_real_map = {}
    con_av.close()

    for dia in dias:
        consumo_dia = {}
        ops_activas = ordenes[
            (ordenes["Fecha_Inicio"] <= str(dia)) &
            (ordenes["Fecha_Fin"]    >= str(dia))
        ]
        for _, op in ops_activas.iterrows():
            f_inicio = date.fromisoformat(str(op["Fecha_Inicio"])[:10])
            cap      = float(op["Cap_Diaria"])
            planif   = float(op["Planificado"])
            belnr    = int(op["BELNR_ID"])
            item     = str(op["ItemCode"]).strip()

            # Pendiente = Planificado - Acumulado_Real
            acum_real = acum_real_map.get(belnr, 0)
            pendiente = max(planif - acum_real, 0)

            if pendiente <= 0:
                continue  # OP ya completada

            # Días hábiles previos al día actual desde f_inicio
            dias_prev = 0
            d = f_inicio
            while d < dia:
                if es_habil(d):
                    dias_prev += 1
                d += _td2(days=1)

            # Plan teórico del día basado en pendiente
            ya_planificado = min(cap * dias_prev, pendiente)
            plan_dia_teo   = min(cap, max(pendiente - ya_planificado, 0))

            if plan_dia_teo <= 0:
                continue

            # Proyección bruta sin descuento de stock
            stk_semi_proy = {}
            reqs = _explotar_bom_op(bom, item, plan_dia_teo, stk_semi_proy)
            for r in reqs:
                comp = r["Codigo_Comp"]
                consumo_dia[comp] = consumo_dia.get(comp, 0) + r["Req_Bruto"]

        consumo_por_dia[dia] = consumo_dia

    # Todos los componentes que aparecen en el plan
    todos_comp = set()
    for d_cons in consumo_por_dia.values():
        todos_comp.update(d_cons.keys())

    # Tipo_Material2 por componente
    spec["ItemCode"] = spec["ItemCode"].astype(str).str.strip()
    tipo_map = spec.set_index("ItemCode")["Tipo_Material2"].to_dict()

    # Desc_Comp por componente
    desc_map = bom.drop_duplicates("Codigo_Comp").set_index("Codigo_Comp")["Desc_Comp"].to_dict()

    # Stock inicial por componente
    stk_map = stk_total.set_index("ItemCode")["Stock_Total"].to_dict()

    registros = []
    for comp in sorted(todos_comp):
        stock_ini = stk_map.get(comp, 0)
        stock_act = stock_ini
        fecha_quiebre = None
        stock_proy = {}   # {fecha: stock al final del día}

        consumo_total = 0
        dias_con_consumo = 0

        for dia in dias:
            consumo_dia = consumo_por_dia[dia].get(comp, 0)
            consumo_total += consumo_dia
            if consumo_dia > 0:
                dias_con_consumo += 1
            stock_act = max(stock_act - consumo_dia, 0)
            stock_proy[dia] = round(stock_act, 0)

            # Detectar primer quiebre
            if fecha_quiebre is None and stock_act <= 0 and consumo_dia > 0:
                fecha_quiebre = dia

        consumo_diario_prom = consumo_total / max(dias_con_consumo, 1)
        # Días de cobertura hábiles reales
        dias_cobertura = 0
        stk_temp = stock_ini
        for dia_c in dias:
            consumo_c = consumo_por_dia[dia_c].get(comp, 0)
            if consumo_c > 0:
                if stk_temp <= 0:
                    break
                stk_temp -= consumo_c
                dias_cobertura += 1
            if stk_temp <= 0:
                break
        if consumo_total == 0:
            dias_cobertura = 999

        # Semáforo: rojo < 7 días, amarillo < 15 días, verde >= 30 días
        if fecha_quiebre is None:
            semaforo = "verde"
        elif dias_cobertura < 7:
            semaforo = "rojo"
        elif dias_cobertura < 15:
            semaforo = "amarillo"
        else:
            semaforo = "verde"

        diferencia = max(round(consumo_total - stock_ini, 0), 0)
        row = {
            "Tipo_Material2":      tipo_map.get(comp, "Sin clasificar"),
            "Codigo_Comp":         comp,
            "Desc_Comp":           desc_map.get(comp, ""),
            "Stock_Inicial":       round(stock_ini, 0),
            "Consumo_Total":       round(consumo_total, 0),
            "Diferencia":          diferencia,
            "Consumo_Diario_Prom": round(consumo_diario_prom, 1),
            "Dias_Cobertura":      dias_cobertura,
            "Fecha_Quiebre":       fecha_quiebre.strftime("%d/%m/%Y") if fecha_quiebre else "✅ Sin quiebre",
            "Semaforo":            semaforo,
        }
        # Agregar stock proyectado por día
        for dia in dias:
            row[dia.strftime("%d/%m")] = stock_proy.get(dia, round(stock_ini, 0))

        registros.append(row)

    df = pd.DataFrame(registros)
    return df.sort_values(["Semaforo","Tipo_Material2","Codigo_Comp"],
                          key=lambda x: x.map({"rojo":0,"amarillo":1,"verde":2})
                          if x.name == "Semaforo" else x).reset_index(drop=True)



def get_cumplimiento_detalle(db_path: str, fecha_ref: date = None) -> dict:
    """
    Cumplimiento semanal completo para la pestaña de cumplimiento.
    Retorna KPIs, detalle diario, detalle por proceso y detalle por orden.
    """
    if fecha_ref is None:
        fecha_ref = date.today()

    lunes = fecha_ref - timedelta(days=fecha_ref.weekday())
    dias_semana = [lunes + timedelta(days=i) for i in range(5)
                   if es_habil(lunes + timedelta(days=i))]

    detalle_dia     = []
    detalle_proceso = {}
    detalle_ordenes = []
    plan_total = 0
    real_total = 0

    for dia in dias_semana:
        df_dia = get_ordenes_activas_en_dia(db_path, dia)
        if df_dia.empty:
            detalle_dia.append({
                "fecha": dia.strftime("%a %d/%m"), "plan": 0, "real": 0, "pct": 0})
            continue

        plan_dia = df_dia["Plan_Dia"].sum()
        real_dia = df_dia["Real_Dia"].sum()
        pct_dia  = round(real_dia / plan_dia * 100, 1) if plan_dia > 0 else 0
        plan_total += plan_dia
        real_total += real_dia

        detalle_dia.append({
            "fecha": dia.strftime("%a %d/%m"),
            "plan":  round(plan_dia),
            "real":  round(real_dia),
            "pct":   pct_dia,
        })

        # Por proceso
        for proc, grp in df_dia.groupby("Proceso"):
            if proc not in detalle_proceso:
                detalle_proceso[proc] = {"plan": 0, "real": 0}
            detalle_proceso[proc]["plan"] += grp["Plan_Dia"].sum()
            detalle_proceso[proc]["real"] += grp["Real_Dia"].sum()

        # Por orden — acumular semana
        for _, row in df_dia.iterrows():
            bid = int(row["BELNR_ID"])
            found = next((x for x in detalle_ordenes if x["BELNR_ID"] == bid), None)
            if found:
                found["plan_sem"] += row["Plan_Dia"]
                found["real_sem"] += row["Real_Dia"]
            else:
                detalle_ordenes.append({
                    "BELNR_ID":   bid,
                    "Maquina":    row["Maquina"],
                    "Proceso":    row["Proceso"],
                    "Descripcion": row["Descripcion"][:45],
                    "Planificado": int(row["Planificado"]),
                    "plan_sem":   row["Plan_Dia"],
                    "real_sem":   row["Real_Dia"],
                })

    # Calcular % por orden
    for o in detalle_ordenes:
        o["pct"] = round(o["real_sem"] / o["plan_sem"] * 100, 1) if o["plan_sem"] > 0 else 0
        o["plan_sem"] = round(o["plan_sem"])
        o["real_sem"] = round(o["real_sem"])
        if o["pct"] >= 95:   o["estado"] = "Completo"
        elif o["pct"] >= 75: o["estado"] = "Parcial"
        elif o["pct"] > 0:   o["estado"] = "Atrasado"
        else:                 o["estado"] = "Pendiente"

    # % por proceso — incluye meta del día (cuánto debería llevar acumulado)
    # Meta del día = (días transcurridos / días totales semana) × 100%
    dias_transcurridos = sum(1 for d in detalle_dia if d["real"] > 0 or d["plan"] > 0)
    dias_habiles_semana = len(dias_semana)
    pct_meta_dia = round(dias_transcurridos / dias_habiles_semana * 100, 1) if dias_habiles_semana > 0 else 0

    proc_list = []
    for proc, v in detalle_proceso.items():
        pct_real = round(v["real"] / v["plan"] * 100, 1) if v["plan"] > 0 else 0
        # Meta acumulada = plan_proceso × (días transcurridos / días totales)
        meta_acum = round(v["plan"] * dias_transcurridos / dias_habiles_semana) if dias_habiles_semana > 0 else 0
        pct_vs_meta = round(v["real"] / meta_acum * 100, 1) if meta_acum > 0 else 0
        proc_list.append({
            "proceso":     proc,
            "plan":        round(v["plan"]),
            "real":        round(v["real"]),
            "pct":         pct_real,          # % real vs plan total
            "meta_acum":   meta_acum,          # cuánto debería llevar
            "pct_vs_meta": pct_vs_meta,        # nota real = real / meta_acum
            "pct_meta_dia": pct_meta_dia,      # % del plan que debería estar hecho
        })
    proc_list.sort(key=lambda x: x["pct_vs_meta"])

    pct_semana = round(real_total / plan_total * 100, 1) if plan_total > 0 else 0
    ordenes_ok = sum(1 for o in detalle_ordenes if o["estado"] == "Completo")

    return {
        "semana":          f"{lunes.strftime('%d/%m')} – {dias_semana[-1].strftime('%d/%m') if dias_semana else ''}",
        "plan_total":      round(plan_total),
        "real_total":      round(real_total),
        "pct_semana":      pct_semana,
        "ordenes_total":   len(detalle_ordenes),
        "ordenes_ok":      ordenes_ok,
        "detalle_dia":     detalle_dia,
        "detalle_proceso": proc_list,
        "detalle_ordenes": detalle_ordenes,
    }

def get_cumplimiento_semanal(db_path: str, fecha_ref: date = None) -> dict:
    """
    Calcula el cumplimiento de producción de la semana actual (lun-vie).
    Retorna dict con plan_semana, real_semana, pct, y detalle por día.
    """
    if fecha_ref is None:
        fecha_ref = date.today()

    # Lunes de la semana
    lunes = fecha_ref - timedelta(days=fecha_ref.weekday())
    dias_semana = [lunes + timedelta(days=i) for i in range(5)
                   if es_habil(lunes + timedelta(days=i))]

    con = sqlite3.connect(db_path)
    detalle = []
    plan_total = 0
    real_total = 0

    for dia in dias_semana:
        df_dia = get_ordenes_activas_en_dia(db_path, dia)
        if df_dia.empty:
            detalle.append({"fecha": str(dia), "plan": 0, "real": 0, "pct": 0})
            continue
        plan_dia = df_dia["Plan_Dia"].sum()
        real_dia = df_dia["Real_Dia"].sum()
        pct_dia  = round(real_dia / plan_dia * 100, 1) if plan_dia > 0 else 0
        plan_total += plan_dia
        real_total += real_dia
        detalle.append({
            "fecha": dia.strftime("%a %d/%m"),
            "plan":  round(plan_dia),
            "real":  round(real_dia),
            "pct":   pct_dia,
        })

    con.close()
    pct_semana = round(real_total / plan_total * 100, 1) if plan_total > 0 else 0
    return {
        "semana":      f"{lunes.strftime('%d/%m')} – {dias_semana[-1].strftime('%d/%m') if dias_semana else ''}",
        "plan_total":  round(plan_total),
        "real_total":  round(real_total),
        "pct_semana":  pct_semana,
        "detalle":     detalle,
    }


def get_semanas_plan(db_path: str) -> list:
    """Devuelve lista de semanas con OPs activas en el plan."""
    con = sqlite3.connect(db_path)
    df = pd.read_sql("SELECT Fecha_Inicio, Fecha_Fin FROM ordenes_plan", con)
    con.close()
    if df.empty:
        return []

    fecha_min = date.fromisoformat(df["Fecha_Inicio"].min()[:10])
    fecha_max = date.fromisoformat(df["Fecha_Fin"].max()[:10])

    # Lunes de cada semana en el rango
    semanas = []
    lunes = fecha_min - timedelta(days=fecha_min.weekday())
    while lunes <= fecha_max:
        viernes = lunes + timedelta(days=4)
        semana_num = lunes.isocalendar()[1]
        semanas.append({
            "label": f"Sem {semana_num} — {lunes.strftime('%d/%m')} al {viernes.strftime('%d/%m/%Y')}",
            "value": str(lunes),
        })
        lunes += timedelta(weeks=1)
    return semanas


def get_resumen_ops(db_path: str) -> pd.DataFrame:
    """
    Resumen de todas las OPs con estado actual basado en avance real.
    """
    con = sqlite3.connect(db_path)
    query = """
        SELECT
            p.BELNR_ID, p.Proceso, p.Linea, p.Maquina, p.Sec,
            p.ItemCode, p.Descripcion,
            p.Fecha_Inicio, p.Hora_Inicio,
            p.Fecha_Fin,   p.Hora_Fin,
            p.Planificado, p.Cap_Diaria, p.Duracion_Horas,
            COALESCE(a.acum, 0) AS Acumulado,
            p.Planificado - COALESCE(a.acum, 0) AS Pendiente
        FROM ordenes_plan p
        LEFT JOIN (
            SELECT BELNR_ID, SUM(Cantidad_Real) AS acum
            FROM avance_real GROUP BY BELNR_ID
        ) a ON p.BELNR_ID = a.BELNR_ID
        ORDER BY p.Proceso, p.Linea, p.Maquina, p.Sec
    """
    df = pd.read_sql(query, con)
    con.close()

    df["Pct_Acum"] = (df["Acumulado"] / df["Planificado"] * 100).round(1).fillna(0)
    df["Inicio"]   = df["Fecha_Inicio"].str[5:].str.replace("-","/") + " " + df["Hora_Inicio"]
    df["Fin"]      = df["Fecha_Fin"].str[5:].str.replace("-","/")    + " " + df["Hora_Fin"]

    def estado(pct, pendiente):
        if pendiente <= 0:      return "Completado"
        if pct >= 75:           return "En curso"
        if pct > 0:             return "Parcial"
        return "Pendiente"

    df["Estado"] = df.apply(lambda r: estado(r["Pct_Acum"], r["Pendiente"]), axis=1)
    return df


def get_avance_campana(db_path: str, fecha_ref: date = None) -> pd.DataFrame:
    """Avance de campana agrupado por Linea, Subcomponente y Maquina."""
    if fecha_ref is None:
        fecha_ref = date.today()

    con = sqlite3.connect(db_path)
    df_plan = pd.read_sql("SELECT * FROM ordenes_plan", con)
    df_av   = pd.read_sql(
        "SELECT BELNR_ID, SUM(Cantidad_Real) as acum FROM avance_real "
        "WHERE Fecha <= ? GROUP BY BELNR_ID",
        con, params=[str(fecha_ref if fecha_ref else date.today())])
    con.close()

    if df_plan.empty:
        return pd.DataFrame()

    df_plan = df_plan.merge(df_av, on="BELNR_ID", how="left")
    df_plan["acum"] = df_plan["acum"].fillna(0)
    # Subcomponente viene directo de ordenes_plan
    if "Subcomponente" not in df_plan.columns:
        df_plan["Subcomponente"] = "Sin clasificar"
    df_plan["Subcomponente"] = df_plan["Subcomponente"].fillna("Sin clasificar")

    fecha_str = str(fecha_ref)
    # Todas las maquinas que ya arrancaron hasta hoy
    arrancadas = set(df_plan[
        df_plan["Fecha_Inicio"] <= fecha_str
    ]["Maquina"].unique())

    if not arrancadas:
        return pd.DataFrame()

    resultado = []
    for (linea_v, maquina), grp in df_plan.groupby(["Linea","Maquina"]):
        if maquina not in arrancadas:
            continue
        # Solo OPs que ya arrancaron
        grp_arr = grp[grp["Fecha_Inicio"] <= fecha_str]
        if grp_arr.empty:
            continue

        planificado = float(grp_arr["Planificado"].sum())
        acumulado   = float(grp["acum"].sum())  # acumulado de todas las OPs
        cap_diaria  = float(grp_arr["Cap_Diaria"].max())
        horas_prog  = int(grp_arr["Horas_Prog"].iloc[0])
        f_inicio    = date.fromisoformat(grp_arr["Fecha_Inicio"].min()[:10])

        dias_trans = 0
        d = f_inicio
        while d <= fecha_ref:
            if es_habil(d):
                dias_trans += 1
            d += timedelta(days=1)

        deberia_ir  = min(cap_diaria * dias_trans, planificado)
        diferencia  = acumulado - deberia_ir
        dias_atraso = round(-diferencia / cap_diaria, 1) if diferencia < 0 and cap_diaria > 0 else 0
        pct_acum    = round(acumulado / planificado * 100, 1) if planificado > 0 else 0
        pct_deberia = round(deberia_ir / planificado * 100, 1) if planificado > 0 else 0

        if diferencia >= 0:     estado = "Al dia"
        elif dias_atraso <= 1:  estado = "Leve retraso"
        elif dias_atraso <= 3:  estado = "Retraso"
        else:                   estado = "Critico"

        # Resumen ejecutivo: "X und atrasado = Y días"
        und_atraso = round(-diferencia) if diferencia < 0 else 0
        resumen = (f"-{und_atraso:,} und = {dias_atraso} días"
                   if diferencia < 0 else "Al día ✅")

        resultado.append({
            "Linea":              linea_v,
            "Maquina":            maquina,
            "Planificado":        round(planificado),
            "Deberia_Ir":         round(deberia_ir),
            "Acumulado":          round(acumulado),
            "Diferencia":         round(diferencia),
            "Pct_Acum":           pct_acum,
            "Pct_Deberia":        pct_deberia,
            "Dias_Atraso":        dias_atraso,
            "Cap_Diaria":         round(cap_diaria),
            "Dias_Transcurridos": dias_trans,
            "Estado":             estado,
            "Resumen":            resumen,
        })

    df = pd.DataFrame(resultado)
    if df.empty:
        return df
    orden = {"Critico":0,"Retraso":1,"Leve retraso":2,"Al dia":3}
    df["_ord"] = df["Estado"].map(orden)
    return df.sort_values(["_ord","Linea","Maquina"]).drop(columns=["_ord"]).reset_index(drop=True)


def get_resumen_ops(db_path: str) -> pd.DataFrame:
    """
    Resumen de todas las OPs con estado actual basado en avance real.
    """
    con = sqlite3.connect(db_path)
    query = """
        SELECT
            p.BELNR_ID, p.Proceso, p.Linea, p.Maquina, p.Sec,
            p.ItemCode, p.Descripcion,
            p.Fecha_Inicio, p.Hora_Inicio,
            p.Fecha_Fin,   p.Hora_Fin,
            p.Planificado, p.Cap_Diaria, p.Duracion_Horas,
            COALESCE(a.acum, 0) AS Acumulado,
            p.Planificado - COALESCE(a.acum, 0) AS Pendiente
        FROM ordenes_plan p
        LEFT JOIN (
            SELECT BELNR_ID, SUM(Cantidad_Real) AS acum
            FROM avance_real GROUP BY BELNR_ID
        ) a ON p.BELNR_ID = a.BELNR_ID
        ORDER BY p.Proceso, p.Linea, p.Maquina, p.Sec
    """
    df = pd.read_sql(query, con)
    con.close()

    df["Pct_Acum"] = (df["Acumulado"] / df["Planificado"] * 100).round(1).fillna(0)
    df["Inicio"]   = df["Fecha_Inicio"].str[5:].str.replace("-","/") + " " + df["Hora_Inicio"]
    df["Fin"]      = df["Fecha_Fin"].str[5:].str.replace("-","/")    + " " + df["Hora_Fin"]

    def estado(pct, pendiente):
        if pendiente <= 0:      return "Completado"
        if pct >= 75:           return "En curso"
        if pct > 0:             return "Parcial"
        return "Pendiente"

    df["Estado"] = df.apply(lambda r: estado(r["Pct_Acum"], r["Pendiente"]), axis=1)
    return df


def get_avance_campana(db_path: str, fecha_ref: date = None) -> pd.DataFrame:
    """
    Avance de campaña agrupado por Línea → Máquina.
    Solo máquinas con OPs activas en fecha_ref.
    Calcula desde el inicio del plan (primera OP de la máquina).
    """
    if fecha_ref is None:
        fecha_ref = date.today()

    con = sqlite3.connect(db_path)
    df_plan = pd.read_sql("SELECT * FROM ordenes_plan", con)
    df_av   = pd.read_sql(
        "SELECT BELNR_ID, SUM(Cantidad_Real) as acum FROM avance_real "
        "WHERE Fecha <= ? GROUP BY BELNR_ID",
        con, params=[str(fecha_ref if fecha_ref else date.today())])
    con.close()

    if df_plan.empty:
        return pd.DataFrame()

    # Merge acumulado real
    df_plan = df_plan.merge(df_av, on="BELNR_ID", how="left")
    df_plan["acum"] = df_plan["acum"].fillna(0)

    # OPs activas hoy
    fecha_str = str(fecha_ref)
    activas_hoy = set(df_plan[
        (df_plan["Fecha_Inicio"] <= fecha_str) &
        (df_plan["Fecha_Fin"]    >= fecha_str)
    ]["Maquina"].unique())

    if not activas_hoy:
        return pd.DataFrame()

    resultado = []
    for maquina, grp in df_plan.groupby("Maquina"):
        if maquina not in activas_hoy:
            continue

        proceso_v = grp["Proceso"].iloc[0] if "Proceso" in grp.columns else ""
        lineas_v  = ", ".join(sorted(grp["Linea"].dropna().unique())) if "Linea" in grp.columns else ""

        planificado = float(grp["Planificado"].sum())
        acumulado   = float(grp["acum"].sum())
        cap_diaria  = float(grp["Cap_Diaria"].max())

        # Fecha inicio de la primera OP de esta máquina
        f_inicio = date.fromisoformat(grp["Fecha_Inicio"].min()[:10])

        # Días hábiles trabajados desde f_inicio hasta fecha_ref
        dias_trans = 0
        d = f_inicio
        while d <= fecha_ref:
            if es_habil(d):
                dias_trans += 1
            d += timedelta(days=1)

        # Debería ir = Cap_Diaria × días hábiles trabajados (sin exceder planificado)
        deberia_ir  = min(cap_diaria * dias_trans, planificado)
        diferencia  = acumulado - deberia_ir
        dias_atraso = round(-diferencia / cap_diaria, 1) if diferencia < 0 and cap_diaria > 0 else 0
        pct_acum    = round(acumulado / planificado * 100, 1) if planificado > 0 else 0
        pct_deberia = round(deberia_ir / planificado * 100, 1) if planificado > 0 else 0

        if diferencia >= 0:     estado = "Al dia"
        elif dias_atraso <= 1:  estado = "Leve retraso"
        elif dias_atraso <= 3:  estado = "Retraso"
        else:                   estado = "Critico"

        resultado.append({
            "Proceso":            proceso_v,
            "Maquina":            maquina,
            "Linea":              lineas_v,
            "Planificado":        round(planificado),
            "Deberia_Ir":         round(deberia_ir),
            "Acumulado":          round(acumulado),
            "Diferencia":         round(diferencia),
            "Pct_Acum":           pct_acum,
            "Pct_Deberia":        pct_deberia,
            "Dias_Atraso":        dias_atraso,
            "Cap_Diaria":         round(cap_diaria),
            "Dias_Transcurridos": dias_trans,
            "Estado":             estado,
        })

    df = pd.DataFrame(resultado)
    if df.empty:
        return df
    orden = {"Critico": 0, "Retraso": 1, "Leve retraso": 2, "Al dia": 3}
    df["_ord"] = df["Estado"].map(orden).fillna(4)
    return df.sort_values(["_ord","Proceso","Maquina"]).drop(columns=["_ord"]).reset_index(drop=True)


def get_avance_campana_bom(db_path: str, excel_path: str) -> pd.DataFrame:
    """
    Lee Bom_Campaña del Excel y cruza con avance_real por Componente (ItemCode).
    Filtra solo Proceso_Interno en: Inyección, Semiterminado, Extrusión, Masas y Tintas, Soplado.
    """
    PROCESOS_OK = ["Inyección","Semiterminado","Extrusión","Masas y Tintas","Soplado"]
    try:
        df_bom = pd.read_excel(excel_path, sheet_name="BOM_Campaña")
        df_bom.columns = df_bom.columns.str.strip()
    except Exception as e:
        print(f"  Warning: Bom_Campaña no cargada: {e}")
        return pd.DataFrame()

    # Filtrar procesos válidos
    if "Proceso_Interno" in df_bom.columns:
        df_bom = df_bom[df_bom["Proceso_Interno"].isin(PROCESOS_OK)].reset_index(drop=True)

    if df_bom.empty:
        return pd.DataFrame()

    # Estandarizar nombre columna Componente
    for alias in ["Componente","Código Semi","Codigo_Semi","Codigo_Comp"]:
        if alias in df_bom.columns:
            df_bom = df_bom.rename(columns={alias: "Componente"})
            break

    df_bom["Componente"] = df_bom["Componente"].astype(str).str.strip()

    # Avance real acumulado por ItemCode desde avance_real
    con = sqlite3.connect(db_path)
    try:
        df_av = pd.read_sql(
            "SELECT ItemCode, SUM(Cantidad_Real) as Avance FROM avance_real GROUP BY ItemCode",
            con)
        df_av["ItemCode"] = df_av["ItemCode"].astype(str).str.strip()
    except Exception:
        df_av = pd.DataFrame(columns=["ItemCode","Avance"])
    con.close()

    # Estandarizar nombre columna cantidad
    cant_col = "Cantidad Total Requerida"
    if cant_col not in df_bom.columns:
        for c in df_bom.columns:
            if "cantidad" in c.lower() and "requerida" in c.lower():
                df_bom = df_bom.rename(columns={c: cant_col})
                break
    df_bom[cant_col] = pd.to_numeric(df_bom.get(cant_col, 0), errors="coerce").fillna(0)

    # Agrupar por Componente sumando Cantidad Total Requerida
    # Mantener columnas descriptivas del primer registro de cada componente
    desc_cols = ["Descripción Componente","Tipo_Material","Proceso_Interno","Líneas","Subcomponente"]
    desc_cols = [c for c in desc_cols if c in df_bom.columns]

    grp_agg = {cant_col: "sum"}
    for c in desc_cols:
        grp_agg[c] = "first"

    df_bom = df_bom.groupby("Componente", as_index=False).agg(grp_agg)

    # Merge avance real
    df_bom = df_bom.merge(df_av, left_on="Componente", right_on="ItemCode", how="left")
    df_bom["Avance"] = df_bom["Avance"].fillna(0)

    # Calcular % avance
    df_bom["Pct_Avance"] = (df_bom["Avance"] / df_bom[cant_col] * 100).round(1).clip(upper=100).fillna(0)

    return df_bom
