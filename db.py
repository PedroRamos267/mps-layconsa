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
            Horas_Prog        INTEGER,
            Fecha_Inicio      TEXT,
            Hora_Inicio       TEXT,
            Fecha_Fin         TEXT,
            Hora_Fin          TEXT,
            Velocidad         REAL,
            Cap_Diaria        REAL,
            Planificado       REAL,
            Duracion_Horas    REAL
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

def cargar_plan_desde_excel(db_path: str, excel_path: str):
    try:
        df = pd.read_excel(excel_path, sheet_name="Ordenes",
                           dtype={"BELNR_ID": int, "ItemCode": str})
    except Exception:
        df = pd.read_excel(excel_path, sheet_name=0,
                           dtype={"BELNR_ID": int, "ItemCode": str})

    df.columns = df.columns.str.strip()
    df = df.rename(columns={"Descripción del artículo": "Descripcion"})

    # Limpiar filas vacías
    df = df.dropna(subset=["BELNR_ID", "Sec", "Cantidad Base", "Planificado"])
    df["Sec"] = df["Sec"].astype(int)

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
    df["Fecha_Inicio_Base"] = pd.to_datetime(df["Fecha Inicio"], dayfirst=True).dt.normalize()

    registros = []
    for maquina, grupo in df.groupby("Maquina", sort=False):
        grupo = grupo.sort_values("Sec").reset_index(drop=True)
        horas_prog = int(grupo.loc[0, "Horas_Prog"])

        primera_fecha = grupo.loc[0, "Fecha_Inicio_Base"]
        cursor = datetime(primera_fecha.year, primera_fecha.month,
                          primera_fecha.day, HORA_INICIO)
        cursor = siguiente_momento_habil(cursor, horas_prog)

        for _, row in grupo.iterrows():
            hp = int(row["Horas_Prog"])
            velocidad      = row["Cantidad_Base"] / row["Horas"]   # und/hora
            cap_diaria     = velocidad * hp                         # und/día
            duracion_horas = row["Planificado"] / velocidad         # horas totales

            dt_inicio = cursor
            dt_fin    = sumar_horas_habiles(cursor, duracion_horas, hp)
            cursor    = dt_fin

            registros.append({
                "BELNR_ID":       int(row["BELNR_ID"]),
                "ItemCode":       row["ItemCode"],
                "Descripcion":    row["Descripcion"],
                "Sec":            int(row["Sec"]),
                "Maquina":        maquina,
                "Proceso":        row["Proceso"],
                "Linea":          row["Linea"],
                "Horas_Prog":     hp,
                "Fecha_Inicio":   dt_inicio.strftime("%Y-%m-%d"),
                "Hora_Inicio":    dt_inicio.strftime("%H:%M"),
                "Fecha_Fin":      dt_fin.strftime("%Y-%m-%d"),
                "Hora_Fin":       dt_fin.strftime("%H:%M"),
                "Velocidad":      round(velocidad, 4),
                "Cap_Diaria":     round(cap_diaria, 0),
                "Planificado":    row["Planificado"],
                "Duracion_Horas": round(duracion_horas, 2),
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
        WHERE p.Fecha_Inicio <= :fecha
          AND (
            p.Fecha_Fin >= :fecha
            OR (
              -- OP con avance real pero aún no completada (real < planificado)
              p.Fecha_Fin < :fecha
              AND COALESCE(acum.total_real, 0) < p.Planificado
              AND COALESCE(acum.total_real, 0) > 0
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

    # BD_Spec — Tipo_Material2 por componente
    try:
        spec = pd.read_excel(excel_path, sheet_name="BD_Spec",
                             usecols=["Código_Material","Tipo_Material2"])
        spec.columns = ["ItemCode","Tipo_Material2"]
        spec["ItemCode"] = spec["ItemCode"].astype(str).str.strip()
        spec["Tipo_Material2"] = spec["Tipo_Material2"].fillna("Sin clasificar").str.strip()
        con = sqlite3.connect(db_path)
        spec.to_sql("bd_spec", con, if_exists="replace", index=False)
        con.close()
        print(f"  → {len(spec)} registros de BD_Spec cargados")
    except Exception as e:
        print(f"  ⚠️  BD_Spec no cargado: {e}")


def _explotar_bom_op(bom_df: pd.DataFrame, codigo: str, cantidad_bruta: float,
                     cantidad_neta: float, stk_map: dict,
                     visitados: set, nivel: int = 1) -> list:
    """
    Explosión recursiva descendente del BOM.
    - cantidad_bruta: req. bruto del nivel (para mostrar)
    - cantidad_neta:  req. neto = bruto - stock (para explotar siguiente nivel)
    - Si un hijo empieza con 231 → semiterminado → explotar su BOM con req. neto
    - Si NO empieza con 231 → materia prima → stop
    """
    if nivel > 10 or cantidad_neta <= 0:
        return []

    hijos = bom_df[bom_df["Codigo_Padre"] == codigo]
    if hijos.empty:
        return []

    resultado = []
    for _, row in hijos.iterrows():
        comp     = str(row["Codigo_Comp"]).strip()
        req_b    = cantidad_bruta * row["Factor"]   # bruto basado en cantidad bruta padre
        req_n    = cantidad_neta  * row["Factor"]   # neto basado en cantidad neta padre
        es_semi  = comp.startswith("231")
        stock_c  = stk_map.get(comp, 0)
        req_neto_c = max(req_n - stock_c, 0)

        resultado.append({
            "Codigo_Comp":  comp,
            "Desc_Comp":    str(row["Desc_Comp"]),
            "Req_Bruto":    req_b,
            "Req_Neto":     req_neto_c,
            "Stock_Comp":   stock_c,
            "Nivel":        nivel,
            "Es_Semi":      es_semi,
        })

        # Si es semiterminado → explotar con req_neto (lo que falta fabricar)
        if es_semi and comp not in visitados and req_neto_c > 0:
            visitados.add(comp)
            sub = _explotar_bom_op(bom_df, comp, req_b, req_neto_c,
                                   stk_map, visitados, nivel + 1)
            resultado.extend(sub)

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

    # Explotar BOM desde cada OP — nivel siguiente usa req. neto
    todos = []
    for _, op in df_plan.iterrows():
        item     = str(op["ItemCode"]).strip()
        plan_dia = op["Plan_Dia"]
        stock_op = stk_map.get(item, 0)
        neto_op  = max(plan_dia - stock_op, 0)
        visitados = {item}
        reqs = _explotar_bom_op(bom, item, plan_dia, neto_op, stk_map, visitados)
        todos.extend(reqs)

    if not todos:
        return pd.DataFrame()

    req_df = pd.DataFrame(todos)

    # Agrupar por componente
    req = (req_df.groupby(["Codigo_Comp","Desc_Comp","Es_Semi"], as_index=False)
           .agg(Req_Bruto=("Req_Bruto","sum"),
                Stock_Comp=("Stock_Comp","first"),
                Nivel=("Nivel","min")))

    # Stock desglosado por almacén
    for s, col in [(s002,"Stock_ALMA002"),(s089,"Stock_ALMA089"),(sprod,"Stock_Prod")]:
        req = req.merge(s, left_on="Codigo_Comp", right_on="ItemCode", how="left").drop(
            columns=["ItemCode"], errors="ignore")
        req[col] = req[col].fillna(0)

    req["Stock_Total"] = req["Stock_ALMA002"] + req["Stock_ALMA089"] + req["Stock_Prod"]
    req["Req_Neto"]    = (req["Req_Bruto"] - req["Stock_Total"]).clip(lower=0)
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

    for dia in dias:
        consumo_dia = {}
        # Solo OPs activas ese día según fechas
        ops_activas = ordenes[
            (ordenes["Fecha_Inicio"] <= str(dia)) &
            (ordenes["Fecha_Fin"]    >= str(dia))
        ]
        for _, op in ops_activas.iterrows():
            f_inicio = date.fromisoformat(str(op["Fecha_Inicio"])[:10])
            cap      = float(op["Cap_Diaria"])
            planif   = float(op["Planificado"])
            item     = str(op["ItemCode"]).strip()

            # Días hábiles previos al día actual desde f_inicio
            dias_prev = 0
            d = f_inicio
            while d < dia:
                if es_habil(d):
                    dias_prev += 1
                d += _td2(days=1)

            ya_planificado = min(cap * dias_prev, planif)
            plan_dia_teo   = min(cap, max(planif - ya_planificado, 0))

            if plan_dia_teo <= 0:
                continue

            # Proyección: explotar BOM con req bruto para ver consumo total
            stk_map_proy = {}  # vacío = sin descontar stock en proyección
            visitados = {item}
            reqs = _explotar_bom_op(bom, item, plan_dia_teo, plan_dia_teo,
                                    stk_map_proy, visitados)
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

        row = {
            "Tipo_Material2":      tipo_map.get(comp, "Sin clasificar"),
            "Codigo_Comp":         comp,
            "Desc_Comp":           desc_map.get(comp, ""),
            "Stock_Inicial":       round(stock_ini, 0),
            "Consumo_Total":       round(consumo_total, 0),
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

    # % por proceso
    proc_list = []
    for proc, v in detalle_proceso.items():
        pct = round(v["real"] / v["plan"] * 100, 1) if v["plan"] > 0 else 0
        proc_list.append({"proceso": proc, "plan": round(v["plan"]),
                           "real": round(v["real"]), "pct": pct})
    proc_list.sort(key=lambda x: x["pct"])

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
