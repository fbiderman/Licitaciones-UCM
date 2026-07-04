#!/usr/bin/env python3
"""Tubería de actualización del panel de Traslados (Mercado Público).

Subcomandos:
  seed     Carga el histórico desde el Excel a la base local (una sola vez).
  resolve  Resuelve el CodigoProveedor de cada RUT vía API (se cachea en providers.csv).
  update   Trae OC nuevas desde la última fecha almacenada hasta hoy (incremental).
  build    Regenera dashboard.html a partir de la base local.
  probe    Vuelca el JSON crudo de una OC para verificar el mapeo de campos.
  run      resolve -> update -> build (lo que se agenda a diario).

Requiere la variable de entorno MP_TICKET (excepto para seed/build).
"""
from __future__ import annotations
import argparse, codecs, csv, datetime as dt, gzip, io, json, os, re, shutil, sqlite3, sys, tempfile, unicodedata, urllib.request, zipfile
csv.field_size_limit(10 * 1024 * 1024)

HERE = os.path.dirname(os.path.abspath(__file__))
DB = os.path.join(HERE, "mp_traslados.db")
PROVIDERS = os.path.join(HERE, "providers.csv")
TEMPLATE = os.path.join(HERE, "dashboard_template.html")
OUT_HTML = os.path.join(HERE, "dashboard.html")

REGIONES_CHILE = {  # fallback región -> se completa desde el Excel al sembrar
}

SCHEMA = """
CREATE TABLE IF NOT EXISTS oc (
  codigo TEXT PRIMARY KEY, nombre TEXT, comprador TEXT, rut_comprador TEXT,
  fecha TEXT, proveedor TEXT, rut_proveedor TEXT, estado TEXT,
  moneda TEXT, conversion_rate REAL, monto_bruto REAL, tipo_orden TEXT,
  anio INTEGER, mes INTEGER, region TEXT
);
CREATE TABLE IF NOT EXISTS region_map (comprador TEXT PRIMARY KEY, region TEXT);
CREATE TABLE IF NOT EXISTS meta (clave TEXT PRIMARY KEY, valor TEXT);
CREATE TABLE IF NOT EXISTS fx (fecha TEXT, indicador TEXT, valor REAL, PRIMARY KEY(fecha,indicador));
CREATE TABLE IF NOT EXISTS candidates (
  rut TEXT PRIMARY KEY, nombre TEXT, oc INTEGER, monto REAL, ultima TEXT, primera TEXT
);
CREATE TABLE IF NOT EXISTS licitacion (
  codigo TEXT PRIMARY KEY, nombre TEXT, comprador TEXT, estado TEXT, fecha TEXT,
  tipo TEXT, moneda TEXT, conversion_rate REAL, monto_estimado REAL,
  anio INTEGER, mes INTEGER, region TEXT, adjudicada INTEGER, origen TEXT DEFAULT 'proveedor'
);
"""


def conn():
    c = sqlite3.connect(DB)
    c.executescript(SCHEMA)
    # migración: agrega 'origen' si la tabla es anterior
    cols = [r[1] for r in c.execute("PRAGMA table_info(licitacion)")]
    if "origen" not in cols:
        c.execute("ALTER TABLE licitacion ADD COLUMN origen TEXT DEFAULT 'proveedor'")
        c.commit()
    return c


def _norm(s: str) -> str:
    s = unicodedata.normalize("NFKD", str(s)).encode("ascii", "ignore").decode()
    return " ".join(s.upper().split())


# ------------------------------------------------------------------ SEED
def cmd_seed(args):
    import pandas as pd
    xls = args.excel
    c = conn()
    n0 = c.execute("SELECT COUNT(*) FROM oc").fetchone()[0]
    if n0 and not args.force:
        print(f"La base ya tiene {n0} OC. Usa --force para resembrar.")
        return
    if args.force:
        c.execute("DELETE FROM oc")

    df = pd.read_excel(xls, sheet_name="Datos")
    df["conversion_rate"] = df["ConversionRate"].fillna(1.0)
    rows = []
    for _, r in df.iterrows():
        f = r["Fecha"]
        anio = int(r["Año"]); mes = int(r["Mes"])
        rows.append((str(r["Codigo de OC"]).strip(), str(r["Nombre de la Orden de Compra"]),
                     str(r["Comprador"]).strip(), "", str(f), str(r["Proveedor"]).strip(), "",
                     str(r["Estado"]).strip(), str(r["MonedaOC"]).strip(),
                     float(r["conversion_rate"]), float(r["MontoOC_BRUTO"]),
                     str(r["TipoOrden"]).strip(), anio, mes, str(r["Region"]).strip()))
    c.executemany("INSERT OR REPLACE INTO oc VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", rows)

    # mapa comprador -> región (hoja Regiones)
    try:
        raw = pd.read_excel(xls, sheet_name="Regiones", header=None)
        hdr = raw.index[raw.apply(lambda x: x.astype(str).str.contains("Comprador").any(), axis=1)][0]
        cols = raw.iloc[hdr].tolist()
        ic, ir = cols.index("Comprador"), cols.index("Region")
        reg = raw.iloc[hdr + 1:, [ic, ir]].dropna()
        for _, r in reg.iterrows():
            c.execute("INSERT OR REPLACE INTO region_map VALUES (?,?)",
                      (_norm(r.iloc[0]), str(r.iloc[1]).strip()))
    except Exception as e:  # noqa: BLE001
        print("Aviso: no se pudo leer la hoja Regiones:", e)
    # también deriva el mapa desde los propios datos
    for comp, region in c.execute("SELECT comprador, region FROM oc GROUP BY comprador").fetchall():
        c.execute("INSERT OR IGNORE INTO region_map VALUES (?,?)", (_norm(comp), region))

    c.execute("INSERT OR REPLACE INTO meta VALUES ('snapshot', ?)", (args.snapshot,))
    c.commit()
    n = c.execute("SELECT COUNT(*) FROM oc").fetchone()[0]
    mx = c.execute("SELECT MAX(fecha) FROM oc").fetchone()[0]
    print(f"Sembradas {n} OC. Última fecha en base: {mx}")


# ------------------------------------------------------------------ RESOLVE
def cmd_resolve(args):
    from mp_api import MPClient
    cli = MPClient()
    filas = list(csv.DictReader(open(PROVIDERS, encoding="utf-8")))
    cambios = 0

    def guardar():
        with open(PROVIDERS, "w", newline="", encoding="utf-8") as f:
            w = csv.DictWriter(f, fieldnames=["alias", "rut", "codigo_proveedor"])
            w.writeheader(); w.writerows(filas)

    for row in filas:
        if row.get("codigo_proveedor"):
            continue
        try:
            emp = cli.buscar_proveedor(row["rut"])
        except Exception as e:  # noqa: BLE001  — un 500 puntual no debe abortar todo
            print(f"  {row['rut']:>12}  ->  ERROR ({row['alias']}): {e}")
            continue
        cod = (emp or {}).get("CodigoEmpresa") or (emp or {}).get("codigoEmpresa")
        if cod:
            row["codigo_proveedor"] = str(cod); cambios += 1
            print(f"  {row['rut']:>12}  ->  {cod}  ({row['alias']})")
            if cambios % 10 == 0:
                guardar()  # persiste avance por si algo falla más adelante
        else:
            print(f"  {row['rut']:>12}  ->  SIN RESULTADO ({row['alias']})")
    guardar()
    print(f"Resueltos {cambios} códigos nuevos. Total con código: "
          f"{sum(1 for r in filas if r['codigo_proveedor'])}/{len(filas)}")


# ------------------------------------------------------------------ UPDATE
def _fx_rate(c, moneda: str, fecha_iso: str) -> float:
    """CLP por unidad de moneda en la fecha. CLP=1. Cachea en la tabla fx."""
    from mp_api import indicador_para_moneda, fetch_indicador
    ind = indicador_para_moneda(moneda)
    if not ind:
        return 1.0
    dkey = str(fecha_iso)[:10]
    row = c.execute("SELECT valor FROM fx WHERE fecha=? AND indicador=?", (dkey, ind)).fetchone()
    if row:
        return float(row[0])
    val = fetch_indicador(ind, dkey)
    if val:
        c.execute("INSERT OR REPLACE INTO fx VALUES (?,?,?)", (dkey, ind, val))
        return float(val)
    print(f"  ! sin tipo de cambio {ind} para {dkey}; se usa 1.0")
    return 1.0


def _region_de(c, comprador: str) -> str:
    r = c.execute("SELECT region FROM region_map WHERE comprador=?", (_norm(comprador),)).fetchone()
    return r[0] if r else "Sin Region"


def cmd_update(args):
    from mp_api import MPClient, extraer_campos, MPQuotaError
    cli = MPClient()
    c = conn()
    provs = [r for r in csv.DictReader(open(PROVIDERS, encoding="utf-8")) if r.get("codigo_proveedor")]
    if not provs:
        print("No hay proveedores con código. Ejecuta primero: resolve"); return

    last = c.execute("SELECT MAX(fecha) FROM oc").fetchone()[0]
    desde = dt.date.fromisoformat(str(last)[:10]) + dt.timedelta(days=1) if last else \
        dt.date.today() - dt.timedelta(days=args.dias)
    if args.desde:
        desde = dt.date.fromisoformat(args.desde)
    hasta = dt.date.fromisoformat(args.hasta) if args.hasta else dt.date.today()
    tramo = getattr(args, "tramo", 0)
    if tramo and (hasta - desde).days > tramo:
        hasta = desde + dt.timedelta(days=tramo)  # procesa por tramos, guardando al final de cada uno
    presupuesto = getattr(args, "max_llamadas", 0) or 9000   # tope de llamadas por corrida (cuota diaria = 10.000)
    print(f"Actualizando {desde} -> {hasta} para {len(provs)} proveedores… (tope {presupuesto} llamadas)", flush=True)

    nuevas = 0; ok = 0; err = 0; llamadas = 0; corte = None
    dia = desde
    while dia <= hasta and corte is None:
        fstr = dia.strftime("%d%m%Y")
        d_ok = 0; d_err = 0
        for p in provs:
            if llamadas >= presupuesto:
                corte = "presupuesto"; break
            try:
                listado = cli.oc_por_proveedor_dia(p["codigo_proveedor"], fstr); llamadas += 1
                d_ok += 1
            except MPQuotaError as e:
                corte = f"CUOTA DIARIA AGOTADA (Codigo 203): {e}"; break
            except Exception:  # noqa: BLE001  (fallo puntual de un proveedor)
                llamadas += 1; d_err += 1; continue
            for item in listado:
                cod = item.get("Codigo") or item.get("codigo")
                if not cod:
                    continue
                if c.execute("SELECT 1 FROM oc WHERE codigo=?", (cod,)).fetchone():
                    continue
                if llamadas >= presupuesto:
                    corte = "presupuesto"; break
                try:
                    det = cli.oc_detalle(cod); llamadas += 1
                except MPQuotaError as e:
                    corte = f"CUOTA DIARIA AGOTADA (Codigo 203): {e}"; break
                except Exception:  # noqa: BLE001
                    llamadas += 1; continue
                if not det:
                    continue
                f = extraer_campos(det)
                fecha = f["fecha"] or dia.isoformat()
                try:
                    d = dt.datetime.fromisoformat(fecha.replace("Z", "").split(".")[0])
                    anio, mes = d.year, d.month
                except Exception:  # noqa: BLE001
                    anio, mes = dia.year, dia.month
                cr = _fx_rate(c, f["moneda"], fecha)  # CLP=1; UF/UTM por fecha
                region = _region_de(c, f["comprador"])
                if region == "Sin Region":
                    region = _region_corta(f.get("region_texto")) or "Sin Region"
                c.execute("INSERT OR REPLACE INTO oc VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
                          (cod, f["nombre"], f["comprador"], f["rut_comprador"], fecha,
                           f["proveedor"] or p["alias"], f["rut_proveedor"] or p["rut"],
                           f["estado"], f["moneda"], cr, f["monto_bruto"], f["tipo_orden"],
                           anio, mes, region))
                nuevas += 1
        ok += d_ok; err += d_err
        c.commit()
        print(f"  {dia}  ok={d_ok} err={d_err}  (+{nuevas} OC nuevas · {llamadas} llamadas)", flush=True)
        if corte:
            break
        dia += dt.timedelta(days=1)
    c.execute("INSERT OR REPLACE INTO meta VALUES ('last_update', ?)",
              (dt.datetime.now().isoformat(timespec="seconds"),))
    c.commit()
    if corte and corte.startswith("CUOTA"):
        print(f"  ! {corte}", flush=True)
        print(f"  ! Detenido en {dia}. La cuota se reinicia cada día; la próxima corrida retoma desde aquí.", flush=True)
    elif corte == "presupuesto":
        print(f"  ! Alcanzado el tope de {presupuesto} llamadas en {dia}. La próxima corrida continúa.", flush=True)
    print(f"Listo. {nuevas} OC nuevas · {llamadas} llamadas · ok={ok} err={err}.", flush=True)


# ------------------------------------------------------------------ CANDIDATES
def cmd_candidates(args):
    """Escanea la lista diaria nacional de OC, filtra por nombre de 'traslados' y
    acumula proveedores que NO están en la nómina, para sugerirlos en el panel."""
    from mp_api import MPClient, extraer_campos, es_traslado
    cli = MPClient()
    c = conn()
    roster = {r["rut"].replace(".", "").upper()
              for r in csv.DictReader(open(PROVIDERS, encoding="utf-8")) if r.get("rut")}

    hasta = dt.date.fromisoformat(args.hasta) if args.hasta else dt.date.today()
    desde = dt.date.fromisoformat(args.desde) if args.desde else hasta - dt.timedelta(days=args.dias)
    print(f"Buscando proveedores de traslados fuera de la nómina {desde} -> {hasta}…")

    vistos = 0
    dia = desde
    while dia <= hasta:
        fstr = dia.strftime("%d%m%Y")
        try:
            listado = cli.oc_del_dia(fstr, estado="todos")
        except Exception as e:  # noqa: BLE001
            print(f"  ! {fstr}: {e}"); dia += dt.timedelta(days=1); continue
        for item in listado:
            nombre = item.get("Nombre") or item.get("nombre") or ""
            if not es_traslado(nombre):
                continue
            cod = item.get("Codigo") or item.get("codigo")
            try:
                det = cli.oc_detalle(cod) if cod else None
            except Exception as e:  # noqa: BLE001
                print(f"  ! detalle {cod}: {e}"); continue
            if not det:
                continue
            f = extraer_campos(det)
            rut = (f["rut_proveedor"] or "").replace(".", "").upper()
            if not rut or rut in roster:
                continue
            monto = f["monto_bruto"] * _fx_rate(c, f["moneda"], f["fecha"] or dia.isoformat())
            prev = c.execute("SELECT oc,monto,primera FROM candidates WHERE rut=?", (rut,)).fetchone()
            if prev:
                c.execute("UPDATE candidates SET nombre=?,oc=oc+1,monto=monto+?,ultima=? WHERE rut=?",
                          (f["proveedor"], monto, (f["fecha"] or dia.isoformat())[:10], rut))
            else:
                c.execute("INSERT INTO candidates VALUES (?,?,?,?,?,?)",
                          (rut, f["proveedor"], 1, monto, (f["fecha"] or dia.isoformat())[:10],
                           (f["fecha"] or dia.isoformat())[:10]))
            vistos += 1
        c.commit()
        if dia.day == 1 or dia == hasta:
            print(f"  … {dia}")
        dia += dt.timedelta(days=1)
    n = c.execute("SELECT COUNT(*) FROM candidates").fetchone()[0]
    print(f"Listo. {vistos} OC de traslados agregadas a {n} proveedores candidatos.")


# ------------------------------------------------------------------ LICITACIONES
def cmd_licitaciones(args):
    """Trae licitaciones donde participaron los proveedores de la nómina (por proveedor y día)."""
    from mp_api import MPClient, extraer_licitacion
    cli = MPClient()
    c = conn()
    provs = [r for r in csv.DictReader(open(PROVIDERS, encoding="utf-8")) if r.get("codigo_proveedor")]
    if not provs:
        print("No hay proveedores con código. Ejecuta primero: resolve"); return
    hasta = dt.date.fromisoformat(args.hasta) if args.hasta else dt.date.today()
    desde = dt.date.fromisoformat(args.desde) if args.desde else hasta - dt.timedelta(days=args.dias)
    print(f"Licitaciones {desde} -> {hasta} para {len(provs)} proveedores…")
    nuevas = 0
    dia = desde
    while dia <= hasta:
        fstr = dia.strftime("%d%m%Y")
        for p in provs:
            try:
                listado = cli.licitaciones_por_proveedor_dia(p["codigo_proveedor"], fstr)
            except Exception as e:  # noqa: BLE001
                print(f"  ! {fstr} {p['alias']}: {e}"); continue
            for item in listado:
                cod = item.get("CodigoExterno") or item.get("Codigo") or item.get("codigo")
                if not cod or c.execute("SELECT 1 FROM licitacion WHERE codigo=?", (cod,)).fetchone():
                    continue
                try:
                    det = cli.licitacion_detalle(cod)
                except Exception as e:  # noqa: BLE001
                    print(f"  ! detalle lic {cod}: {e}"); continue
                if not det:
                    continue
                f = extraer_licitacion(det)
                fecha = f["fecha"] or dia.isoformat()
                try:
                    d = dt.datetime.fromisoformat(fecha.replace("Z", "").split(".")[0])
                    anio, mes = d.year, d.month
                except Exception:  # noqa: BLE001
                    anio, mes = dia.year, dia.month
                cr = _fx_rate(c, f["moneda"], fecha)
                c.execute("INSERT OR REPLACE INTO licitacion VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
                          (cod, f["nombre"], f["comprador"], f["estado"], fecha, f["tipo"],
                           f["moneda"], cr, f["monto_estimado"], anio, mes,
                           _region_de(c, f["comprador"]), f["adjudicada"], "proveedor"))
                nuevas += 1
        if dia.day == 1 or dia == hasta:
            c.commit(); print(f"  … {dia}")
        dia += dt.timedelta(days=1)
    c.commit()
    n = c.execute("SELECT COUNT(*) FROM licitacion").fetchone()[0]
    print(f"Listo. {nuevas} licitaciones nuevas ({n} en total).")


# ------------------------------------------------------------------ BUILD
def cmd_build(args):
    c = conn()
    rows = c.execute("""SELECT anio,mes,proveedor,comprador,region,estado,tipo_orden,
                        monto_bruto*COALESCE(conversion_rate,1) FROM oc""").fetchall()
    prov, comp, reg, est, tip = {}, {}, {}, {}, {}
    def idx(d, v):
        v = v or ""
        return d.setdefault(v, len(d))
    out = []
    total = 0
    for a, m, p, cp, rg, es, tp, monto in rows:
        monto = float(monto or 0); total += monto
        out.append([int(a), int(m), idx(prov, p), idx(comp, cp), idx(reg, rg),
                    idx(est, es), idx(tip, tp), round(monto)])
    inv = lambda d: [k for k, _ in sorted(d.items(), key=lambda kv: kv[1])]
    snap = (c.execute("SELECT valor FROM meta WHERE clave='snapshot'").fetchone() or
            [dt.date.today().isoformat()])[0]

    # nómina actual (para regenerar providers.csv desde el panel)
    roster = [[r["alias"], r["rut"]] for r in csv.DictReader(open(PROVIDERS, encoding="utf-8"))]
    roster_ruts = {r[1].replace(".", "").upper() for r in roster}

    # proveedores sugeridos (fuera de la nómina), ordenados por monto
    cand = []
    for rut, nombre, oc, monto, ultima, primera in c.execute(
            "SELECT rut,nombre,oc,monto,ultima,primera FROM candidates ORDER BY monto DESC"):
        if rut.replace(".", "").upper() in roster_ruts:
            continue
        cand.append({"rut": rut, "nombre": nombre, "oc": oc,
                     "monto": round(monto or 0), "ultima": ultima, "primera": primera})

    # tipo de cambio aplicado (últimos valores) + conteo de OC en UF/UTM
    def last_fx(ind):
        r = c.execute("SELECT fecha,valor FROM fx WHERE indicador=? ORDER BY fecha DESC LIMIT 1",
                      (ind,)).fetchone()
        if not r:  # respaldo: último conversion_rate del histórico para esa moneda
            monedas = ("CLF", "UF") if ind == "uf" else ("UTM",)
            q = ("SELECT fecha,conversion_rate FROM oc WHERE UPPER(moneda) IN (%s) "
                 "AND conversion_rate>1 ORDER BY fecha DESC LIMIT 1" %
                 ",".join("?" * len(monedas)))
            r = c.execute(q, monedas).fetchone()
        return {"fecha": str(r[0])[:10], "valor": round(r[1])} if r else None
    n_uf = c.execute("SELECT COUNT(*) FROM oc WHERE UPPER(moneda) IN ('CLF','UF')").fetchone()[0]
    n_utm = c.execute("SELECT COUNT(*) FROM oc WHERE UPPER(moneda)='UTM'").fetchone()[0]
    fx = {"uf": last_fx("uf"), "utm": last_fx("utm"), "n_uf": n_uf, "n_utm": n_utm}
    uf_ref = last_fx("uf")  # UF de referencia (respaldo)

    # ---- UF de cierre por año (31-dic); año en curso usa la UF más reciente ----
    # Valores de respaldo (UF al 31-dic) por si no hay red al construir.
    _UF_EOY = {2016:26348,2017:26798,2018:27566,2019:28310,2020:29070,
               2021:30992,2022:35111,2023:36789,2024:38363,2025:40197,2026:40828}
    def _uf_en(fecha_iso):
        r = c.execute("SELECT valor FROM fx WHERE fecha=? AND indicador='uf'", (fecha_iso,)).fetchone()
        if r:
            return r[0]
        try:
            from mp_api import fetch_indicador
            v = fetch_indicador("uf", fecha_iso)
        except Exception:  # noqa: BLE001
            v = None
        if v:
            c.execute("INSERT OR REPLACE INTO fx VALUES (?,?,?)", (fecha_iso, "uf", v)); c.commit()
        return v
    cur_year = dt.date.today().year
    hoy_iso = dt.date.today().isoformat()
    uf_by_year = {}
    anios_oc = [r[0] for r in c.execute("SELECT DISTINCT anio FROM oc WHERE anio>0 ORDER BY anio")]
    for y in anios_oc:
        fecha = hoy_iso if y >= cur_year else f"{y}-12-31"
        val = _uf_en(fecha) or _UF_EOY.get(y) or (uf_ref or {}).get("valor") or 40000
        uf_by_year[str(y)] = round(val)

    updated = (c.execute("SELECT valor FROM meta WHERE clave='last_update'").fetchone() or
               [dt.datetime.now().isoformat(timespec="seconds")])[0]
    data_through = c.execute("SELECT MAX(fecha) FROM oc").fetchone()[0] or snap

    # ---- licitaciones (bloque compacto propio) ----
    lic_rows_db = c.execute("""SELECT anio,mes,comprador,region,estado,tipo,
        monto_estimado*COALESCE(conversion_rate,1), adjudicada FROM licitacion""").fetchall()
    lc, lr, le, lt = {}, {}, {}, {}
    lic_rows = []
    for a, mm, cp, rg, es, tp, monto, adj in lic_rows_db:
        lic_rows.append([int(a), int(mm), idx(lc, cp), idx(lr, rg), idx(le, es),
                         idx(lt, tp), round(float(monto or 0)), int(adj or 0)])
    lic_recent = [{"codigo": r[0], "nombre": r[1], "comp": r[2], "est": r[3],
                   "fecha": r[4], "monto": round(float(r[5] or 0))}
                  for r in c.execute("""SELECT codigo,nombre,comprador,estado,fecha,
                      monto_estimado*COALESCE(conversion_rate,1) FROM licitacion
                      ORDER BY fecha DESC LIMIT 40""")]
    lic = {"comp": inv(lc), "reg": inv(lr), "est": inv(le), "tip": inv(lt),
           "rows": lic_rows, "recent": lic_recent}

    data = {"meta": {"source": "Mercado Publico - Traslados", "snapshot": snap,
                     "updated": updated, "data_through": data_through,
                     "rows": len(out), "total_clp": int(total), "fx": fx, "uf_ref": uf_ref,
                     "uf_by_year": uf_by_year},
            "prov": inv(prov), "comp": inv(comp), "reg": inv(reg),
            "est": inv(est), "tip": inv(tip), "rows": out,
            "roster": roster, "candidates": cand, "lic": lic}
    payload = json.dumps(data, ensure_ascii=False, separators=(",", ":"))
    tpl = open(TEMPLATE, encoding="utf-8").read()
    html = tpl.replace("__DATA_JSON__", payload)
    open(OUT_HTML, "w", encoding="utf-8").write(html)
    # copia para GitHub Pages (se sirve como index.html del sitio)
    os.makedirs(os.path.join(HERE, "site"), exist_ok=True)
    open(os.path.join(HERE, "site", "index.html"), "w", encoding="utf-8").write(html)
    print(f"dashboard.html generado: {len(out)} OC, total MM$ {total/1e6:,.0f}")


# ------------------------------------------------------------------ PROBE
def cmd_probe(args):
    import mp_api
    from mp_api import MPClient
    cli = MPClient()

    # modo 1: detalle de una OC (comportamiento original)
    if args.codigo:
        det = cli.oc_detalle(args.codigo)
        print(json.dumps(det, ensure_ascii=False, indent=2)[:6000])
        print("\n--- Revisa que estas claves existan y ajusta mp_api.extraer_campos() si difieren ---")
        return

    # modo 2: diagnóstico de la consulta diaria por proveedor
    dia = dt.date.fromisoformat(args.dia) if args.dia else dt.date.today() - dt.timedelta(days=7)
    f = dia.strftime("%d%m%Y")
    cod = args.codigo_proveedor
    if not cod:
        provs = [r for r in csv.DictReader(open(PROVIDERS, encoding="utf-8")) if r.get("codigo_proveedor")]
        cod = provs[0]["codigo_proveedor"] if provs else None
    ep = f"{mp_api.BASE}/ordenesdecompra.json"
    print(f"== Diagnóstico API · día {dia} (fecha={f}) · proveedor {cod} ==\n", flush=True)

    d1 = cli._get(cli._url(ep, fecha=f, CodigoProveedor=str(cod)))
    l1 = d1.get("Listado") or []
    print(f"[A] fecha={f} & CodigoProveedor={cod}")
    print("    claves:", list(d1.keys()), "| Cantidad:", d1.get("Cantidad"), "| len(Listado):", len(l1))
    print("    RESPUESTA CRUDA:", json.dumps(d1, ensure_ascii=False)[:800], flush=True)

    d2 = cli._get(cli._url(ep, fecha=f))
    l2 = d2.get("Listado") or []
    print(f"\n[B] fecha={f}  (todas las OC del país ese día)")
    print("    Cantidad:", d2.get("Cantidad"), "| len(Listado):", len(l2))
    print("    RESPUESTA CRUDA:", json.dumps(d2, ensure_ascii=False)[:800])
    if l2:
        print("    campos de un registro:", list(l2[0].keys()), flush=True)

    d3 = cli._get(cli._url(ep, CodigoProveedor=str(cod)))
    l3 = d3.get("Listado") or []
    print(f"\n[C] CodigoProveedor={cod}  (sin fecha)")
    print("    Cantidad:", d3.get("Cantidad"), "| len(Listado):", len(l3))
    print("    RESPUESTA CRUDA:", json.dumps(d3, ensure_ascii=False)[:800], flush=True)

    print("\n--- Interpretación ---")
    if any(("Mensaje" in d and not (d.get("Listado"))) for d in (d1, d2, d3)):
        print("La API respondió con un MENSAJE de error (clave 'Mensaje'), no con datos.")
        print("Lee el texto de 'Mensaje' arriba: suele indicar ticket inválido, límite de")
        print("consultas superado, o fecha/endpoint fuera de rango.")
    elif not l1 and l2:
        print("La combinación fecha+CodigoProveedor NO trae datos, pero fecha sola sí.")
        print("=> Hay que consultar por día [B] y filtrar por proveedor, o usar [C] sin fecha.")
    elif not l1 and not l2:
        print("Ni por proveedor ni por día hay datos: revisa el ticket, la fecha o el endpoint.")
    else:
        print("La consulta [A] sí trae datos; el problema estaría en el parseo, no en la consulta.")


# ------------------------------------------------------------ IMPORT DATOS ABIERTOS
# Alias de columnas de los CSV de datos-abiertos.chilecompra.cl (Órdenes de Compra).
# Se comparan tras normalizar (mayúsculas, sin tildes, solo alfanumérico).
_DA_ALIASES = {
    "codigo":        ["Codigo", "CodigoOrdenCompra", "CodigoOC", "IdOrdenCompra", "OrdenDeCompra", "codigoOrdenCompra"],
    "nombre":        ["Nombre", "NombreOC", "NombreOrdenCompra"],
    "comprador":     ["OrganismoPublico", "NombreOrganismo", "Organismo", "UnidadCompra", "NombreUnidadCompra",
                      "UnidadDeCompra", "Institucion", "NombreInstitucion", "Comprador"],
    "rut_comprador": ["RutUnidadCompra", "RutOrganismo", "RutInstitucion", "RutComprador"],
    "fecha":         ["FechaEnvio", "FechaCreacion", "FechaOrdenCompra", "Fecha", "FechaAceptacion", "FechaEnvioOC"],
    "proveedor":     ["NombreProveedor", "RazonSocialProveedor", "RazonSocialSucursal", "Proveedor", "NombreSucursal"],
    "rut_proveedor": ["RutProveedor", "RutSucursal", "RutEmpresaProveedor", "RUTProveedor", "Rut"],
    "estado":        ["Estado", "EstadoOC", "NombreEstado", "CodigoEstado"],
    "monto":         ["MontoTotalOC_PesosChilenos", "MontoTotalPesos", "MontoTotal", "MontoBruto",
                      "MontoOC", "Monto", "MontoTotalBruto", "montoTotal"],
    "tipo":          ["TipoModalidad", "Modalidad", "Tipo", "TipoOrdenCompra", "TipoDeCompra", "TipoContrato"],
    "region":        ["RegionUnidadCompra", "Region", "RegionOrganismo"],
}


def _colnorm(s: str) -> str:
    s = unicodedata.normalize("NFKD", str(s)).encode("ascii", "ignore").decode()
    return re.sub(r"[^A-Z0-9]", "", s.upper())


def _rut_norm(s: str) -> str:
    return re.sub(r"[^0-9K]", "", str(s or "").upper())


def _mapear_columnas(header):
    """Devuelve {campo_interno: nombre_real_de_columna} según los alias conocidos."""
    norm2real = {_colnorm(h): h for h in header}
    mapa = {}
    for campo, alias in _DA_ALIASES.items():
        for a in alias:
            real = norm2real.get(_colnorm(a))
            if real is not None:
                mapa[campo] = real
                break
    return mapa


def _parse_monto(v):
    s = str(v or "").strip()
    if not s:
        return 0.0
    m = re.match(r"^\s*([0-9]+[.,]?[0-9]*)\s*[eE]\s*([+-]?[0-9]+)\s*$", s)  # "5,7e+07"
    if m:
        try:
            return float(m.group(1).replace(",", ".")) * (10 ** int(m.group(2)))
        except ValueError:
            return 0.0
    s = re.sub(r"[^0-9,.\-]", "", s)
    if "," in s and "." in s:          # 1.234.567,89  -> punto miles, coma decimal
        s = s.replace(".", "").replace(",", ".")
    elif "," in s:                      # 1234567,89    -> coma decimal
        s = s.replace(",", ".")
    else:                               # 1.234.567     -> punto como separador de miles
        if s.count(".") > 1 or re.match(r"^\d{1,3}(\.\d{3})+$", s):
            s = s.replace(".", "")
    try:
        return float(s)
    except ValueError:
        return 0.0


def _parse_fecha(v):
    s = str(v or "").strip()
    if not s:
        return None
    s = s.replace("T", " ").split(".")[0].split(" ")[0]
    for fmt in ("%Y-%m-%d", "%d-%m-%Y", "%d/%m/%Y", "%Y/%m/%d", "%m/%d/%Y"):
        try:
            return dt.datetime.strptime(s, fmt).date().isoformat()
        except ValueError:
            continue
    return None


# Regiones: del texto largo del archivo -> forma corta del histórico. Orden importa.
_REGION_KEYS = [
    ("aricayparinacota", "Arica"), ("arica", "Arica"),
    ("tarapaca", "Tarapaca"), ("antofagasta", "Antofagasta"), ("atacama", "Atacama"),
    ("coquimbo", "Coquimbo"), ("valparaiso", "Valparaiso"), ("metropolitana", "Metropolitana"),
    ("higgins", "O'Higgins"), ("maule", "Maule"), ("nuble", "Ñuble"),
    ("biobio", "Bio Bio"), ("araucania", "Araucania"),
    ("losrios", "Los Rios"), ("loslagos", "Los Lagos"),
    ("aysen", "Aysen"), ("magallanes", "Magallanes"), ("nacional", "Nacional"),
]


def _region_corta(s):
    t = unicodedata.normalize("NFKD", str(s or "")).encode("ascii", "ignore").decode().lower()
    t = re.sub(r"[^a-z]", "", t)
    for key, val in _REGION_KEYS:
        if key in t:
            return val
    return None


def _tipo_da(row):
    def g(k):
        return str(row.get(k, "") or "").strip().lower()
    if g("EsTratoDirecto") in ("si", "sí", "1", "true"):
        return "Trato Directo"
    if g("EsCompraAgil") in ("si", "sí", "1", "true"):
        return "Compra Agil"
    proc = unicodedata.normalize("NFKD", g("ProcedenciaOC")).encode("ascii", "ignore").decode()
    if "convenio marco" in proc:
        return "Convenio Marco"
    if "microcompra" in proc:
        return "Microcompra"
    if "licitacion" in proc:
        return "Licitacion"
    return "Sin Clasificacion"


def _preparar_csv(path):
    """Descomprime a disco si hace falta y devuelve (ruta_csv, [temporales_a_borrar])."""
    with open(path, "rb") as fh:
        magic = fh.read(6)
    low = path.lower()
    if magic[:2] == b"\x1f\x8b" or low.endswith(".gz"):
        out = tempfile.mkstemp(suffix=".csv")[1]
        with gzip.open(path, "rb") as g, open(out, "wb") as o:
            shutil.copyfileobj(g, o)
        return out, [out]
    if magic[:2] == b"PK" or low.endswith(".zip"):
        zf = zipfile.ZipFile(path)
        nombre = next((n for n in zf.namelist() if n.lower().endswith(".csv")), zf.namelist()[0])
        print(f"    (zip) miembro: {nombre}", flush=True)
        d = tempfile.mkdtemp()
        zf.extract(nombre, d); zf.close()
        return os.path.join(d, nombre), [os.path.join(d, nombre)]
    if magic[:2] == b"7z\xbc" or low.endswith(".7z"):
        try:
            import py7zr
        except ImportError:
            raise RuntimeError("El archivo es .7z; añade 'py7zr' a requirements.txt.")
        d = tempfile.mkdtemp()
        with py7zr.SevenZipFile(path, "r") as z:
            z.extractall(d)
        csvs = [os.path.join(r, f) for r, _, fs in os.walk(d) for f in fs if f.lower().endswith(".csv")]
        if not csvs:
            raise RuntimeError("El .7z no contiene ningún .csv")
        print(f"    (7z) archivo: {os.path.basename(csvs[0])}", flush=True)
        return csvs[0], [csvs[0]]
    return path, []


def _sniff_encoding(path):
    """Detecta UTF-8 vs Latin-1/Windows-1252 leyendo una muestra."""
    with open(path, "rb") as fh:
        sample = fh.read(200_000)
    try:
        codecs.getincrementaldecoder("utf-8")().decode(sample, False)  # tolera corte al final
        return "utf-8"
    except UnicodeDecodeError:
        return "cp1252"


def _origen_a_path(origen):
    """Si 'origen' es URL, la descarga a un temporal y devuelve el path; si es ruta, la devuelve."""
    if origen.lower().startswith(("http://", "https://")):
        print(f"  Descargando {origen} …", flush=True)
        fd, tmp = tempfile.mkstemp()
        os.close(fd)
        req = urllib.request.Request(origen, headers={"User-Agent": "traslados-dashboard/1.0"})
        with urllib.request.urlopen(req, timeout=600) as r, open(tmp, "wb") as out:
            while True:
                chunk = r.read(1 << 20)
                if not chunk:
                    break
                out.write(chunk)
        print(f"  Descargado ({os.path.getsize(tmp)/1e6:.1f} MB).", flush=True)
        return tmp
    return origen


def cmd_import_da(args):
    c = conn()
    ruts = {_rut_norm(r["rut"]): r["alias"] for r in csv.DictReader(open(PROVIDERS, encoding="utf-8")) if r.get("rut")}
    if not ruts:
        print("No hay RUT en providers.csv."); return
    print(f"Filtrando por {len(ruts)} RUT de proveedores.", flush=True)

    origenes = []
    for u in (args.url or []):
        u = str(u).strip()
        if u.lower().startswith(("http://", "https://")):
            u = u.split()[0]      # descarta texto pegado tras la URL (p.ej. "Descargar archivo")
        if u:
            origenes.append(u)
    if args.dir and os.path.isdir(args.dir):
        for f in sorted(os.listdir(args.dir)):
            if f.lower().endswith((".csv", ".gz", ".zip", ".7z")):
                origenes.append(os.path.join(args.dir, f))
    if not origenes:
        print("Sin fuentes. Usa --url <URL> (repetible) o --dir <carpeta>.")
        return

    total_ins = 0
    for origen in origenes:
        descargado = _origen_a_path(origen)
        csv_path, temporales = _preparar_csv(descargado)
        enc = _sniff_encoding(csv_path)
        print(f"    codificación detectada: {enc}", flush=True)
        fh = open(csv_path, encoding=enc, errors="replace")
        try:
            muestra = fh.readline()
            sep = ";" if muestra.count(";") >= muestra.count(",") else ","
            header = next(csv.reader([muestra], delimiter=sep))
            mapa = _mapear_columnas(header)
            print(f"\n  Fuente: {os.path.basename(str(origen))}")
            print(f"  Separador detectado: '{sep}' · {len(header)} columnas")
            print(f"  Mapeo de columnas: {mapa}", flush=True)
            if args.dry_run:
                print(f"  CABECERA COMPLETA: {header}", flush=True)
            faltan = [k for k in ("codigo", "rut_proveedor", "monto", "fecha") if k not in mapa]
            if faltan:
                print(f"  ! No pude ubicar columnas clave: {faltan}")
                print(f"  ! Cabecera real: {header}")
                print("  ! Ajusta _DA_ALIASES con estos nombres y reintenta. No se insertó nada de esta fuente.")
                continue

            reader = csv.DictReader(fh, fieldnames=header, delimiter=sep)
            leidas = 0; coincide = 0; ins = 0
            for row in reader:
                leidas += 1
                if args.limit and leidas > args.limit:
                    break
                rutp = _rut_norm(row.get(mapa["rut_proveedor"]))
                if rutp not in ruts:
                    continue
                coincide += 1
                if args.dry_run and coincide == 1:
                    print("\n  EJEMPLO (primera fila de un proveedor tuyo):", flush=True)
                    for k, v in row.items():
                        vs = str(v or "").strip()
                        if vs:
                            print(f"      {k} = {vs[:60]}", flush=True)
                cod = str(row.get(mapa["codigo"]) or "").strip()
                if not cod:
                    continue
                fecha = _parse_fecha(row.get(mapa["fecha"])) or ""
                anio = int(fecha[:4]) if fecha[:4].isdigit() else 0
                mes = int(fecha[5:7]) if fecha[5:7].isdigit() else 0
                comprador = str(row.get(mapa.get("comprador", ""), "") or "").strip()
                region = _region_de(c, comprador)
                if region == "Sin Region":
                    region = (_region_corta(row.get(mapa["region"])) if mapa.get("region") else None) or "Sin Region"
                if not args.dry_run:
                    c.execute("INSERT OR REPLACE INTO oc VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", (
                        cod,
                        str(row.get(mapa.get("nombre", ""), "") or "").strip(),
                        comprador,
                        str(row.get(mapa.get("rut_comprador", ""), "") or "").strip(),
                        fecha,
                        str(row.get(mapa.get("proveedor", ""), "") or ruts[rutp]).strip(),
                        rutp,
                        str(row.get(mapa.get("estado", ""), "") or "").strip(),
                        "CLP", 1.0, _parse_monto(row.get(mapa["monto"])),
                        _tipo_da(row),
                        anio, mes, region))
                    ins += 1
            if not args.dry_run:
                c.commit()
            total_ins += ins
            estado_ins = "(dry-run, 0 insertadas)" if args.dry_run else f"{ins} insertadas/actualizadas"
            print(f"  Leídas {leidas} filas · {coincide} de tus proveedores · {estado_ins}", flush=True)
        finally:
            fh.close()
            for t in temporales:
                d = os.path.dirname(t)
                if os.path.exists(t):
                    os.remove(t)
                if d.startswith(tempfile.gettempdir()) and os.path.isdir(d) and not os.listdir(d):
                    os.rmdir(d)
            if descargado != origen and os.path.exists(descargado):
                os.remove(descargado)

    if not args.dry_run:
        c.execute("INSERT OR REPLACE INTO meta VALUES ('last_update', ?)",
                  (dt.datetime.now().isoformat(timespec="seconds"),))
        c.commit()
        n = c.execute("SELECT COUNT(*) FROM oc").fetchone()[0]
        mx = c.execute("SELECT MAX(fecha) FROM oc").fetchone()[0]
        print(f"\nListo. {total_ins} filas incorporadas. Base ahora: {n} OC, máx fecha {mx}.", flush=True)
    else:
        print("\nModo prueba (dry-run): nada se guardó. Revisa el mapeo de columnas arriba.", flush=True)


def cmd_resolve_da(args):
    """Extrae el CodigoProveedor correcto desde archivos de Datos Abiertos y lo escribe en providers.csv."""
    from collections import Counter
    filas = list(csv.DictReader(open(PROVIDERS, encoding="utf-8")))
    rutset = {_rut_norm(p["rut"]) for p in filas if p.get("rut")}
    if not rutset:
        print("No hay RUT en providers.csv."); return

    origenes = []
    for u in (args.url or []):
        u = str(u).strip()
        if u.lower().startswith(("http://", "https://")):
            u = u.split()[0]
        if u:
            origenes.append(u)
    if args.dir and os.path.isdir(args.dir):
        for f in sorted(os.listdir(args.dir)):
            if f.lower().endswith((".csv", ".gz", ".zip", ".7z")):
                origenes.append(os.path.join(args.dir, f))
    if not origenes:
        print("Sin fuentes. Usa --url <URL> o --dir <carpeta>."); return

    encontrados = {}  # rut_norm -> Counter de CodigoProveedor
    for origen in origenes:
        try:
            descargado = _origen_a_path(origen)
        except Exception as e:  # noqa: BLE001  — un mes inexistente (404) no debe abortar todo
            print(f"  ! No se pudo descargar {origen}: {e}", flush=True)
            continue
        csv_path, temporales = _preparar_csv(descargado)
        enc = _sniff_encoding(csv_path)
        fh = open(csv_path, encoding=enc, errors="replace")
        try:
            muestra = fh.readline()
            sep = ";" if muestra.count(";") >= muestra.count(",") else ","
            header = next(csv.reader([muestra], delimiter=sep))
            mapa = _mapear_columnas(header)
            col_rut = mapa.get("rut_proveedor")
            norm2real = {_colnorm(h): h for h in header}
            col_cod = norm2real.get(_colnorm("CodigoProveedor"))
            if not col_rut or not col_cod:
                print(f"  ! No encontré columnas de RUT y/o CodigoProveedor en {os.path.basename(str(origen))}. Cabecera: {header}")
                continue
            reader = csv.DictReader(fh, fieldnames=header, delimiter=sep)
            for row in reader:
                rn = _rut_norm(row.get(col_rut))
                if rn in rutset:
                    cod = str(row.get(col_cod) or "").strip()
                    if cod:
                        encontrados.setdefault(rn, Counter())[cod] += 1
        finally:
            fh.close()
            for t in temporales:
                d = os.path.dirname(t)
                if os.path.exists(t):
                    os.remove(t)
                if d.startswith(tempfile.gettempdir()) and os.path.isdir(d) and not os.listdir(d):
                    os.rmdir(d)
            if descargado != origen and os.path.exists(descargado):
                os.remove(descargado)

    cambios = 0
    for p in filas:
        rn = _rut_norm(p["rut"])
        if rn in encontrados and encontrados[rn]:
            cod = encontrados[rn].most_common(1)[0][0]
            if str(p.get("codigo_proveedor", "")) != str(cod):
                print(f"  {p['rut']:>12}  {p.get('codigo_proveedor','') or '(vacío)':>10} -> {cod}  ({p['alias']})")
                p["codigo_proveedor"] = cod; cambios += 1
    with open(PROVIDERS, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=["alias", "rut", "codigo_proveedor"])
        w.writeheader(); w.writerows(filas)
    resueltos = sum(1 for r in filas if r.get("codigo_proveedor"))
    sin = [r["alias"] for r in filas if _rut_norm(r["rut"]) not in encontrados]
    print(f"\nListo. {cambios} códigos corregidos. Con código: {resueltos}/{len(filas)}.")
    if sin:
        print(f"Sin aparición en las fuentes ({len(sin)}): {', '.join(sin[:15])}{' …' if len(sin) > 15 else ''}")
        print("(Esos proveedores no tuvieron órdenes en el/los mes(es) usados; prueba con otro mes si te interesa resolverlos.)")


# ------------------------------------------------------------ IMPORT LICITACIONES (Datos Abiertos)
_LIC_ALIASES = {
    "codigo":    ["CodigoExterno", "Codigo", "CodigoLicitacion", "IdLicitacion"],
    "nombre":    ["Nombre", "NombreLicitacion", "Descripcion"],
    "comprador": ["OrganismoPublico", "NombreOrganismo", "Organismo", "UnidadCompra", "NombreUnidadCompra", "Comprador", "sector"],
    "fecha":     ["FechaPublicacion", "FechaCreacion", "FechaInicio", "Fecha", "FechaCierre"],
    "estado":    ["Estado", "EstadoLicitacion", "CodigoEstado", "NombreEstado"],
    "monto":     ["MontoEstimado", "MontoEstimadoLicitacion", "Monto", "MontoTotal", "MontoTotalEstimado"],
    "tipo":      ["Tipo", "TipoLicitacion", "TipoConvocatoria", "Modalidad"],
    "region":    ["RegionUnidad", "RegionUnidadCompra", "Region", "RegionOrganismo"],
    "adjudicada":["FechaAdjudicacion", "NumeroAdjudicaciones"],
}
# rubro/categoría del ítem, para detectar traslados
_LIC_RUBRO_COLS = ["CodigoProductoONU", "codigoProductoONU", "Nombre producto genrico",
                   "Nombre linea Adquisicion", "Rubro3", "Rubro2", "Rubro1",
                   "Nombre", "Descripcion", "Descripcion linea Adquisicion"]
# servicios de traslado de pacientes (el mercado de UCM)
_SERVICIO_ONU = {"92101902"}          # "Servicios de ambulancia"
_VEHICULO_ONU = {"25101703"}          # "Ambulancias" (vehículo)
# contexto médico que convierte un "traslado/transporte" en traslado de pacientes
_MED = ["paciente", "clinico", "asistido", "sanitario", "prehospital", "enfermo",
        "dializ", "hospital", "samu", "urgencia", "medic", "cardio", "neonatal", "hemodial"]
# la ambulancia aparece pero el objeto NO es el servicio de traslado
_VEH_WORDS = ["compra", "adquisicion", "mantencion", "mantenimiento", "reparacion", "renovacion", "repotenc"]
_OTRO_WORDS = ["seguro", "poliza", "aseo", "limpieza", "repuesto", "suministro", "insumo",
               "camilla", "pintura", "gps", "desabolla", "neumatico"]


def _txtnorm(s):
    return unicodedata.normalize("NFKD", str(s or "")).encode("ascii", "ignore").decode().lower()


def _clasificar_traslado(row, obj_cols, line_cols):
    """'servicio' (traslado de pacientes), 'incidental' (ambulancia como línea de otro contrato o
    accesorio), 'vehiculo' (compra/mantención) o None. Se basa en el NOMBRE de la licitación."""
    name = " ".join(_txtnorm(row.get(c, "")) for c in obj_cols)
    codes = set()
    for col in line_cols:
        raw = str(row.get(col, "") or "").strip()
        if raw:
            codes.add(raw)
    if "ambulancia" in name:
        if any(v in name for v in _VEH_WORDS):
            return "vehiculo"
        if any(o in name for o in _OTRO_WORDS):
            return "incidental"
        return "servicio"
    if ("traslado" in name or "transporte" in name) and any(m in name for m in _MED):
        return "servicio"
    if (_VEHICULO_ONU & codes) and any(v in name for v in _VEH_WORDS):
        return "vehiculo"
    if _SERVICIO_ONU & codes:
        return "incidental"       # línea de ambulancia dentro de un contrato de otra cosa
    return None


def cmd_import_lic_da(args):
    c = conn()
    origenes = []
    _LIC_BASE = "https://transparenciachc.blob.core.windows.net/lic-da/"
    for u in (args.url or []):
        u = str(u).strip()
        if u.lower().startswith(("http://", "https://")):
            u = u.split()[0]
        elif u and not os.path.exists(u):
            m = re.match(r"^(?:lic-da/)?(\d{4}-\d{1,2})(?:\.zip)?$", u)
            if m:
                u = _LIC_BASE + m.group(1) + ".zip"
            elif u.startswith("lic-da/"):
                u = "https://transparenciachc.blob.core.windows.net/" + u
        if u:
            origenes.append(u)
    if args.dir and os.path.isdir(args.dir):
        for f in sorted(os.listdir(args.dir)):
            if f.lower().endswith((".csv", ".gz", ".zip", ".7z")):
                origenes.append(os.path.join(args.dir, f))
    desde = getattr(args, "desde", None); hasta = getattr(args, "hasta", None)
    if desde and hasta:
        try:
            y, m = (int(x) for x in desde.split("-")); y2, m2 = (int(x) for x in hasta.split("-"))
            while (y, m) <= (y2, m2):
                origenes.append(f"https://transparenciachc.blob.core.windows.net/lic-da/{y}-{m}.zip")
                m += 1
                if m > 12:
                    m = 1; y += 1
        except ValueError:
            print("  ! --desde/--hasta deben ser AÑO-MES, ej: 2023-1"); return
    if not origenes:
        print("Sin fuentes. Usa --url <URL o AÑO-MES>, --desde/--hasta, o --dir <carpeta>."); return

    total_ins = 0
    for origen in origenes:
        try:
            descargado = _origen_a_path(origen)
        except Exception as e:  # noqa: BLE001
            print(f"  ! No se pudo descargar {origen}: {e}", flush=True); continue
        temporales = []; fh = None
        try:
            csv_path, temporales = _preparar_csv(descargado)
            enc = _sniff_encoding(csv_path)
            print(f"    codificación: {enc}", flush=True)
            fh = open(csv_path, encoding=enc, errors="replace")
            muestra = fh.readline()
            sep = ";" if muestra.count(";") >= muestra.count(",") else ","
            header = next(csv.reader([muestra], delimiter=sep))
            mapa = _mapear_columnas_lic(header)
            # columnas de rubro presentes en este archivo
            norm2real = {_colnorm(h): h for h in header}
            mapa_rubro = [norm2real[_colnorm(x)] for x in _LIC_RUBRO_COLS if _colnorm(x) in norm2real]
            obj_cols = [mapa[k] for k in ("nombre",) if k in mapa]
            print(f"\n  Fuente: {os.path.basename(str(origen))} · sep '{sep}' · {len(header)} columnas")
            print(f"  Mapeo: {mapa}")
            print(f"  Objeto (nombre/desc): {obj_cols} · rubro de línea: {mapa_rubro}", flush=True)
            if args.dry_run:
                print(f"  CABECERA COMPLETA: {header}", flush=True)
            faltan = [k for k in ("codigo", "monto", "fecha") if k not in mapa]
            if faltan:
                print(f"  ! Faltan columnas clave: {faltan}. Cabecera: {header}")
                print("  ! Ajusta _LIC_ALIASES y reintenta."); continue

            reader = csv.DictReader(fh, fieldnames=header, delimiter=sep)
            PRIOR = {"servicio": 3, "incidental": 2, "vehiculo": 1}
            cat = {}         # codigo -> mejor categoría vista
            filas = {}       # codigo -> última fila vista
            nombres = {}     # codigo -> nombre
            leidas = 0
            for row in reader:
                leidas += 1
                if args.limit and leidas > args.limit:
                    break
                cod = str(row.get(mapa["codigo"]) or "").strip()
                if not cod:
                    continue
                cl = _clasificar_traslado(row, obj_cols, mapa_rubro)
                if not cl:
                    continue
                if PRIOR[cl] > PRIOR.get(cat.get(cod, ""), 0):
                    cat[cod] = cl
                filas[cod] = row
                nombres[cod] = str(row.get(mapa.get("nombre", ""), "") or "").strip()
            serv = [k for k, v in cat.items() if v == "servicio"]
            inci = [k for k, v in cat.items() if v == "incidental"]
            veh = [k for k, v in cat.items() if v == "vehiculo"]
            incluir_veh = getattr(args, "incluir_vehiculos", False)
            incluir_inci = getattr(args, "incluir_incidentales", False)
            cargar = serv + (inci if incluir_inci else []) + (veh if incluir_veh else [])
            print(f"  Leídas {leidas} filas.")
            print(f"  Licitaciones → {len(serv)} traslado de pacientes · {len(inci)} incidental (ambulancia en otro contrato) · {len(veh)} vehículo (compra/mantención)")
            sel = []
            if not incluir_inci:
                sel.append("sin incidentales")
            if not incluir_veh:
                sel.append("sin vehículos")
            print(f"  Se cargarán {len(cargar)} ({', '.join(sel) or 'todo'}).", flush=True)
            if args.dry_run:
                print("  NOMBRES de 'traslado de pacientes' (hasta 25):")
                for cod in serv[:25]:
                    print(f"      · [{cod}] {nombres[cod][:80]}")
                if inci:
                    print("  NOMBRES marcados 'incidental' (hasta 10, para verificar el corte):")
                    for cod in inci[:10]:
                        print(f"      · [{cod}] {nombres[cod][:80]}")
                continue

            ins = 0
            for cod in cargar:
                row = filas[cod]
                fecha = _parse_fecha(row.get(mapa["fecha"])) or ""
                anio = int(fecha[:4]) if fecha[:4].isdigit() else 0
                mes = int(fecha[5:7]) if fecha[5:7].isdigit() else 0
                comprador = str(row.get(mapa.get("comprador", ""), "") or "").strip()
                region = _region_de(c, comprador)
                if region == "Sin Region":
                    region = (_region_corta(row.get(mapa["region"])) if mapa.get("region") else None) or "Sin Region"
                estado = str(row.get(mapa.get("estado", ""), "") or "").strip()
                adjudicada = 1 if _txtnorm(estado) == "adjudicada" else 0
                c.execute("INSERT OR REPLACE INTO licitacion VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?)", (
                    cod,
                    nombres.get(cod, ""),
                    comprador,
                    estado,
                    fecha,
                    str(row.get(mapa.get("tipo", ""), "") or "").strip(),
                    "CLP", 1.0, _parse_monto(row.get(mapa["monto"])),
                    anio, mes, region, adjudicada, "mercado"))
                ins += 1
            c.commit(); total_ins += ins
            print(f"  Insertadas/actualizadas {ins} licitaciones de traslado.", flush=True)
        except Exception as e:  # noqa: BLE001
            print(f"  ! Error procesando {origen}: {e} (se omite y se continúa)", flush=True)
        finally:
            if fh:
                fh.close()
            for t in temporales:
                d = os.path.dirname(t)
                if os.path.exists(t):
                    os.remove(t)
                if d.startswith(tempfile.gettempdir()) and os.path.isdir(d) and not os.listdir(d):
                    os.rmdir(d)
            if descargado != origen and os.path.exists(descargado):
                os.remove(descargado)
    if not args.dry_run:
        n = c.execute("SELECT COUNT(*) FROM licitacion WHERE origen='mercado'").fetchone()[0]
        print(f"\nListo. {total_ins} licitaciones incorporadas. Total de mercado: {n}.", flush=True)
    else:
        print("\nModo prueba (dry-run): nada se guardó. Revisa el mapeo y el conteo de traslado.", flush=True)


def _mapear_columnas_lic(header):
    norm2real = {_colnorm(h): h for h in header}
    mapa = {}
    for campo, alias in _LIC_ALIASES.items():
        for a in alias:
            real = norm2real.get(_colnorm(a))
            if real is not None:
                mapa[campo] = real; break
    return mapa


def cmd_run(args):
    cmd_resolve(args); cmd_update(args)
    if getattr(args, "sugerencias", 0):
        cmd_candidates(argparse.Namespace(desde=None, hasta=None, dias=args.sugerencias))
    if getattr(args, "licitaciones", 0):
        cmd_licitaciones(argparse.Namespace(desde=None, hasta=None, dias=args.licitaciones))
    cmd_build(args)


def main():
    ap = argparse.ArgumentParser(description="Tubería del panel de Traslados")
    sub = ap.add_subparsers(dest="cmd", required=True)
    s = sub.add_parser("seed"); s.add_argument("--excel", required=True)
    s.add_argument("--snapshot", default=dt.date.today().isoformat()); s.add_argument("--force", action="store_true")
    s.set_defaults(func=cmd_seed)
    sub.add_parser("resolve").set_defaults(func=cmd_resolve)
    u = sub.add_parser("update"); u.add_argument("--desde"); u.add_argument("--hasta")
    u.add_argument("--dias", type=int, default=14); u.add_argument("--tramo", type=int, default=0)
    u.add_argument("--max-llamadas", dest="max_llamadas", type=int, default=0,
                   help="tope de llamadas a la API por corrida (cuota diaria del ticket = 10.000)")
    u.set_defaults(func=cmd_update)
    sub.add_parser("build").set_defaults(func=cmd_build)
    rda = sub.add_parser("resolve-da", help="Corrige codigo_proveedor en providers.csv usando CodigoProveedor de Datos Abiertos.")
    rda.add_argument("--url", action="append", help="URL de un CSV/zip de Datos Abiertos (repetible).")
    rda.add_argument("--dir", default="datos_abiertos", help="Carpeta con archivos (por defecto: datos_abiertos).")
    rda.set_defaults(func=cmd_resolve_da)
    da = sub.add_parser("import-da", help="Importa OC desde CSV de Datos Abiertos (URL o carpeta), filtrando por tus RUT.")
    da.add_argument("--url", action="append", help="URL de un CSV/zip/7z de Datos Abiertos (repetible).")
    da.add_argument("--dir", default="datos_abiertos", help="Carpeta con archivos a importar (por defecto: datos_abiertos).")
    da.add_argument("--dry-run", action="store_true", help="Solo muestra el mapeo de columnas y cuenta, sin guardar.")
    da.add_argument("--limit", type=int, default=0, help="Procesa solo las primeras N filas (para pruebas).")
    da.set_defaults(func=cmd_import_da)
    lda = sub.add_parser("import-lic-da", help="Importa licitaciones de traslado desde Datos Abiertos (lic-da), filtrando por rubro.")
    lda.add_argument("--url", action="append", help="URL o AÑO-MES de un archivo lic-da (repetible). Ej: 2026-6")
    lda.add_argument("--desde", help="Rango: mes inicial AÑO-MES (ej: 2023-1).")
    lda.add_argument("--hasta", help="Rango: mes final AÑO-MES (ej: 2026-6).")
    lda.add_argument("--dir", default="lic_abiertos", help="Carpeta con archivos (por defecto: lic_abiertos).")
    lda.add_argument("--dry-run", action="store_true", help="Solo muestra mapeo, cabecera y conteo de traslado.")
    lda.add_argument("--incluir-vehiculos", action="store_true", help="Incluye también compra/mantención de ambulancias.")
    lda.add_argument("--incluir-incidentales", action="store_true", help="Incluye contratos donde la ambulancia es solo una línea (eventos, etc.).")
    lda.add_argument("--limit", type=int, default=0, help="Procesa solo las primeras N filas (para pruebas).")
    lda.set_defaults(func=cmd_import_lic_da)
    ca = sub.add_parser("candidates"); ca.add_argument("--desde"); ca.add_argument("--hasta")
    ca.add_argument("--dias", type=int, default=30); ca.set_defaults(func=cmd_candidates)
    li = sub.add_parser("licitaciones"); li.add_argument("--desde"); li.add_argument("--hasta")
    li.add_argument("--dias", type=int, default=30); li.set_defaults(func=cmd_licitaciones)
    pr = sub.add_parser("probe")
    pr.add_argument("codigo", nargs="?", default=None)
    pr.add_argument("--dia"); pr.add_argument("--codigo-proveedor", dest="codigo_proveedor")
    pr.set_defaults(func=cmd_probe)
    r = sub.add_parser("run"); r.add_argument("--desde"); r.add_argument("--hasta")
    r.add_argument("--dias", type=int, default=14); r.add_argument("--tramo", type=int, default=0)
    r.add_argument("--max-llamadas", dest="max_llamadas", type=int, default=0,
                   help="tope de llamadas a la API por corrida (cuota diaria del ticket = 10.000)")
    r.add_argument("--sugerencias", type=int, default=0,
                   help="si >0, escanea ese N de días buscando proveedores candidatos")
    r.add_argument("--licitaciones", type=int, default=0,
                   help="si >0, trae licitaciones de ese N de días")
    r.set_defaults(func=cmd_run)
    args = ap.parse_args()
    try:
        args.func(args)
    except RuntimeError as e:
        print("Error:", e); return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
