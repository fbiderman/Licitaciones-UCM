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
import argparse, csv, datetime as dt, json, os, sqlite3, sys, unicodedata

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
  anio INTEGER, mes INTEGER, region TEXT, adjudicada INTEGER
);
"""


def conn():
    c = sqlite3.connect(DB)
    c.executescript(SCHEMA)
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
    from mp_api import MPClient, extraer_campos
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
    print(f"Actualizando {desde} -> {hasta} para {len(provs)} proveedores…", flush=True)

    nuevas = 0
    dia = desde
    while dia <= hasta:
        fstr = dia.strftime("%d%m%Y")
        for p in provs:
            try:
                listado = cli.oc_por_proveedor_dia(p["codigo_proveedor"], fstr)
            except Exception as e:  # noqa: BLE001
                print(f"  ! {fstr} {p['alias']}: {e}"); continue
            for item in listado:
                cod = item.get("Codigo") or item.get("codigo")
                if not cod:
                    continue
                if c.execute("SELECT 1 FROM oc WHERE codigo=?", (cod,)).fetchone():
                    continue
                try:
                    det = cli.oc_detalle(cod)
                except Exception as e:  # noqa: BLE001
                    print(f"  ! detalle {cod}: {e}"); continue
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
                c.execute("INSERT OR REPLACE INTO oc VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
                          (cod, f["nombre"], f["comprador"], f["rut_comprador"], fecha,
                           f["proveedor"] or p["alias"], f["rut_proveedor"] or p["rut"],
                           f["estado"], f["moneda"], cr, f["monto_bruto"], f["tipo_orden"],
                           anio, mes, _region_de(c, f["comprador"])))
                nuevas += 1
        if dia.day % 7 == 0 or dia == hasta:
            c.commit(); print(f"  … {dia}  (+{nuevas} OC nuevas)", flush=True)
        dia += dt.timedelta(days=1)
    c.execute("INSERT OR REPLACE INTO meta VALUES ('last_update', ?)",
              (dt.datetime.now().isoformat(timespec="seconds"),))
    c.commit()
    print(f"Listo. {nuevas} OC nuevas incorporadas.")


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
                c.execute("INSERT OR REPLACE INTO licitacion VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)",
                          (cod, f["nombre"], f["comprador"], f["estado"], fecha, f["tipo"],
                           f["moneda"], cr, f["monto_estimado"], anio, mes,
                           _region_de(c, f["comprador"]), f["adjudicada"]))
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
    uf_ref = last_fx("uf")  # UF de referencia para el selector $/UF

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
                     "rows": len(out), "total_clp": int(total), "fx": fx, "uf_ref": uf_ref},
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
    from mp_api import MPClient
    cli = MPClient()
    det = cli.oc_detalle(args.codigo)
    print(json.dumps(det, ensure_ascii=False, indent=2)[:6000])
    print("\n--- Revisa que estas claves existan y ajusta mp_api.extraer_campos() si difieren ---")


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
    u.set_defaults(func=cmd_update)
    sub.add_parser("build").set_defaults(func=cmd_build)
    ca = sub.add_parser("candidates"); ca.add_argument("--desde"); ca.add_argument("--hasta")
    ca.add_argument("--dias", type=int, default=30); ca.set_defaults(func=cmd_candidates)
    li = sub.add_parser("licitaciones"); li.add_argument("--desde"); li.add_argument("--hasta")
    li.add_argument("--dias", type=int, default=30); li.set_defaults(func=cmd_licitaciones)
    pr = sub.add_parser("probe"); pr.add_argument("codigo"); pr.set_defaults(func=cmd_probe)
    r = sub.add_parser("run"); r.add_argument("--desde"); r.add_argument("--hasta")
    r.add_argument("--dias", type=int, default=14); r.add_argument("--tramo", type=int, default=0)
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
