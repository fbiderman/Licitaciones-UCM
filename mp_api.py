"""Cliente para la API pública de Mercado Público (ChileCompra).

Documentación: https://api.mercadopublico.cl/modules/api.aspx
El `ticket` se solicita gratis en ChileCompra y se entrega por variable de entorno MP_TICKET.

NOTA IMPORTANTE SOBRE EL ESQUEMA
--------------------------------
La API no publica un esquema JSON formal y estable de la Orden de Compra. La consulta
por `codigo` devuelve el detalle completo; la consulta por `fecha` devuelve datos básicos.
`extraer_campos()` es deliberadamente defensiva: prueba varios nombres de campo probables.
Antes de la primera corrida real, ejecuta `python pipeline.py probe <codigo_oc>` para volcar
una respuesta cruda y ajustar el mapeo si algún campo no coincide.
"""
from __future__ import annotations
import os, time, json, unicodedata, urllib.parse, urllib.request

BASE = "https://api.mercadopublico.cl/servicios/v1/publico"
BASE_EMP = "https://api.mercadopublico.cl/servicios/v1/Publico/Empresas"


class MPQuotaError(Exception):
    """La API respondió que el ticket superó su cuota diaria (Codigo 203)."""


class MPAPIError(Exception):
    """La API respondió con un sobre de error {Codigo, Mensaje} distinto de cuota."""


class MPClient:
    def __init__(self, ticket: str | None = None, min_interval: float = 0.3,
                 max_retries: int = 2, timeout: int = 25):
        self.ticket = ticket or os.environ.get("MP_TICKET")
        if not self.ticket:
            raise RuntimeError("Falta el ticket. Define la variable de entorno MP_TICKET.")
        self.min_interval = min_interval          # segundos entre llamadas (rate limit)
        self.max_retries = max_retries
        self.timeout = timeout
        self._last = 0.0

    # -------- transporte --------
    def _get(self, url: str) -> dict:
        for intento in range(1, self.max_retries + 1):
            wait = self.min_interval - (time.time() - self._last)
            if wait > 0:
                time.sleep(wait)
            try:
                req = urllib.request.Request(url, headers={"User-Agent": "traslados-dashboard/1.0"})
                with urllib.request.urlopen(req, timeout=self.timeout) as r:
                    self._last = time.time()
                    raw = r.read().decode("utf-8", "replace")
                data = json.loads(raw)
            except Exception as e:                 # noqa: BLE001  (fallo de red/parseo)
                self._last = time.time()
                if intento == self.max_retries:
                    raise
                time.sleep(min(2 * intento, 4))    # backoff acotado
                continue
            # respuesta válida: detectar el sobre de error {Codigo, Mensaje} y NO reintentar
            if (isinstance(data, dict) and "Codigo" in data and "Mensaje" in data
                    and "Listado" not in data and "listaEmpresas" not in data and "empresas" not in data):
                cod = data.get("Codigo"); msg = str(data.get("Mensaje", ""))
                if cod == 203 or "cuota" in msg.lower():
                    raise MPQuotaError(msg)
                raise MPAPIError(f"{cod}: {msg}")
            return data
        return {}

    def _url(self, endpoint: str, **params) -> str:
        params["ticket"] = self.ticket
        return f"{endpoint}?{urllib.parse.urlencode(params)}"

    # -------- métodos --------
    def buscar_proveedor(self, rut: str) -> dict | None:
        """Devuelve el primer registro {CodigoEmpresa, NombreEmpresa,...} para un RUT."""
        url = self._url(f"{BASE_EMP}/BuscarProveedor", rutempresaproveedor=_rut_con_puntos(rut))
        data = self._get(url)
        listado = data.get("listaEmpresas") or data.get("Listado") or data.get("empresas") or []
        if isinstance(listado, dict):
            listado = [listado]
        return listado[0] if listado else None

    def oc_por_proveedor_dia(self, codigo_proveedor: str | int, fecha_ddmmaaaa: str) -> list[dict]:
        """Lista básica de OC de un proveedor en un día (formato de fecha: ddmmaaaa)."""
        url = self._url(f"{BASE}/ordenesdecompra.json",
                        fecha=fecha_ddmmaaaa, CodigoProveedor=str(codigo_proveedor))
        data = self._get(url)
        return data.get("Listado") or []

    def oc_del_dia(self, fecha_ddmmaaaa: str, estado: str = "todos") -> list[dict]:
        """Lista básica de TODAS las OC de un día (alternativa a la consulta por proveedor)."""
        url = self._url(f"{BASE}/ordenesdecompra.json", fecha=fecha_ddmmaaaa, estado=estado)
        data = self._get(url)
        return data.get("Listado") or []

    def oc_detalle(self, codigo_oc: str) -> dict | None:
        """Detalle completo de una OC por su código."""
        url = self._url(f"{BASE}/ordenesdecompra.json", codigo=codigo_oc)
        data = self._get(url)
        listado = data.get("Listado") or []
        return listado[0] if listado else None

    # ---- licitaciones ----
    def licitaciones_por_proveedor_dia(self, codigo_proveedor: str | int, fecha_ddmmaaaa: str) -> list[dict]:
        url = self._url(f"{BASE}/licitaciones.json",
                        fecha=fecha_ddmmaaaa, CodigoProveedor=str(codigo_proveedor))
        return (self._get(url) or {}).get("Listado") or []

    def licitacion_detalle(self, codigo_lic: str) -> dict | None:
        url = self._url(f"{BASE}/licitaciones.json", codigo=codigo_lic)
        listado = (self._get(url) or {}).get("Listado") or []
        return listado[0] if listado else None


# ---------------- utilidades de parseo ----------------
def _rut_con_puntos(rut: str) -> str:
    """'77912490-8' -> '77.912.490-8' (formato que espera BuscarProveedor)."""
    rut = rut.replace(".", "").replace(" ", "").upper()
    if "-" not in rut:
        rut = rut[:-1] + "-" + rut[-1]
    num, dv = rut.split("-")
    partes = []
    while num:
        partes.insert(0, num[-3:]); num = num[:-3]
    return ".".join(partes) + "-" + dv


def _first(d: dict, *keys, default=None):
    for k in keys:
        if isinstance(d, dict) and d.get(k) not in (None, ""):
            return d[k]
    return default


# estados -> texto consistente con el Excel histórico
ESTADO_TXT = {
    4: "Enviada a Proveedor", 5: "En proceso", 6: "Aceptada", 8: "No Aceptada",
    9: "Cancelada", 12: "Recepcion Conforme", 13: "Pendiente de Recepcionar",
    14: "Recepcionada Parcialmente", 15: "Recepcion Conforme Incompleta",
}


def extraer_campos(oc: dict) -> dict:
    """Normaliza el detalle crudo de una OC al esquema del histórico (columnas del Excel).

    Ajusta aquí los nombres de campo si `probe` muestra otras claves en tu respuesta real.
    """
    comprador = _first(oc, "Comprador", default={}) or {}
    proveedor = _first(oc, "Proveedor", default={}) or {}
    fechas = _first(oc, "Fechas", default={}) or {}

    cod_estado = _first(oc, "CodigoEstado", "codigoEstado")
    estado = ESTADO_TXT.get(cod_estado) or _first(oc, "Estado", "estado", default="")

    moneda = _first(oc, "TipoMoneda", "Moneda", "MonedaOrdenCompra", "codigoMoneda", default="CLP")
    total = _first(oc, "Total", "TotalOC", "MontoTotalOC", "totalOC", default=0) or 0

    fecha_txt = _first(fechas, "FechaEnvio", "FechaCreacion", "FechaAceptacion") \
        or _first(oc, "FechaEnvio", "FechaCreacion", default="")

    return {
        "codigo": _first(oc, "Codigo", "codigo", default=""),
        "nombre": _first(oc, "Nombre", "nombre", default=""),
        "comprador": _first(comprador, "NombreOrganismo", "Nombre", "nombreOrganismo", default=""),
        "rut_comprador": _first(comprador, "RutUnidad", "RutSucursal", "Rut", default=""),
        "fecha": str(fecha_txt),
        "proveedor": _first(proveedor, "Nombre", "NombreEmpresa", "nombre", default=""),
        "rut_proveedor": _first(proveedor, "RutSucursal", "RutEmpresa", "Rut", default=""),
        "estado": estado,
        "moneda": str(moneda).upper(),
        "monto_bruto": float(total),
        "tipo_orden": _first(oc, "Tipo", "tipo", "TipoDespachoOC", default="") or "Sin Clasificacion",
    }


# estados de licitación -> texto
ESTADO_LIC = {5: "Publicada", 6: "Cerrada", 7: "Desierta", 8: "Adjudicada",
              18: "Readjudicada", 19: "Suspendida", 4: "Revocada"}


def extraer_licitacion(lic: dict) -> dict:
    """Normaliza el detalle de una licitación. Verifica claves con `probe` si difieren."""
    comprador = _first(lic, "Comprador", default={}) or {}
    fechas = _first(lic, "Fechas", default={}) or {}
    cod_estado = _first(lic, "CodigoEstado", "codigoEstado")
    estado = ESTADO_LIC.get(cod_estado) or _first(lic, "Estado", "estado", default="")
    fecha = _first(fechas, "FechaPublicacion", "FechaCreacion", "FechaCierre") \
        or _first(lic, "FechaPublicacion", "FechaCreacion", default="")
    return {
        "codigo": _first(lic, "CodigoExterno", "Codigo", "codigo", default=""),
        "nombre": _first(lic, "Nombre", "nombre", default=""),
        "comprador": _first(comprador, "NombreOrganismo", "Nombre", default=""),
        "estado": estado,
        "fecha": str(fecha),
        "tipo": _first(lic, "Tipo", "tipo", default="") or "Sin Tipo",
        "moneda": str(_first(lic, "Moneda", "TipoMoneda", default="CLP")).upper(),
        "monto_estimado": float(_first(lic, "MontoEstimado", "Estimacion", "montoEstimado", default=0) or 0),
        "adjudicada": 1 if str(estado).lower().startswith("adjud") else 0,
    }


# ---------------- tipo de cambio UF/UTM (mindicador.cl) ----------------
# La API de OC entrega el monto en la moneda de la orden. Para llevar UF/UTM a CLP
# se usa el valor del indicador en la fecha de la OC.
MINDICADOR = "https://mindicador.cl/api"
_MONEDA_IND = {"CLF": "uf", "UF": "uf", "UTM": "utm"}  # CLP -> factor 1 (no se consulta)


def indicador_para_moneda(moneda: str) -> str | None:
    return _MONEDA_IND.get(str(moneda).upper())


def fetch_indicador(indicador: str, fecha_iso: str, timeout: int = 30) -> float | None:
    """Valor de 'uf' o 'utm' (CLP por unidad) en una fecha. None si no se puede obtener."""
    import datetime as _dt
    try:
        d = _dt.date.fromisoformat(str(fecha_iso)[:10])
        url = f"{MINDICADOR}/{indicador}/{d.strftime('%d-%m-%Y')}"
        req = urllib.request.Request(url, headers={"User-Agent": "traslados-dashboard/1.0"})
        with urllib.request.urlopen(req, timeout=timeout) as r:
            data = json.loads(r.read().decode("utf-8", "replace"))
        serie = data.get("serie") or []
        return float(serie[0]["valor"]) if serie else None
    except Exception:  # noqa: BLE001
        return None


# ---------------- detección de OC de traslados por nombre ----------------
_TRASLADO_KW = ("traslado", "ambulancia", "transporte de paciente", "transporte de pacientes",
                "traslado de paciente", "rescate", "atencion prehospitalaria",
                "urgencia prehospitalaria", "movilizacion de paciente")


def es_traslado(nombre: str) -> bool:
    n = unicodedata.normalize("NFKD", str(nombre)).encode("ascii", "ignore").decode().lower()
    return any(k in n for k in _TRASLADO_KW)
