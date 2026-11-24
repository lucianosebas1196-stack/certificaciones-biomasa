"""Lista blanca de IP: solo las PC autorizadas de la planta pueden conectarse al programa.

La lista está en trabajo/ips_autorizadas.txt (texto plano, una entrada por línea):

    # comentario
    192.168.1.50            ← una PC puntual
    192.168.1.0/24          ← toda una red (notación CIDR)

· Si el archivo no existe, se crea con las redes locales comunes y la red de esta PC.
· Si alguien edita el archivo, el cambio vale desde el próximo pedido (no hace falta reiniciar).
· Las líneas mal escritas se ignoran y se avisa en la ventana negra; el resto de la lista sigue valiendo.
· La propia PC servidor (127.0.0.1 / ::1) siempre puede entrar: así nunca queda nadie afuera sin arreglo.
· Sin --red (EN_RED = False) solo se acepta la propia PC, aunque la lista diga otra cosa.

Integración en servidor.py (DESPUÉS de definir el middleware _seguridad: el último middleware que se registra es
el primero que atiende, así la IP se controla antes que todo lo demás):

    import lista_blanca
    lista_blanca.instalar(app_web, TRABAJO / "ips_autorizadas.txt", en_red=lambda: EN_RED,
                          ips_propias=_ips_locales, registrar=_registrar_acceso)
"""
from __future__ import annotations

import ipaddress
import threading
import time
from pathlib import Path
from typing import Callable, Iterable

from fastapi.responses import JSONResponse

__all__ = ["ListaBlanca", "instalar", "MENSAJE"]

MENSAJE = "Acceso denegado desde esta ubicación"
LOCALES = ("127.0.0.1", "::1")
REDES_POR_DEFECTO = ("127.0.0.1", "::1", "192.168.1.0/24", "10.0.0.0/8")

_Red = ipaddress.IPv4Network | ipaddress.IPv6Network


def _ip(texto: str) -> ipaddress.IPv4Address | ipaddress.IPv6Address | None:
    """IP del cliente como objeto; None si no es una IP (ej. un socket local sin dirección).
    Las IPv4 que llegan como IPv6 (::ffff:192.168.1.5) se tratan como IPv4."""
    try:
        ip = ipaddress.ip_address(str(texto).strip().strip("[]").split("%")[0])
    except ValueError:
        return None
    if isinstance(ip, ipaddress.IPv6Address) and ip.ipv4_mapped:
        return ip.ipv4_mapped
    return ip


def _es_local(ip) -> bool:
    return ip is not None and ip.is_loopback


class ListaBlanca:
    """Lee y mantiene al día la lista de IP/redes autorizadas."""

    def __init__(self, ruta, ips_propias: Callable[[], Iterable[str]] | None = None):
        self.ruta = Path(ruta)
        self._ips_propias = ips_propias or (lambda: [])
        self._redes: list[_Red] = []
        self._marca = None                      # (fecha de modificación, tamaño) de la última lectura
        self._lock = threading.Lock()
        self._crear_si_falta()
        self._recargar(forzar=True)

    # ------------------------------------------------------------------ archivo
    def _crear_si_falta(self) -> None:
        if self.ruta.exists():
            return
        propias = []
        for ip in self._ips_propias():
            try:
                red = ipaddress.ip_network(f"{ip}/24", strict=False)       # la red de esta PC (ej. 192.168.1.0/24)
            except ValueError:
                continue
            if str(red) not in REDES_POR_DEFECTO and str(red) not in propias:
                propias.append(str(red))
        lineas = [
            "# IP y redes autorizadas a usar el programa de certificaciones (una por línea).",
            "# Ejemplos:  192.168.1.50   (una PC)      192.168.1.0/24   (toda la red 192.168.1.x)",
            "# Las líneas que empiezan con # no cuentan. Los cambios valen al instante (sin reiniciar).",
            "# La PC servidor (127.0.0.1) siempre puede entrar.",
            "",
            *REDES_POR_DEFECTO,
            *(["", "# Red de la PC servidor (agregada sola al crear este archivo)"] + propias if propias else []),
            "",
        ]
        try:
            self.ruta.parent.mkdir(parents=True, exist_ok=True)
            tmp = self.ruta.with_suffix(".tmp")
            tmp.write_text("\n".join(lineas), encoding="utf-8")
            tmp.replace(self.ruta)
        except OSError as e:
            print(f"⚠ No pude crear {self.ruta}: {e}. Solo se permite la propia PC.")

    def _recargar(self, forzar: bool = False) -> None:
        """Vuelve a leer el archivo si cambió desde la última vez."""
        try:
            st = self.ruta.stat()
            marca = (st.st_mtime_ns, st.st_size)
        except OSError:
            marca = None
        if not forzar and marca == self._marca:
            return
        with self._lock:
            if not forzar and marca == self._marca:
                return
            redes, malas = [], []
            try:
                texto = self.ruta.read_text(encoding="utf-8-sig")
            except (OSError, UnicodeDecodeError) as e:
                print(f"⚠ No pude leer {self.ruta.name} ({e}): solo se permite la propia PC.")
                texto = ""
            for n, linea in enumerate(texto.splitlines(), 1):
                linea = linea.split("#", 1)[0].strip()
                if not linea:
                    continue
                try:
                    redes.append(ipaddress.ip_network(linea, strict=False))    # "192.168.1.50" → /32
                except ValueError:
                    malas.append(f"línea {n}: '{linea}'")
            if malas:
                print(f"⚠ {self.ruta.name}: no entiendo {', '.join(malas)} (se ignoran).")
            self._redes, self._marca = redes, marca

    # ------------------------------------------------------------------ consulta
    def permitida(self, host: str, en_red: bool) -> bool:
        """¿Puede conectarse esta IP? Sin red: solo la propia PC. En red: la propia PC o algo de la lista."""
        if str(host).strip().casefold() == "localhost":
            return True
        ip = _ip(host)
        if _es_local(ip):
            return True
        if not en_red or ip is None:
            return False
        self._recargar()
        return any(ip.version == red.version and ip in red for red in self._redes)

    @property
    def redes(self) -> list[str]:
        self._recargar()
        return [str(r) for r in self._redes]


def instalar(app, ruta, en_red: Callable[[], bool], ips_propias=None,
             registrar: Callable[[str, str, str], None] | None = None) -> ListaBlanca:
    """Agrega el middleware a la app. Registrarlo DESPUÉS de los otros middleware para que atienda primero."""
    lista = ListaBlanca(ruta, ips_propias)
    avisadas: dict[str, float] = {}             # para no llenar el registro con la misma IP rechazada

    @app.middleware("http")
    async def _lista_blanca(request, call_next):
        host = request.client.host if request.client else ""
        try:
            ok = lista.permitida(host, bool(en_red()))
        except Exception as e:                  # ante cualquier falla, se cierra (no se abre)
            print(f"⚠ Error en la lista blanca de IP: {e}")
            ok = _es_local(_ip(host))
        if not ok:
            ahora = time.time()
            if registrar and ahora - avisadas.get(host, 0) > 600:          # una vez cada 10 minutos por IP
                avisadas[host] = ahora
                try:
                    registrar("-", host or "?", "RECHAZADO: IP NO AUTORIZADA")
                except Exception:
                    pass
            return JSONResponse({"detail": MENSAJE}, status_code=403)
        return await call_next(request)

    return lista
