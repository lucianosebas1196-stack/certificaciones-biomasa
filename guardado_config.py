"""Guardado del config (trabajo/config_proveedores.xlsx) que no se pierde aunque alguien lo tenga abierto en Excel.

Problema: Excel bloquea el archivo mientras está abierto. Antes, si el programa quería guardar (precios, CUIT,
numeración...), fallaba con PermissionError y el cambio quedaba solo en memoria: si se reiniciaba el servidor
antes del próximo guardado, se perdía.

Cómo funciona ahora:
  1. Siempre se escribe primero un archivo aparte y después se reemplaza el original de un solo paso
     (nunca queda un config escrito a medias).
  2. Si el original está bloqueado, lo escrito se deja en trabajo/config_proveedores_PENDIENTE.xlsx y se crea la
     bandera trabajo/config_proveedores_PENDIENTE.json (desde cuándo, por qué). No se corta nada.
  3. Un hilo revisa cada 60 segundos: apenas se cierra el Excel, el PENDIENTE pasa a ser el config y se borran
     el temporal y la bandera. Si mientras tanto alguien guardó cambios EN Excel, esa versión no se pisa sin
     más: se guarda una copia en config_respaldos/ y se avisa.
  4. estado() dice si hay algo pendiente: el servidor lo agrega a sus respuestas y la página muestra un banner.

Conexión con servidor.py (ver el bloque "Guardado del config" allí):
    import guardado_config
    GUARDADO = guardado_config.GuardadoConfig(CONFIG_TRABAJO, escribir=cb.guardar_config)
    cb.guardar_config = GUARDADO.envolver(cb.guardar_config)   # todo lo que guarde el config pasa por acá
    GUARDADO.iniciar_sincronizacion()                           # el hilo de los 60 segundos
    ... y en las respuestas JSON: GUARDADO.estado()
"""
from __future__ import annotations

import json
import os
import shutil
import threading
import time
import uuid
from datetime import datetime
from pathlib import Path
from typing import Callable

__all__ = ["GuardadoConfig"]


class GuardadoConfig:
    def __init__(self, ruta, escribir: Callable, intervalo: float = 60.0, registrar: Callable[[str], None] | None = None):
        """ruta: el config oficial (trabajo/config_proveedores.xlsx).
        escribir(ruta, conf, precios, par): la función que arma el Excel (cb.guardar_config original).
        registrar(texto): opcional, para dejar constancia (accesos.log / ventana negra)."""
        self.ruta = Path(ruta)
        self.pendiente = self.ruta.with_name(self.ruta.stem + "_PENDIENTE.xlsx")
        self.bandera = self.ruta.with_name(self.ruta.stem + "_PENDIENTE.json")
        self.respaldos = self.ruta.parent / "config_respaldos"
        self._escribir = escribir
        self.intervalo = float(intervalo)
        self._registrar = registrar or (lambda texto: print(texto))
        self._lock = threading.RLock()                # guardar y sincronizar nunca a la vez
        self._parar = threading.Event()
        self._hilo: threading.Thread | None = None
        self._ultimo_aviso = None                     # mensaje de "ya se pasó" para mostrar una vez
        self._reemplazar = os.replace                 # se puede cambiar en pruebas

    # ------------------------------------------------------------------------------------------ guardar
    def envolver(self, original: Callable) -> Callable:
        """Devuelve una función con la misma firma que cb.guardar_config: si la ruta es el config oficial usa el
        guardado protegido; cualquier otra ruta (por ejemplo, un archivo para descargar) se escribe como siempre."""
        def guardar_config(ruta, conf, precios, par):
            if Path(ruta).resolve() == self.ruta.resolve():
                return self.guardar(conf, precios, par)
            return original(ruta, conf, precios, par)
        guardar_config.__doc__ = original.__doc__
        return guardar_config

    def guardar(self, conf, precios, par) -> bool:
        """Guarda el config. True = quedó en el archivo oficial; False = quedó PENDIENTE (archivo bloqueado).
        Solo lanza excepción si no se pudo guardar en NINGÚN lado (disco lleno, sin permisos en la carpeta)."""
        with self._lock:
            tmp = self.ruta.with_name(f"~{self.ruta.stem}.{uuid.uuid4().hex[:8]}.xlsx")   # .xlsx: lo exige openpyxl
            try:
                self._escribir(tmp, conf, precios, par)          # 1) el Excel completo en un archivo aparte
            except BaseException:
                self._borrar(tmp)
                raise
            try:
                self._reemplazar(tmp, self.ruta)                 # 2) reemplazo de un solo paso
            except OSError as e:                                 # PermissionError incluido: Excel lo tiene abierto
                return self._dejar_pendiente(tmp, e)
            # se guardó en el oficial: lo pendiente (si había) quedó superado por esta versión más nueva
            if self.bandera.exists() or self.pendiente.exists():
                self._borrar(self.pendiente)
                self._borrar(self.bandera)
                self._ultimo_aviso = "✅ config_proveedores.xlsx se pudo guardar: ya no hay cambios pendientes."
            return True

    def _dejar_pendiente(self, tmp: Path, error: OSError) -> bool:
        try:
            self._reemplazar(tmp, self.pendiente)                # la versión más nueva siempre gana
        except OSError:
            try:                                                 # el PENDIENTE también abierto: se copia encima
                shutil.copyfile(tmp, self.pendiente)
            finally:
                self._borrar(tmp)
        datos = self._leer_bandera() or {"desde": datetime.now().strftime("%d/%m/%Y %H:%M"),
                                         "desde_ts": time.time(), "original_mtime": self._mtime(self.ruta)}
        datos.update({"ultimo_intento": datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
                      "motivo": f"{type(error).__name__}: {error}"})
        self._escribir_bandera(datos)
        self._registrar(f"CONFIG BLOQUEADO (abierto en otro programa): cambios guardados en {self.pendiente.name}")
        return False

    # ------------------------------------------------------------------------------------------ sincronizar
    def sincronizar(self) -> bool:
        """Si hay algo pendiente y el original ya se liberó, lo pasa. True si quedó todo al día."""
        with self._lock:
            if not self.pendiente.exists():
                self._borrar(self.bandera)                       # bandera sin archivo: no hay nada que pasar
                return True
            datos = self._leer_bandera() or {}
            editado_en_excel = (self.ruta.exists() and datos.get("original_mtime") is not None
                                and self._mtime(self.ruta) != datos["original_mtime"])
            copia_excel = None
            if editado_en_excel:                                 # alguien guardó cambios en Excel mientras tanto
                try:
                    self.respaldos.mkdir(exist_ok=True)
                    copia_excel = self.respaldos / f"config_proveedores_EDITADO_EN_EXCEL_{datetime.now():%Y-%m-%d_%H%M%S}.xlsx"
                    shutil.copy2(self.ruta, copia_excel)
                except OSError:
                    return False                                 # todavía bloqueado para leer: se reintenta luego
            try:
                self._reemplazar(self.pendiente, self.ruta)
            except OSError:
                if copia_excel:
                    self._borrar(copia_excel)
                return False                                     # sigue abierto: se reintenta en 60 s
            self._borrar(self.bandera)
            self._ultimo_aviso = ("✅ Se cerró el Excel: los cambios pendientes ya están en config_proveedores.xlsx."
                                  + (f" Ojo: el archivo se había editado en Excel; esa versión quedó en "
                                     f"config_respaldos\\{copia_excel.name} (revisala)." if copia_excel else ""))
            self._registrar("CONFIG PENDIENTE SINCRONIZADO" + (f" (versión de Excel en {copia_excel.name})" if copia_excel else ""))
            return True

    def iniciar_sincronizacion(self) -> None:
        """Hilo en segundo plano: al arrancar y después cada `intervalo` segundos (solo actúa si hay pendiente)."""
        if self._hilo and self._hilo.is_alive():
            return

        def bucle():
            while True:
                try:
                    if self.pendiente.exists() or self.bandera.exists():
                        self.sincronizar()
                except Exception as e:                           # el hilo no se muere por un error puntual
                    print(f"⚠ Sincronización del config: {e}")
                if self._parar.wait(self.intervalo):
                    return

        self._parar.clear()
        self._hilo = threading.Thread(target=bucle, name="sincronizar-config", daemon=True)
        self._hilo.start()

    def detener(self) -> None:
        self._parar.set()

    # ------------------------------------------------------------------------------------------ estado
    def ruta_vigente(self) -> Path:
        """El config más actualizado: el PENDIENTE si existe (aún no pudo pasar al oficial), si no el oficial.
        Usarlo para leer o descargar el config, así nunca se usa una versión vieja."""
        return self.pendiente if self.pendiente.exists() else self.ruta

    def estado(self) -> dict:
        """Para agregar a las respuestas JSON (la página muestra el banner con esto)."""
        datos = self._leer_bandera() if self.pendiente.exists() else None
        resuelto, self._ultimo_aviso = self._ultimo_aviso, None
        if not datos:
            return {"bloqueado": False, **({"resuelto": resuelto} if resuelto else {})}
        return {"bloqueado": True, "desde": datos.get("desde", ""), "archivo": self.ruta.name,
                "pendiente": self.pendiente.name,
                "mensaje": (f"⚠ {self.ruta.name} está abierto en Excel (u otro programa) desde las "
                            f"{datos.get('desde', '?')[-5:]}. Los cambios NO se perdieron: están guardados en "
                            f"{self.pendiente.name} y pasan solos apenas cierres el Excel (se revisa cada minuto).")}

    # ------------------------------------------------------------------------------------------ auxiliares
    def _leer_bandera(self) -> dict | None:
        try:
            return json.loads(self.bandera.read_text(encoding="utf-8"))
        except (OSError, ValueError):
            return None

    def _escribir_bandera(self, datos: dict) -> None:
        tmp = self.bandera.with_suffix(".tmp")
        try:
            tmp.write_text(json.dumps(datos, ensure_ascii=False, indent=1), encoding="utf-8")
            os.replace(tmp, self.bandera)
        except OSError as e:                                     # sin bandera igual funciona: manda el PENDIENTE
            print(f"⚠ No pude escribir {self.bandera.name}: {e}")

    @staticmethod
    def _mtime(p: Path):
        try:
            return p.stat().st_mtime_ns
        except OSError:
            return None

    @staticmethod
    def _borrar(p: Path | None) -> None:
        if p is None:
            return
        try:
            p.unlink(missing_ok=True)
        except OSError:
            pass
