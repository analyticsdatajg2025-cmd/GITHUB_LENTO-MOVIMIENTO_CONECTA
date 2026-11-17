"""Resumen de cada ejecución para el monitor de MKT SISTEMA (módulo «Lento y Obsoletos»).

El script de flyers va anotando aquí lo que pasa (tiendas, productos, imágenes, mes usado...) y cada
anotación se guarda al momento en resumen.json. Al final del job, el workflow ejecuta
`python resumen.py cerrar <resultado del paso>` y sube el archivo al release «resumenes» del repo,
de donde lo lee MKT SISTEMA. Solo informa: si algo de aquí falla, el proceso de flyers sigue igual.
"""
import json
import os
import sys
import traceback
from datetime import datetime, timezone

ARCHIVO = "resumen.json"
_r = None


def _ahora():
    return datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")


def _guardar():
    try:
        tmp = ARCHIVO + ".tmp"
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(_r, f, ensure_ascii=False)
        os.replace(tmp, ARCHIVO)
    except Exception as e:  # nunca frenar el proceso por el resumen
        print(f"(resumen) no se pudo guardar: {e}")


def _excepcion(tipo, valor, tb):
    # Error no controlado: queda anotado antes de que Python termine
    if _r is not None and not issubclass(tipo, KeyboardInterrupt):
        error(f"{tipo.__name__}: {valor}")
    sys.__excepthook__(tipo, valor, tb)


def iniciar(proceso, bloque="0"):
    """Empieza el resumen de esta máquina (borra el de un intento anterior)."""
    global _r
    _r = {
        "v": 1,
        "proceso": proceso,
        "bloque": str(bloque),
        "run_id": os.environ.get("GITHUB_RUN_ID", ""),
        "run_attempt": os.environ.get("GITHUB_RUN_ATTEMPT", ""),
        "inicio": _ahora(),
        "fin": "",
        "resultado": "en curso",
        "error": "",
        "datos": {},
        "tiendas": [],
    }
    sys.excepthook = _excepcion
    _guardar()


def dato(**kv):
    """Anota datos sueltos: mes_usado, imagenes_faltantes, total_tiendas..."""
    if _r is None:
        return
    for k, v in kv.items():
        try:
            json.dumps(v)
        except TypeError:
            v = str(v)
        _r["datos"][k] = v
    _guardar()


def tienda(nombre, ok, productos=0, url="", error_txt=""):
    """Anota una tienda: si su PDF salió bien, cuántos productos lleva y su link."""
    if _r is None:
        return
    _r["tiendas"] = [t for t in _r["tiendas"] if t["n"] != str(nombre)]
    _r["tiendas"].append({"n": str(nombre), "ok": bool(ok), "productos": int(productos or 0),
                          "url": str(url or ""), "error": str(error_txt or "")[:300]})
    _guardar()


def error(msg):
    """Anota el motivo por el que la ejecución no terminó bien."""
    if _r is None:
        return
    _r["error"] = str(msg)[:500]
    _guardar()


def cerrar(resultado_paso):
    """Lo llama el workflow al final del job, con el resultado del paso del script
    (success, failure, cancelled o skipped)."""
    global _r
    try:
        with open(ARCHIVO, encoding="utf-8") as f:
            _r = json.load(f)
    except Exception:
        # El script no llegó a empezar el resumen (falló al instalar, etc.)
        _r = {"v": 1, "proceso": os.environ.get("RESUMEN_PROCESO", ""), "bloque": os.environ.get("RESUMEN_BLOQUE", "0"),
              "run_id": os.environ.get("GITHUB_RUN_ID", ""), "run_attempt": os.environ.get("GITHUB_RUN_ATTEMPT", ""),
              "inicio": "", "fin": "", "resultado": "", "error": "El script no llegó a ejecutarse.",
              "datos": {}, "tiendas": []}
    _r["fin"] = _ahora()
    _r["resultado"] = {"success": "ok", "failure": "error", "cancelled": "cancelado"}.get(resultado_paso, resultado_paso or "error")
    if _r["resultado"] == "ok" and _r.get("error"):
        _r["error"] = ""  # un reintento que salió bien borra el error del intento anterior
    _guardar()


if __name__ == "__main__":
    if len(sys.argv) >= 3 and sys.argv[1] == "cerrar":
        try:
            cerrar(sys.argv[2])
            print(f"Resumen cerrado: {_r['resultado']} · {len(_r['tiendas'])} tiendas")
        except Exception:
            traceback.print_exc()
    else:
        print("Uso: python resumen.py cerrar <success|failure|cancelled>")
