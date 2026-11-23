"""Prueba la busqueda SIN microfono: escribe lo que 'oyo' el programa y mira que hace.
Uso:  python probar_busqueda.py          (o: python probar_busqueda.py "cuaderno kinota cosido")
"""
import os
import sys

from buscador_voz import Buscador, SENSIBILIDAD, decidir_coincidencia, leer_alias

AQUI = os.path.dirname(os.path.abspath(__file__))


def cargar():
    import openpyxl
    wb = openpyxl.load_workbook(os.path.join(AQUI, "similitud.xlsx"), read_only=True, data_only=True)
    prods = [str(r[0]).strip() for r in wb.worksheets[0].iter_rows(max_col=1, values_only=True) if r and r[0]]
    return Buscador(prods, leer_alias(os.path.join(AQUI, "alias.txt")))


def probar(b, frase):
    r = decidir_coincidencia(b, frase, SENSIBILIDAD["normal"])
    print(f"\n  oido     : {frase}")
    print(f"  escribe  : {r['texto'] or r['corregido']}   [{r['modo']} {int(r['score'] * 100)}%]")
    for a in r["alts"][:5]:
        print("     -", a)


if __name__ == "__main__":
    b = cargar()
    if len(sys.argv) > 1:
        probar(b, " ".join(sys.argv[1:]))
    else:
        print("Escribe lo que oyo el programa (vacio para salir).")
        while True:
            f = input("\n> ").strip()
            if not f:
                break
            probar(b, f)
