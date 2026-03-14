import tempfile
import pandas as pd
from pathlib import Path
import sys

ROOT_DIR = Path(__file__).resolve().parents[1]
if str(ROOT_DIR) not in sys.path:
    sys.path.insert(0, str(ROOT_DIR))

import reportstock_core as core_mod


def make_excel(tmp_path, df, suffix='.xlsx'):
    p = tmp_path / f"sample{suffix}"
    df.to_excel(p, index=False)
    return p


def test_obtener_unicos_and_agotados(tmp_path):
    df = pd.DataFrame([
        {'BODEGA': 'PRINCIPAL', 'REFERENCIA': 'R1', 'CANTIDAD': 2, 'NOMBRE DEL GRUPO': 'G1'},
        {'BODEGA': 'B1', 'REFERENCIA': 'R1', 'CANTIDAD': 5, 'NOMBRE DEL GRUPO': 'G1'},
        {'BODEGA': 'B1', 'REFERENCIA': 'R2', 'CANTIDAD': 1, 'NOMBRE DEL GRUPO': 'G2'},
        {'BODEGA': 'PRINCIPAL', 'REFERENCIA': 'R3', 'CANTIDAD': 10, 'NOMBRE DEL GRUPO': 'G3'},
        {'BODEGA': 'B2', 'REFERENCIA': 'R3', 'CANTIDAD': 2, 'NOMBRE DEL GRUPO': 'G3'},
    ])

    p = make_excel(tmp_path, df)

    uniques = core_mod.obtener_unicos(p, 'BODEGA')
    assert 'PRINCIPAL' in uniques
    assert 'B1' in uniques

    agotados = core_mod.obtener_agotados_por_bodega(p, cantidad_minima=3)
    # B1 should list R1 (principal has 2 <= 3) and R2 as principal 0
    assert 'B1' in agotados
    refs_b1 = {r['referencia'] for r in agotados['B1']}
    assert 'R1' in refs_b1

    # B2 should not include R3 because principal has 10 > 3
    refs_b2 = {r['referencia'] for r in agotados.get('B2', [])}
    assert 'R3' not in refs_b2


def test_referencia_en_nueva_coleccion_normaliza_formato():
    assert core_mod.referencia_en_nueva_coleccion('  l1290   c01 ')
    assert core_mod.referencia_en_nueva_coleccion('OB070 C1')
    assert core_mod.referencia_en_nueva_coleccion('OB070 C01')
    assert core_mod.referencia_en_nueva_coleccion('OB077 C04')
    assert not core_mod.referencia_en_nueva_coleccion('NO EXISTE')


def test_construir_configuracion_resaltado_usa_paleta_por_defecto():
    configuracion = core_mod.construir_configuracion_resaltado(True, 'Paleta inexistente')

    assert configuracion['habilitado'] is True
    assert configuracion['nombre_paleta'] == core_mod.DEFAULT_PALETA_RESALTADO
    assert configuracion['fondo_hex'] == core_mod.PALETAS_RESALTADO[core_mod.DEFAULT_PALETA_RESALTADO]['fondo']
