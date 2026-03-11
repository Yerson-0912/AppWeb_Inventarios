import tempfile
import pandas as pd
from pathlib import Path
from ReportStock_main import main as main_mod


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

    uniques = main_mod.obtener_unicos(p, 'BODEGA')
    assert 'PRINCIPAL' in uniques
    assert 'B1' in uniques

    agotados = main_mod.obtener_agotados_por_bodega(p, cantidad_minima=3)
    # B1 should list R1 (principal has 2 <= 3) and R2 as principal 0
    assert 'B1' in agotados
    refs_b1 = {r['referencia'] for r in agotados['B1']}
    assert 'R1' in refs_b1

    # B2 should not include R3 because principal has 10 > 3
    refs_b2 = {r['referencia'] for r in agotados.get('B2', [])}
    assert 'R3' not in refs_b2
