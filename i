En el **archivo que te entregué** (`phisycal_risk_penultimo_fen.py`) ya está todo insertado. Es tu código padre completo + las adiciones.

Si quieres aplicarlo **manualmente** a tu original, son exactamente dos lugares:

---

**PARTE 1 — Al final del archivo** (después de la última línea `_HTML_TEMPLATE = r"""...`), pega todo el bloque nuevo desde:

```python
# ==============================================================================
# 11) FEN APOYO — MAPA DEPARTAMENTAL + TABLA INTERACTIVA
# ==============================================================================
```

hasta el final.

---

**PARTE 2 — Dentro de `main()`**, busca este bloque que ya existe:

```python
    # ====== EXTENSIÓN ACTIVOS FÍSICOS ======
    try:
        generar_dashboard_activos(...)
    except Exception as e:
        log(...)

    log("=" * 60)          ← aquí termina el main actual
    log("PROCESO COMPLETADO")
    return xlsx_out, html_out
```

Y agrega **entre** el bloque de activos y el `log("=")`:

```python
    # ====== CHART FEN APOYO ======
    try:
        _generar_fen_chart(
            risk_path=risk_path,
            out_dir=OUT_DIR,
            districts=districts,
            exp_clase=exp_clase,
            aum_cartera=aum_cartera,
            emp_sector=emp_sector,
            geojson=geojson,
            scores=scores,
            stamp=stamp,
        )
    except Exception as _e_fen:
        log(f"FEN chart no generado (no afecta el resto): {_e_fen}", "WARN")
```

---

Pero lo más simple es usar directamente el archivo entregado — es tu código padre sin ninguna modificación, con esas dos cosas ya en su lugar.
