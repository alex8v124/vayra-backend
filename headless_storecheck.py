import pandas as pd
import numpy as np
import openpyxl as xl
import xlsxwriter
import os
import sys
import json
import re

def generar_storecheck(config_path):
    with open(config_path, 'r', encoding='utf-8') as f:
        config = json.load(f)

    base_file = config.get("base_file")
    crono_data = config.get("cronograma_data", [])
    informe_ant_file = config.get("informe_anterior_file")
    promocion_elegida = config.get("promocion_elegida", "Bolivar")
    numero_storecheck = config.get("numero_storecheck", "1ER")
    skus_seleccionados = config.get("skus_seleccionados", [])
    mapeo_mercados = config.get("mapeo_mercados", {})
    output_file = config.get("output_file", "output.xlsx")

    try:
        df = pd.read_excel(base_file)
    except Exception as e:
        print(json.dumps({"error": f"Error leyendo base: {e}"}))
        sys.exit(1)

    df_crono = pd.DataFrame(crono_data) if crono_data else None

    # Filtrar por promoción
    columna = "Act. Promocional"
    if columna not in df.columns:
        print(json.dumps({"error": f"La columna '{columna}' no existe en el excel base"}))
        sys.exit(1)

    df_f = df[df[columna] == promocion_elegida]
    
    seleccionados = list(mapeo_mercados.keys())
    if not seleccionados:
        seleccionados = df_f["PDV_nombre"].dropna().unique().tolist()
    
    df_modificado = df_f[df_f["PDV_nombre"].isin(seleccionados)]
    df_modificado_v2 = df_modificado[df_modificado["PRODUCTO"].notna()].copy()

    # (Lógica de inyectar informe anterior omitida por brevedad en esta versión headless básica, 
    # se asume que si quieren inyectar, se mapea igual. Pero añadiremos soporte si es necesario).
    _df_informe_ant = None
    if informe_ant_file and os.path.exists(informe_ant_file):
        try:
            xl_inf = pd.ExcelFile(informe_ant_file)
            nombre_hoja_real = next((sheet for sheet in xl_inf.sheet_names if sheet.strip().upper() == "STOCK SIN ACTIVIDAD"), None)
            if nombre_hoja_real:
                df_sa = pd.read_excel(informe_ant_file, sheet_name=nombre_hoja_real, header=[0, 1])
                new_cols = []
                for c in df_sa.columns:
                    if isinstance(c, tuple) and (str(c[1]).strip() == '' or str(c[1]).lower().startswith('unnamed')):
                        new_cols.append(str(c[0]).strip())
                    else:
                        new_cols.append(c)
                df_sa.columns = new_cols
                _df_informe_ant = df_sa
        except Exception:
            pass

    # Inyección simplificada
    if _df_informe_ant is not None:
        def _fc(df_, n): return next((c for c in df_.columns if (c[0] if isinstance(c, tuple) else c) == n), None)
        _ic_merc = _fc(_df_informe_ant, 'MERCADO')
        _ic_cli = _fc(_df_informe_ant, 'NOMBRE CLIENTE')
        _ic_puesto = _fc(_df_informe_ant, 'N° de puesto')
        _ic_fecha = _fc(_df_informe_ant, 'FECHA')
        _ic_sf_cols = [(c, c[1]) for c in _df_informe_ant.columns if isinstance(c, tuple) and c[0] == 'STOCK FINAL' and c[1] != '']
        _ic_si_cols = [(c, c[1]) for c in _df_informe_ant.columns if isinstance(c, tuple) and c[0] == 'STOCK INICIAL' and c[1] != '']
        _ic_sf_dict = {sn: sc for sc, sn in _ic_sf_cols}
        _ic_si_dict = {sn: sc for sc, sn in _ic_si_cols}

        _mercados_a_inyectar = set()
        for mc_crono, info in mapeo_mercados.items():
            if info.get("DESDE_INFORME"):
                _mercados_a_inyectar.add(str(mc_crono).strip())
        
        if _ic_merc is not None and _mercados_a_inyectar:
            _df_inf_det = _df_informe_ant.copy()
            if _ic_cli is not None:
                _df_inf_det = _df_inf_det[_df_inf_det[_ic_cli].apply(lambda x: isinstance(x, str) and x.strip() != '')]
            _df_inf_det = _df_inf_det[_df_inf_det[_ic_merc].astype(str).str.strip().isin(_mercados_a_inyectar)]
            
            if len(_df_inf_det) > 0:
                _rows_inf = []
                for _, _inf_row in _df_inf_det.iterrows():
                    _merc_val   = str(_inf_row[_ic_merc]).strip() if _ic_merc else ''
                    _cli_val    = str(_inf_row[_ic_cli]).strip()  if _ic_cli  else ''
                    _puesto_val = str(_inf_row[_ic_puesto]).strip() if _ic_puesto else ''
                    _fecha_val  = _inf_row[_ic_fecha] if _ic_fecha else pd.NaT

                    for _sc, _sn in _ic_sf_cols:
                        _sf = pd.to_numeric(_inf_row.get(_sc, np.nan), errors='coerce')
                        _si_col = _ic_si_dict.get(_sn)
                        _si = pd.to_numeric(_inf_row.get(_si_col, np.nan), errors='coerce') if _si_col else np.nan
                        if pd.isna(_sf) and pd.isna(_si):
                            continue
                        _rows_inf.append({
                            "PDV_nombre":        _merc_val,
                            "FECHA":             _fecha_val,
                            "PUESTO DE MERCADO": f"{_puesto_val}: {_cli_val}",
                            "Act. Promocional":  promocion_elegida,
                            "PRODUCTO":          _sn,
                            "STOCK FINAL":       0 if pd.isna(_sf) else _sf,
                            "STOCK INICIAL":     0 if pd.isna(_si) else _si,
                        })

                if _rows_inf:
                    _df_inf_inject = pd.DataFrame(_rows_inf)
                    df_modificado_v2 = pd.concat([df_modificado_v2, _df_inf_inject], ignore_index=True)


    if skus_seleccionados:
        df_modificado_v2 = df_modificado_v2[df_modificado_v2["PRODUCTO"].isin(skus_seleccionados)]

    # === INYECTAR CRONOGRAMA ===
    if df_crono is not None and not df_crono.empty and "fecha" in df_crono.columns:
        _rows_crono = []
        for _, _cr_row in df_crono.iterrows():
            _c_fecha = _cr_row.get("fecha")
            _c_merc = str(_cr_row.get("MERCADO", "")).strip()
            _c_puesto = str(_cr_row.get("nroPuesto", "")).strip()
            _c_encarg = str(_cr_row.get("encargado", "")).strip()
            _c_puesto_full = f"{_c_puesto}: {_c_encarg}"

            if not _c_merc:
                continue

            for _sku in (skus_seleccionados if skus_seleccionados else ["DUMMY_SKU"]):
                _rows_crono.append({
                    "PDV_nombre":        _c_merc, # Usamos el mercado como PDV simulado temporalmente
                    "FECHA":             _c_fecha,
                    "PUESTO DE MERCADO": _c_puesto_full,
                    "Act. Promocional":  promocion_elegida,
                    "PRODUCTO":          _sku,
                    "STOCK FINAL":       np.nan,
                    "STOCK INICIAL":     np.nan,
                })
        
        if _rows_crono:
            _df_crono_inject = pd.DataFrame(_rows_crono)
            df_modificado_v2 = pd.concat([df_modificado_v2, _df_crono_inject], ignore_index=True)
            # Removemos el DUMMY_SKU si se agregó
            df_modificado_v2 = df_modificado_v2[df_modificado_v2["PRODUCTO"] != "DUMMY_SKU"]

    required_cols = ["PDV_nombre", "FECHA", "PUESTO DE MERCADO", "Act. Promocional", "PRODUCTO", "STOCK FINAL", "STOCK INICIAL"]
    missing = [c for c in required_cols if c not in df_modificado_v2.columns]
    if missing:
        print(json.dumps({"error": f"Faltan columnas en el excel base: {missing}"}))
        sys.exit(1)

    df_modificado_v3 = df_modificado_v2[required_cols]
    df_modificado_v4 = df_modificado_v3.groupby(["PDV_nombre", "FECHA", "PUESTO DE MERCADO", "Act. Promocional", "PRODUCTO"])[["STOCK INICIAL", "STOCK FINAL"]].sum().reset_index()

    df_modificado_v5 = df_modificado_v4.copy()
    df_modificado_v5 = df_modificado_v5.rename(columns={'PUESTO DE MERCADO': 'NOMBRE CLIENTE'})
    split_cols = df_modificado_v5['NOMBRE CLIENTE'].str.split(':', n=1, expand=True)
    df_modificado_v5['N° de puesto'] = split_cols[0].str.strip()
    df_modificado_v5['NOMBRE CLIENTE'] = split_cols[1].str.strip() if 1 in split_cols.columns else df_modificado_v5['NOMBRE CLIENTE']

    df_modificado_v7 = df_modificado_v5.copy()
    df_modificado_v7["PROVINCIA"] = "N/A"
    df_modificado_v7["DEX"] = "N/A"

    def _get_mercado_nombre(pdv):
        pdv_str = str(pdv).strip()
        if pdv_str in mapeo_mercados:
            return mapeo_mercados[pdv_str].get("MERCADO_CRONO", pdv_str)
        return pdv_str

    df_modificado_v7["MERCADO"] = df_modificado_v7["PDV_nombre"].map(_get_mercado_nombre)

    def _clean_val(v, fallback="N/A"):
        s = str(v).strip()
        return fallback if s.upper() in ("NAN", "NONE", "N/A", "") else s

    _prov_dex_map = {}
    for _pdv_m, _info_m in mapeo_mercados.items():
        _pv = _clean_val(_info_m.get("Ciudad", "N/A"))
        _dv = _clean_val(_info_m.get("DEX",    "N/A"))
        _prov_dex_map[str(_pdv_m).strip()] = (_pv, _dv)
        _mc = str(_info_m.get("MERCADO_CRONO", "")).strip()
        if _mc:
            _prov_dex_map[_mc] = (_pv, _dv)
            
    # Also add from crono_data directly if available
    if df_crono is not None and not df_crono.empty:
        for _, _cr in df_crono.iterrows():
            _c_merc = str(_cr.get("MERCADO", "")).strip()
            _c_ciu = _clean_val(_cr.get("Ciudad", "N/A"))
            _c_dex = _clean_val(_cr.get("DEX", "N/A"))
            if _c_merc and _c_merc not in _prov_dex_map:
                _prov_dex_map[_c_merc] = (_c_ciu, _c_dex)

    def _get_prov(pdv): return _prov_dex_map.get(str(pdv).strip(), ("N/A", "N/A"))[0]
    def _get_dex(pdv): return _prov_dex_map.get(str(pdv).strip(), ("N/A", "N/A"))[1]

    df_modificado_v7["PROVINCIA"] = df_modificado_v7["PDV_nombre"].map(_get_prov).fillna("N/A")
    df_modificado_v7["DEX"]       = df_modificado_v7["PDV_nombre"].map(_get_dex).fillna("N/A")

    if df_crono is not None and not df_crono.empty and "MERCADO" in df_crono.columns:
        _crono_order = {mc: i for i, mc in enumerate(df_crono["MERCADO"].dropna().tolist())}
        def _get_crono_pos(m): return _crono_order.get(m, 9999)
        _mercados_en_data = df_modificado_v7["MERCADO"].unique().tolist()
        _orden_mercados = sorted(_mercados_en_data, key=lambda m: _get_crono_pos(m))
        _orden_map = {m: i for i, m in enumerate(_orden_mercados)}
        df_modificado_v7["_orden_crono"] = df_modificado_v7["MERCADO"].map(_orden_map)
    else:
        df_modificado_v7["_orden_crono"] = 0
        _orden_mercados = None

    pv = pd.pivot_table(
        data=df_modificado_v7,
        index=["FECHA", "PROVINCIA", "DEX", 'MERCADO', 'NOMBRE CLIENTE', 'N° de puesto'],
        columns='PRODUCTO',
        values=["STOCK INICIAL", 'STOCK FINAL'],
        aggfunc='sum',
        fill_value=0
    )

    if isinstance(pv.columns, pd.MultiIndex):
        cols = pv.columns.to_list()
        level0_order = ['STOCK INICIAL', 'STOCK FINAL']
        sorted_cols = sorted(cols, key=lambda c: (level0_order.index(c[0]) if c[0] in level0_order else 99, c[1]))
        pv = pv[sorted_cols]

    pv = pv.reset_index()

    new_cols = []
    for c in pv.columns:
        if isinstance(c, tuple) and c[1] == '':
            new_cols.append(c[0])
        else:
            new_cols.append(c)
    pv.columns = new_cols

    cols_pv = list(pv.columns)
    nombre_idx = next((i for i, c in enumerate(cols_pv) if (c[0] if isinstance(c, tuple) else c) == 'NOMBRE CLIENTE'), None)
    npuesto_col = next((c for c in cols_pv if (c[0] if isinstance(c, tuple) else c) == 'N° de puesto'), None)
    if nombre_idx is not None and npuesto_col is not None:
        cols_pv.remove(npuesto_col)
        cols_pv.insert(nombre_idx + 1, npuesto_col)
        pv = pv[cols_pv]

    cols_numericas = [c for c in pv.select_dtypes(include='number').columns]
    pv[cols_numericas] = pv[cols_numericas].replace(0, np.nan)

    if _orden_mercados:
        _pv_orden = {m: i for i, m in enumerate(_orden_mercados)}
        pv["_ord"] = pv["MERCADO"].map(_pv_orden).fillna(9999)
        pv = pv.sort_values(by=["_ord", "N° de puesto"]).drop(columns=["_ord"]).reset_index(drop=True)
    else:
        pv = pv.sort_values(by=['MERCADO', 'N° de puesto']).reset_index(drop=True)

    numeration = pv.groupby('MERCADO').cumcount() + 1
    pv.insert(0, 'N°', numeration)

    cols_pv2 = list(pv.columns)
    sf_idx = next((i for i, c in enumerate(cols_pv2) if isinstance(c, tuple) and c[0] == 'STOCK FINAL'), None)
    if sf_idx is not None:
        pv.insert(sf_idx, '_separador_', np.nan)

    def _ciudad_sort_key(ciudad):
        c = str(ciudad).strip().upper()
        return (0, c) if c == 'LIMA' else (1, c)

    _skus_sf   = [c for c in pv.columns if isinstance(c, tuple) and c[0] == 'STOCK FINAL' and c[1] != '']
    _sku_names = [c[1] for c in _skus_sf]

    _resumen_rows = []
    for _mercado, _grp in pv.groupby('MERCADO'):
        _prov    = _grp['PROVINCIA'].iloc[0]
        _puestos = _grp['N° de puesto'].count()
        _row = {
            'CIUDAD':  _prov if str(_prov).strip().upper() != 'N/A' else '',
            'MERCADO': _mercado,
            'PUESTOS DE MERCADO': _puestos,
        }
        for _sc, _sn in zip(_skus_sf, _sku_names):
            _col_data = pd.to_numeric(_grp[_sc], errors='coerce')
            _s = _col_data.sum(min_count=1)
            _row[_sn] = _s if not pd.isna(_s) else 0
        _resumen_rows.append(_row)

    resumen_stock_final = pd.DataFrame(_resumen_rows)

    if _orden_mercados:
        resumen_stock_final["_ord"] = resumen_stock_final["MERCADO"].map(
            {m: i for i, m in enumerate(_orden_mercados)}).fillna(9999)
        resumen_stock_final["_ciudad_key"] = resumen_stock_final["CIUDAD"].map(_ciudad_sort_key)
        resumen_stock_final = resumen_stock_final.sort_values(
            ["_ciudad_key", "_ord"]).drop(columns=["_ord", "_ciudad_key"]).reset_index(drop=True)
    else:
        resumen_stock_final["_ciudad_key"] = resumen_stock_final["CIUDAD"].map(_ciudad_sort_key)
        resumen_stock_final = resumen_stock_final.sort_values(
            ["_ciudad_key", "MERCADO"]).drop(columns=["_ciudad_key"]).reset_index(drop=True)

    resumen_stock_final['STOCK TOTAL'] = resumen_stock_final[_sku_names].sum(axis=1)

    _total_row = {'CIUDAD': '', 'MERCADO': 'TOTAL',
                  'PUESTOS DE MERCADO': resumen_stock_final['PUESTOS DE MERCADO'].sum()}
    for _sn in _sku_names:
        _total_row[_sn] = resumen_stock_final[_sn].sum()
    _total_row['STOCK TOTAL'] = resumen_stock_final['STOCK TOTAL'].sum()
    resumen_stock_final = pd.concat([resumen_stock_final, pd.DataFrame([_total_row])], ignore_index=True)

    df_modificado_v8 = df_modificado_v7.copy()
    df_modificado_v8["STOCK INICIAL"] = df_modificado_v8["STOCK INICIAL"].replace(0, np.nan)
    df_modificado_v8["STOCK FINAL"]   = df_modificado_v8["STOCK FINAL"].replace(0, np.nan)
    df_modificado_v8_drop = df_modificado_v8[["MERCADO", "PROVINCIA", "NOMBRE CLIENTE"]].drop_duplicates()
    df_modificado_v10 = df_modificado_v8_drop.groupby(["PROVINCIA", "MERCADO"]).size().reset_index(name='PUESTOS DE MERCADO')
    df_modificado_v10["PUESTOS ATENDIDOS POR ALICORP"] = df_modificado_v10["PUESTOS DE MERCADO"]
    df_modificado_v10["PRESENCIA  DEL PRODUCTO"]       = df_modificado_v10["PUESTOS DE MERCADO"]
    df_modificado_v10["COBERTURA TOTAL (Puestos Atendidos por Alicorp)"] = '100%'

    MARCAS_DICT = {
        "ALA": "Alacena", "ALP": "Alpesa", "AMA": "Amaras", "ANG": "Angel",
        "AVA": "Aval", "BFL": "Blancaflor", "BOL": "Bolivar", "DEN": "Dento",
        "DON": "Don Vittorio", "DVI": "Dvi", "LAV": "Lavaggi", "LIF": "Life",
        "MAN": "Manty", "MAR": "Marsella", "MIR": "Mirasol", "MMC": "Multimarca",
        "NIC": "Nicolini", "OPA": "Opal", "PAT": "Patito", "PRI": "Primor",
        "PUR": "Chocolate Puro", "SAP": "Sapolio", "SAY": "Sayon",
        "SDO": "Sello De Oro", "TRO": "Trome", "UMS": "Umsha", "VIC": "Victoria",
    }
    
    _nombre_corto = promocion_elegida
    _codigos = re.findall(r'(?<![A-Z])([A-Z]{3})(?![A-Z])', promocion_elegida)
    for _cod in _codigos:
        if _cod in MARCAS_DICT:
            _nombre_corto = MARCAS_DICT[_cod]
            break

    columna_promocion_interna = f"Vende {promocion_elegida} ¿Si/No?"
    columna_promocion_display = f"Vende {_nombre_corto} ¿Si/No?"

    skus_desde_pv = [c for c in pv.columns if isinstance(c, tuple) and c[1] != '']
    productos_unicos = list(dict.fromkeys([c[1] for c in skus_desde_pv]))
    marca_storecheck = _nombre_corto

    _cols_id = ['FECHA', 'PROVINCIA', 'DEX', 'MERCADO', 'NOMBRE CLIENTE', 'N° de puesto']
    pv_tb2 = pv[[c for c in pv.columns if c in _cols_id]].copy()
    pv_tb2 = pv_tb2.rename(columns={'N° de puesto': 'N° Puesto'})

    pv_tb2[columna_promocion_interna] = "1"
    pv_tb2["Alicorp"] = 1
    pv_tb2["Otros"] = ""
    pv_tb2["Presencia en Puestos Alicorp"] = "1"
    pv_tb2["Presencia Producto"] = "1"

    _skus_cols = [c for c in pv.columns if isinstance(c, tuple) and c[0] == 'STOCK FINAL' and c[1] != '']
    for _sku_col in _skus_cols:
        _sku_name = _sku_col[1]
        pv_tb2[_sku_name] = pv[_sku_col].apply(lambda x: 1 if pd.notna(x) and x != 0 else np.nan)

    if _orden_mercados:
        _pvt_orden = {m: i for i, m in enumerate(_orden_mercados)}
        pv_tb2["_ord"] = pv_tb2["MERCADO"].map(_pvt_orden).fillna(9999)
        pv_tb2 = pv_tb2.sort_values(by=["_ord", "N° Puesto"]).drop(columns=["_ord"]).reset_index(drop=True)
    else:
        pv_tb2 = pv_tb2.sort_values(by=['MERCADO', 'N° Puesto']).reset_index(drop=True)

    pv_tb2.insert(0, 'N°', pv_tb2.groupby('MERCADO').cumcount() + 1)

    _cols_fijas_sc = ['N°', 'FECHA', 'PROVINCIA', 'MERCADO', 'NOMBRE CLIENTE', 'N° Puesto', columna_promocion_interna, 'Alicorp', 'Otros', 'Presencia en Puestos Alicorp', 'Presencia Producto']
    _cols_excluir  = ['DEX', 'N°']
    _skus_sc = [c for c in pv_tb2.columns if c not in _cols_fijas_sc and c not in _cols_excluir and not isinstance(c, tuple)]
    _cols_order_sc = [c for c in _cols_fijas_sc if c in pv_tb2.columns] + _skus_sc
    pv_tb2 = pv_tb2[_cols_order_sc]

    _c1_rows = []
    for _merc, _grp in pv_tb2.groupby('MERCADO'):
        _prov = _grp['PROVINCIA'].iloc[0] if 'PROVINCIA' in _grp.columns else ''
        _prov = '' if str(_prov).strip().upper() == 'N/A' else _prov
        _puestos = _grp['N° Puesto'].count()
        _c1_rows.append({
            'CIUDAD': _prov,
            'MERCADO': _merc,
            'PUESTOS DE MERCADO': _puestos,
            'PUESTOS ATENDIDOS POR ALICORP': _puestos,
            'PRESENCIA  DEL PRODUCTO': _puestos,
            'COBERTURA TOTAL (Puestos Atendidos por Alicorp)': '100%',
        })
    df_modificado_v10 = pd.DataFrame(_c1_rows)

    if _orden_mercados:
        df_modificado_v10["_ord"] = df_modificado_v10["MERCADO"].map({m: i for i, m in enumerate(_orden_mercados)}).fillna(9999)
        df_modificado_v10["_ciudad_key"] = df_modificado_v10["CIUDAD"].map(_ciudad_sort_key)
        df_modificado_v10 = df_modificado_v10.sort_values(["_ciudad_key", "_ord"]).drop(columns=["_ord", "_ciudad_key"]).reset_index(drop=True)
    else:
        df_modificado_v10["_ciudad_key"] = df_modificado_v10["CIUDAD"].map(_ciudad_sort_key)
        df_modificado_v10 = df_modificado_v10.sort_values(["_ciudad_key", "MERCADO"]).drop(columns=["_ciudad_key"]).reset_index(drop=True)

    _parte_final = marca_storecheck
    directorio = os.path.dirname(os.path.abspath(output_file))
    backend_dir = os.path.dirname(os.path.abspath(__file__))
    def _col_letter(n):
        if n <= 0:
            return 'A'
        s = ""
        while n > 0:
            n, r = divmod(n - 1, 26)
            s = chr(65 + r) + s
        return s
    gcl = _col_letter

    hojas = {
        "STOCK SIN ACTIVIDAD": pv,
        "STORECHECK":          pv_tb2,
        "RESUMEN":             df_modificado_v10,
    }
    _orden_pestanas = ["Resumen", "Storecheck", "Stock sin actividad"]
    
    for nombre_hoja, df in hojas.items():
        if df is None:
            print(f"Error: DataFrame para hoja '{nombre_hoja}' es None")
            sys.exit()
        if len(df) == 0:
            print(f"Advertencia: DataFrame para hoja '{nombre_hoja}' está vacío")
    
    # ─── Generar Excel con formato personalizado ──────────────────────────────────
    # _nombre_archivo = f"{numero_storecheck.lower()} Storecheck {_parte_final}.xlsx"
    # output_file = os.path.join(directorio, _nombre_archivo)
    
    
    print(f"Generando archivo Excel con formato: {output_file}")
    
    _sc_row_ranges = {}
    _sa_row_ranges = {}
    
    def _merge_ciudad_col(worksheet, row_start, row_end, col_idx, df_ref, fmt):
        if row_start > row_end or col_idx < 0:
            return
        col_name = 'CIUDAD'
        if col_name not in df_ref.columns:
            return
        valores = df_ref[df_ref.get('MERCADO', pd.Series([''] * len(df_ref))) != 'TOTAL'][col_name].tolist()
        if not valores:
            return
        i = 0
        current_excel_row = row_start
        while i < len(valores) and current_excel_row <= row_end:
            ciudad_val = str(valores[i]) if pd.notna(valores[i]) else ''
            count = 1
            while (i + count < len(valores) and
                   str(valores[i + count]) == ciudad_val and
                   current_excel_row + count <= row_end):
                count += 1
            if count > 1:
                worksheet.merge_range(current_excel_row, col_idx,
                                       current_excel_row + count - 1, col_idx,
                                       ciudad_val, fmt)
            else:
                worksheet.write(current_excel_row, col_idx, ciudad_val, fmt)
            i += count
            current_excel_row += count
    
    with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
        workbook = writer.book
    
        _sc_total_rows     = {}
        _sa_total_rows     = {}
        _sc_col_indices    = {}
        _sa_col_indices    = {}
        _sa_sku_cols       = {}
        _sa_sku_cols_sf    = {}
        _sa_detail_row_map = {}
        _sc_col_map        = {}
        _sa_col_map        = {}
    
        header_format = workbook.add_format({
            'bold': True, 'text_wrap': True, 'valign': 'vcenter', 'align': 'center',
            'fg_color': '#1F497D', 'font_color': 'white', 'border': 1,
            'font_size': 12, 'font_name': 'Aptos'
        })
        rotated_header_format = workbook.add_format({
            'bold': False, 'text_wrap': True, 'align': 'center', 'valign': 'vcenter',
            'fg_color': '#FFFFFF', 'font_color': '#000000', 'border': 1,
            'rotation': 90, 'font_size': 12, 'font_name': 'Aptos'
        })
        flat_sku_header_format = workbook.add_format({
            'bold': True, 'text_wrap': True, 'align': 'center', 'valign': 'vcenter',
            'fg_color': '#1F497D', 'font_color': 'white', 'border': 1,
            'font_size': 12, 'font_name': 'Aptos'
        })
        date_format = workbook.add_format({
            'num_format': 'dd/mm/yyyy', 'align': 'center', 'border': 1,
            'valign': 'vcenter', 'font_size': 14, 'font_name': 'Aptos'
        })
        body_format = workbook.add_format({
            'border': 1, 'align': 'center', 'valign': 'vcenter',
            'font_size': 14, 'font_name': 'Aptos'
        })
        number_format = workbook.add_format({
            'border': 1, 'align': 'center', 'valign': 'vcenter',
            'num_format': '#,##0', 'font_size': 14, 'font_name': 'Aptos'
        })
        sa_header_format = workbook.add_format({
            'bold': True, 'text_wrap': True, 'valign': 'vcenter', 'align': 'center',
            'fg_color': '#1F497D', 'font_color': 'white', 'border': 1,
            'font_size': 14, 'font_name': 'Aptos'
        })
        sa_group_header_format = workbook.add_format({
            'bold': True, 'text_wrap': True, 'valign': 'vcenter', 'align': 'center',
            'fg_color': '#1F497D', 'font_color': 'white', 'border': 1,
            'font_size': 14, 'font_name': 'Aptos'
        })
        sa_sku_format = workbook.add_format({
            'bold': False, 'text_wrap': True, 'align': 'center', 'valign': 'vcenter',
            'fg_color': '#FFFFFF', 'font_color': '#000000', 'border': 1,
            'rotation': 90, 'font_size': 14, 'font_name': 'Aptos'
        })
        sa_total_format = workbook.add_format({
            'bold': False, 'border': 1, 'align': 'center', 'valign': 'vcenter',
            'font_size': 14, 'fg_color': '#1F497D', 'font_color': 'white', 'font_name': 'Aptos'
        })
        sa_total_number_format = workbook.add_format({
            'bold': False, 'border': 1, 'align': 'center', 'valign': 'vcenter',
            'num_format': '#,##0', 'font_size': 14,
            'fg_color': '#1F497D', 'font_color': 'white', 'font_name': 'Aptos'
        })
        sa_total_date_format = workbook.add_format({
            'bold': False, 'num_format': 'dd/mm/yyyy', 'align': 'center', 'border': 1,
            'valign': 'vcenter', 'font_size': 14,
            'fg_color': '#1F497D', 'font_color': 'white', 'font_name': 'Aptos'
        })
    
        SA_WIDTHS = {
            'N°': 7.3, 'FECHA': 21.3, 'PROVINCIA': 21.3, 'DEX': 21.3,
            'MERCADO': 74.3, 'NOMBRE CLIENTE': 69.0, 'N° de puesto': 21.1
        }
        SA_SKU_WIDTH = 10.9
        SA_SEP_WIDTH = 2
    
        sc_body_format = workbook.add_format({
            'border': 1, 'align': 'center', 'valign': 'vcenter',
            'font_size': 12, 'font_name': 'Aptos'
        })
        sc_number_format = workbook.add_format({
            'border': 1, 'align': 'center', 'valign': 'vcenter',
            'num_format': '#,##0', 'font_size': 12, 'font_name': 'Aptos'
        })
        sc_date_format = workbook.add_format({
            'num_format': 'dd/mm/yyyy', 'align': 'center', 'border': 1,
            'valign': 'vcenter', 'font_size': 12, 'font_name': 'Aptos'
        })
        hdr_presencia_alicorp = workbook.add_format({
            'bold': True, 'text_wrap': True, 'align': 'center', 'valign': 'vcenter',
            'fg_color': '#404040', 'font_color': 'white', 'border': 1,
            'rotation': 90, 'font_size': 12, 'font_name': 'Aptos'
        })
        data_presencia_alicorp = workbook.add_format({
            'border': 1, 'align': 'center', 'valign': 'vcenter',
            'fg_color': '#808080', 'font_size': 12, 'font_name': 'Aptos'
        })
        hdr_presencia_producto = workbook.add_format({
            'bold': True, 'text_wrap': True, 'align': 'center', 'valign': 'vcenter',
            'fg_color': '#D9D9D9', 'font_color': '#000000', 'border': 1,
            'rotation': 90, 'font_size': 12, 'font_name': 'Aptos'
        })
        data_presencia_producto = workbook.add_format({
            'border': 1, 'align': 'center', 'valign': 'vcenter',
            'fg_color': '#D9D9D9', 'font_size': 12, 'font_name': 'Aptos'
        })
        hdr_rotated_sc = workbook.add_format({
            'bold': True, 'text_wrap': True, 'align': 'center', 'valign': 'vcenter',
            'fg_color': '#1F497D', 'font_color': 'white', 'border': 1,
            'rotation': 90, 'font_size': 12, 'font_name': 'Aptos'
        })
        hdr_marca_sc = workbook.add_format({
            'bold': True, 'text_wrap': True, 'align': 'center', 'valign': 'vcenter',
            'fg_color': '#1F497D', 'font_color': 'white', 'border': 1,
            'font_size': 12, 'font_name': 'Aptos'
        })
        hdr_sku_sc = workbook.add_format({
            'bold': False, 'text_wrap': True, 'align': 'center', 'valign': 'vcenter',
            'fg_color': '#FFFFFF', 'font_color': '#000000', 'border': 1,
            'rotation': 90, 'font_size': 14, 'font_name': 'Aptos'
        })
        pct_format_sc = workbook.add_format({
            'bold': False, 'border': 1, 'align': 'center', 'valign': 'vcenter',
            'num_format': '0%', 'font_size': 12, 'font_name': 'Aptos',
            'fg_color': '#003366', 'font_color': 'white'
        })
    
        for sheet_name, df in hojas.items():
            print(f"Procesando hoja: {sheet_name}")
            if df.empty:
                print(f"  Advertencia: La hoja '{sheet_name}' está vacía.")
                continue
    
            _nombres_display = {
                "RESUMEN":             "Resumen",
                "STORECHECK":          "Storecheck",
                "STOCK SIN ACTIVIDAD": "Stock sin actividad",
            }
            worksheet = workbook.add_worksheet(_nombres_display.get(sheet_name, sheet_name))
            worksheet.hide_gridlines(2)
    
            if sheet_name in ("STORECHECK", "STOCK SIN ACTIVIDAD"):
                worksheet.outline_settings(visible=True, symbols_below=True,
                                            symbols_right=True, auto_style=False)
    
            if sheet_name == "STORECHECK":
                worksheet.set_row(0, 25)
                worksheet.set_row(1, 152)
    
                cols_index_sc = ['N°', 'FECHA', 'PROVINCIA', 'DEX', 'MERCADO', 'NOMBRE CLIENTE',
                                 'N° Puesto', columna_promocion_interna, 'Alicorp', 'Otros',
                                 'Presencia en Puestos Alicorp', 'Presencia Producto']
                productos_sc = [c for c in df.columns if c not in cols_index_sc and not isinstance(c, tuple)]
    
                col_num = 0
                simple_cols    = ['N°', 'FECHA', 'PROVINCIA', 'MERCADO', 'NOMBRE CLIENTE',
                                   'N° Puesto', columna_promocion_interna]
                simple_widths  = [8.3, 19.7, 27.6, 112.9, 62.6, 19.1, 16.1]
                simple_display = ['N°', 'FECHA', 'PROVINCIA', 'MERCADO', 'NOMBRE CLIENTE',
                                   'N° Puesto', columna_promocion_display]
                for name, display, w in zip(simple_cols, simple_display, simple_widths):
                    if name in df.columns:
                        worksheet.merge_range(0, col_num, 1, col_num, display, header_format)
                        worksheet.set_column(col_num, col_num, w)
                        col_num += 1
    
                abast_start = col_num
                worksheet.merge_range(0, abast_start, 0, abast_start + 1, 'Abastecido por:', header_format)
                worksheet.write(1, abast_start,     'Alicorp', hdr_rotated_sc)
                worksheet.set_column(abast_start,     abast_start,     10.9)
                worksheet.write(1, abast_start + 1, 'Otros',   hdr_rotated_sc)
                worksheet.set_column(abast_start + 1, abast_start + 1, 10.9)
                col_num += 2
    
                worksheet.merge_range(0, col_num, 1, col_num, 'Presencia en Puestos Alicorp', hdr_presencia_alicorp)
                worksheet.set_column(col_num, col_num, 10.9)
                col_num += 1
    
                worksheet.merge_range(0, col_num, 1, col_num, 'Presencia Producto', hdr_presencia_producto)
                worksheet.set_column(col_num, col_num, 10.9)
                col_num += 1
    
                if productos_sc:
                    sku_start = col_num
                    sku_count = len(productos_sc)
                    if sku_count > 1:
                        worksheet.merge_range(0, sku_start, 0, sku_start + sku_count - 1,
                                              marca_storecheck, hdr_marca_sc)
                    else:
                        worksheet.write(0, sku_start, marca_storecheck, hdr_marca_sc)
                    for sku in productos_sc:
                        worksheet.write(1, col_num, sku, hdr_sku_sc)
                        worksheet.set_column(col_num, col_num, 11.1)
                        col_num += 1
    
                header_offset = 2
                _col_num_sc = 0
                for _cn_sc in simple_cols:
                    if _cn_sc in df.columns:
                        from xlsxwriter.utility import xl_col_to_name
                        _sc_col_map[_cn_sc] = xl_col_to_name(_col_num_sc)
                        _col_num_sc += 1
                _sc_col_map['Alicorp']   = xl_col_to_name(_col_num_sc)
                _sc_col_map['Otros']     = xl_col_to_name(_col_num_sc + 1)
                _sc_col_map['Presencia en Puestos Alicorp'] = xl_col_to_name(_col_num_sc + 2)
                _sc_col_map['Presencia Producto']           = xl_col_to_name(_col_num_sc + 3)
                _sku_col_start_sc = _col_num_sc + 4
                for _si_sc, _sku_sc in enumerate(productos_sc):
                    _sc_col_map[_sku_sc] = xl_col_to_name(_sku_col_start_sc + _si_sc)
    
            else:
                is_pivot = any(isinstance(c, tuple) for c in df.columns)
    
                if sheet_name == "STOCK SIN ACTIVIDAD":
                    worksheet.set_row(0, 25)
                    worksheet.set_row(1, 190)
                    _hdr_main  = sa_header_format
                    _hdr_group = sa_group_header_format
                    _hdr_sku   = sa_sku_format
                else:
                    worksheet.set_row(0, 25)
                    worksheet.set_row(1, 60)
                    _hdr_main  = header_format
                    _hdr_group = header_format
                    _hdr_sku   = rotated_header_format
    
                if not is_pivot:
                    if sheet_name != "RESUMEN":
                        for col_num, value in enumerate(df.columns.values):
                            worksheet.write(0, col_num, value, _hdr_main)
                            col_width = max(len(str(value)) + 2, 10)
                            worksheet.set_column(col_num, col_num, col_width)
                    header_offset = 1
                else:
                    for col_num, col_name in enumerate(df.columns):
                        if col_name == '_separador_':
                            worksheet.write(0, col_num, '', None)
                            worksheet.write(1, col_num, '', None)
                            worksheet.set_column(col_num, col_num,
                                                 SA_SEP_WIDTH if sheet_name == "STOCK SIN ACTIVIDAD" else 2)
                        elif isinstance(col_name, tuple) and len(col_name) > 1 and col_name[0] not in [
                                'N°', 'FECHA', 'PROVINCIA', 'DEX', 'MERCADO', 'NOMBRE CLIENTE',
                                'N° de puesto', 'Alicorp', 'Otros',
                                'Presencia en Puestos Alicorp', 'Presencia Producto']:
                            worksheet.write(0, col_num, col_name[0], _hdr_group)
                            worksheet.write(1, col_num, col_name[1], _hdr_sku)
                            worksheet.set_column(col_num, col_num,
                                                 SA_SKU_WIDTH if sheet_name == "STOCK SIN ACTIVIDAD" else 5)
                        else:
                            display_name = (col_name[0] if isinstance(col_name, tuple)
                                            and len(col_name) > 0 else col_name)
                            worksheet.merge_range(0, col_num, 1, col_num, display_name, _hdr_main)
                            w = (SA_WIDTHS.get(display_name, SA_SKU_WIDTH)
                                 if sheet_name == "STOCK SIN ACTIVIDAD"
                                 else max(len(str(display_name)) + 2, 12))
                            worksheet.set_column(col_num, col_num, w)
    
                    groups = []
                    current_header = None
                    current_start  = None
                    current_count  = 0
                    for col_num, col_name in enumerate(df.columns):
                        if isinstance(col_name, tuple) and len(col_name) == 2 and col_name[1] != '':
                            level0 = col_name[0]
                            if level0 == current_header:
                                current_count += 1
                            else:
                                if current_header is not None:
                                    groups.append((current_header, current_start, current_count))
                                current_header = level0
                                current_start  = col_num
                                current_count  = 1
                        else:
                            if current_header is not None:
                                groups.append((current_header, current_start, current_count))
                                current_header = None
                                current_start  = None
                                current_count  = 0
                    if current_header is not None:
                        groups.append((current_header, current_start, current_count))
    
                    for header, start, count in groups:
                        if count > 1:
                            worksheet.merge_range(0, start, 0, start + count - 1, header, _hdr_group)
                        else:
                            worksheet.write(0, start, header, _hdr_group)
                    header_offset = 2
    
                    from xlsxwriter.utility import xl_col_to_name as _xl_cn
                    for _ci_sa, _cn_sa in enumerate(df.columns):
                        if isinstance(_cn_sa, tuple) and _cn_sa[0] == 'STOCK FINAL' and _cn_sa[1] != '':
                            _sa_col_map[_cn_sa] = _xl_cn(_ci_sa)
                        elif isinstance(_cn_sa, tuple) and _cn_sa[0] == 'STOCK INICIAL' and _cn_sa[1] != '':
                            _sa_col_map[_cn_sa] = _xl_cn(_ci_sa)
                        elif not isinstance(_cn_sa, tuple):
                            _sa_col_map[_cn_sa] = _xl_cn(_ci_sa)
    
                    if sheet_name == "STOCK SIN ACTIVIDAD":
                        _vta_skus   = [c[1] for c in df.columns
                                       if isinstance(c, tuple) and c[0] == 'STOCK FINAL' and c[1] != '']
                        _n_cols_df  = len(df.columns)
                        _vta_start  = _n_cols_df + 2
                        _vta_mercado_col = _vta_start
                        _vta_skus_start  = _vta_start + 1
                        _vta_skus_end    = _vta_start + len(_vta_skus)
                        worksheet.merge_range(0, _vta_mercado_col, 1, _vta_mercado_col, '', None)
                        worksheet.set_column(_vta_mercado_col, _vta_mercado_col,
                                             SA_WIDTHS.get('MERCADO', 74.3))
                        if len(_vta_skus) > 1:
                            worksheet.merge_range(0, _vta_skus_start, 0, _vta_skus_end,
                                                  'Vta Declarada SA', sa_header_format)
                        else:
                            worksheet.write(0, _vta_skus_start, 'Vta Declarada SA', sa_header_format)
                        for _vi, _vsku in enumerate(_vta_skus):
                            worksheet.write(1, _vta_skus_start + _vi, _vsku, sa_sku_format)
                            worksheet.set_column(_vta_skus_start + _vi, _vta_skus_start + _vi, SA_SKU_WIDTH)
    
            if sheet_name == "RESUMEN":
                _R_COL = 2
                _R_ROW = 9
    
                _aptos11   = workbook.add_format({'font_name': 'Aptos', 'font_size': 11})
                _white_fmt = workbook.add_format({'fg_color': '#FFFFFF', 'font_name': 'Aptos', 'font_size': 11})
                worksheet.set_column_pixels(0, 50, None, _white_fmt)
                worksheet.hide_gridlines(2)
                worksheet.set_column_pixels(0, 0, 26, _aptos11)
                worksheet.set_column_pixels(1, 1, 26, _aptos11)
    
                r_header_fmt = workbook.add_format({
                    'bold': True, 'text_wrap': True, 'valign': 'vcenter', 'align': 'center',
                    'fg_color': '#F2F2F2', 'font_color': '#000000', 'border': 1,
                    'font_size': 11, 'font_name': 'Aptos'
                })
                r_body_fmt = workbook.add_format({
                    'border': 1, 'align': 'center', 'valign': 'vcenter',
                    'font_size': 11, 'font_name': 'Aptos'
                })
                r_number_fmt = workbook.add_format({
                    'border': 1, 'align': 'center', 'valign': 'vcenter',
                    'num_format': '#,##0', 'font_size': 11, 'font_name': 'Aptos'
                })
                r_pct_fmt = workbook.add_format({
                    'border': 1, 'align': 'center', 'valign': 'vcenter',
                    'num_format': '0%', 'font_size': 11, 'font_name': 'Aptos'
                })
                r_total_fmt = workbook.add_format({
                    'bold': True, 'border': 1, 'align': 'center', 'valign': 'vcenter',
                    'font_size': 11, 'fg_color': '#F2F2F2', 'font_color': '#000000', 'font_name': 'Aptos'
                })
                r_total_num_fmt = workbook.add_format({
                    'bold': True, 'border': 1, 'align': 'center', 'valign': 'vcenter',
                    'num_format': '#,##0', 'font_size': 11,
                    'fg_color': '#F2F2F2', 'font_color': '#000000', 'font_name': 'Aptos'
                })
                r_total_pct_fmt = workbook.add_format({
                    'bold': True, 'border': 1, 'align': 'center', 'valign': 'vcenter',
                    'num_format': '0%', 'font_size': 11,
                    'fg_color': '#F2F2F2', 'font_color': '#000000', 'font_name': 'Aptos'
                })
                r_flat_sku_fmt = workbook.add_format({
                    'bold': True, 'text_wrap': True, 'align': 'center', 'valign': 'vcenter',
                    'fg_color': '#D9D9D9', 'font_color': '#000000', 'border': 1,
                    'font_size': 11, 'font_name': 'Aptos'
                })
                r_total_col_fmt = workbook.add_format({
                    'bold': True, 'text_wrap': True, 'align': 'center', 'valign': 'vcenter',
                    'fg_color': '#F2F2F2', 'font_color': '#000000', 'border': 1,
                    'font_size': 11, 'font_name': 'Aptos'
                })
                r_stock_header_fmt = workbook.add_format({
                    'bold': True, 'text_wrap': True, 'align': 'center', 'valign': 'vcenter',
                    'fg_color': '#1F497D', 'font_color': 'white', 'border': 1,
                    'font_size': 11, 'font_name': 'Aptos'
                })
                _titulo_fmt = workbook.add_format({
                    'bold': True, 'align': 'center', 'valign': 'vcenter',
                    'fg_color': '#1F497D', 'font_color': 'white', 'border': 1,
                    'font_size': 11, 'font_name': 'Aptos'
                })
                _titulo_label = f"RESUMEN STORE CHECK - {marca_storecheck.upper()}"
                _ncols = len(df.columns) - 1
                if _ncols > 0:
                    worksheet.merge_range(_R_ROW, _R_COL, _R_ROW, _R_COL + _ncols, _titulo_label, _titulo_fmt)
                else:
                    worksheet.write(_R_ROW, _R_COL, _titulo_label, _titulo_fmt)
                worksheet.set_row(0, 15)
    
                _img_candidates = [
                    os.path.join(backend_dir, 'xplora.jpeg'),
                    os.path.join(backend_dir, 'logo-xplora-cuadrado.png'),
                    os.path.join(directorio, 'xplora.jpeg'),
                ]
                for _img_path in _img_candidates:
                    if os.path.exists(_img_path):
                        worksheet.insert_image(0, 1, _img_path, {
                            'x_offset': 0, 'y_offset': 10,
                            'x_scale': 174 / 330, 'y_scale': 174 / 330,
                            'object_position': 1,
                        })
                        break
    
                _bold11 = workbook.add_format({'bold': True,  'font_size': 11, 'font_name': 'Aptos', 'fg_color': '#FFFFFF'})
                _norm11 = workbook.add_format({'bold': False, 'font_size': 11, 'font_name': 'Aptos', 'fg_color': '#FFFFFF'})
    
                worksheet.write_rich_string(5, 2, _bold11, 'País: ', _norm11, 'Perú', _norm11)
    
                _ciu_series = df_modificado_v10['CIUDAD'].dropna()
                _ciudades_unicas = list(dict.fromkeys(
                    c.strip().upper() for c in _ciu_series
                    if str(c).strip() and str(c).strip().upper() not in ('N/A', 'NAN', 'NONE', '')
                ))
                _ciudades_str = ', '.join(_ciudades_unicas) if _ciudades_unicas else ''
                if _ciudades_str:
                    worksheet.write_rich_string(6, 2, _bold11, 'Ciudad: ', _norm11, _ciudades_str, _norm11)
                else:
                    worksheet.write(6, 2, 'Ciudad:', _bold11)
    
                worksheet.set_row(7, 26.3)
                worksheet.write(7, 2, '1.- Resumen.', _bold11)
    
                _ciudad_val = _ciudades_unicas[0] if _ciudades_unicas else ''
                _c5_label = f"{numero_storecheck} STORECHECK - {_parte_final.upper()} - {_ciudades_str}"
                _c5_fmt = workbook.add_format({
                    'bold': True, 'align': 'center', 'valign': 'vcenter',
                    'font_size': 22, 'font_name': 'Aptos', 'fg_color': '#FFFFFF'
                })
                if _ncols > 0:
                    worksheet.merge_range(4, 2, 4, 2 + _ncols, _c5_label, _c5_fmt)
                else:
                    worksheet.write(4, 2, _c5_label, _c5_fmt)
                worksheet.set_row(_R_ROW, 25)
                worksheet.set_row(_R_ROW + 1, 86.3)
    
                _c1_widths_px = {
                    'PROVINCIA': 141, 'MERCADO': 700,
                    'PUESTOS DE MERCADO': 229,
                    'PUESTOS ATENDIDOS POR ALICORP': 180,
                    'PRESENCIA  DEL PRODUCTO': 180,
                    'COBERTURA TOTAL (Puestos Atendidos por Alicorp)': 180,
                }
                for _ci, _cn in enumerate(df.columns.values):
                    worksheet.write(_R_ROW + 1, _R_COL + _ci, _cn, r_header_fmt)
                    _px = _c1_widths_px.get(_cn, 120)
                    worksheet.set_column_pixels(_R_COL + _ci, _R_COL + _ci, _px, _aptos11)
                current_row = _R_ROW + 2
            else:
                current_row = header_offset
    
            _w_col_offset = _R_COL if sheet_name == "RESUMEN" else 0
            date_col_idx  = -1
            for i, col in enumerate(df.columns):
                col_name = col[0] if isinstance(col, tuple) and len(col) > 0 else col
                if col_name == 'FECHA':
                    date_col_idx = i
                    break
    
            market_col    = next((c for c in df.columns
                                  if (c[0] if isinstance(c, tuple) and len(c) > 0 else c) == 'MERCADO'), None)
            apply_spacing = sheet_name in ["STORECHECK", "STOCK SIN ACTIVIDAD"]
    
            if market_col and apply_spacing:
                _mercados_orden = list(dict.fromkeys(df[market_col].dropna().tolist()))
                data_to_write   = [(m, df[df[market_col] == m]) for m in _mercados_orden]
            else:
                data_to_write = [(None, df)]
    
            if sheet_name == "STORECHECK":
                _sc_market_ranges = {}
            if sheet_name == "STOCK SIN ACTIVIDAD":
                _sa_market_ranges = {}
    
            sep_col_idx       = next((i for i, c in enumerate(df.columns) if c == '_separador_'), -1)
            _sc_use_formulas  = False
            _sc_formula_first = 0
            _sc_formula_last  = 0
    
            col_names_flat = [(c[0] if isinstance(c, tuple) else c) for c in df.columns]
            idx_n         = col_names_flat.index('N°')             if 'N°'             in col_names_flat else -1
            idx_fecha     = col_names_flat.index('FECHA')          if 'FECHA'          in col_names_flat else -1
            idx_provincia = col_names_flat.index('PROVINCIA')      if 'PROVINCIA'      in col_names_flat else -1
            idx_dex       = col_names_flat.index('DEX')            if 'DEX'            in col_names_flat else -1
            idx_mercado   = col_names_flat.index('MERCADO')        if 'MERCADO'        in col_names_flat else -1
            idx_cliente   = col_names_flat.index('NOMBRE CLIENTE') if 'NOMBRE CLIENTE' in col_names_flat else -1
            idx_npuesto   = col_names_flat.index('N° de puesto')   if 'N° de puesto'   in col_names_flat else -1
            stock_col_idxs = [i for i, c in enumerate(df.columns) if isinstance(c, tuple) and c[1] != '']
    
            total_format = workbook.add_format({
                'bold': False, 'border': 1, 'align': 'center', 'valign': 'vcenter',
                'font_size': 12, 'fg_color': '#1F497D', 'font_color': 'white', 'font_name': 'Aptos'
            })
            total_number_format = workbook.add_format({
                'bold': False, 'border': 1, 'align': 'center', 'valign': 'vcenter',
                'num_format': '#,##0', 'font_size': 12,
                'fg_color': '#1F497D', 'font_color': 'white', 'font_name': 'Aptos'
            })
            total_date_format = workbook.add_format({
                'bold': False, 'num_format': 'dd/mm/yyyy', 'align': 'center', 'border': 1,
                'valign': 'vcenter', 'font_size': 12,
                'fg_color': '#1F497D', 'font_color': 'white', 'font_name': 'Aptos'
            })
    
            if sheet_name == 'STORECHECK':
                for _ci, _cn in enumerate(df.columns):
                    _sc_col_indices[str(_cn)] = _ci
            if sheet_name == 'STOCK SIN ACTIVIDAD':
                for _ci, _cn in enumerate(df.columns):
                    _sa_col_indices[str(_cn)] = _ci
                    if isinstance(_cn, tuple) and _cn[0] == 'STOCK INICIAL' and _cn[1] != '':
                        _sa_sku_cols[_cn[1]] = _ci
                    if isinstance(_cn, tuple) and _cn[0] == 'STOCK FINAL' and _cn[1] != '':
                        _sa_sku_cols_sf[_cn[1]] = _ci
    
            if sheet_name == "STORECHECK":
                _sc_fixed = ['N°', 'FECHA', 'PROVINCIA', 'DEX', 'MERCADO', 'NOMBRE CLIENTE',
                             'N° Puesto', columna_promocion_interna, 'Alicorp', 'Otros',
                             'Presencia en Puestos Alicorp', 'Presencia Producto']
                sku_col_idxs_sc = [i for i, c in enumerate(df.columns)
                                    if c not in _sc_fixed and not isinstance(c, tuple)]
                idx_npuesto_sc  = col_names_flat.index('N° Puesto') if 'N° Puesto' in col_names_flat else -1
                idx_vende_sc    = (col_names_flat.index(columna_promocion_interna)
                                   if columna_promocion_interna in col_names_flat else -1)
            else:
                sku_col_idxs_sc = []
                idx_npuesto_sc  = -1
                idx_vende_sc    = -1
    
            if sheet_name == "STORECHECK":
                idx_pres_alicorp_data  = (col_names_flat.index('Presencia en Puestos Alicorp')
                                          if 'Presencia en Puestos Alicorp' in col_names_flat else -1)
                idx_pres_producto_data = (col_names_flat.index('Presencia Producto')
                                          if 'Presencia Producto' in col_names_flat else -1)
            else:
                idx_pres_alicorp_data  = -1
                idx_pres_producto_data = -1
    
            if sheet_name == "STORECHECK":
                _sa_header_off = 2
                _sa_row_map    = {}
                _sa_cur        = _sa_header_off
                for _sa_merc, _sa_grp in [(m, pv[pv['MERCADO'] == m]) for m in
                                           list(dict.fromkeys(pv['MERCADO'].tolist()))]:
                    for _sa_idx in _sa_grp.index:
                        _sa_row_map[_sa_idx] = _sa_cur
                        _sa_cur += 1
                    _sa_cur += 1 + 3
    
                from xlsxwriter.utility import xl_col_to_name
                _sa_sf_cols = {}
                for _ci, _cn in enumerate(pv.columns):
                    if isinstance(_cn, tuple) and _cn[0] == 'STOCK FINAL' and _cn[1] != '':
                        _sa_sf_cols[_cn[1]] = xl_col_to_name(_ci)
                _pv_idx_list = list(pv.index)
    
            for _, group in data_to_write:
                for _, row_data in group.iterrows():
                    if sheet_name in ("STORECHECK", "STOCK SIN ACTIVIDAD"):
                        worksheet.set_row(current_row, None, None,
                                          {'level': 1, 'collapsed': True, 'hidden': True})
    
                    if sheet_name == "STOCK SIN ACTIVIDAD":
                        _sa_detail_row_map[row_data.name] = current_row + 1
    
                    for col_num, cell_value in enumerate(row_data):
                        if col_num == sep_col_idx:
                            worksheet.write(current_row, col_num, '', None)
                            continue
    
                        if sheet_name == "STORECHECK":
                            val = cell_value
                            if col_num == idx_vende_sc:
                                if idx_pres_producto_data >= 0:
                                    _pres_col = gcl(idx_pres_producto_data + 1)
                                    _r1b = current_row + 1
                                    formula = '=IF(' + _pres_col + str(_r1b) + '=1,"SI","NO")'
                                    worksheet.write_formula(current_row, col_num, formula, sc_body_format)
                                else:
                                    worksheet.write(current_row, col_num, 'NO', sc_body_format)
                                continue
                            if col_num == idx_pres_alicorp_data:
                                _r1b = current_row + 1
                                idx_alicorp_data = (col_names_flat.index('Alicorp')
                                                    if 'Alicorp' in col_names_flat else -1)
                                if idx_alicorp_data >= 0 and idx_pres_producto_data >= 0:
                                    _alic_col = gcl(idx_alicorp_data + 1)
                                    _pres_col = gcl(idx_pres_producto_data + 1)
                                    formula = ('=IF(AND(' + _alic_col + str(_r1b) + '=1,'
                                               + _pres_col + str(_r1b) + '=1),1," ")')
                                    worksheet.write_formula(current_row, col_num, formula, data_presencia_alicorp)
                                else:
                                    worksheet.write(current_row, col_num, '', data_presencia_alicorp)
                                continue
                            if col_num == idx_pres_producto_data:
                                if sku_col_idxs_sc:
                                    _first_sku = gcl(sku_col_idxs_sc[0] + 1)
                                    _last_sku  = gcl(sku_col_idxs_sc[-1] + 1)
                                    _r1b = current_row + 1
                                    formula = f"=IF(SUM({_first_sku}{_r1b}:{_last_sku}{_r1b})>0,1,\"\")"
                                    worksheet.write_formula(current_row, col_num, formula, data_presencia_producto)
                                else:
                                    worksheet.write(current_row, col_num, '', data_presencia_producto)
                                continue
                            if col_num in sku_col_idxs_sc:
                                _sku_name = df.columns[col_num]
                                _sa_r     = _sa_detail_row_map.get(row_data.name, None)
                                _sa_ci    = _sa_sku_cols.get(_sku_name, _sa_sku_cols_sf.get(_sku_name, None))
                                if _sa_r is not None and _sa_ci is not None:
                                    _sa_col_l = gcl(_sa_ci + 1)
                                    formula = f"=IF('STOCK SIN ACTIVIDAD'!${_sa_col_l}${_sa_r}>0,1,\"\")"
                                    worksheet.write_formula(current_row, col_num, formula, sc_number_format)
                                else:
                                    val = 1 if pd.notna(cell_value) and cell_value != 0 else ''
                                    fmt = sc_number_format if val != '' else sc_body_format
                                    worksheet.write(current_row, col_num, val, fmt)
                                continue
                            if col_num == date_col_idx and pd.notna(val) and isinstance(val, pd.Timestamp):
                                worksheet.write(current_row, col_num, val, sc_date_format)
                                continue
                            if pd.isna(val):
                                val = ''
                            fmt = (sc_number_format if isinstance(val, (int, float)) and val != ''
                                   else sc_body_format)
                            worksheet.write(current_row, col_num, val, fmt)
                            continue
    
                        if sheet_name == "RESUMEN":
                            _cn   = df.columns[col_num]
                            _merc = (row_data.get('MERCADO', '') if hasattr(row_data, 'get')
                                     else row_data['MERCADO'])
                            _sc_row = _sc_total_rows.get(str(_merc), None)
    
                            if _cn == 'CIUDAD' or _cn == 'MERCADO':
                                val = cell_value if not pd.isna(cell_value) else ''
                                worksheet.write(current_row, _R_COL + col_num, val, r_body_fmt)
                            elif _cn == 'PUESTOS DE MERCADO' and _sc_row:
                                _sc_npuesto_col = gcl(_sc_col_indices.get('N° Puesto', 5) + 1)
                                formula = f"='STORECHECK'!${_sc_npuesto_col}${_sc_row}"
                                worksheet.write_formula(current_row, _R_COL + col_num, formula, r_number_fmt)
                            elif _cn == 'PUESTOS ATENDIDOS POR ALICORP' and _sc_row:
                                _sc_alicorp_col = gcl(_sc_col_indices.get('Alicorp', 8) + 1)
                                formula = f"='STORECHECK'!${_sc_alicorp_col}${_sc_row}"
                                worksheet.write_formula(current_row, _R_COL + col_num, formula, r_number_fmt)
                            elif 'PRESENCIA' in str(_cn).upper() and _sc_row:
                                _sc_pres_col = gcl(_sc_col_indices.get('Presencia Producto', 10) + 1)
                                formula = f"='STORECHECK'!${_sc_pres_col}${_sc_row}"
                                worksheet.write_formula(current_row, _R_COL + col_num, formula, r_number_fmt)
                            elif 'COBERTURA' in str(_cn).upper() and _sc_row:
                                _sc_npuesto_col = gcl(_sc_col_indices.get('N° Puesto', 5) + 1)
                                _sc_alicorp_col = gcl(_sc_col_indices.get('Alicorp', 8) + 1)
                                formula = (f"=IFERROR('STORECHECK'!${_sc_alicorp_col}${_sc_row}"
                                           f"/'STORECHECK'!${_sc_npuesto_col}${_sc_row},0)")
                                worksheet.write_formula(current_row, _R_COL + col_num, formula, r_pct_fmt)
                            else:
                                val = cell_value if not pd.isna(cell_value) else ''
                                worksheet.write(current_row, _R_COL + col_num, val, r_body_fmt)
                            continue
                        else:
                            fmt = body_format
                            val = cell_value
                            if col_num == date_col_idx and pd.notna(val) and isinstance(val, pd.Timestamp):
                                fmt = date_format
                            elif pd.isna(val):
                                val = ''
                            elif isinstance(val, (int, float, np.integer, np.floating)) and val != '':
                                fmt = number_format
                                val = (int(val) if isinstance(val, np.integer) else
                                       float(val) if isinstance(val, np.floating) else val)
                        worksheet.write(current_row, _w_col_offset + col_num, val, fmt)
                    current_row += 1
    
                if sheet_name == "STORECHECK" and _ is not None:
                    _sc_market_ranges[_] = {
                        'first': current_row - len(group),
                        'last':  current_row - 1,
                        'total': current_row,
                        'pct':   current_row + 1,
                    }
                if sheet_name == "STOCK SIN ACTIVIDAD" and _ is not None:
                    _sa_market_ranges[_] = {
                        'first': current_row - len(group),
                        'last':  current_row - 1,
                        'total': current_row,
                    }
    
                if sheet_name in ("STOCK SIN ACTIVIDAD", "STORECHECK"):
                    total_row = [''] * len(df.columns)
    
                    if idx_n >= 0:
                        total_row[idx_n] = len(group)
                    if idx_fecha >= 0:
                        fechas = group.iloc[:, idx_fecha].dropna()
                        total_row[idx_fecha] = fechas.max() if not fechas.empty else ''
                    if idx_provincia >= 0:
                        val_prov = group.iloc[0, idx_provincia]
                        total_row[idx_provincia] = ('' if pd.isna(val_prov) or
                                                    str(val_prov).strip().upper() == 'N/A' else val_prov)
                    if idx_dex >= 0:
                        val_dex = group.iloc[0, idx_dex]
                        total_row[idx_dex] = ('' if pd.isna(val_dex) or
                                              str(val_dex).strip().upper() == 'N/A' else val_dex)
                    if idx_mercado >= 0:
                        total_row[idx_mercado] = group.iloc[0, idx_mercado]
                    if idx_cliente >= 0:
                        total_row[idx_cliente] = group.iloc[:, idx_cliente].count()
    
                    _vta_body_fmt = workbook.add_format({
                        'border': 1, 'align': 'center', 'valign': 'vcenter',
                        'font_size': 14, 'font_name': 'Aptos',
                        'fg_color': '#FFFFFF', 'font_color': '#000000'
                    })
                    _vta_num_fmt = workbook.add_format({
                        'border': 1, 'align': 'center', 'valign': 'vcenter',
                        'num_format': '#,##0', 'font_size': 14, 'font_name': 'Aptos',
                        'fg_color': '#FFFFFF', 'font_color': '#000000'
                    })
    
                    if sheet_name == "STORECHECK":
                        _sc_first_1b = current_row - len(group) + 1
                        _sc_last_1b  = current_row
                        _sc_use_formulas  = True
                        _sc_formula_first = _sc_first_1b
                        _sc_formula_last  = _sc_last_1b
    
                        if idx_n >= 0:         total_row[idx_n]        = len(group)
                        if idx_fecha >= 0:
                            fechas = group.iloc[:, idx_fecha].dropna()
                            total_row[idx_fecha] = fechas.max() if not fechas.empty else ''
                        if idx_provincia >= 0: total_row[idx_provincia] = group.iloc[0, idx_provincia]
                        if idx_dex >= 0:       total_row[idx_dex]       = group.iloc[0, idx_dex]
                        if idx_mercado >= 0:   total_row[idx_mercado]   = group.iloc[0, idx_mercado]
                        if idx_cliente >= 0:   total_row[idx_cliente]   = group.iloc[0, idx_cliente]
    
                    _mercado_nombre = group.iloc[0, idx_mercado] if idx_mercado >= 0 else ''
                    if sheet_name == "STORECHECK":
                        _sc_total_rows[_mercado_nombre] = current_row + 1
                    elif sheet_name == "STOCK SIN ACTIVIDAD":
                        _sa_total_rows[_mercado_nombre] = current_row + 1
    
                    if sheet_name in ("STORECHECK", "STOCK SIN ACTIVIDAD"):
                        worksheet.set_row(current_row, None, None, {'level': 0, 'collapsed': True})
    
                    _tf  = sa_total_format        if sheet_name == "STOCK SIN ACTIVIDAD" else total_format
                    _tnf = sa_total_number_format if sheet_name == "STOCK SIN ACTIVIDAD" else total_number_format
                    _tdf = sa_total_date_format   if sheet_name == "STOCK SIN ACTIVIDAD" else total_date_format
    
                    if sheet_name == "STORECHECK" and _sc_use_formulas:
                        _f1 = _sc_formula_first
                        _f2 = _sc_formula_last
                        for col_num in range(len(df.columns)):
                            if col_num == sep_col_idx:
                                worksheet.write(current_row, col_num, '', None)
                                continue
                            _cl  = gcl(col_num + 1)
                            val  = total_row[col_num] if col_num < len(total_row) else ''
                            if col_num == idx_n:
                                worksheet.write(current_row, col_num, val, _tnf)
                            elif col_num == idx_fecha:
                                worksheet.write(current_row, col_num,
                                                val if isinstance(val, pd.Timestamp) else '', _tdf)
                            elif col_num in (idx_provincia, idx_dex, idx_mercado, idx_cliente):
                                worksheet.write(current_row, col_num, val if val != '' else '', _tf)
                            elif col_num == idx_npuesto_sc:
                                worksheet.write_formula(current_row, col_num,
                                    f"=COUNTA({_cl}{_f1}:{_cl}{_f2})", _tnf)
                            elif col_num == idx_vende_sc:
                                worksheet.write_formula(current_row, col_num,
                                    '=COUNTIF(' + _cl + str(_f1) + ':' + _cl + str(_f2) + ',"SI")', _tnf)
                            elif col_num == idx_pres_alicorp_data:
                                worksheet.write_formula(current_row, col_num,
                                    f"=COUNTIF({_cl}{_f1}:{_cl}{_f2},1)", _tnf)
                            elif col_num == idx_pres_producto_data:
                                worksheet.write_formula(current_row, col_num,
                                    f"=COUNTIF({_cl}{_f1}:{_cl}{_f2},1)", _tnf)
                            elif col_num in sku_col_idxs_sc:
                                worksheet.write_formula(current_row, col_num,
                                    f"=COUNTIF({_cl}{_f1}:{_cl}{_f2},1)", _tnf)
                            elif col_names_flat[col_num] == 'Alicorp':
                                worksheet.write_formula(current_row, col_num,
                                    f"=COUNTIF({_cl}{_f1}:{_cl}{_f2},1)", _tnf)
                            elif col_names_flat[col_num] == 'Otros':
                                worksheet.write_formula(current_row, col_num,
                                    f"=COUNTIF({_cl}{_f1}:{_cl}{_f2},1)", _tnf)
                            else:
                                if isinstance(val, (int, float)) and val != '':
                                    worksheet.write(current_row, col_num, val, _tnf)
                                else:
                                    worksheet.write(current_row, col_num,
                                                    val if val != '' else '', _tf)
                    elif sheet_name == "STOCK SIN ACTIVIDAD":
                        _sa_f1 = current_row - len(group) + 1
                        _sa_f2 = current_row
                        for col_num in range(len(df.columns)):
                            if col_num == sep_col_idx:
                                worksheet.write(current_row, col_num, '', None)
                                continue
                            _cl  = gcl(col_num + 1)
                            val  = total_row[col_num] if col_num < len(total_row) else ''
                            cn   = col_names_flat[col_num]
    
                            if col_num == idx_n:
                                worksheet.write_formula(current_row, col_num,
                                    f"=COUNTA({_cl}{_sa_f1}:{_cl}{_sa_f2})", _tnf)
                            elif col_num == idx_fecha:
                                worksheet.write_formula(current_row, col_num,
                                    f"=MAX({_cl}{_sa_f1}:{_cl}{_sa_f2})", _tdf)
                            elif col_num in (idx_provincia, idx_dex, idx_mercado):
                                worksheet.write(current_row, col_num,
                                                val if val != '' else '', _tf)
                            elif col_num == idx_cliente:
                                worksheet.write_formula(current_row, col_num,
                                    f"=COUNTA({_cl}{_sa_f1}:{_cl}{_sa_f2})", _tnf)
                            elif col_num == idx_npuesto:
                                worksheet.write_formula(current_row, col_num,
                                    f"=COUNTA({_cl}{_sa_f1}:{_cl}{_sa_f2})", _tnf)
                            elif col_num in stock_col_idxs:
                                worksheet.write_formula(current_row, col_num,
                                    f"=SUM({_cl}{_sa_f1}:{_cl}{_sa_f2})", _tnf)
                            else:
                                if isinstance(val, (int, float)) and val != '':
                                    worksheet.write(current_row, col_num, val, _tnf)
                                else:
                                    worksheet.write(current_row, col_num,
                                                    val if val != '' else '', _tf)
                        _vta_skus_f  = [c[1] for c in df.columns
                                        if isinstance(c, tuple) and c[0] == 'STOCK FINAL' and c[1] != '']
                        _vta_start_f = len(df.columns) + 2
                        _mercado_val = group.iloc[0, idx_mercado] if idx_mercado >= 0 else ''
                        worksheet.write(current_row, _vta_start_f, _mercado_val, _vta_body_fmt)
                        for _vi, _vsku in enumerate(_vta_skus_f):
                            _si_c = next((c for c in df.columns
                                          if isinstance(c, tuple) and c[0] == 'STOCK INICIAL'
                                          and c[1] == _vsku), None)
                            _sf_c = next((c for c in df.columns
                                          if isinstance(c, tuple) and c[0] == 'STOCK FINAL'
                                          and c[1] == _vsku), None)
                            if _si_c is not None and _sf_c is not None:
                                _si_letter = gcl(df.columns.get_loc(_si_c) + 1)
                                _sf_letter = gcl(df.columns.get_loc(_sf_c) + 1)
                                worksheet.write_formula(
                                    current_row, _vta_start_f + 1 + _vi,
                                    f"=SUM({_si_letter}{_sa_f1}:{_si_letter}{_sa_f2})"
                                    f"-SUM({_sf_letter}{_sa_f1}:{_sf_letter}{_sa_f2})",
                                    _vta_num_fmt)
                            else:
                                worksheet.write(current_row, _vta_start_f + 1 + _vi, 0, _vta_num_fmt)
                    else:
                        for col_num, val in enumerate(total_row):
                            if col_num == sep_col_idx:
                                worksheet.write(current_row, col_num, '', None)
                                continue
                            if col_num == idx_fecha and isinstance(val, pd.Timestamp):
                                worksheet.write(current_row, col_num, val, _tdf)
                            elif isinstance(val, (int, float)) and val != '':
                                worksheet.write(current_row, col_num, val, _tnf)
                            else:
                                worksheet.write(current_row, col_num, val, _tf)
                    current_row += 1
    
                    if sheet_name == "STORECHECK":
                        worksheet.set_row(current_row, None, None, {'level': 0})
                        pct_empty_format = workbook.add_format({'border': 0, 'fg_color': '#FFFFFF'})
                        idx_alicorp = (col_names_flat.index('Alicorp')
                                       if 'Alicorp' in col_names_flat else -1)
                        _total_row_1b  = current_row
                        _npuesto_col_l = gcl(idx_npuesto_sc + 1) if idx_npuesto_sc >= 0 else 'A'
                        for col_num in range(len(df.columns)):
                            if col_num < idx_alicorp:
                                worksheet.write(current_row, col_num, '', pct_empty_format)
                            elif col_num == sep_col_idx:
                                worksheet.write(current_row, col_num, '', None)
                            else:
                                _cl = gcl(col_num + 1)
                                formula = ('=IFERROR(' + _cl + str(_total_row_1b) +
                                           '/' + _npuesto_col_l + str(_total_row_1b) + ',"")')
                                worksheet.write_formula(current_row, col_num, formula, pct_format_sc)
                        current_row += 1
    
                if market_col and apply_spacing:
                    current_row += 3
    
            if sheet_name == "RESUMEN":
                _c1_cols    = list(df.columns)
                idx_prov_c1 = _c1_cols.index('CIUDAD')   if 'CIUDAD'   in _c1_cols else 0
                idx_merc_c1 = _c1_cols.index('MERCADO')  if 'MERCADO'  in _c1_cols else 1
    
                if idx_prov_c1 != idx_merc_c1:
                    worksheet.merge_range(current_row, _R_COL + idx_prov_c1,
                                           current_row, _R_COL + idx_merc_c1, 'TOTAL', r_total_fmt)
                else:
                    worksheet.write(current_row, _R_COL + idx_prov_c1, 'TOTAL', r_total_fmt)
    
                _c1_data_first = _R_ROW + 2 + 1
                _c1_data_last  = current_row
                for _ci, _cn in enumerate(_c1_cols):
                    if _ci in (idx_prov_c1, idx_merc_c1):
                        continue
                    _col_letter = gcl(_R_COL + _ci + 1)
                    if 'COBERTURA' in str(_cn).upper():
                        _pm_col = (gcl(_R_COL + _c1_cols.index('PUESTOS DE MERCADO') + 1)
                                   if 'PUESTOS DE MERCADO' in _c1_cols else _col_letter)
                        _pa_col = (gcl(_R_COL + _c1_cols.index('PUESTOS ATENDIDOS POR ALICORP') + 1)
                                   if 'PUESTOS ATENDIDOS POR ALICORP' in _c1_cols else _col_letter)
                        formula = (f"=IFERROR(SUM({_pa_col}{_c1_data_first}:{_pa_col}{_c1_data_last})"
                                   f"/SUM({_pm_col}{_c1_data_first}:{_pm_col}{_c1_data_last}),0)")
                        worksheet.write_formula(current_row, _R_COL + _ci, formula, r_total_pct_fmt)
                    else:
                        formula = f"=SUM({_col_letter}{_c1_data_first}:{_col_letter}{_c1_data_last})"
                        worksheet.write_formula(current_row, _R_COL + _ci, formula, r_total_num_fmt)
                current_row += 1
    
                _merge_ciudad_col(worksheet, _R_ROW + 2, current_row - 1,
                                   _R_COL + idx_prov_c1, df, r_body_fmt)
                current_row += 3
    
                _r_fixed = ['CIUDAD', 'MERCADO', 'PUESTOS DE MERCADO']
                _r_skus  = [c for c in resumen_stock_final.columns
                            if c not in _r_fixed and c != 'STOCK TOTAL']
                _r_cols  = _r_fixed + _r_skus + ['STOCK TOTAL']
    
                _hdr_row     = current_row
                _stock_label = f"STOCK {marca_storecheck}"
    
                worksheet.set_row(_hdr_row,     20)
                worksheet.set_row(_hdr_row + 1, 86.3)
    
                for _ci in range(len(_r_fixed)):
                    worksheet.write(_hdr_row, _R_COL + _ci, '', None)
    
                if _r_skus:
                    _sf_start = len(_r_fixed)
                    _sf_end   = _sf_start + len(_r_skus)
                    if _sf_end > _sf_start:
                        worksheet.merge_range(_hdr_row, _R_COL + _sf_start,
                                               _hdr_row, _R_COL + _sf_end,
                                               _stock_label, r_stock_header_fmt)
                    else:
                        worksheet.write(_hdr_row, _R_COL + _sf_start, _stock_label, r_stock_header_fmt)
    
                _c2_widths_px = {'CIUDAD': 141, 'MERCADO': 700, 'PUESTOS DE MERCADO': 229}
                for _ci, _cn in enumerate(_r_fixed):
                    worksheet.write(_hdr_row + 1, _R_COL + _ci, _cn, r_header_fmt)
                    _px2 = _c2_widths_px.get(_cn, 120)
                    worksheet.set_column_pixels(_R_COL + _ci, _R_COL + _ci, _px2, _aptos11)
    
                if _r_skus:
                    for _si, _sn in enumerate(_r_skus):
                        _col_pos = _R_COL + _sf_start + _si
                        worksheet.write(_hdr_row + 1, _col_pos, _sn, r_flat_sku_fmt)
                        worksheet.set_column_pixels(_col_pos, _col_pos, 180, _aptos11)
                    _total_col_idx = _R_COL + _sf_start + len(_r_skus)
                    worksheet.write(_hdr_row + 1, _total_col_idx, 'STOCK TOTAL', r_total_col_fmt)
                    worksheet.set_column_pixels(_total_col_idx, _total_col_idx, 180, _aptos11)
    
                current_row += 2
    
                _idx_prov_r2  = _r_cols.index('CIUDAD')      if 'CIUDAD'      in _r_cols else -1
                _idx_merc_r2  = _r_cols.index('MERCADO')     if 'MERCADO'     in _r_cols else -1
                _idx_total_r2 = _r_cols.index('STOCK TOTAL') if 'STOCK TOTAL' in _r_cols else -1
    
                _c2_data_start = current_row
                _resumen_data  = resumen_stock_final[
                    resumen_stock_final['MERCADO'].str.strip().str.upper() != 'TOTAL']
    
                for _ri, _rd in _resumen_data.iterrows():
                    _merc   = str(_rd.get('MERCADO', ''))
                    _sa_row = _sa_total_rows.get(_merc, None)
    
                    for _ci, _cn in enumerate(_r_cols):
                        _abs_col = _R_COL + _ci
                        if _cn == 'CIUDAD':
                            _val = _rd.get('CIUDAD', '')
                            worksheet.write(current_row, _abs_col,
                                            '' if pd.isna(_val) else _val, r_body_fmt)
                        elif _cn == 'MERCADO':
                            worksheet.write(current_row, _abs_col, _merc, r_body_fmt)
                        elif _cn == 'PUESTOS DE MERCADO' and _sa_row:
                            _sa_np_col = gcl(_sa_col_indices.get('N° de puesto', 5) + 1)
                            formula = f"='STOCK SIN ACTIVIDAD'!${_sa_np_col}${_sa_row}"
                            worksheet.write_formula(current_row, _abs_col, formula, r_number_fmt)
                        elif _cn == 'STOCK TOTAL':
                            if _r_skus:
                                _first_sku_col = gcl(_R_COL + _sf_start + 1)
                                _last_sku_col  = gcl(_R_COL + _sf_start + len(_r_skus))
                                _this_row_1b   = current_row + 1
                                formula = f"=SUM({_first_sku_col}{_this_row_1b}:{_last_sku_col}{_this_row_1b})"
                                worksheet.write_formula(current_row, _abs_col, formula, r_number_fmt)
                            else:
                                worksheet.write(current_row, _abs_col, 0, r_number_fmt)
                        elif _cn in _sku_names and _sa_row:
                            _sa_sku_ci = _sa_sku_cols_sf.get(_cn, None)
                            if _sa_sku_ci is not None:
                                _sa_sku_col_letter = gcl(_sa_sku_ci + 1)
                                formula = f"='STOCK SIN ACTIVIDAD'!${_sa_sku_col_letter}${_sa_row}"
                                worksheet.write_formula(current_row, _abs_col, formula, r_number_fmt)
                            else:
                                worksheet.write(current_row, _abs_col, 0, r_number_fmt)
                        else:
                            _val = _rd.get(_cn, '')
                            worksheet.write(current_row, _abs_col,
                                            '' if pd.isna(_val) else _val, r_body_fmt)
                    current_row += 1
    
                _c2_data_end = current_row
                _c2_first_1b = _c2_data_start + 1
                _c2_last_1b  = _c2_data_end
    
                if _idx_prov_r2 >= 0 and _idx_merc_r2 >= 0 and _idx_prov_r2 != _idx_merc_r2:
                    worksheet.merge_range(current_row, _R_COL + _idx_prov_r2,
                                           current_row, _R_COL + _idx_merc_r2, 'TOTAL', r_total_fmt)
                elif _idx_merc_r2 >= 0:
                    worksheet.write(current_row, _R_COL + _idx_merc_r2, 'TOTAL', r_total_fmt)
    
                for _ci, _cn in enumerate(_r_cols):
                    if _ci in (_idx_prov_r2, _idx_merc_r2):
                        continue
                    _abs_col    = _R_COL + _ci
                    _col_letter = gcl(_abs_col + 1)
                    if _cn in ('STOCK TOTAL', 'PUESTOS DE MERCADO') or _cn in _sku_names:
                        formula = f"=SUM({_col_letter}{_c2_first_1b}:{_col_letter}{_c2_last_1b})"
                        worksheet.write_formula(current_row, _abs_col, formula, r_total_num_fmt)
                    else:
                        worksheet.write(current_row, _abs_col, '', r_total_fmt)
                current_row += 1
    
                _ciudad_col_c2 = _R_COL + _idx_prov_r2
                _merge_ciudad_col(worksheet, _hdr_row + 2, current_row - 1,
                                   _ciudad_col_c2, resumen_stock_final, r_body_fmt)
    
                if _r_skus and _idx_total_r2 >= 0:
                    _data_first = _hdr_row + 2
                    _data_last  = current_row - 2
                    if _data_last >= _data_first:
                        worksheet.conditional_format(
                            _data_first, _total_col_idx,
                            _data_last,  _total_col_idx,
                            {
                                'type': 'icon_set',
                                'icon_style': '3_arrows',
                                'icons': [
                                    {'type': 'number', 'value': 300, 'criteria': '>'},
                                    {'type': 'number', 'value': 150, 'criteria': '>'},
                                ]
                            }
                        )
    
    
        # ── Reordenar pestañas: RESUMEN | STORECHECK | STOCK SIN ACTIVIDAD ───────
        workbook.worksheets_objs.sort(
            key=lambda ws: _orden_pestanas.index(ws.name)
            if ws.name in _orden_pestanas else len(_orden_pestanas)
        )
    

    print(json.dumps({"success": True, "output_file": output_file}))

def main():
    if len(sys.argv) < 2:
        print(json.dumps({"error": "Falta ruta del config.json"}))
        sys.exit(1)
    generar_storecheck(sys.argv[1])

if __name__ == '__main__':
    main()
