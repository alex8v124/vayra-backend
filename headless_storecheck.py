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

    hojas = {
        "RESUMEN": df_modificado_v10,
        "STORECHECK": pv_tb2,
        "STOCK SIN ACTIVIDAD": pv,
    }

    with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
        workbook = writer.book
        
        header_format = workbook.add_format({
            'bold': True, 'text_wrap': True, 'valign': 'vcenter', 'align': 'center',
            'fg_color': '#1F497D', 'font_color': 'white', 'border': 1, 'font_size': 12, 'font_name': 'Aptos'
        })
        rotated_header_format = workbook.add_format({
            'bold': False, 'text_wrap': True, 'align': 'center', 'valign': 'vcenter',
            'fg_color': '#FFFFFF', 'font_color': '#000000', 'border': 1, 'rotation': 90, 'font_size': 12, 'font_name': 'Aptos'
        })
        sa_header_format = workbook.add_format({
            'bold': True, 'text_wrap': True, 'valign': 'vcenter', 'align': 'center',
            'fg_color': '#1F497D', 'font_color': 'white', 'border': 1, 'font_size': 14, 'font_name': 'Aptos'
        })
        sa_group_header_format = workbook.add_format({
            'bold': True, 'text_wrap': True, 'valign': 'vcenter', 'align': 'center',
            'fg_color': '#1F497D', 'font_color': 'white', 'border': 1, 'font_size': 14, 'font_name': 'Aptos'
        })
        sa_sku_format = workbook.add_format({
            'bold': False, 'text_wrap': True, 'align': 'center', 'valign': 'vcenter',
            'fg_color': '#FFFFFF', 'font_color': '#000000', 'border': 1, 'rotation': 90, 'font_size': 14, 'font_name': 'Aptos'
        })
        hdr_presencia_alicorp = workbook.add_format({
            'bold': True, 'text_wrap': True, 'align': 'center', 'valign': 'vcenter',
            'fg_color': '#404040', 'font_color': 'white', 'border': 1, 'rotation': 90, 'font_size': 12, 'font_name': 'Aptos'
        })
        hdr_presencia_producto = workbook.add_format({
            'bold': True, 'text_wrap': True, 'align': 'center', 'valign': 'vcenter',
            'fg_color': '#D9D9D9', 'font_color': '#000000', 'border': 1, 'rotation': 90, 'font_size': 12, 'font_name': 'Aptos'
        })
        hdr_rotated_sc = workbook.add_format({
            'bold': True, 'text_wrap': True, 'align': 'center', 'valign': 'vcenter',
            'fg_color': '#1F497D', 'font_color': 'white', 'border': 1, 'rotation': 90, 'font_size': 12, 'font_name': 'Aptos'
        })
        hdr_marca_sc = workbook.add_format({
            'bold': True, 'text_wrap': True, 'align': 'center', 'valign': 'vcenter',
            'fg_color': '#1F497D', 'font_color': 'white', 'border': 1, 'font_size': 12, 'font_name': 'Aptos'
        })
        hdr_sku_sc = workbook.add_format({
            'bold': False, 'text_wrap': True, 'align': 'center', 'valign': 'vcenter',
            'fg_color': '#FFFFFF', 'font_color': '#000000', 'border': 1, 'rotation': 90, 'font_size': 14, 'font_name': 'Aptos'
        })

        SA_WIDTHS = {'N°': 7.3, 'FECHA': 21.3, 'PROVINCIA': 21.3, 'DEX': 21.3, 'MERCADO': 74.3, 'NOMBRE CLIENTE': 69.0, 'N° de puesto': 21.1}
        SA_SKU_WIDTH = 10.9

        for sheet_name, df in hojas.items():
            if df.empty:
                continue
            
            _nombres_display = {
                "RESUMEN": "Resumen",
                "STORECHECK": "Storecheck",
                "STOCK SIN ACTIVIDAD": "Stock sin actividad",
            }
            worksheet = workbook.add_worksheet(_nombres_display.get(sheet_name, sheet_name))
            worksheet.hide_gridlines(2)
            
            if sheet_name == "STORECHECK":
                worksheet.set_row(0, 25)
                worksheet.set_row(1, 152)
                cols_index_sc = ['N°', 'FECHA', 'PROVINCIA', 'DEX', 'MERCADO', 'NOMBRE CLIENTE', 'N° Puesto', columna_promocion_interna, 'Alicorp', 'Otros', 'Presencia en Puestos Alicorp', 'Presencia Producto']
                productos_sc = [c for c in df.columns if c not in cols_index_sc and not isinstance(c, tuple)]
                
                col_num = 0
                simple_cols = ['N°', 'FECHA', 'PROVINCIA', 'MERCADO', 'NOMBRE CLIENTE', 'N° Puesto', columna_promocion_interna]
                simple_widths = [8.3, 19.7, 27.6, 112.9, 62.6, 19.1, 16.1]
                simple_display = ['N°', 'FECHA', 'PROVINCIA', 'MERCADO', 'NOMBRE CLIENTE', 'N° Puesto', columna_promocion_display]
                
                for name, display, w in zip(simple_cols, simple_display, simple_widths):
                    if name in df.columns:
                        worksheet.merge_range(0, col_num, 1, col_num, display, header_format)
                        worksheet.set_column(col_num, col_num, w)
                        col_num += 1

                abast_start = col_num
                worksheet.merge_range(0, abast_start, 0, abast_start + 1, 'Abastecido por:', header_format)
                worksheet.write(1, abast_start, 'Alicorp', hdr_rotated_sc)
                worksheet.set_column(abast_start, abast_start, 10.9)
                worksheet.write(1, abast_start + 1, 'Otros', hdr_rotated_sc)
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
                        worksheet.merge_range(0, sku_start, 0, sku_start + sku_count - 1, marca_storecheck, hdr_marca_sc)
                    else:
                        worksheet.write(0, sku_start, marca_storecheck, hdr_marca_sc)
                    for sku in productos_sc:
                        worksheet.write(1, col_num, sku, hdr_sku_sc)
                        worksheet.set_column(col_num, col_num, 11.1)
                        col_num += 1

                for row_num, row_data in enumerate(df.values):
                    for i, val in enumerate(row_data):
                        worksheet.write(row_num + 2, i, val if not pd.isna(val) else '')

            else:
                is_pivot = any(isinstance(c, tuple) for c in df.columns)
                if sheet_name == "STOCK SIN ACTIVIDAD":
                    worksheet.set_row(0, 25)
                    worksheet.set_row(1, 190)
                    _hdr_main = sa_header_format
                    _hdr_group = sa_group_header_format
                    _hdr_sku = sa_sku_format
                else:
                    worksheet.set_row(0, 25)
                    worksheet.set_row(1, 60)
                    _hdr_main = header_format
                    _hdr_group = header_format
                    _hdr_sku = rotated_header_format

                if not is_pivot:
                    if sheet_name != "RESUMEN":
                        for col_num, value in enumerate(df.columns.values):
                            worksheet.write(0, col_num, value, _hdr_main)
                            col_width = max(len(str(value)) + 2, 10)
                            worksheet.set_column(col_num, col_num, col_width)
                    for row_num, row_data in enumerate(df.values):
                        for i, val in enumerate(row_data):
                            worksheet.write(row_num + 1, i, val if not pd.isna(val) else '')
                else:
                    for col_num, col_name in enumerate(df.columns):
                        if col_name == '_separador_':
                            worksheet.write(0, col_num, '', None)
                            worksheet.write(1, col_num, '', None)
                            worksheet.set_column(col_num, col_num, 2)
                        elif isinstance(col_name, tuple) and len(col_name) > 1 and col_name[0] not in ['N°', 'FECHA', 'PROVINCIA', 'DEX', 'MERCADO', 'NOMBRE CLIENTE', 'N° de puesto', 'Alicorp', 'Otros', 'Presencia en Puestos Alicorp', 'Presencia Producto']:
                            worksheet.write(0, col_num, col_name[0], _hdr_group)
                            worksheet.write(1, col_num, col_name[1], _hdr_sku)
                            worksheet.set_column(col_num, col_num, SA_SKU_WIDTH)
                        else:
                            display_name = col_name[0] if isinstance(col_name, tuple) and len(col_name) > 0 else col_name
                            worksheet.merge_range(0, col_num, 1, col_num, display_name, _hdr_main)
                            w = SA_WIDTHS.get(display_name, SA_SKU_WIDTH) if sheet_name == "STOCK SIN ACTIVIDAD" else max(len(str(display_name)) + 2, 12)
                            worksheet.set_column(col_num, col_num, w)

                    groups = []
                    current_header = None
                    current_start = None
                    current_count = 0
                    for col_num, col_name in enumerate(df.columns):
                        if isinstance(col_name, tuple) and len(col_name) == 2 and col_name[1] != '':
                            level0 = col_name[0]
                            if level0 == current_header:
                                current_count += 1
                            else:
                                if current_header is not None:
                                    groups.append((current_header, current_start, current_count))
                                current_header = level0
                                current_start = col_num
                                current_count = 1
                        else:
                            if current_header is not None:
                                groups.append((current_header, current_start, current_count))
                                current_header = None
                                current_start = None
                                current_count = 0
                    if current_header is not None:
                        groups.append((current_header, current_start, current_count))

                    for header, start, count in groups:
                        if count > 1:
                            worksheet.merge_range(0, start, 0, start + count - 1, header, _hdr_group)
                        else:
                            worksheet.write(0, start, header, _hdr_group)

                    for row_num, row_data in enumerate(df.values):
                        for i, val in enumerate(row_data):
                            worksheet.write(row_num + 2, i, val if not pd.isna(val) else '')

    print(json.dumps({"success": True, "output_file": output_file}))

if __name__ == '__main__':
    if len(sys.argv) < 2:
        print(json.dumps({"error": "Falta la ruta al archivo JSON de configuración"}))
        sys.exit(1)
    
    config_path = sys.argv[1]
    if not os.path.exists(config_path):
        print(json.dumps({"error": f"Archivo no encontrado: {config_path}"}))
        sys.exit(1)
        
    generar_storecheck(config_path)
