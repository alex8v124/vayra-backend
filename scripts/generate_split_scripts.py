import os
import random
import json
import datetime

def generate_split():
    base_dir = r"c:\Users\alex0\Downloads\Trabajo\Integrador 2\xplora-backend\scripts\sql"
    os.makedirs(base_dir, exist_ok=True)

    # 00. Limpieza
    s00 = [
        "-- ==========================================================",
        "-- 00_limpieza.sql: Limpieza en orden inverso de dependencias",
        "-- ==========================================================",
        "DELETE FROM pm_act;",
        "DELETE FROM act;",
        "DELETE FROM reporte;",
        "DELETE FROM cronogramas;",
        "DELETE FROM planning;",
        "DELETE FROM pm;",
        "DELETE FROM producto;",
        "DELETE FROM pdv;",
        "DELETE FROM usuario_rol WHERE usuario_id <> 1;",
        "DELETE FROM usuario WHERE usuario_id <> 1;",
        "-- No borramos roles ni equipo_comercial base, pero los actualizaremos con ON CONFLICT",
        ""
    ]
    with open(os.path.join(base_dir, "00_limpieza.sql"), "w", encoding="utf-8") as f:
        f.write("\n".join(s00))

    # 01. Roles
    s01 = [
        "-- ==========================================================",
        "-- 01_roles.sql: Tabla roles",
        "-- ==========================================================",
        "INSERT INTO roles (rol_id, rol_nombre) VALUES",
        "(1, 'ADMIN'),",
        "(2, 'GERENTE'),",
        "(3, 'SUPERVISOR'),",
        "(4, 'MERCADERISTA')",
        "ON CONFLICT (rol_id) DO UPDATE SET rol_nombre = EXCLUDED.rol_nombre;",
        "",
        "-- Actualizar secuencia de PostgreSQL",
        "SELECT setval('roles_rol_id_seq', (SELECT MAX(rol_id) FROM roles));",
        ""
    ]
    with open(os.path.join(base_dir, "01_roles.sql"), "w", encoding="utf-8") as f:
        f.write("\n".join(s01))

    # 02. Equipo Comercial
    s02 = [
        "-- ==========================================================",
        "-- 02_equipo_comercial.sql: Tabla equipo_comercial",
        "-- ==========================================================",
        "INSERT INTO equipo_comercial (equipo_id, nombre, descripcion) VALUES"
    ]
    equipos = []
    for i in range(1, 11):
        equipos.append(f"({i}, 'Equipo Zona {i}', 'Descripción de Zona {i}')")
    s02.append(",\n".join(equipos))
    s02.append("ON CONFLICT (equipo_id) DO UPDATE SET nombre = EXCLUDED.nombre;")
    s02.append("")
    s02.append("-- Actualizar secuencia de PostgreSQL")
    s02.append("SELECT setval('equipo_comercial_equipo_id_seq', (SELECT MAX(equipo_id) FROM equipo_comercial));")
    s02.append("")
    with open(os.path.join(base_dir, "02_equipo_comercial.sql"), "w", encoding="utf-8") as f:
        f.write("\n".join(s02))

    # 03 & 04. Usuario y Usuario_Rol
    names = ['Juan', 'Maria', 'Carlos', 'Ana', 'Luis', 'Elena', 'Pedro', 'Laura', 'Jorge', 'Sofia', 'Diego', 'Lucia', 'Miguel', 'Carmen', 'Jose', 'Paula', 'Fernando', 'Rosa', 'Ricardo', 'Gabriela']
    surnames = ['Rodriguez', 'Lopez', 'Garcia', 'Torres', 'Martinez', 'Perez', 'Sanchez', 'Romero', 'Diaz', 'Fernandez', 'Gomez', 'Ruiz', 'Alvarez', 'Molina', 'Cano', 'Vargas', 'Castro', 'Ramos', 'Chavez']
    
    usuarios = []
    usuario_roles = []
    mercaderistas = []
    supervisores = []
    user_id = 10
    for i in range(1, 201):
        fname = random.choice(names)
        lname = random.choice(surnames)
        full_name = f"{fname} {lname}"
        username = f"{fname[:1].lower()}{lname.lower()}{i}"
        email = f"{username}@trademart.com"
        eq = f"Equipo Zona {random.randint(1, 10)}"
        
        usuarios.append(f"({user_id}, '{username}', '$2a$10$xyz', '{fname}', '{lname}', '{email}', true, '{eq}')")
        
        rol_id = 3 if i % 6 == 0 else 4
        usuario_roles.append(f"({user_id}, {user_id}, {rol_id})")
        if rol_id == 4:
            mercaderistas.append((user_id, full_name))
        else:
            supervisores.append((user_id, full_name))
        user_id += 1

    s03 = [
        "-- ==========================================================",
        "-- 03_usuario.sql: Tabla usuario (Mercaderistas y Supervisores)",
        "-- ==========================================================",
        "INSERT INTO usuario (usuario_id, username, password, firstname, lastname, email, estado, equipo_comercial) VALUES",
        ",\n".join(usuarios),
        "ON CONFLICT DO NOTHING;",
        "",
        "-- Actualizar secuencia de PostgreSQL",
        "SELECT setval('usuario_usuario_id_seq', (SELECT MAX(usuario_id) FROM usuario));",
        ""
    ]
    with open(os.path.join(base_dir, "03_usuario.sql"), "w", encoding="utf-8") as f:
        f.write("\n".join(s03))

    s04 = [
        "-- ==========================================================",
        "-- 04_usuario_rol.sql: Relación usuario - rol",
        "-- ==========================================================",
        "INSERT INTO usuario_rol (usu_rol_id, usuario_id, rol_id) VALUES",
        ",\n".join(usuario_roles),
        "ON CONFLICT DO NOTHING;",
        "",
        "-- Actualizar secuencia de PostgreSQL",
        "SELECT setval('usuario_rol_usu_rol_id_seq', (SELECT MAX(usu_rol_id) FROM usuario_rol));",
        ""
    ]
    with open(os.path.join(base_dir, "04_usuario_rol.sql"), "w", encoding="utf-8") as f:
        f.write("\n".join(s04))

    # 05. PDV
    pdvs = []
    pdv_map_names = {}
    distritos = ['Surco', 'Miraflores', 'San Isidro', 'Lima Cercado', 'Los Olivos', 'San Borja', 'La Molina', 'Chorrillos', 'Magdalena', 'Jesus Maria']
    tipos = ['Mercado Tradicional', 'Bodega', 'Mayorista', 'Supermercado']
    estados_pdv = ['Activo', 'Activo', 'Activo', 'Activo', 'Activo', 'Inactivo']
    
    for i in range(1, 501):
        dist = random.choice(distritos)
        tipo = random.choice(tipos)
        estado = random.choice(estados_pdv)
        pend = random.choice(['true', 'false'])
        vis = random.randint(0, 15)
        nombre_pdv = f"MERCADO {dist.upper()} {i}"
        pdv_map_names[i] = nombre_pdv
        pdvs.append(f"({i}, '{nombre_pdv}', 'COD-{i:03d}', '{dist}', '{tipo}', '{estado}', {vis}, {pend})")

    s05 = [
        "-- ==========================================================",
        "-- 05_pdv.sql: Puntos de Venta (PDVs / Mercados)",
        "-- ==========================================================",
        "INSERT INTO pdv (pdv_id, pdv_nombre, codigo, distrito, tipo, estado, visitas, pendiente) VALUES",
        ",\n".join(pdvs),
        "ON CONFLICT DO NOTHING;",
        "",
        "-- Actualizar secuencia de PostgreSQL",
        "SELECT setval('pdv_pdv_id_seq', (SELECT MAX(pdv_id) FROM pdv));",
        ""
    ]
    with open(os.path.join(base_dir, "05_pdv.sql"), "w", encoding="utf-8") as f:
        f.write("\n".join(s05))

    # 06. PM
    pms = []
    pm_counter = 1
    pdv_puestos_map = {}
    for pdv_id in range(1, 501):
        num_puestos = random.randint(2, 6)
        pdv_puestos_map[pdv_id] = []
        for p in range(1, num_puestos + 1):
            num_str = f"P-{p:02d}"
            pms.append(f"({pm_counter}, 'Puesto {num_str} Abarrotes', {pdv_id})")
            pdv_puestos_map[pdv_id].append((pm_counter, num_str))
            pm_counter += 1

    s06 = [
        "-- ==========================================================",
        "-- 06_pm.sql: Puestos de Mercado (PM)",
        "-- ==========================================================",
        "INSERT INTO pm (pm_id, pm_nombre, pdv_id) VALUES",
        ",\n".join(pms),
        "ON CONFLICT DO NOTHING;",
        "",
        "-- Actualizar secuencia de PostgreSQL",
        "SELECT setval('pm_pm_id_seq', (SELECT MAX(pm_id) FROM pm));",
        ""
    ]
    with open(os.path.join(base_dir, "06_pm.sql"), "w", encoding="utf-8") as f:
        f.write("\n".join(s06))

    # 07. Producto
    productos = []
    categorias = ['Cuidado Personal', 'Detergentes', 'Alimentos', 'Limpieza', 'Bebidas']
    marcas = ['Bolivar', 'Opal', 'AlaCena', 'Primor', 'Sayón']
    for i in range(1, 21):
        mar = random.choice(marcas)
        cat = random.choice(categorias)
        productos.append(f"({i}, '{mar} Producto Especial {i}', 'Descripción de producto SKU-{i:03d}', '{mar}', '{cat}', {random.randint(5, 45)}.50, {random.randint(100, 500)}, {random.randint(20, 80)}, true)")

    s07 = [
        "-- ==========================================================",
        "-- 07_producto.sql: SKUs y Productos",
        "-- ==========================================================",
        "INSERT INTO producto (producto_id, nombre, descripcion, marca, categoria, precio, stock_inicial, stock_final, estado) VALUES",
        ",\n".join(productos),
        "ON CONFLICT DO NOTHING;",
        "",
        "-- Actualizar secuencia de PostgreSQL",
        "SELECT setval('producto_producto_id_seq', (SELECT MAX(producto_id) FROM producto));",
        ""
    ]
    with open(os.path.join(base_dir, "07_producto.sql"), "w", encoding="utf-8") as f:
        f.write("\n".join(s07))

    # 08. Act
    acts = []
    for i in range(1, 11):
        acts.append(f"({i}, 'Promoción Especial {i}', {i}, true, 'Exhibición', 'Vigente', '2026-01-01', '2026-12-31', '1,2,3', 'Exhibición en cabecera de góndola y verificación')")

    s08 = [
        "-- ==========================================================",
        "-- 08_act.sql: Actividades Promocionales",
        "-- ==========================================================",
        "INSERT INTO act (act_id, act_promocional, producto_id, estado, tipo, estado_string, inicio, fin, sku_ids, descripcion) VALUES",
        ",\n".join(acts),
        "ON CONFLICT DO NOTHING;",
        "",
        "-- Actualizar secuencia de PostgreSQL",
        "SELECT setval('act_act_id_seq', (SELECT MAX(act_id) FROM act));",
        ""
    ]
    with open(os.path.join(base_dir, "08_act.sql"), "w", encoding="utf-8") as f:
        f.write("\n".join(s08))

    # 09. PM_ACT
    pm_acts = []
    pm_act_id = 1
    for pm_id in range(1, min(pm_counter, 1501)):
        for act_id in [1, 2, 3]:
            pm_acts.append(f"({pm_act_id}, {pm_id}, {act_id})")
            pm_act_id += 1

    s09 = [
        "-- ==========================================================",
        "-- 09_pm_act.sql: Relación puestos y actividades",
        "-- ==========================================================",
        "INSERT INTO pm_act (pm_act_id, pm_id, act_id) VALUES",
        ",\n".join(pm_acts),
        "ON CONFLICT DO NOTHING;",
        "",
        "-- Actualizar secuencia de PostgreSQL",
        "SELECT setval('pm_act_pm_act_id_seq', (SELECT MAX(pm_act_id) FROM pm_act));",
        ""
    ]
    with open(os.path.join(base_dir, "09_pm_act.sql"), "w", encoding="utf-8") as f:
        f.write("\n".join(s09))

    # 10. Planning
    plannings = []
    estados_plan = ['Pendiente', 'Completado', 'En curso', 'Pendiente']
    dias_nombres = ['Lunes', 'Martes', 'Miercoles', 'Jueves', 'Viernes', 'Sabado']
    for idx, (m_id, m_name) in enumerate(mercaderistas, 1):
        pdv_id = ((idx - 1) % 500) + 1
        estado = random.choice(estados_plan)
        puestos = pdv_puestos_map.get(pdv_id, [])
        pm_ids_list = [str(pm[0]) for pm in puestos[:3]] if puestos else ['1']
        pm_ids_str = ",".join(pm_ids_list)
        dias_map = {pm_id: random.choice(dias_nombres) for pm_id in pm_ids_list}
        dias_json = json.dumps(dias_map).replace("'", "''")
        plannings.append(f"({idx}, {pdv_id}, {m_id}, '{pm_ids_str}', '{dias_json}', '1,2,3', '2026-01-01', '2026-12-31', '{estado}')")

    s10 = [
        "-- ==========================================================",
        "-- 10_planning.sql: Plannings (1 por mercaderista)",
        "-- ==========================================================",
        "INSERT INTO planning (planning_id, pdv_id, usuario_id, pm_ids, dias_semana_pms, act_ids, fecha_inicio, fecha_fin, estado) VALUES",
        ",\n".join(plannings),
        "ON CONFLICT DO NOTHING;",
        "",
        "-- Actualizar secuencia de PostgreSQL",
        "SELECT setval('planning_planning_id_seq', (SELECT MAX(planning_id) FROM planning));",
        ""
    ]
    with open(os.path.join(base_dir, "10_planning.sql"), "w", encoding="utf-8") as f:
        f.write("\n".join(s10))

    # 11. Cronogramas
    cronos = []
    for idx, (m_id, m_name) in enumerate(mercaderistas[:30], 1):
        pdv_id = ((idx - 1) % 500) + 1
        puestos = pdv_puestos_map.get(pdv_id, [(1, "P-01")])
        nombre_mercado = pdv_map_names.get(pdv_id, "MERCADO CALIENTES")
        datos_list = []
        base_date = datetime.date(2026, 1, 5)
        for w in range(26):
            curr_date = base_date + datetime.timedelta(weeks=w)
            fecha_str = curr_date.strftime("%Y-%m-%d")
            for p_id, p_num in puestos[:2]:
                datos_list.append({
                    "fecha": fecha_str,
                    "Ciudad": "",
                    "MERCADO": nombre_mercado,
                    "DEX": "",
                    "nroPuesto": p_num,
                    "encargado": m_name
                })
        json_data = json.dumps(datos_list, ensure_ascii=False).replace("'", "''")
        cronos.append(f"({idx}, 'Cronograma Mercaderista {m_name} - 2026', '2026-01-01', '2026-12-31', '{idx}', '{json_data}')")

    s11 = [
        "-- ==========================================================",
        "-- 11_cronogramas.sql: Cronogramas con formato exacto",
        "-- ==========================================================",
        "INSERT INTO cronogramas (id, nombre, fecha_inicio, fecha_fin, planning_ids, datos_json) VALUES",
        ",\n".join(cronos),
        "ON CONFLICT DO NOTHING;",
        "",
        "-- Actualizar secuencia de PostgreSQL",
        "SELECT setval('cronogramas_id_seq', (SELECT MAX(id) FROM cronogramas));",
        ""
    ]
    with open(os.path.join(base_dir, "11_cronogramas.sql"), "w", encoding="utf-8") as f:
        f.write("\n".join(s11))

    # 12. Reportes
    reportes = []
    estados_sc = ['Completado', 'Pendiente', 'En proceso', 'Completado', 'Observado']
    start_date = datetime.date(2026, 1, 1)
    marcas_sc = ['Primor', 'Opal', 'AlaCena', 'Bolivar', 'Sayón']
    for i in range(1, 971):
        day_offset = random.randint(0, 300)
        sc_date = start_date + datetime.timedelta(days=day_offset)
        f_str = sc_date.strftime("%Y-%m-%d")
        f_ts = f"{f_str} 10:{random.randint(10, 59):02d}:00"
        pdv_id = random.randint(1, 500)
        nombre_pdv = pdv_map_names.get(pdv_id, f"MERCADO {pdv_id}")
        puestos = pdv_puestos_map.get(pdv_id, [(1, "P-01")])
        puesto_elegido = random.choice(puestos)
        pm_id = puesto_elegido[0]
        p_num = puesto_elegido[1]
        merc = random.choice(mercaderistas)
        merc_id = merc[0]
        merc_name = merc[1]
        estado = random.choice(estados_sc)
        obs = "Todo en orden durante la revisión" if estado == 'Completado' else ("Revisión en curso" if estado == 'En proceso' else "Falta verificación de stock")
        
        # Generar JSON exacto con formato [{nombre:..., stockInicial:..., stockFinal:...}]
        skus_num = random.randint(3, 6)
        productos_json = []
        for p_idx in range(1, skus_num + 1):
            s_init = random.randint(4, 30)
            s_final = random.randint(0, s_init)
            m_elegida = random.choice(marcas_sc)
            productos_json.append({
                "nombre": f"{m_elegida} Producto Especial {p_idx}",
                "stockInicial": s_init,
                "stockFinal": s_final
            })
        reporte_json_str = json.dumps(productos_json, ensure_ascii=False).replace("'", "''")
        
        reportes.append(f"({i}, '{nombre_pdv}', '{p_num}', '{f_str}', '{merc_name}', '{estado}', {skus_num}, true, 'Variación de stock', '{obs}', '{reporte_json_str}', 'R1', {1000+i}, {merc_id}, {pm_id}, '{f_ts}')")

    s12 = [
        "-- ==========================================================",
        "-- 12_reporte.sql: Reportes (Storechecks)",
        "-- ==========================================================",
        "INSERT INTO reporte (registro_reporte_id, pdv, puesto, fecha_string, mercaderista, estado, skus, tiene_foto, actividad, observaciones, reporte, tipo_reporte, cod_lucky, usuario_id, pm_id, fecha) VALUES",
        ",\n".join(reportes),
        "ON CONFLICT DO NOTHING;",
        "",
        "-- Actualizar secuencia de PostgreSQL",
        "SELECT setval('reporte_registro_reporte_id_seq', (SELECT MAX(registro_reporte_id) FROM reporte));",
        ""
    ]
    with open(os.path.join(base_dir, "12_reporte.sql"), "w", encoding="utf-8") as f:
        f.write("\n".join(s12))

    # Master script that joins them all cleanly
    master = [
        "-- ==============================================================================",
        "-- 99_master_seed.sql: Ejecución consolidada de todos los scripts individuales",
        "-- Carpeta oficial: xplora-backend/scripts/sql/",
        "-- ==============================================================================",
        ""
    ]
    for s_list in [s00, s01, s02, s03, s04, s05, s06, s07, s08, s09, s10, s11, s12]:
        master.extend(s_list)
        master.append("\n")

    with open(os.path.join(base_dir, "99_master_seed.sql"), "w", encoding="utf-8") as f:
        f.write("\n".join(master))

    print(f"Generated 13 split SQL files perfectly inside {base_dir}")

if __name__ == "__main__":
    generate_split()
