-- ==========================================================
-- 08_act.sql: Actividades Promocionales
-- ==========================================================
INSERT INTO act (act_id, act_promocional, producto_id, estado, tipo, estado_string, inicio, fin, sku_ids, descripcion) VALUES
(1, 'Promoción Especial 1', 1, true, 'Exhibición', 'Vigente', '2026-01-01', '2026-12-31', '1,2,3', 'Exhibición en cabecera de góndola y verificación'),
(2, 'Promoción Especial 2', 2, true, 'Exhibición', 'Vigente', '2026-01-01', '2026-12-31', '1,2,3', 'Exhibición en cabecera de góndola y verificación'),
(3, 'Promoción Especial 3', 3, true, 'Exhibición', 'Vigente', '2026-01-01', '2026-12-31', '1,2,3', 'Exhibición en cabecera de góndola y verificación'),
(4, 'Promoción Especial 4', 4, true, 'Exhibición', 'Vigente', '2026-01-01', '2026-12-31', '1,2,3', 'Exhibición en cabecera de góndola y verificación'),
(5, 'Promoción Especial 5', 5, true, 'Exhibición', 'Vigente', '2026-01-01', '2026-12-31', '1,2,3', 'Exhibición en cabecera de góndola y verificación'),
(6, 'Promoción Especial 6', 6, true, 'Exhibición', 'Vigente', '2026-01-01', '2026-12-31', '1,2,3', 'Exhibición en cabecera de góndola y verificación'),
(7, 'Promoción Especial 7', 7, true, 'Exhibición', 'Vigente', '2026-01-01', '2026-12-31', '1,2,3', 'Exhibición en cabecera de góndola y verificación'),
(8, 'Promoción Especial 8', 8, true, 'Exhibición', 'Vigente', '2026-01-01', '2026-12-31', '1,2,3', 'Exhibición en cabecera de góndola y verificación'),
(9, 'Promoción Especial 9', 9, true, 'Exhibición', 'Vigente', '2026-01-01', '2026-12-31', '1,2,3', 'Exhibición en cabecera de góndola y verificación'),
(10, 'Promoción Especial 10', 10, true, 'Exhibición', 'Vigente', '2026-01-01', '2026-12-31', '1,2,3', 'Exhibición en cabecera de góndola y verificación')
ON CONFLICT DO NOTHING;

-- Actualizar secuencia de PostgreSQL
SELECT setval('act_act_id_seq', (SELECT MAX(act_id) FROM act));
