-- ==========================================================
-- 02_equipo_comercial.sql: Tabla equipo_comercial
-- ==========================================================
INSERT INTO equipo_comercial (equipo_id, nombre, descripcion) VALUES
(1, 'Equipo Zona 1', 'Descripción de Zona 1'),
(2, 'Equipo Zona 2', 'Descripción de Zona 2'),
(3, 'Equipo Zona 3', 'Descripción de Zona 3'),
(4, 'Equipo Zona 4', 'Descripción de Zona 4'),
(5, 'Equipo Zona 5', 'Descripción de Zona 5'),
(6, 'Equipo Zona 6', 'Descripción de Zona 6'),
(7, 'Equipo Zona 7', 'Descripción de Zona 7'),
(8, 'Equipo Zona 8', 'Descripción de Zona 8'),
(9, 'Equipo Zona 9', 'Descripción de Zona 9'),
(10, 'Equipo Zona 10', 'Descripción de Zona 10')
ON CONFLICT (equipo_id) DO UPDATE SET nombre = EXCLUDED.nombre;

-- Actualizar secuencia de PostgreSQL
SELECT setval('equipo_comercial_equipo_id_seq', (SELECT MAX(equipo_id) FROM equipo_comercial));
