-- ==========================================================
-- 01_roles.sql: Tabla roles
-- ==========================================================
INSERT INTO roles (rol_id, rol_nombre) VALUES
(1, 'ADMIN'),
(2, 'GERENTE'),
(3, 'SUPERVISOR'),
(4, 'MERCADERISTA')
ON CONFLICT (rol_id) DO UPDATE SET rol_nombre = EXCLUDED.rol_nombre;

-- Actualizar secuencia de PostgreSQL
SELECT setval('roles_rol_id_seq', (SELECT MAX(rol_id) FROM roles));
