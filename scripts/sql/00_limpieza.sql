-- ==========================================================
-- 00_limpieza.sql: Limpieza en orden inverso de dependencias
-- ==========================================================
DELETE FROM pm_act;
DELETE FROM act;
DELETE FROM reporte;
DELETE FROM cronogramas;
DELETE FROM planning;
DELETE FROM pm;
DELETE FROM producto;
DELETE FROM pdv;
DELETE FROM usuario_rol WHERE usuario_id <> 1;
DELETE FROM usuario WHERE usuario_id <> 1;
-- No borramos roles ni equipo_comercial base, pero los actualizaremos con ON CONFLICT
