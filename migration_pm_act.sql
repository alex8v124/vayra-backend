-- =========================================
-- MIGRACIÓN: Relación Muchos a Muchos PM ↔ ACT
-- =========================================
-- Paso 1: Eliminar la FK act_id de la tabla pm
-- (Primero eliminamos la constraint, luego la columna)

ALTER TABLE pm DROP CONSTRAINT IF EXISTS fk_pm_act;
ALTER TABLE pm DROP COLUMN IF EXISTS act_id;

-- Paso 2: Crear tabla intermedia pm_act (N:M)
CREATE TABLE IF NOT EXISTS pm_act (
    pm_act_id SERIAL PRIMARY KEY,
    pm_id INT NOT NULL,
    act_id INT NOT NULL,

    CONSTRAINT fk_pm_act_pm FOREIGN KEY (pm_id)
        REFERENCES pm(pm_id) ON DELETE CASCADE,

    CONSTRAINT fk_pm_act_act FOREIGN KEY (act_id)
        REFERENCES act(act_id) ON DELETE CASCADE,

    CONSTRAINT uq_pm_act UNIQUE (pm_id, act_id)
);

-- Paso 3: Índices para rendimiento
CREATE INDEX IF NOT EXISTS idx_pm_act_pm ON pm_act(pm_id);
CREATE INDEX IF NOT EXISTS idx_pm_act_act ON pm_act(act_id);

-- =========================================
-- ESQUEMA FINAL RESULTANTE:
-- =========================================
-- pm (pm_id, pm_nombre, pdv_id)            → 1 puesto pertenece a 1 PDV
-- pm_act (pm_act_id, pm_id, act_id)        → N:M entre puestos y actividades
-- act (act_id, act_promocional, producto_id, estado, ...)
-- =========================================
