-- ==========================================================
-- 07_producto.sql: SKUs y Productos
-- ==========================================================
INSERT INTO producto (producto_id, nombre, descripcion, marca, categoria, precio, stock_inicial, stock_final, estado) VALUES
(1, 'Opal Producto Especial 1', 'Descripción de producto SKU-001', 'Opal', 'Limpieza', 44.50, 459, 57, true),
(2, 'Opal Producto Especial 2', 'Descripción de producto SKU-002', 'Opal', 'Bebidas', 39.50, 499, 27, true),
(3, 'AlaCena Producto Especial 3', 'Descripción de producto SKU-003', 'AlaCena', 'Bebidas', 9.50, 297, 42, true),
(4, 'Primor Producto Especial 4', 'Descripción de producto SKU-004', 'Primor', 'Cuidado Personal', 38.50, 454, 57, true),
(5, 'AlaCena Producto Especial 5', 'Descripción de producto SKU-005', 'AlaCena', 'Bebidas', 33.50, 284, 29, true),
(6, 'Bolivar Producto Especial 6', 'Descripción de producto SKU-006', 'Bolivar', 'Cuidado Personal', 14.50, 105, 73, true),
(7, 'Sayón Producto Especial 7', 'Descripción de producto SKU-007', 'Sayón', 'Detergentes', 39.50, 470, 57, true),
(8, 'AlaCena Producto Especial 8', 'Descripción de producto SKU-008', 'AlaCena', 'Cuidado Personal', 27.50, 484, 47, true),
(9, 'Primor Producto Especial 9', 'Descripción de producto SKU-009', 'Primor', 'Cuidado Personal', 34.50, 422, 57, true),
(10, 'Sayón Producto Especial 10', 'Descripción de producto SKU-010', 'Sayón', 'Alimentos', 16.50, 205, 78, true),
(11, 'Opal Producto Especial 11', 'Descripción de producto SKU-011', 'Opal', 'Limpieza', 34.50, 363, 29, true),
(12, 'Primor Producto Especial 12', 'Descripción de producto SKU-012', 'Primor', 'Bebidas', 13.50, 228, 40, true),
(13, 'Bolivar Producto Especial 13', 'Descripción de producto SKU-013', 'Bolivar', 'Limpieza', 31.50, 332, 48, true),
(14, 'Bolivar Producto Especial 14', 'Descripción de producto SKU-014', 'Bolivar', 'Bebidas', 31.50, 428, 57, true),
(15, 'Sayón Producto Especial 15', 'Descripción de producto SKU-015', 'Sayón', 'Alimentos', 7.50, 221, 63, true),
(16, 'Bolivar Producto Especial 16', 'Descripción de producto SKU-016', 'Bolivar', 'Limpieza', 44.50, 166, 62, true),
(17, 'AlaCena Producto Especial 17', 'Descripción de producto SKU-017', 'AlaCena', 'Cuidado Personal', 36.50, 376, 20, true),
(18, 'Primor Producto Especial 18', 'Descripción de producto SKU-018', 'Primor', 'Detergentes', 25.50, 281, 21, true),
(19, 'Opal Producto Especial 19', 'Descripción de producto SKU-019', 'Opal', 'Alimentos', 19.50, 186, 29, true),
(20, 'Primor Producto Especial 20', 'Descripción de producto SKU-020', 'Primor', 'Cuidado Personal', 9.50, 168, 67, true)
ON CONFLICT DO NOTHING;

-- Actualizar secuencia de PostgreSQL
SELECT setval('producto_producto_id_seq', (SELECT MAX(producto_id) FROM producto));
