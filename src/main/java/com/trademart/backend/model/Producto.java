package com.trademart.backend.model;

import jakarta.persistence.*;
import lombok.Data;
import java.math.BigDecimal;
import java.time.LocalDateTime;

@Data
@Entity
@Table(name = "producto")
public class Producto {
    @Id
    @GeneratedValue(strategy = GenerationType.IDENTITY)
    @Column(name = "producto_id")
    private Integer productoId;

    @Column(nullable = false)
    private String nombre;

    private String descripcion;
    private String marca;

    @Column(name = "categoria")
    private String categoria;
    
    @Column(nullable = false)
    private BigDecimal precio;

    @Column(name = "stock_inicial")
    private Integer stockInicial = 0;

    @Column(name = "stock_final")
    private Integer stockFinal = 0;

    private Boolean estado = true;

    @Column(name = "fecha_registro", insertable = false, updatable = false)
    private LocalDateTime fechaRegistro;
}
