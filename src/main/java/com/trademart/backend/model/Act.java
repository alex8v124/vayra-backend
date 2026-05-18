package com.trademart.backend.model;

import jakarta.persistence.*;
import lombok.Data;

@Data
@Entity
@Table(name = "act")
public class Act {
    @Id
    @GeneratedValue(strategy = GenerationType.IDENTITY)
    @Column(name = "act_id")
    private Integer actId;

    @Column(name = "act_promocional")
    private String nombre;

    @Column(name = "producto_id", nullable = false)
    private Integer productoId;

    @Column(name = "estado")
    private Boolean estadoBool = true;

    @Column(name = "tipo")
    private String tipo;

    @Column(name = "estado_string")
    private String estado;

    @Column(name = "inicio")
    private String inicio;

    @Column(name = "fin")
    private String fin;

    @Column(name = "sku_ids")
    private String skuIds;

    @Column(name = "descripcion")
    private String descripcion;
}
