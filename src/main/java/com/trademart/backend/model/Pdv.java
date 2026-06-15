package com.trademart.backend.model;

import jakarta.persistence.*;
import lombok.Data;

@Data
@Entity
@Table(name = "pdv")
public class Pdv {
    @Id
    @GeneratedValue(strategy = GenerationType.IDENTITY)
    @Column(name = "pdv_id")
    private Integer pdvId;

    @Column(name = "pdv_nombre", nullable = false)
    private String pdvNombre;

    @Column(name = "codigo")
    private String codigo;

    @Column(name = "distrito")
    private String distrito;

    @Column(name = "tipo")
    private String tipo;

    @Column(name = "estado")
    private String estado;

    @Column(name = "visitas")
    private Integer visitas;

    @Column(name = "pendiente")
    private Boolean pendiente;
}
