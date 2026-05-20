package com.trademart.backend.model;

import jakarta.persistence.*;
import lombok.Data;

@Data
@Entity
@Table(name = "equipo_comercial")
public class EquipoComercial {
    @Id
    @GeneratedValue(strategy = GenerationType.IDENTITY)
    @Column(name = "equipo_id")
    private Integer equipoId;

    @Column(nullable = false, unique = true)
    private String nombre;

    private String descripcion;
}
