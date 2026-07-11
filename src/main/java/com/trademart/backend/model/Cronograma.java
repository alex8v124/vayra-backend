package com.trademart.backend.model;

import jakarta.persistence.*;
import lombok.Data;

@Data
@Entity
@Table(name = "cronogramas")
public class Cronograma {
    @Id
    @GeneratedValue(strategy = GenerationType.IDENTITY)
    private Long id;

    @Column(nullable = false)
    private String nombre;

    @Column(name = "planning_ids")
    private String planningIds;
    
    private String fechaInicio;
    
    private String fechaFin;

    @Column(columnDefinition = "TEXT", nullable = false)
    private String datosJson;
}
