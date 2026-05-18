package com.trademart.backend.model;

import jakarta.persistence.*;
import lombok.Data;

@Data
@Entity
@Table(name = "pm")
public class Pm {
    @Id
    @GeneratedValue(strategy = GenerationType.IDENTITY)
    @Column(name = "pm_id")
    private Integer pmId;

    @Column(name = "pm_nombre")
    private String pmNombre;

    @Column(name = "pdv_id", nullable = false)
    private Integer pdvId;
}
