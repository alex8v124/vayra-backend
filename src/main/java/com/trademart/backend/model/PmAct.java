package com.trademart.backend.model;

import jakarta.persistence.*;
import lombok.Data;

@Data
@Entity
@Table(name = "pm_act")
public class PmAct {
    @Id
    @GeneratedValue(strategy = GenerationType.IDENTITY)
    @Column(name = "pm_act_id")
    private Integer pmActId;

    @Column(name = "pm_id", nullable = false)
    private Integer pmId;

    @Column(name = "act_id", nullable = false)
    private Integer actId;
}
