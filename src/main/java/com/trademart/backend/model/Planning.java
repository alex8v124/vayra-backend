package com.trademart.backend.model;

import jakarta.persistence.*;
import lombok.Data;

@Data
@Entity
@Table(name = "planning")
public class Planning {
    @Id
    @GeneratedValue(strategy = GenerationType.IDENTITY)
    @Column(name = "planning_id")
    private Integer planningId;

    @Column(name = "usuario_id", nullable = false)
    private Integer usuarioId; // Mercaderista asignado

    @Column(name = "pdv_id", nullable = false)
    private Integer pdvId; // PDV asignado

    @Column(name = "pm_ids")
    private String pmIds; // Lista de IDs de puestos de mercado asignados, ej: "1,2,3"

    @Column(name = "dias_semana_pms", columnDefinition = "TEXT")
    private String diasSemanaPms; // JSON mapeando pmId a día de la semana, ej: "{\"1\":\"Lunes\", \"2\":\"Martes\"}"

    @Column(name = "act_ids")
    private String actIds; // Lista de IDs de actividades promocionales asignadas, ej: "10,11"

    @Column(name = "fecha_inicio")
    private String fechaInicio; // Formato YYYY-MM-DD

    @Column(name = "fecha_fin")
    private String fechaFin; // Formato YYYY-MM-DD

    @Column(name = "estado")
    private String estado = "Pendiente"; // Pendiente, Completado
}
