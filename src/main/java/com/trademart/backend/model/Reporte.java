package com.trademart.backend.model;

import jakarta.persistence.*;
import lombok.Data;
import java.time.LocalDateTime;
import com.fasterxml.jackson.annotation.JsonIgnoreProperties;

@Data
@Entity
@Table(name = "reporte")
public class Reporte {
    @Id
    @GeneratedValue(strategy = GenerationType.IDENTITY)
    @Column(name = "registro_reporte_id")
    private Integer registroReporteId;

    @Column(name = "pdv")
    private String pdv;

    @Column(name = "puesto")
    private String puesto;

    @Column(name = "fecha_string")
    private String fechaStr;

    @Column(name = "mercaderista")
    private String mercaderista;

    @Column(name = "estado")
    private String estado;

    @Column(name = "skus")
    private Integer skus;

    @Column(name = "tiene_foto")
    private Boolean foto;

    @Column(name = "actividad")
    private String actividad;

    @Column(name = "observaciones", columnDefinition = "TEXT")
    private String observaciones;

    // Legacy fields mapped just in case Supabase schema needs them (nullable)
    @Column(name = "reporte", columnDefinition = "TEXT")
    private String reporte;
    
    @Column(name = "tipo_reporte")
    private String tipoReporte;
    
    @Column(name = "cod_lucky")
    private Integer codLucky;

    @ManyToOne(fetch = FetchType.LAZY)
    @JoinColumn(name = "usuario_id", nullable = true)
    @JsonIgnoreProperties({"hibernateLazyInitializer", "handler"})
    private Usuario usuario;

    @Column(name = "fecha", insertable = false, updatable = false)
    private LocalDateTime fecha;

    @ManyToOne(fetch = FetchType.LAZY)
    @JoinColumn(name = "pm_id", nullable = true)
    @JsonIgnoreProperties({"hibernateLazyInitializer", "handler"})
    private Pm pm;

    @Column(name = "fotos", columnDefinition = "TEXT")
    private String fotos;
}
