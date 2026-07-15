package com.trademart.backend.model;

import jakarta.persistence.*;
import lombok.Data;
import java.time.LocalDateTime;

@Data
@Entity
@Table(name = "usuario")
public class Usuario {
    @Id
    @GeneratedValue(strategy = GenerationType.IDENTITY)
    @Column(name = "usuario_id")
    private Integer usuarioId;

    @Column(nullable = false, unique = true)
    private String username;

    @Column(nullable = false)
    private String password;

    private String firstname;
    private String lastname;
    
    @Column(unique = true)
    private String email;

    private Boolean estado = true;

    @Column(name = "equipo_comercial")
    private String equipoComercial;

    @Column(name = "pdvs_asignados", columnDefinition = "TEXT")
    private String pdvsAsignados;

    @Column(name = "fecha_creacion", insertable = false, updatable = false)
    private LocalDateTime fechaCreacion;

    @Column(name = "ultimo_login")
    private LocalDateTime ultimoLogin;
}
