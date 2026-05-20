package com.trademart.backend.dto;

import lombok.AllArgsConstructor;
import lombok.Data;
import java.util.List;

@Data
@AllArgsConstructor
public class LoginResponse {
    private String token;
    private Integer id;
    private String username;
    private String email;
    private String firstname;
    private String lastname;
    private String name;
    private String role;
    private List<String> roles;
    private String equipoComercial;
}
