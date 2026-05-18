package com.trademart.backend.controller;

import com.trademart.backend.dto.LoginRequest;
import com.trademart.backend.dto.LoginResponse;
import com.trademart.backend.security.JwtProvider;
import com.trademart.backend.security.UserDetailsImpl;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.ResponseEntity;
import org.springframework.security.authentication.AuthenticationManager;
import org.springframework.security.authentication.UsernamePasswordAuthenticationToken;
import org.springframework.security.core.Authentication;
import org.springframework.security.core.GrantedAuthority;
import org.springframework.security.core.context.SecurityContextHolder;
import org.springframework.web.bind.annotation.*;

import java.util.List;
import java.util.stream.Collectors;

@RestController
@RequestMapping("/api/auth")
public class AuthController {

    @Autowired
    private AuthenticationManager authenticationManager;

    @Autowired
    private JwtProvider jwtProvider;

    @PostMapping("/login")
    public ResponseEntity<?> authenticateUser(@RequestBody LoginRequest loginRequest) {
        try {
            Authentication authentication = authenticationManager.authenticate(
                    new UsernamePasswordAuthenticationToken(loginRequest.getUsername(), loginRequest.getPassword()));

            SecurityContextHolder.getContext().setAuthentication(authentication);
            String jwt = jwtProvider.generateToken((UserDetailsImpl) authentication.getPrincipal());
            
            UserDetailsImpl userDetails = (UserDetailsImpl) authentication.getPrincipal();
            List<String> roles = userDetails.getAuthorities().stream()
                    .map(GrantedAuthority::getAuthority)
                    .collect(Collectors.toList());

            String primaryRole = roles.isEmpty() ? "usuario" : roles.get(0).replace("ROLE_", "").toLowerCase();

            return ResponseEntity.ok(new LoginResponse(
                    jwt, 
                    userDetails.getUsername(), 
                    userDetails.getUsuario().getEmail(),
                    userDetails.getUsuario().getFirstname(),
                    userDetails.getUsuario().getLastname(),
                    userDetails.getUsuario().getFirstname() + " " + userDetails.getUsuario().getLastname(),
                    primaryRole,
                    roles
            ));
        } catch (org.springframework.security.core.AuthenticationException e) {
            return ResponseEntity.status(org.springframework.http.HttpStatus.UNAUTHORIZED)
                .body(java.util.Map.of("error", "Credenciales incorrectas o usuario deshabilitado"));
        }
    }
}
