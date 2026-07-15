package com.trademart.backend.controller;

import com.trademart.backend.model.Usuario;
import com.trademart.backend.model.Rol;
import com.trademart.backend.model.UsuarioRol;
import com.trademart.backend.repository.UsuarioRepository;
import com.trademart.backend.repository.RolRepository;
import com.trademart.backend.repository.UsuarioRolRepository;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.security.crypto.password.PasswordEncoder;
import org.springframework.web.bind.annotation.*;
import java.util.List;
import java.util.Map;
import java.util.stream.Collectors;

@RestController
@RequestMapping("/api/usuarios")
public class UsuarioController {

    @Autowired
    private UsuarioRepository usuarioRepository;
    
    @Autowired
    private RolRepository rolRepository;
    
    @Autowired
    private UsuarioRolRepository usuarioRolRepository;
    
    @Autowired
    private PasswordEncoder passwordEncoder;

    @GetMapping
    public List<Map<String, Object>> getUsuarios() {
        return usuarioRepository.findAll().stream().map(u -> {
            List<UsuarioRol> userRoles = usuarioRolRepository.findByUsuario_UsuarioId(u.getUsuarioId());
            String roleName = userRoles.isEmpty() ? "mercaderista" : userRoles.get(0).getRol().getRolNombre().toLowerCase();
            if (u.getUsuarioId() != null && u.getUsuarioId() == 1) {
                roleName = "admin";
            } else if (userRoles.isEmpty()) {
                String str = ((u.getUsername() == null ? "" : u.getUsername()) + " " + (u.getEmail() == null ? "" : u.getEmail()) + " " + (u.getFirstname() == null ? "" : u.getFirstname())).toLowerCase();
                if (str.contains("admin")) roleName = "admin";
                else if (str.contains("controller")) roleName = "controller";
                else if (str.contains("analista")) roleName = "analista";
                else if (str.contains("gerente")) roleName = "gerente";
                else if (str.contains("supervisor")) roleName = "supervisor";
            }
            
            java.util.Map<String, Object> map = new java.util.HashMap<>();
            map.put("id", u.getUsuarioId());
            map.put("name", (u.getFirstname() == null ? "" : u.getFirstname()) + " " + (u.getLastname() == null ? "" : u.getLastname()));
            map.put("email", u.getEmail() == null ? "" : u.getEmail());
            map.put("role", roleName);
            map.put("status", u.getEstado() != null && u.getEstado() ? "Activo" : "Inactivo");
            map.put("equipoComercial", u.getEquipoComercial() == null ? "" : u.getEquipoComercial());
            map.put("pdvsAsignados", u.getPdvsAsignados() == null ? "" : u.getPdvsAsignados());
            return map;
        }).collect(Collectors.toList());
    }

    @PostMapping
    public Map<String, Object> createUsuario(@RequestBody Map<String, String> payload) {
        Usuario u = new Usuario();
        u.setUsername(payload.get("email")); // using email as username for simplicity
        u.setEmail(payload.get("email"));
        u.setFirstname(payload.get("name"));
        String rawPassword = payload.getOrDefault("password", "vayra2026");
        u.setPassword(passwordEncoder.encode(rawPassword));
        u.setEstado(true);
        if (payload.containsKey("equipoComercial")) {
            u.setEquipoComercial(payload.get("equipoComercial"));
        }
        if (payload.containsKey("pdvsAsignados")) {
            u.setPdvsAsignados(payload.get("pdvsAsignados"));
        }
        Usuario saved = usuarioRepository.save(u);

        String role = payload.getOrDefault("role", "mercaderista").toLowerCase();
        Rol r = rolRepository.findByRolNombreIgnoreCase(role).orElseGet(() -> {
            Rol newRol = new Rol();
            newRol.setRolNombre(role);
            return rolRepository.save(newRol);
        });

        UsuarioRol ur = new UsuarioRol();
        ur.setUsuario(saved);
        ur.setRol(r);
        usuarioRolRepository.save(ur);

        return Map.<String, Object>of("success", true, "id", saved.getUsuarioId());
    }

    @PutMapping("/{id}/status")
    public Map<String, Object> updateStatus(@PathVariable Integer id, @RequestBody Map<String, Boolean> payload) {
        Usuario u = usuarioRepository.findById(id).orElseThrow();
        u.setEstado(payload.get("active"));
        usuarioRepository.save(u);
        return Map.<String, Object>of("success", true);
    }

    @PutMapping("/{id}")
    public Map<String, Object> updateUsuario(@PathVariable Integer id, @RequestBody Map<String, String> payload) {
        Usuario u = usuarioRepository.findById(id).orElseThrow();
        
        if (payload.containsKey("name")) {
            u.setFirstname(payload.get("name"));
        }
        if (payload.containsKey("email")) {
            u.setEmail(payload.get("email"));
            u.setUsername(payload.get("email"));
        }
        if (payload.containsKey("password") && !payload.get("password").trim().isEmpty()) {
            u.setPassword(passwordEncoder.encode(payload.get("password")));
        }
        if (payload.containsKey("equipoComercial")) {
            u.setEquipoComercial(payload.get("equipoComercial"));
        }
        if (payload.containsKey("pdvsAsignados")) {
            u.setPdvsAsignados(payload.get("pdvsAsignados"));
        }
        
        usuarioRepository.save(u);

        if (payload.containsKey("role")) {
            String roleStr = payload.get("role").toLowerCase();
            Rol r = rolRepository.findByRolNombreIgnoreCase(roleStr).orElseGet(() -> {
                Rol newRol = new Rol();
                newRol.setRolNombre(roleStr);
                return rolRepository.save(newRol);
            });

            // Update role relation
            List<UsuarioRol> userRoles = usuarioRolRepository.findByUsuario_UsuarioId(id);
            if (!userRoles.isEmpty()) {
                UsuarioRol ur = userRoles.get(0);
                ur.setRol(r);
                usuarioRolRepository.save(ur);
            } else {
                UsuarioRol ur = new UsuarioRol();
                ur.setUsuario(u);
                ur.setRol(r);
                usuarioRolRepository.save(ur);
            }
        }
        
        return Map.<String, Object>of("success", true);
    }

    @DeleteMapping("/{id}")
    @org.springframework.transaction.annotation.Transactional
    public Map<String, Object> deleteUsuario(@PathVariable Integer id) {
        // Delete child relations first
        List<UsuarioRol> userRoles = usuarioRolRepository.findByUsuario_UsuarioId(id);
        usuarioRolRepository.deleteAll(userRoles);
        
        // Delete main user
        usuarioRepository.deleteById(id);
        
        return Map.<String, Object>of("success", true);
    }
}
