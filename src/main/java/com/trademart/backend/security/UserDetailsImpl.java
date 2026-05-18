package com.trademart.backend.security;

import com.trademart.backend.model.Usuario;
import com.trademart.backend.model.UsuarioRol;
import org.springframework.security.core.GrantedAuthority;
import org.springframework.security.core.authority.SimpleGrantedAuthority;
import org.springframework.security.core.userdetails.UserDetails;

import java.util.Collection;
import java.util.List;
import java.util.stream.Collectors;

public class UserDetailsImpl implements UserDetails {

    private final Usuario usuario;
    private final List<UsuarioRol> roles;

    public UserDetailsImpl(Usuario usuario, List<UsuarioRol> roles) {
        this.usuario = usuario;
        this.roles = roles;
    }

    @Override
    public Collection<? extends GrantedAuthority> getAuthorities() {
        return roles.stream()
                .map(ur -> new SimpleGrantedAuthority("ROLE_" + ur.getRol().getRolNombre().toUpperCase()))
                .collect(Collectors.toList());
    }

    @Override
    public String getPassword() {
        return usuario.getPassword();
    }

    @Override
    public String getUsername() {
        return usuario.getUsername();
    }

    @Override
    public boolean isAccountNonExpired() {
        return true;
    }

    @Override
    public boolean isAccountNonLocked() {
        return true;
    }

    @Override
    public boolean isCredentialsNonExpired() {
        return true;
    }

    @Override
    public boolean isEnabled() {
        return usuario.getEstado();
    }
    
    public Usuario getUsuario() {
        return this.usuario;
    }
}
