package com.trademart.backend.security;

import com.trademart.backend.model.Usuario;
import com.trademart.backend.model.UsuarioRol;
import com.trademart.backend.repository.UsuarioRepository;
import com.trademart.backend.repository.UsuarioRolRepository;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.security.core.userdetails.UserDetails;
import org.springframework.security.core.userdetails.UserDetailsService;
import org.springframework.security.core.userdetails.UsernameNotFoundException;
import org.springframework.stereotype.Service;

import java.util.List;

@Service
public class UserDetailsServiceImpl implements UserDetailsService {

    @Autowired
    private UsuarioRepository usuarioRepository;

    @Autowired
    private UsuarioRolRepository usuarioRolRepository;

    @Override
    public UserDetails loadUserByUsername(String username) throws UsernameNotFoundException {
        Usuario usuario = usuarioRepository.findByUsername(username)
                .or(() -> usuarioRepository.findByEmail(username))
                .orElseThrow(() -> new UsernameNotFoundException("Usuario no encontrado con username o email: " + username));
        
        List<UsuarioRol> roles = usuarioRolRepository.findByUsuario_UsuarioId(usuario.getUsuarioId());
        
        return new UserDetailsImpl(usuario, roles);
    }
}
