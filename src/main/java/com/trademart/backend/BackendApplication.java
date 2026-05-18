package com.trademart.backend;

import com.trademart.backend.model.Rol;
import com.trademart.backend.model.Usuario;
import com.trademart.backend.model.UsuarioRol;
import com.trademart.backend.repository.RolRepository;
import com.trademart.backend.repository.UsuarioRepository;
import com.trademart.backend.repository.UsuarioRolRepository;
import org.springframework.boot.CommandLineRunner;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.context.annotation.Bean;
import org.springframework.security.crypto.password.PasswordEncoder;

@SpringBootApplication
public class BackendApplication {
	public static void main(String[] args) {
		SpringApplication.run(BackendApplication.class, args);
	}

	@Bean
	CommandLineRunner initDatabase(UsuarioRepository usuarioRepository, RolRepository rolRepository, UsuarioRolRepository usuarioRolRepository, PasswordEncoder passwordEncoder) {
		return args -> {
			if (usuarioRepository.count() == 0) {
				Rol adminRol = new Rol();
				adminRol.setRolNombre("admin");
				adminRol = rolRepository.save(adminRol);

				Usuario admin = new Usuario();
				admin.setUsername("admin@trademart.pe");
				admin.setEmail("admin@trademart.pe");
				admin.setFirstname("Administrador");
				admin.setLastname("Sistema");
				admin.setPassword(passwordEncoder.encode("admin123"));
				admin.setEstado(true);
				admin = usuarioRepository.save(admin);

				UsuarioRol ur = new UsuarioRol();
				ur.setUsuario(admin);
				ur.setRol(adminRol);
				usuarioRolRepository.save(ur);
				
				System.out.println("Usuario administrador por defecto creado.");
			}
		};
	}
}
