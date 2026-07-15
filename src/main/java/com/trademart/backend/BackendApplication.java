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
import org.springframework.jdbc.core.JdbcTemplate;

@SpringBootApplication
public class BackendApplication {
	public static void main(String[] args) {
		SpringApplication.run(BackendApplication.class, args);
	}

	@Bean
	CommandLineRunner initDatabase(UsuarioRepository usuarioRepository, RolRepository rolRepository, UsuarioRolRepository usuarioRolRepository, PasswordEncoder passwordEncoder, JdbcTemplate jdbcTemplate) {
		return args -> {
			try {
				jdbcTemplate.execute("CREATE TABLE IF NOT EXISTS plan_rutas (id BIGSERIAL PRIMARY KEY, nombre VARCHAR(255) NOT NULL, datos_json TEXT NOT NULL);");
				jdbcTemplate.execute("ALTER TABLE cronogramas ADD COLUMN IF NOT EXISTS plan_ruta_id BIGINT;");
				jdbcTemplate.execute("ALTER TABLE cronogramas ADD COLUMN IF NOT EXISTS fecha_inicio VARCHAR(255);");
				jdbcTemplate.execute("ALTER TABLE cronogramas ADD COLUMN IF NOT EXISTS fecha_fin VARCHAR(255);");
				jdbcTemplate.execute("ALTER TABLE cronogramas ADD COLUMN IF NOT EXISTS planning_ids TEXT;");
				jdbcTemplate.execute("ALTER TABLE planning ADD COLUMN IF NOT EXISTS dias_semana_pms TEXT;");
			} catch (Exception e) {
				System.err.println("Error initializing schema manually: " + e.getMessage());
			}

			// 1. Crear o actualizar usuario admin por defecto si no existe
			if (usuarioRepository.count() == 0) {
				Rol adminRol = new Rol();
				adminRol.setRolNombre("admin");
				adminRol = rolRepository.save(adminRol);

				Usuario admin = new Usuario();
				admin.setUsername("admin");
				admin.setEmail("admin@trademart.com");
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
			} else {
				// 2. Si ya existen usuarios del seed SQL con hash dummy "$2a$10$xyz", actualizarlos a un BCrypt real de "admin123"
				usuarioRepository.findAll().forEach(u -> {
					boolean changed = false;
					if (u.getPassword() == null || u.getPassword().equals("$2a$10$xyz") || !u.getPassword().startsWith("$2a$")) {
						u.setPassword(passwordEncoder.encode("admin123"));
						changed = true;
					}
					if ("admin".equalsIgnoreCase(u.getUsername()) || "admin@trademart.com".equalsIgnoreCase(u.getEmail()) || "admin@trademart.pe".equalsIgnoreCase(u.getEmail())) {
						u.setPassword(passwordEncoder.encode("admin123"));
						if (!"admin@trademart.com".equalsIgnoreCase(u.getEmail())) {
							u.setEmail("admin@trademart.com");
						}
						changed = true;
					}
					if (changed) {
						usuarioRepository.save(u);
					}
				});
				System.out.println("Hashes de usuarios verificados/actualizados para login exitoso.");
			}
		};
	}
}
