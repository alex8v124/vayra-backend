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
				Rol adminRol = rolRepository.findByRolNombreIgnoreCase("admin").orElseGet(() -> {
					Rol r = new Rol();
					r.setRolNombre("admin");
					return rolRepository.save(r);
				});

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

			// 3. Asegurar que existan los roles y usuarios Controller y Analista en la base de datos
			if (usuarioRepository.findByUsername("controller@trademart.com").isEmpty()) {
				Rol ctrlRol = rolRepository.findByRolNombreIgnoreCase("controller").orElseGet(() -> {
					Rol r = new Rol();
					r.setRolNombre("controller");
					return rolRepository.save(r);
				});
				Usuario uCtrl = new Usuario();
				uCtrl.setUsername("controller@trademart.com");
				uCtrl.setEmail("controller@trademart.com");
				uCtrl.setFirstname("Carlos");
				uCtrl.setLastname("Controller");
				uCtrl.setPassword(passwordEncoder.encode("vayra2026"));
				uCtrl.setEstado(true);
				uCtrl = usuarioRepository.save(uCtrl);

				UsuarioRol ur = new UsuarioRol();
				ur.setUsuario(uCtrl);
				ur.setRol(ctrlRol);
				usuarioRolRepository.save(ur);
				System.out.println("Usuario Carlos Controller creado.");
			}

			if (usuarioRepository.findByUsername("analista@trademart.com").isEmpty()) {
				Rol anRol = rolRepository.findByRolNombreIgnoreCase("analista").orElseGet(() -> {
					Rol r = new Rol();
					r.setRolNombre("analista");
					return rolRepository.save(r);
				});
				Usuario uAn = new Usuario();
				uAn.setUsername("analista@trademart.com");
				uAn.setEmail("analista@trademart.com");
				uAn.setFirstname("Lucia");
				uAn.setLastname("Analista");
				uAn.setPassword(passwordEncoder.encode("vayra2026"));
				uAn.setEstado(true);
				uAn = usuarioRepository.save(uAn);

				UsuarioRol ur = new UsuarioRol();
				ur.setUsuario(uAn);
				ur.setRol(anRol);
				usuarioRolRepository.save(ur);
				System.out.println("Usuario Lucia Analista creado.");
			}
		};
	}
}
