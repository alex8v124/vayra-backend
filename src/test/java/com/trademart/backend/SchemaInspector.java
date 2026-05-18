package com.trademart.backend;

import org.springframework.boot.CommandLineRunner;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.context.annotation.Bean;
import org.springframework.jdbc.core.JdbcTemplate;

import java.util.List;
import java.util.Map;

import org.springframework.boot.builder.SpringApplicationBuilder;

@SpringBootApplication
public class SchemaInspector {

    public static void main(String[] args) {
        new SpringApplicationBuilder(SchemaInspector.class)
            .web(org.springframework.boot.WebApplicationType.NONE)
            .run(args);
    }

    @Bean
    public CommandLineRunner run(JdbcTemplate jdbcTemplate) {
        return args -> {
            System.out.println("=== DB TABLES ===");
            List<Map<String, Object>> tables = jdbcTemplate.queryForList(
                "SELECT table_name FROM information_schema.tables WHERE table_schema='public'"
            );
            for (Map<String, Object> table : tables) {
                String tableName = (String) table.get("table_name");
                System.out.println("TABLE: " + tableName);
                List<Map<String, Object>> columns = jdbcTemplate.queryForList(
                    "SELECT column_name, data_type FROM information_schema.columns WHERE table_name='" + tableName + "'"
                );
                for (Map<String, Object> col : columns) {
                    System.out.println("  - " + col.get("column_name") + " (" + col.get("data_type") + ")");
                }
            }
            System.out.println("=== END DB TABLES ===");
        };
    }
}
