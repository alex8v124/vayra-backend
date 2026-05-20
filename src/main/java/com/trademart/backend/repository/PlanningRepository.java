package com.trademart.backend.repository;

import com.trademart.backend.model.Planning;
import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.stereotype.Repository;

import java.util.List;

@Repository
public interface PlanningRepository extends JpaRepository<Planning, Integer> {
    List<Planning> findByUsuarioId(Integer usuarioId);
}
