package com.trademart.backend.repository;

import com.trademart.backend.model.EquipoComercial;
import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.stereotype.Repository;

@Repository
public interface EquipoComercialRepository extends JpaRepository<EquipoComercial, Integer> {
}
