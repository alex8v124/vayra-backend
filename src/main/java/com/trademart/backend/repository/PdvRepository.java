package com.trademart.backend.repository;

import com.trademart.backend.model.Pdv;
import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.stereotype.Repository;

@Repository
public interface PdvRepository extends JpaRepository<Pdv, Integer> {
}
