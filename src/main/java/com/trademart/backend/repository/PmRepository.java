package com.trademart.backend.repository;

import com.trademart.backend.model.Pm;
import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.stereotype.Repository;

@Repository
public interface PmRepository extends JpaRepository<Pm, Integer> {
}
