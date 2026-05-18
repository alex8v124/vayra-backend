package com.trademart.backend.repository;

import com.trademart.backend.model.PmAct;
import org.springframework.data.jpa.repository.JpaRepository;
import java.util.List;

public interface PmActRepository extends JpaRepository<PmAct, Integer> {
    List<PmAct> findByPmId(Integer pmId);
    void deleteByPmId(Integer pmId);
}
