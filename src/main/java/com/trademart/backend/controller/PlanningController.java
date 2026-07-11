package com.trademart.backend.controller;

import com.trademart.backend.model.Planning;
import com.trademart.backend.service.TrademartService;
import com.trademart.backend.repository.PdvRepository;
import com.trademart.backend.repository.UsuarioRepository;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.*;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

@RestController
@RequestMapping("/api/data/plannings")
public class PlanningController {

    @Autowired
    private TrademartService trademartService;

    @Autowired
    private UsuarioRepository usuarioRepository;

    @Autowired
    private PdvRepository pdvRepository;

    @GetMapping
    public List<Map<String, Object>> getAllPlannings() {
        List<Planning> plannings = trademartService.getPlannings();
        List<Map<String, Object>> response = new ArrayList<>();

        // Fetch all users and PDVs into maps to avoid N+1 queries
        Map<Integer, String> userNamesMap = new HashMap<>();
        usuarioRepository.findAll().forEach(u -> 
            userNamesMap.put(u.getUsuarioId(), (u.getFirstname() == null ? "" : u.getFirstname()) + " " + (u.getLastname() == null ? "" : u.getLastname()))
        );

        Map<Integer, String> pdvNamesMap = new HashMap<>();
        pdvRepository.findAll().forEach(pdv -> 
            pdvNamesMap.put(pdv.getPdvId(), pdv.getPdvNombre())
        );

        for (Planning p : plannings) {
            Map<String, Object> map = new HashMap<>();
            map.put("planningId", p.getPlanningId());
            map.put("usuarioId", p.getUsuarioId());
            map.put("pdvId", p.getPdvId());
            map.put("pmIds", p.getPmIds());
            map.put("diasSemanaPms", p.getDiasSemanaPms());
            map.put("actIds", p.getActIds());
            map.put("fechaInicio", p.getFechaInicio());
            map.put("fechaFin", p.getFechaFin());
            map.put("estado", p.getEstado());

            // Get names from maps (O(1)) instead of querying the DB (N+1)
            map.put("mercaderistaName", userNamesMap.getOrDefault(p.getUsuarioId(), "Desconocido"));
            map.put("pdvNombre", pdvNamesMap.getOrDefault(p.getPdvId(), "Desconocido"));

            response.add(map);
        }
        return response;
    }

    @PostMapping
    public Planning createPlanning(@RequestBody Planning planning) {
        return trademartService.savePlanning(planning);
    }

    @PutMapping("/{id}")
    public Planning updatePlanning(@PathVariable Integer id, @RequestBody Planning planningDetails) {
        Planning planning = trademartService.getPlannings().stream()
                .filter(p -> p.getPlanningId().equals(id))
                .findFirst()
                .orElseThrow(() -> new RuntimeException("Planning not found with id " + id));

        planning.setUsuarioId(planningDetails.getUsuarioId());
        planning.setPdvId(planningDetails.getPdvId());
        planning.setPmIds(planningDetails.getPmIds());
        planning.setDiasSemanaPms(planningDetails.getDiasSemanaPms());
        planning.setActIds(planningDetails.getActIds());
        planning.setFechaInicio(planningDetails.getFechaInicio());
        planning.setFechaFin(planningDetails.getFechaFin());
        if (planningDetails.getEstado() != null) {
            planning.setEstado(planningDetails.getEstado());
        }

        return trademartService.savePlanning(planning);
    }

    @DeleteMapping("/{id}")
    public Map<String, Boolean> deletePlanning(@PathVariable Integer id) {
        trademartService.deletePlanning(id);
        Map<String, Boolean> response = new HashMap<>();
        response.put("deleted", Boolean.TRUE);
        return response;
    }
}
