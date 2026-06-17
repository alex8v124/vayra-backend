package com.trademart.backend.controller;

import com.trademart.backend.model.Cronograma;
import com.trademart.backend.repository.CronogramaRepository;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.*;

import java.util.List;

@RestController
@RequestMapping("/api/cronogramas")
@CrossOrigin(origins = "*")
public class CronogramaController {

    @Autowired
    private CronogramaRepository cronogramaRepository;

    @GetMapping
    public List<Cronograma> listar() {
        return cronogramaRepository.findAll();
    }

    @PostMapping
    public Cronograma guardar(@RequestBody Cronograma cronograma) {
        return cronogramaRepository.save(cronograma);
    }

    @PutMapping("/{id}")
    public Cronograma actualizar(@PathVariable Long id, @RequestBody Cronograma cronograma) {
        cronograma.setId(id);
        return cronogramaRepository.save(cronograma);
    }

    @DeleteMapping("/{id}")
    public void eliminar(@PathVariable Long id) {
        cronogramaRepository.deleteById(id);
    }
}
