package com.trademart.backend.controller;

import com.trademart.backend.model.*;
import com.trademart.backend.service.TrademartService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.*;
import java.util.*;

@RestController
@RequestMapping("/api/data")
public class DataController {

    @Autowired
    private TrademartService trademartService;

    // ==== PRODUCTOS ====
    @GetMapping("/productos")
    public List<Producto> getProductos() { return trademartService.getProductos(); }
    @PostMapping("/productos")
    public Producto addProducto(@RequestBody Producto p) { return trademartService.saveProducto(p); }
    @PutMapping("/productos/{id}")
    public Producto updateProducto(@PathVariable Integer id, @RequestBody Producto p) { p.setProductoId(id); return trademartService.saveProducto(p); }
    @DeleteMapping("/productos/{id}")
    public void deleteProducto(@PathVariable Integer id) { trademartService.deleteProducto(id); }

    // ==== PDVs ====
    @GetMapping("/pdvs")
    public List<Pdv> getPdvs() { return trademartService.getPdvs(); }
    @PostMapping("/pdvs")
    public Pdv addPdv(@RequestBody Pdv p) { return trademartService.savePdv(p); }
    @PutMapping("/pdvs/{id}")
    public Pdv updatePdv(@PathVariable Integer id, @RequestBody Pdv p) { p.setPdvId(id); return trademartService.savePdv(p); }
    @DeleteMapping("/pdvs/{id}")
    public void deletePdv(@PathVariable Integer id) { trademartService.deletePdv(id); }

    // ==== ACTIVIDADES ====
    @GetMapping("/actividades")
    public List<Act> getActividades() { return trademartService.getActividades(); }
    @PostMapping("/actividades")
    public Act addActividad(@RequestBody Act a) { 
        if (a.getSkuIds() != null && !a.getSkuIds().isEmpty()) {
            try {
                String firstId = a.getSkuIds().split(",")[0].trim();
                a.setProductoId(Integer.parseInt(firstId));
            } catch (Exception e) {
                a.setProductoId(1);
            }
        } else {
            a.setProductoId(1);
        }
        return trademartService.saveActividad(a); 
    }
    @PutMapping("/actividades/{id}")
    public Act updateActividad(@PathVariable Integer id, @RequestBody Act a) { 
        a.setActId(id); 
        if (a.getSkuIds() != null && !a.getSkuIds().isEmpty()) {
            try {
                String firstId = a.getSkuIds().split(",")[0].trim();
                a.setProductoId(Integer.parseInt(firstId));
            } catch (Exception e) {
                a.setProductoId(1);
            }
        } else {
            a.setProductoId(1);
        }
        return trademartService.saveActividad(a); 
    }
    @DeleteMapping("/actividades/{id}")
    public void deleteActividad(@PathVariable Integer id) { trademartService.deleteActividad(id); }

    // ==== PM (Puestos de Mercado) ====
    @GetMapping("/pms")
    public List<Map<String, Object>> getPms() {
        List<Pm> pms = trademartService.getPms();
        List<PmAct> allPmActs = trademartService.getAllPmActs();
        List<Map<String, Object>> result = new ArrayList<>();
        for (Pm pm : pms) {
            Map<String, Object> map = new HashMap<>();
            map.put("pmId", pm.getPmId());
            map.put("pmNombre", pm.getPmNombre());
            map.put("pdvId", pm.getPdvId());
            // Agregar los actIds asociados
            List<Integer> actIds = allPmActs.stream()
                .filter(pa -> pa.getPmId().equals(pm.getPmId()))
                .map(PmAct::getActId)
                .toList();
            map.put("actIds", actIds);
            result.add(map);
        }
        return result;
    }

    @PostMapping("/pms")
    public Map<String, Object> addPm(@RequestBody Map<String, Object> body) {
        Pm pm = new Pm();
        pm.setPmNombre((String) body.get("pmNombre"));
        pm.setPdvId((Integer) body.get("pdvId"));
        Pm saved = trademartService.savePm(pm);
        
        // Guardar relaciones N:M con actividades
        @SuppressWarnings("unchecked")
        List<Integer> actIds = (List<Integer>) body.get("actIds");
        if (actIds != null && !actIds.isEmpty()) {
            trademartService.replacePmActs(saved.getPmId(), actIds);
        }
        
        Map<String, Object> result = new HashMap<>();
        result.put("pmId", saved.getPmId());
        result.put("pmNombre", saved.getPmNombre());
        result.put("pdvId", saved.getPdvId());
        result.put("actIds", actIds != null ? actIds : List.of());
        return result;
    }

    @PutMapping("/pms/{id}")
    public Map<String, Object> updatePm(@PathVariable Integer id, @RequestBody Map<String, Object> body) {
        Pm pm = new Pm();
        pm.setPmId(id);
        pm.setPmNombre((String) body.get("pmNombre"));
        pm.setPdvId((Integer) body.get("pdvId"));
        Pm saved = trademartService.savePm(pm);
        
        // Reemplazar relaciones N:M
        @SuppressWarnings("unchecked")
        List<Integer> actIds = (List<Integer>) body.get("actIds");
        if (actIds != null) {
            trademartService.replacePmActs(id, actIds);
        }
        
        Map<String, Object> result = new HashMap<>();
        result.put("pmId", saved.getPmId());
        result.put("pmNombre", saved.getPmNombre());
        result.put("pdvId", saved.getPdvId());
        result.put("actIds", actIds != null ? actIds : List.of());
        return result;
    }

    @DeleteMapping("/pms/{id}")
    public void deletePm(@PathVariable Integer id) { trademartService.deletePm(id); }

    // ==== REPORTES ====
    @GetMapping("/reportes")
    public List<Reporte> getReportes() { return trademartService.getReportes(); }
    @PostMapping("/reportes")
    public Reporte addReporte(@RequestBody Reporte r) { return trademartService.saveReporte(r); }
    @PutMapping("/reportes/{id}")
    public Reporte updateReporte(@PathVariable Integer id, @RequestBody Reporte r) { r.setRegistroReporteId(id); return trademartService.saveReporte(r); }
    @DeleteMapping("/reportes/{id}")
    public void deleteReporte(@PathVariable Integer id) { trademartService.deleteReporte(id); }

    // ==== USUARIOS ====
    @GetMapping("/usuarios")
    public List<Usuario> getUsuarios() { return trademartService.getUsuarios(); }

    // ==== EQUIPOS COMERCIALES ====
    @GetMapping("/equipos")
    public List<EquipoComercial> getEquipos() { return trademartService.getEquiposComerciales(); }
    @PostMapping("/equipos")
    public EquipoComercial addEquipo(@RequestBody EquipoComercial eq) { return trademartService.saveEquipoComercial(eq); }
    @PutMapping("/equipos/{id}")
    public EquipoComercial updateEquipo(@PathVariable Integer id, @RequestBody EquipoComercial eq) { eq.setEquipoId(id); return trademartService.saveEquipoComercial(eq); }
    @DeleteMapping("/equipos/{id}")
    public void deleteEquipo(@PathVariable Integer id) { trademartService.deleteEquipoComercial(id); }
}
