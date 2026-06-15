package com.trademart.backend.service;

import com.trademart.backend.model.*;
import com.trademart.backend.repository.*;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Transactional;
import java.util.List;
import org.springframework.security.core.Authentication;
import org.springframework.security.core.context.SecurityContextHolder;
import com.trademart.backend.security.UserDetailsImpl;

@Service
public class TrademartService {

    @Autowired
    private ProductoRepository productoRepository;
    @Autowired
    private PdvRepository pdvRepository;
    @Autowired
    private ActRepository actRepository;
    @Autowired
    private PmRepository pmRepository;
    @Autowired
    private PmActRepository pmActRepository;
    @Autowired
    private ReporteRepository reporteRepository;
    @Autowired
    private UsuarioRepository usuarioRepository;
    @Autowired
    private PlanningRepository planningRepository;
    @Autowired
    private EquipoComercialRepository equipoComercialRepository;

    public List<EquipoComercial> getEquiposComerciales() { return equipoComercialRepository.findAll(); }
    public EquipoComercial saveEquipoComercial(EquipoComercial eq) { return equipoComercialRepository.save(eq); }
    public void deleteEquipoComercial(Integer id) { equipoComercialRepository.deleteById(id); }

    public List<Producto> getProductos() { return productoRepository.findAll(); }
    public Producto saveProducto(Producto p) { return productoRepository.save(p); }
    public List<Producto> saveAllProductos(List<Producto> pList) { return productoRepository.saveAll(pList); }
    public void deleteProducto(Integer id) { productoRepository.deleteById(id); }

    // ==== PLANNING ====
    public List<Planning> getPlannings() { return planningRepository.findAll(); }
    public List<Planning> getPlanningsByUsuarioId(Integer usuarioId) { return planningRepository.findByUsuarioId(usuarioId); }
    public Planning savePlanning(Planning p) { return planningRepository.save(p); }
    public void deletePlanning(Integer id) { planningRepository.deleteById(id); }

    public List<Pdv> getPdvs() { return pdvRepository.findAll(); }
    public Pdv savePdv(Pdv p) { return pdvRepository.save(p); }
    public void deletePdv(Integer id) { pdvRepository.deleteById(id); }

    public List<Act> getActividades() { return actRepository.findAll(); }
    public Act saveActividad(Act a) { return actRepository.save(a); }
    public void deleteActividad(Integer id) { actRepository.deleteById(id); }

    // ==== PM (Puestos de Mercado) ====
    public List<Pm> getPms() { return pmRepository.findAll(); }
    public Pm savePm(Pm pm) { return pmRepository.save(pm); }
    
    @Transactional
    public void deletePm(Integer id) { 
        pmActRepository.deleteByPmId(id);
        pmRepository.deleteById(id); 
    }

    // ==== PM_ACT (Relación N:M) ====
    public List<PmAct> getAllPmActs() { return pmActRepository.findAll(); }
    public List<PmAct> getPmActsByPmId(Integer pmId) { return pmActRepository.findByPmId(pmId); }
    
    @Transactional
    public void replacePmActs(Integer pmId, List<Integer> actIds) {
        pmActRepository.deleteByPmId(pmId);
        for (Integer actId : actIds) {
            PmAct pa = new PmAct();
            pa.setPmId(pmId);
            pa.setActId(actId);
            pmActRepository.save(pa);
        }
    }

    public List<Reporte> getReportes() { return reporteRepository.findAll(); }
    public Reporte saveReporte(Reporte r) {
        if (r.getUsuario() == null) {
            Authentication authentication = SecurityContextHolder.getContext().getAuthentication();
            if (authentication != null && authentication.getPrincipal() instanceof UserDetailsImpl) {
                UserDetailsImpl userDetails = (UserDetailsImpl) authentication.getPrincipal();
                r.setUsuario(userDetails.getUsuario());
            } else {
                List<Usuario> usuarios = usuarioRepository.findAll();
                if (!usuarios.isEmpty()) {
                    r.setUsuario(usuarios.get(0));
                }
            }
        }
        if (r.getPm() == null) {
            List<Pm> pms = pmRepository.findAll();
            if (!pms.isEmpty()) {
                r.setPm(pms.get(0));
            }
        }
        return reporteRepository.save(r);
    }
    public void deleteReporte(Integer id) { reporteRepository.deleteById(id); }

    public List<Usuario> getUsuarios() { return usuarioRepository.findAll(); }
}
