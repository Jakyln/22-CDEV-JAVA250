package com.example.demo.service.impl;

import com.example.demo.dto.FactureDto;
import com.example.demo.entity.Facture;
import com.example.demo.repository.FactureRepository;
import com.example.demo.service.FactureService;
import com.example.demo.service.mapper.FactureMapper;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Transactional;

import java.util.ArrayList;
import java.util.List;

/**
 * Service contenant les actions métiers liées aux articles.
 */
@Service
@Transactional
public class FactureServiceImpl implements FactureService {

    private FactureRepository factureRepository;
    private FactureMapper factureMapper;

    public FactureServiceImpl(FactureRepository factureRepository, FactureMapper factureMapper) {
        this.factureRepository = factureRepository;
        this.factureMapper = factureMapper;
    }

    @Override
    public List<FactureDto> findAll() {
        List<Facture> factures = factureRepository.findAll();

        List<FactureDto> list = new ArrayList<>();
        for (Facture facture : factures) {
            FactureDto factureDto = factureMapper.factureDto(facture);
            list.add(factureDto);
        }
        return list;
    }

}
