package com.example.demo.service.mapper;

import com.example.demo.dto.ClientDto;
import com.example.demo.dto.FactureDto;
import com.example.demo.entity.Facture;
import com.example.demo.entity.LigneFacture;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Component;

@Component
public class FactureMapper {

    @Autowired
    private ClientMapper clientMapper;

    public FactureDto factureDto(Facture facture) {
        ClientDto clientDto = clientMapper.clientDto(facture.getClient());

        double montant = 0d;
        for (LigneFacture ligneFacture : facture.getLigneFactures()) {
            double prixUnitaire = ligneFacture.getArticle().getPrix();
            montant += prixUnitaire * ligneFacture.getQuantite();
        }

        return new FactureDto(facture.getId(), clientDto, montant);
    }

}
