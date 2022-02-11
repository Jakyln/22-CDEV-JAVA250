package com.example.demo.test;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

@Service
public class FactureService {

    @Autowired
    private RegleNationalCalculService regleNationalCalculService;

    public void creerFacture(/*...*/) {
        //calculer
        double montantHt = 112;//..

        double v = regleNationalCalculService.calculerPrixTTC(montantHt);
    }
}
