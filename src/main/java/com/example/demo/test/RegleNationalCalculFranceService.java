package com.example.demo.test;

import org.springframework.boot.autoconfigure.condition.ConditionalOnProperty;
import org.springframework.stereotype.Service;

@Service
@ConditionalOnProperty(name = "pays", havingValue = "france")
public class RegleNationalCalculFranceService implements RegleNationalCalculService {
    @Override
    public double calculerPrixTTC(double prixHt) {
        return prixHt * 1.20;
    }
}
