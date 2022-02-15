package com.example.demo.dto;

/**
 * Classe permettant d'exposer des données au format JSON au client.
 */
public class FactureDto {
    private Long id;
    private ClientDto client;
    private double montant;

    public FactureDto(Long id, ClientDto client, double montant) {
        this.id = id;
        this.client = client;
        this.montant = montant;
    }

    public Long getId() {
        return id;
    }

    public void setId(Long id) {
        this.id = id;
    }

    public ClientDto getClient() {
        return client;
    }


    public void setClient(ClientDto client) {
        this.client = client;
    }

    public double getMontant() {
        return montant;
    }

    public void setMontant(double montant) {
        this.montant = montant;
    }
}
