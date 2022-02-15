package com.example.demo.service;

import com.example.demo.entity.Article;
import com.example.demo.entity.Client;
import com.example.demo.entity.Facture;
import com.example.demo.entity.LigneFacture;
import org.springframework.boot.context.event.ApplicationReadyEvent;
import org.springframework.context.ApplicationListener;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Transactional;

import javax.persistence.EntityManager;
import java.time.LocalDate;
import java.util.ArrayList;

/**
 * Classe permettant d'insérer des données dans l'application.
 */
@Service
@Transactional
public class InitData implements ApplicationListener<ApplicationReadyEvent> {

    private EntityManager entityManager;

    public InitData(EntityManager entityManager) {
        this.entityManager = entityManager;
    }

    @Override
    public void onApplicationEvent(ApplicationReadyEvent applicationReadyEvent) {
        insertTestData();
    }

    private void insertTestData() {
        Article a1 = createArticle("Les conserves de viande de licorne", 22.98, 9, "https://static.hitek.fr/img/actualite/2016/08/26/41gn6tpvqtl.jpg");
        Article a2 = createArticle("Wenger Couteau suisse géant", 46.39, 2, "https://static.hitek.fr/img/actualite/2016/08/26/61abqa-gt8s-sx522.jpg");
        Article a3 = createArticle("PAPIER TOILETTE ; DONALD TRUMP", 4.99, 0, "https://static.hitek.fr/img/actualite/2016/08/26/61cb4xnrbol-sx522.jpg");
        Article a4 = createArticle("Grattoir pour Chat en Forme de Platine de DJ", 23.14, 10, "https://static.hitek.fr/img/actualite/2016/08/26/61griray9-l-sx522.jpg");
        Article a5 = createArticle("Jay nothing", 2, 0, "https://static.hitek.fr/img/actualite/2016/08/26/61vu-jqjygl-sy679.jpg");
        Article a6 = createArticle("UN AFFINEUR DE VISAGE", 52, 0, "https://static.hitek.fr/img/actualite/2016/08/26/w_41r-1yapf5l.jpg");

        Client c1 = createClient("John", "Doe", LocalDate.of(2000, 6, 4));
        Client c2 = createClient("Jacqueline", "Parenteau", LocalDate.of(1947, 7, 23));
        Client c3 = createClient("Franck", "Guernon", LocalDate.of(1965, 11, 22));

        createFacture(c1, a1, 2);
        createFacture(c1, a1, 2, a3, 1);

        createFacture(c2, a6, 10);
    }

    private Facture createFacture(Client c1, Article a1, int i) {
        Facture facture = new Facture();
        facture.setClient(c1);
        facture.setLigneFactures(new ArrayList<>());

        LigneFacture ligneFacture = new LigneFacture();
        ligneFacture.setArticle(a1);
        ligneFacture.setQuantite(i);
        ligneFacture.setFacture(facture);

        entityManager.persist(facture);
        entityManager.persist(ligneFacture);
        return facture;
    }

    private void createFacture(Client c1, Article a, int ia, Article b, int ib) {
        Facture facture = createFacture(c1, a, ia);
        //je rajoute une ligne
        LigneFacture ligneFacture = new LigneFacture();
        ligneFacture.setArticle(b);
        ligneFacture.setQuantite(ib);
        ligneFacture.setFacture(facture);
        entityManager.persist(ligneFacture);
    }

    private Client createClient(String prenom, String nom, LocalDate dateNaissance) {
        Client client = new Client();
        client.setPrenom(prenom);
        client.setNom(nom);
        client.setDateNaissance(dateNaissance);
        entityManager.persist(client);
        return client;
    }

    private Article createArticle(String libelle, double prix, int stock, String imageUrl) {
        Article a1 = new Article();
        a1.setLibelle(libelle);
        a1.setPrix(prix);
        a1.setStock(stock);
        a1.setImageUrl(imageUrl);
        entityManager.persist(a1);
        return a1;
    }

}
