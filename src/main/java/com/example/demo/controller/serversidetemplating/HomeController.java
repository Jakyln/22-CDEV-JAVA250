package com.example.demo.controller.serversidetemplating;

import com.example.demo.dto.ArticleDto;
import com.example.demo.service.ArticleService;
import com.example.demo.service.ClientService;
import com.example.demo.service.FactureService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.servlet.ModelAndView;

import java.util.List;

/**
 * Controller principal pour affichage des clients / factures sur la page d'accueil dans le cas d'une application server side templating.
 */
@Controller
public class HomeController {

    @Autowired
    private ArticleService articleService;

    @Autowired
    private ClientService clientService;

    @Autowired
    private FactureService factureService;

    /**
     * Point d'entrée lorsque url http://localhost:8080/ sera déclenchée.
     */
    @GetMapping("/")
    public ModelAndView home() {
        // C'est la view "home" qui sera utilisée pour constuire la réponse => fichier home.mustache
        ModelAndView modelAndView = new ModelAndView("home");

        // Chargement des articles
        List<ArticleDto> articles = articleService.findAll();
        // et "stockage" dans une variable nommé "articles"
        modelAndView.addObject("articles", articles);
        modelAndView.addObject("clients", clientService.findAll());
        modelAndView.addObject("factures", factureService.findAll());
        return modelAndView;
    }

}
