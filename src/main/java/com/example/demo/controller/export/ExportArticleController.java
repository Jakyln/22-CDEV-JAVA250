package com.example.demo.controller.export;

import com.example.demo.service.export.ArticleExportCVSService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.io.PrintWriter;

/**
 * Controller pour réaliser l'export des articles.
 */
@Controller
@RequestMapping("export/articles")
public class ExportArticleController {

    @Autowired
    private ArticleExportCVSService articleExportCVSService;

    /**
     * Export des articles au format CSV.
     */
    @GetMapping("csv")
    public void exportCSV(HttpServletRequest request, HttpServletResponse response) throws IOException {
        response.setContentType("text/csv"); //ce sont ces 2 lignes qui disent que le navigateur devra telecharger un fichier csv, si on les enleves il essaie de l'afficher dans le browser
        response.setHeader("Content-Disposition", "attachment; filename=\"export-articles.csv\"");
        PrintWriter writer = response.getWriter(); // permet d'écrire directement dans la page que l'on veut afficher a l'utilisateur (la réponse)
        articleExportCVSService.export(writer);
    }

}
