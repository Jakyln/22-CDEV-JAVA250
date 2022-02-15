package com.example.demo.service.export;

import com.example.demo.entity.Article;
import com.example.demo.repository.ArticleRepository;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import java.io.PrintWriter;
import java.text.DecimalFormat;
import java.util.Currency;
import java.util.List;
import java.util.Locale;

@Service
public class ArticleExportCVSService {

    @Autowired
    private ArticleRepository articleRepository;

    public void export(PrintWriter writer) {
        List<Article> articles = articleRepository.findAll();
        writer.println("Libellé;Prix");
        for (Article article : articles) {
            StringBuilder stringBuilder = new StringBuilder();
            String libelle = article.getLibelle();
            if (libelle.contains(";")) {
                stringBuilder.append("\"");
                stringBuilder.append(libelle);
                stringBuilder.append("\"");
            } else {
                stringBuilder.append(libelle);
            }
            stringBuilder.append(";");
            DecimalFormat df = new DecimalFormat("###,###,###.00");
            df.setCurrency(Currency.getInstance(Locale.FRANCE));
            stringBuilder.append(df.format(article.getPrix()));
            writer.println(stringBuilder.toString());
        }
    }


}
