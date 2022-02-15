package com.example.demo._explication;

import com.example.demo.entity.Article;

public class Equal {
    public static void main(String[] args) {
        Article a;
        a = new Article();
        a.setId(1L);
        a.setLibelle("mon article");

        Article b;
        b = a;

        // Modifie l'objet pointé par b (et donc pointé par a !)
        b.setLibelle("nouveau libellé");

        boolean test1 = (a == b); //=> true
        boolean test2 = (a.equals(b)); //=> true

        Article c = new Article();
        c.setId(1L);
        c.setLibelle("mon article");

        boolean test3 = (a == c); // ==> false
        boolean test4 = (a.equals(c));// ==> dépendra de la surcharge ?

    }
}
