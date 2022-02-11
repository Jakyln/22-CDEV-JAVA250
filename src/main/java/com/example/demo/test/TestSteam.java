package com.example.demo.test;

import java.util.Arrays;
import java.util.List;
import java.util.stream.Collectors;

import static java.util.Comparator.comparing;

public class TestSteam {

    public static void main(String[] args) {
        List<String> list1 = Arrays.asList("abc", "alexandre", "foo");

        List<String> list2 = list1.stream()
                .map(txt -> txt.toUpperCase())
                .map(txt -> "<" + txt + ">")
                .sorted(comparing(txt -> txt.length()))
                .collect(Collectors.toList());

        // List<String> list2 = new ArrayList<>();
        // for (String txt : list1) {
        ////     String txt2 = txt.toUpperCase();
        ////     list2.add(txt2);
        //// }


        System.out.println(list2);
    }
}
