package com.example.jxls;

import org.jxls.common.Context;
import org.jxls.util.JxlsHelper;
import org.springframework.core.io.ClassPathResource;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.List;


public class JxlsDemoApplication {

    public static List<Price> prices = new ArrayList<>();

    public static void main(String[] args) {

        prices.add(new Price(12d, TypeOfValue.LOW));
        prices.add(new Price(25d, TypeOfValue.HIGH));
        prices.add(new Price(25d, TypeOfValue.HIGH));
        prices.add(new Price(15d, TypeOfValue.LOW));
        prices.add(new Price(100d, TypeOfValue.HIGH));
        prices.add(new Price(1d, TypeOfValue.LOW));


        generateExcel(prices);

    }

    private static void generateExcel(List<Price> items) {
        try (InputStream is = new ClassPathResource("template_price.xlsx").getInputStream()) {
            try (OutputStream os = new FileOutputStream("target/export.xlsx")) {
                Context context = new Context();
                context.putVar("content", items);
                JxlsHelper.getInstance().processTemplate(is, os, context);
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}




