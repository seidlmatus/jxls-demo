package com.example.jxls;

import org.jxls.area.Area;
import org.jxls.builder.AreaBuilder;
import org.jxls.builder.xls.XlsCommentAreaBuilder;
import org.jxls.common.CellRef;
import org.jxls.common.Context;
import org.jxls.transform.Transformer;
import org.jxls.util.JxlsHelper;
import org.jxls.util.TransformerFactory;
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
        generateSyledExcel(prices);

    }

    private static void generateSyledExcel(List<Price> items) {
        try (InputStream is = new ClassPathResource("template_price.xlsx").getInputStream()) {
            try (OutputStream os = new FileOutputStream("target/export_style.xlsx")) {
                Transformer transformer = TransformerFactory.createTransformer(is, os);

                AreaBuilder areaBuilder = new XlsCommentAreaBuilder(transformer);
                List<Area> xlsAreaList = areaBuilder.build();
                xlsAreaList.get(0)
                        .getCommandDataList().get(0)
                        .getCommand()
                        .getAreaList().get(0)
                        .getAreaListeners().add(new SimpleAreaListener(transformer));




                Context context = new Context();
                context.putVar("content", items);
                xlsAreaList.get(0).applyAt(new CellRef("Hárok1!A1"), context);
                xlsAreaList.get(0).processFormulas();
                transformer.write();

               /* System.out.println("Creating area");

                XlsArea xlsArea = new XlsArea("Hárok1!A1:B2", transformer);


                XlsArea priceArea = new XlsArea("Hárok1!A2:B2", transformer);

                XlsArea ifArea = new XlsArea("Hárok1!A2:A2", transformer);
                IfCommand ifCommand = new IfCommand("c.value <= 20",
                        ifArea,ifArea);

                ifArea.addAreaListener(new SimpleAreaListener(ifArea));
                priceArea.addCommand(new AreaRef("Hárok1!A2:B2"), ifCommand);
                Command priceEachCommand = new EachCommand("c", "content", priceArea);
                xlsArea.addCommand(new AreaRef("Hárok1!A2:B2"), priceEachCommand);

                Context context = new Context();
                context.putVar("content", items);

                JxlsHelper.getInstance().processTemplate(context,transformer);
*/





            }
        } catch (IOException e) {
            e.printStackTrace();
        }
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




