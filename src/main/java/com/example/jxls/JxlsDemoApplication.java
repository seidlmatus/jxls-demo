package com.example.jxls;

import lombok.AllArgsConstructor;
import lombok.Data;
import org.jxls.common.Context;
import org.jxls.transform.poi.PoiContext;
import org.jxls.util.JxlsHelper;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.core.io.ClassPathResource;
import org.springframework.core.io.Resource;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RestController;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.Optional;

@SpringBootApplication
public class JxlsDemoApplication {

    public static void main(String[] args) {
        SpringApplication.run(JxlsDemoApplication.class, args);
    }

}

@RestController
class BaseController {

    private static List<Price> prices;


    static {
        prices = new ArrayList<>();
        prices.add(new Price("12",TypeOfValue.LOW));
//        prices.add(new Price(25d,TypeOfValue.HIGH));
//        prices.add(new Price(25d,TypeOfValue.HIGH));
//        prices.add(new Price(15d,TypeOfValue.LOW));
//        prices.add(new Price(100d,TypeOfValue.HIGH));
//        prices.add(new Price(1d,TypeOfValue.LOW));
    }

    @GetMapping(value = "/")

    public ResponseEntity<Resource> getIndex(){
        return  ResponseEntity.ok()
                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=export")
                .body(generateExcel(Optional.ofNullable(prices).orElseThrow(NullPointerException::new)));
    }

    private Resource generateExcel(List<Price> items){
        try(InputStream is = new ClassPathResource("template_price.xlsx").getInputStream() ) {
            try (OutputStream os = new FileOutputStream("target/export.xlsx")) {
                Context context = new Context();
                context.putVar("content", items);
                JxlsHelper.getInstance().processTemplate(is, os, context);
            }
        } catch (IOException e) {
            e.printStackTrace();
        }

        return null;
    }




}


class Price{
    private String value;
    private TypeOfValue typeOfValue;

    public Price(String value, TypeOfValue typeOfValue) {
        this.value = value;
        this.typeOfValue = typeOfValue;
    }

    public String getValue() {
        return value;
    }

    public void setValue(String value) {
        this.value = value;
    }

    public TypeOfValue getTypeOfValue() {
        return typeOfValue;
    }

    public void setTypeOfValue(TypeOfValue typeOfValue) {
        this.typeOfValue = typeOfValue;
    }
}

enum TypeOfValue{
    HIGH, LOW
}