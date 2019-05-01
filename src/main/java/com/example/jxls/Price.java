package com.example.jxls;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

@Data
@AllArgsConstructor
@NoArgsConstructor
public class Price {

    private Double value;
    private TypeOfValue typeOfValue;

}
