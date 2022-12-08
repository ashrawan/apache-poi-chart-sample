package com.projectsample.libapachepoi.playground.additional;

import lombok.AllArgsConstructor;
import lombok.Data;

import java.util.Map;

@Data
@AllArgsConstructor
public class SampleDataDTO {

    private String label;
    private Map<String, ?> valuesMap;
}
