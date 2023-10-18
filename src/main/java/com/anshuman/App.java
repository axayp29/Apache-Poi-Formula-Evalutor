package com.anshuman;


import java.util.Map;

public class App {

    public static void main(String[] args) {
        // all the formulae should be passed into this map
        Map<String, String> formulaMap = Map.of(
            "DA", "ROUND(CTC * 12%, 0)",
            "CTC", "50000",
            "HRA", "ROUND(BASIC * 30%, 0)",
            "PB", "ROUND((BASIC + DA) * 12%, 0)",
            "BASIC", "ROUND(CTC * 40%, 0)"
            );
        System.out.println("Input: " + formulaMap);

        // validate the formulae
        System.out.println("Validation: " + ExcelFormulaEvaluator.validateFormulae(formulaMap, "FormulaValidationSheet"));

        // evaluate the formulae
        System.out.println("Output: " + ExcelFormulaEvaluator.processFormulae(formulaMap, "FormulaEvaluationSheet"));
    }

}
