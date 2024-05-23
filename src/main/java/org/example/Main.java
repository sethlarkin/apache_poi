package org.example;


import java.nio.file.Paths;
import java.util.*;

public class Main {
    public static void main(String[] args) {

        String filePath = Paths.get("test_data", "BritecoreAutomatedTestCases.xlsx").toString();
        BritecoreTestDataExtractor bte1 = new BritecoreTestDataExtractor(filePath);

        List<String> allSheets = bte1.getSheetNames();

        if (!allSheets.isEmpty()) {
            for (String sheet : allSheets) {
                System.out.println("Sheet name: " + sheet);
                System.out.println("Columns: " + bte1.getColumnNames(sheet));
            }

            List<Map<String, String>> quote1Eligibility = bte1.getRowDataById(1, "NB - Eligibility");
            System.out.println("NB Eligibility Matching rows: " + quote1Eligibility);
            List<Map<String, String>> quote1Setup = bte1.getRowDataById(1, "NB - Setup ");
            System.out.println("NS Setup Matching rows: " + quote1Setup);
            List<Map<String, String>> quote2Setup = bte1.getRowDataById(2, "NB - Setup ");
            System.out.println("NS Setup Matching rows: " + quote2Setup);

            Set<String> policyTypes1 = bte1.getAvailablePolicyTypes(1, "NB - Setup ");
            System.out.println(policyTypes1);
        } else {
            System.err.println("No sheets found");
        }
        List<String> testSheets = new ArrayList<>(Arrays.asList("NB - Eligibility", "NB - Setup "));
        Map<String, Object> testData = bte1.createPolicyTestData(1, testSheets);
        System.out.println(testData);
    }
}