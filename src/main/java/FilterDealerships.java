import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;

public class FilterDealerships {

    public static void main(String[] args) {
        String inputFile = "C:/Users/HP/Downloads/autosave (Autosaved).xlsx";
        String outputFile = "Dub_Abu_Shar_Ajm_AlA_Riya_Jedd_Dam_Doha_Kuw_Man.csv";

        try (Workbook wb = new XSSFWorkbook(new FileInputStream(inputFile));
             BufferedWriter writer = new BufferedWriter(new FileWriter(outputFile))) {

            Sheet sheet = wb.getSheetAt(0);

            DataFormatter formatter = new DataFormatter();
            FormulaEvaluator evaluator = wb.getCreationHelper().createFormulaEvaluator();

            int whatsappCol = 3;
            int hasWebsiteCol = 4;

            int count = 0;

            for (int i = 0; i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);
                if (row == null) continue;

                // HEADER
                if (i == 0) {
                    writer.write(convertRow(row, formatter, evaluator) + ",Tags");
                    writer.newLine();
                    continue;
                }

                String whatsapp = getValue(row.getCell(whatsappCol), formatter, evaluator);
                String hasWebsite = getValue(row.getCell(hasWebsiteCol), formatter, evaluator);

                if ("No".equalsIgnoreCase(hasWebsite) &&
                        whatsapp != null && !whatsapp.trim().isEmpty()) {

                    writer.write(convertRow(row, formatter, evaluator) + ",first");
                    writer.newLine();
                    count++;
                }
            }

            System.out.println("Done. Rows written: " + count);

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private static String getValue(Cell cell, DataFormatter formatter, FormulaEvaluator evaluator) {
        if (cell == null) return "";
        return formatter.formatCellValue(cell, evaluator).trim();
    }

    private static String convertRow(Row row, DataFormatter formatter, FormulaEvaluator evaluator) {
        StringBuilder sb = new StringBuilder();

        int lastCol = row.getLastCellNum();

        for (int i = 0; i < lastCol; i++) {
            String value = getValue(row.getCell(i), formatter, evaluator);

            if (value.contains(",") || value.contains("\"") || value.contains("\n")) {
                value = "\"" + value.replace("\"", "\"\"") + "\"";
            }

            sb.append(value);

            if (i < lastCol - 1) {
                sb.append(",");
            }
        }

        return sb.toString();
    }
}