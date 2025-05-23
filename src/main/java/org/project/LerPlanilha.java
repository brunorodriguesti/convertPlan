package org.project;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class LerPlanilha {

    public static List<Map<String, String>> ler(String caminhoArquivo) throws IOException {
        List<Map<String, String>> linhas = new ArrayList<>();
        Workbook workbook = null;
        FileInputStream fist = new FileInputStream(caminhoArquivo);
        System.out.println(fist.toString());
        try (FileInputStream fis = new FileInputStream(caminhoArquivo);) { // Usa XSSFWorkbook para .xlsx
            if (caminhoArquivo.toLowerCase().endsWith(".xlsx")) {
                workbook = new XSSFWorkbook(fis);
            } else if (caminhoArquivo.toLowerCase().endsWith(".xls")) {
                workbook = new HSSFWorkbook(fis);
            } else {
                throw new IllegalArgumentException("Formato de arquivo não suportado. Use .xls ou .xlsx.");
            }
            Sheet sheet = workbook.getSheetAt(0); // Tenta pegar a primeira planilha

            if (sheet == null) {
                System.err.println("Erro: A primeira planilha (índice 0) não foi encontrada no arquivo Excel.");
                return linhas;
            }

            int primeiraLinha = 10;
            int ultimaColuna = 6;

            for (int i = primeiraLinha; i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);

                if (row != null) {
                    Map<String, String> linhaData = new HashMap<>();
                    for (int j = 0; j <= ultimaColuna; j++) {
                        Cell cell = row.getCell(j);
                        String valor = "";
                        if (cell != null) {
                            switch (cell.getCellType()) {
                                case Cell.CELL_TYPE_STRING:
                                    valor = cell.getStringCellValue();
                                    break;
                                case Cell.CELL_TYPE_NUMERIC:
                                    if (DateUtil.isCellDateFormatted(cell)) {
                                        valor = cell.getDateCellValue().toString();
                                    } else {
                                        valor = String.valueOf(cell.getNumericCellValue());
                                    }
                                    break;
                                case Cell.CELL_TYPE_BOOLEAN:
                                    valor = String.valueOf(cell.getBooleanCellValue());
                                    break;
                                case Cell.CELL_TYPE_FORMULA:
                                    try {
                                        valor = String.valueOf(cell.getNumericCellValue());
                                    } catch (IllegalStateException e) {
                                        valor = cell.getStringCellValue();
                                    } catch (Exception e) {
                                        valor = cell.getCellFormula();
                                    }
                                    break;
                                case Cell.CELL_TYPE_BLANK:
                                    valor = "";
                                    break;
                                default:
                                    valor = "";
                                    break;
                            }
                        }

                        linhaData.put(String.valueOf(j), valor);
                    }
                    linhas.add(linhaData);
                }
            }
        }
        return linhas;
    }

}
