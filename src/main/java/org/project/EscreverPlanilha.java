package org.project;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class EscreverPlanilha {


    public static void escrever(List<Map<String, String>> dados, String caminhoArquivo) throws IOException {
        try (Workbook workbook = new XSSFWorkbook(); // Cria um novo workbook .xlsx
             FileOutputStream fos = new FileOutputStream(caminhoArquivo)) {
            Sheet sheet = workbook.createSheet("Dados Processados"); // Cria uma nova planilha

            Row headerRow = sheet.createRow(0);

            Map<String, String> dePara = DePara.getDePara(); // Obtém o mapeamento de colunas
            Map<String, Integer> colunaEntradaParaIndiceSaida = new HashMap<>(); // Mapeia índice de entrada para índice de saída
            int colunaIndiceSaida = 0;

            // Escreve o cabeçalho na primeira linha
            for (Map.Entry<String, String> entry : dePara.entrySet()) {
                Cell headerCell = headerRow.createCell(colunaIndiceSaida);
                headerCell.setCellValue(entry.getValue()); // Define o nome da coluna de saída
                colunaEntradaParaIndiceSaida.put(entry.getKey(), colunaIndiceSaida); // Armazena o mapeamento de índices
                colunaIndiceSaida++;
            }

            // Escreve os dados nas linhas seguintes
            for (int i = 0; i < dados.size(); i++) {
                Row dataRow = sheet.createRow(i + 1);
                Map<String, String> linhaData = dados.get(i);

                // Itera sobre os mapeamentos definidos para preencher as células
                for (Map.Entry<String, String> entry : dePara.entrySet()) {
                    String colunaEntradaIndice = entry.getKey();

                    if (colunaEntradaParaIndiceSaida.containsKey(colunaEntradaIndice)) {
                        int colunaSaidaIndice = colunaEntradaParaIndiceSaida.get(colunaEntradaIndice);
                        Cell dataCell = dataRow.createCell(colunaSaidaIndice);

                        dataCell.setCellValue(linhaData.getOrDefault(colunaEntradaIndice, ""));
                    } else {
                        System.err.println("Aviso: Mapeamento para o índice de coluna de entrada '" + colunaEntradaIndice + "' não encontrado no ColunaDePara. Esta coluna será ignorada na saída.");
                    }
                }
            }

            // Ajusta a largura das colunas para melhor visualização
            for (int i = 0; i < dePara.size(); i++) {
                sheet.autoSizeColumn(i);
            }

            workbook.write(fos);
        }
    }
}
