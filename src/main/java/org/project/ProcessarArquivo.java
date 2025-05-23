package org.project;

import java.io.IOException;
import java.util.List;
import java.util.Map;

public class ProcessarArquivo {

    public static void main(String[] args) {
        String arquivoEntrada = "C:\\files\\plan1.xlsx";
        String arquivoSaida = "C:\\files\\saida.xlsx";

        try {


            List<Map<String, String>> dados = LerPlanilha.ler(arquivoEntrada);
            EscreverPlanilha.escrever(dados, arquivoSaida);
            System.out.println("Processamento concluído com sucesso!");
        } catch (IOException e) {
            System.err.println("Ocorreu um erro durante o processamento do Excel: " + e.getMessage());
            e.printStackTrace();
        }
    }

}
