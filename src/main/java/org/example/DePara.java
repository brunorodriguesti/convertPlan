package org.example;
import java.util.HashMap;
import java.util.Map;


public class DePara {

        private static final Map<String, String> DE_PARA = new HashMap<>();

        static {
            DE_PARA.put("0", "Apolice");
            DE_PARA.put("1", "Nome");
            DE_PARA.put("2", "Documento");
            DE_PARA.put("3", "Cobertura");
            DE_PARA.put("4", "Grupo");
            DE_PARA.put("4", "Ramo");
            DE_PARA.put("6", "Cobertura");
            DE_PARA.put("6", "Status_Requisicao");

        }

        public static Map<String, String> getDePara() {
            return DE_PARA;

    }

}

