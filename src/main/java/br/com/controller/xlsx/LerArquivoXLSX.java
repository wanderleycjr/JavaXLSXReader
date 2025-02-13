package br.com.controller.xlsx;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.Iterator;

public class LerArquivoXLSX {

    public static void main(String[] args) {
        String caminhoArquivo = "C:\\Excel\\Planilhas\\excel_java.xlsx";
        
        try (
        	//Recupera o arquivo	
        	FileInputStream planilha = new FileInputStream(new File(caminhoArquivo));
        		
        	//Lê todas as abas da planilha
            XSSFWorkbook workbook = new XSSFWorkbook(planilha)) {
            
        	//Recupera a primeira aba => (0)
            XSSFSheet sheet = workbook.getSheetAt(0);
            
            //Retorna todas as linhas da primeira aba 1
            Iterator<Row> rowIterator = sheet.iterator();
            
            while (rowIterator.hasNext()) {
                Row row = rowIterator.next(); //Recupera cada linha da planilha
                Iterator<Cell> cellIterator = row.cellIterator();
                
                while (cellIterator.hasNext()) {//Varrendo as linhas da planilha
                    Cell cell = cellIterator.next(); //Próximas células
                    
                    //Switch/Case para tratar todos os tipos de células e separando por tabulação (\t)
                    switch (cell.getCellType()) {
                        case STRING:
                            System.out.print(cell.getStringCellValue() + "\t");
                            break;
                        case NUMERIC:
                            System.out.print(cell.getNumericCellValue() + "\t");
                            break;
                        case BOOLEAN:
                            System.out.print(cell.getBooleanCellValue() + "\t");
                            break;
                        case FORMULA:
                            System.out.print(cell.getCellFormula() + "\t");
                            break;
                        default:
                            System.out.print("\t");
                    }
                }
                System.out.println(); // Nova linha para próxima linha da planilha
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}