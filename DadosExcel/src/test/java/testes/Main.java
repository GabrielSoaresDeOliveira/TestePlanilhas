package testes;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Iterator;

public class Main{

    @Test
    public void Main(){

        FileInputStream fisPlanilha = null;
        try {
            File file = new File("C:\\Users\\gabriel.soares\\Desktop\\PlanilhaTeste.xlsx");

            fisPlanilha = new FileInputStream(file);

            //Cria um workbook = planilha com todas as abas.
            XSSFWorkbook workbook = new XSSFWorkbook(fisPlanilha);

            //Primeira aba.
            XSSFSheet sheet = workbook.getSheetAt(0);

            //Retorna todas as linhas da planilha 0.
            Iterator<Row> rowIterator = sheet.iterator();

            //Varre todas as linhas da planilha 0.
            while (rowIterator.hasNext()){

                //Recebe cada linha da planilha.
                Row row = rowIterator.next();

                //Pega todas as celulas dessa linha.
                Iterator<Cell> cellIterator = row.iterator();

                //Varre todas as celulas dessa linha.
                while (cellIterator.hasNext()){

                    //Criamos uma celula.
                    Cell cell = cellIterator.next();

                    switch (cell.getCellType()){

                        case Cell.CELL_TYPE_STRING:
                            System.out.println("String: " +cell.getStringCellValue());
                            break;

                        case Cell.CELL_TYPE_NUMERIC:
                            System.out.println("Numero: " +cell.getNumericCellValue());
                            break;

                        case Cell.CELL_TYPE_FORMULA:
                            System.out.println("Formular: " +cell.getCellFormula());
                            break;

                    }

                }

            }




        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }


    }

}