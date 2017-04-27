/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package excelread;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.logging.Level;
import java.util.logging.Logger;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author Bilal
 */
public class ExcelRead {

    /**
     * @param args the command line arguments
     */
    public static void main(String[] args) {
        try {
            FileInputStream file = new FileInputStream(new File("E:\\5th semester\\DB system\\CS204-Database+Systems+-Fall+2016-Semester+Work-Sec+A.xlsx"));
            XSSFWorkbook hssf = new XSSFWorkbook(file);
            XSSFSheet hsheet = hssf.getSheetAt(0);
            
            FormulaEvaluator formulaevaluator = hssf.getCreationHelper().createFormulaEvaluator();
            
            for(Row row : hsheet)
            {
                for(Cell cell : row)
                {
                    switch(formulaevaluator.evaluateInCell(cell).getCellType())
                    {
                        case Cell.CELL_TYPE_NUMERIC:
                            System.out.print(cell.getNumericCellValue()+" ");
                            break;
                        
                        case Cell.CELL_TYPE_STRING:
                            System.out.print(cell.getStringCellValue()+" ");
                            break;
                        
                    }                  
                }
                System.out.println("");
            }
            
            // TODO code application logic here
        } catch (FileNotFoundException ex) {
            Logger.getLogger(ExcelRead.class.getName()).log(Level.SEVERE, null, ex);
        } catch (IOException ex) {
            Logger.getLogger(ExcelRead.class.getName()).log(Level.SEVERE, null, ex);
        }
    }    
}
