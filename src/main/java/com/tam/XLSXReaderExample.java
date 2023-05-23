package com.tam;

import org.apache.poi.ss.usermodel.*;

import java.io.FileInputStream;
import java.util.Iterator;

public class XLSXReaderExample {
    public static void main(String[] args) {
        try {
            int a = 0;
            for (int i = 1; i <= 9; i++) {
                Workbook workbook = WorkbookFactory.create(new FileInputStream("C:\\Users\\r.ramezani\\Desktop\\bank\\" + i + ".xlsx"));   //creating a new file instance
                Sheet sheet = workbook.getSheetAt(0);
                Iterator<Row> itr = sheet.iterator();
                DataFormatter dataFormatter = new DataFormatter();
                FormulaEvaluator formulaEvaluator = workbook.getCreationHelper().createFormulaEvaluator();

                while (itr.hasNext()) {
                    Row row = itr.next();
                    Iterator<Cell> cellIterator = row.cellIterator();   //iterating over each column
                    while (cellIterator.hasNext()) {
                        Cell cell = cellIterator.next();
                        String cellValue = dataFormatter.formatCellValue(cell, formulaEvaluator);
                        System.out.println(cellValue);
                    }
                }
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}  