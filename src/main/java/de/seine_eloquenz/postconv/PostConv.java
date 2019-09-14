package de.seine_eloquenz.postconv;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.swing.*;
import javax.swing.filechooser.FileNameExtensionFilter;
import java.io.*;

public class PostConv {

    private static final String SHEET_NAME = "BZA";
    private static final int COLUMN_INDEX = 6;

    public static void main(String[] args) throws IOException {
        JFileChooser chooser = new JFileChooser();
        chooser.setFileFilter(new FileNameExtensionFilter("Excel (.xlsx)", "xlsx"));        int state = chooser.showOpenDialog(null);
        if (state == JFileChooser.APPROVE_OPTION) {
            File file = chooser.getSelectedFile();
            PostConv conv = new PostConv();
            conv.convert(file);

        }
    }

    public void convert(File file) throws IOException {
        Workbook book = readExcel(file);
        replaceColons(book);
        writeExcel(new File(file.getName().replace(".xlsx", "_converted.xlsx")), book);
    }

    public void replaceColons(Workbook book) {
         Sheet sheet = book.getSheet(SHEET_NAME); //TODO
        sheet.rowIterator().forEachRemaining(
                r -> r.getCell(COLUMN_INDEX).setCellValue(Double.parseDouble(r.getCell(COLUMN_INDEX).getStringCellValue().replace("-",""))));
    }

    public Workbook readExcel(File file) throws IOException {
        FileInputStream fis = new FileInputStream(file);
        return new XSSFWorkbook(fis);
    }

    public void writeExcel(File file, Workbook book) throws IOException {
        OutputStream os = new FileOutputStream(file);
        book.write(os);
    }
}
