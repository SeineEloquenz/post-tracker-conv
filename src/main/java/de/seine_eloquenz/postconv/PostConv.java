package de.seine_eloquenz.postconv;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.swing.*;
import javax.swing.filechooser.FileNameExtensionFilter;
import java.io.*;

/**
 * Class for converting tracker data from one format used by the Deutsche Post AG in another format for further usage
 */
public class PostConv {

    /**
     * Name of the excel sheet to convert
     */
    private static final String SHEET_NAME = "BZA";

    /**
     * Index of the column to convert
     */
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

    /**
     * Converts the given tracker file
     * @param file file of the excel sheet
     * @throws IOException if an error occurs computing the changes
     */
    public void convert(File file) throws IOException {
        Workbook book = readExcel(file);
        replaceColons(book);
        writeExcel(new File(file.getName().replace(".xlsx", "_converted.xlsx")), book);
    }

    /**
     * Removes all colons in the column
     * @param book workbook to edit
     */
    public void replaceColons(Workbook book) {
         Sheet sheet = book.getSheet(SHEET_NAME);
        sheet.rowIterator().forEachRemaining(
                r -> r.getCell(COLUMN_INDEX).setCellValue(Double.parseDouble(r.getCell(COLUMN_INDEX).getStringCellValue().replace("-",""))));
    }

    /**
     * Reads in an excel {@link File} and returns a {@link Workbook}
     * @param file file to read
     * @return Workbook representation of the excel file
     * @throws IOException thrown if errors occur reading the file
     */
    public Workbook readExcel(File file) throws IOException {
        FileInputStream fis = new FileInputStream(file);
        return new XSSFWorkbook(fis);
    }

    /**
     * Writes out the given {@link Workbook} to he given {@link File}
     * @param file file to write to
     * @param book book to write
     * @throws IOException if an error occurs writing the file
     */
    public void writeExcel(File file, Workbook book) throws IOException {
        OutputStream os = new FileOutputStream(file);
        book.write(os);
    }
}
