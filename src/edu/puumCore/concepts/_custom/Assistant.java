package edu.puumCore.concepts._custom;

import edu.puumCore.concepts._interfaces.Excel;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTTable;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTTableColumn;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTTableColumns;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTTableStyleInfo;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * @author Puum Core (Mandela Muriithi)<br>
 * <a href = "https://github.com/puumCore">GitHub: Mandela Muriithi</a>
 * @version 1.0
 * @since 29/06/2022
 */

public class Assistant implements Excel {

    @Override
    public void write_to_file(File targetFile) throws IOException {
        Workbook workbook = new XSSFWorkbook();

        Sheet sheet = workbook.createSheet("Persons");
        sheet.setColumnWidth(0, 6000);
        sheet.setColumnWidth(1, 4000);

        Row header = sheet.createRow(0);

        CellStyle headerStyle = workbook.createCellStyle();
        headerStyle.setFillForegroundColor(IndexedColors.LIGHT_BLUE.getIndex());
        headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        XSSFFont font = (XSSFFont) workbook.createFont();
        font.setFontName("Arial");
        font.setFontHeightInPoints((short) 16);
        font.setBold(true);
        headerStyle.setFont(font);

        Cell headerCell = header.createCell(0);
        headerCell.setCellValue("Name");
        headerCell.setCellStyle(headerStyle);

        headerCell = header.createCell(1);
        headerCell.setCellValue("Age");
        headerCell.setCellStyle(headerStyle);

        CellStyle style = workbook.createCellStyle();
        style.setWrapText(true);

        Row row = sheet.createRow(2);
        Cell cell = row.createCell(0);
        cell.setCellValue("John Smith");
        cell.setCellStyle(style);

        cell = row.createCell(1);
        cell.setCellValue(20);
        cell.setCellStyle(style);

        FileOutputStream outputStream = new FileOutputStream(targetFile);
        workbook.write(outputStream);
        workbook.close();
    }

    @Override
    public void read_file(File targetFile) throws IOException {
        Workbook workbook = new XSSFWorkbook(new FileInputStream(targetFile));
        Sheet sheet = workbook.getSheetAt(0);

        Map<Integer, List<String>> data = new HashMap<>();
        int i = 0;
        for (Row row : sheet) {
            data.put(i, new ArrayList<>());
            for (Cell cell : row) {
                switch (cell.getCellType()) {
                    case STRING:
                        data.get(i).add(cell.getRichStringCellValue().getString());
                        break;
                    case NUMERIC:
                        if (DateUtil.isCellDateFormatted(cell)) {
                            data.get(i).add(cell.getDateCellValue() + "");
                        } else {
                            data.get(i).add(cell.getNumericCellValue() + "");
                        }
                        break;
                    case BOOLEAN:
                        data.get(i).add(cell.getBooleanCellValue() + "");
                        break;
                    case FORMULA:
                        data.get(i).add(cell.getCellFormula() + "");
                        break;
                    default:
                        data.get(i).add(" ");
                }
            }
            i++;
        }



    }

    @Override
    public File create_table(File targetFile) throws IOException {
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("Architecture");

        XSSFTable table = sheet.createTable(null); //since apache poi 4.0.0
        CTTable cttable = table.getCTTable();

        cttable.setDisplayName("Table1");
        cttable.setId(1);
        cttable.setName("Test");
        cttable.setRef("A1:C11");//io 11 should be replaced na the no of rows being created
        cttable.setTotalsRowShown(true);

        CTTableStyleInfo styleInfo = cttable.addNewTableStyleInfo();
        styleInfo.setName("TableStyleMedium2");
        styleInfo.setShowColumnStripes(false);
        styleInfo.setShowRowStripes(true);

        CTTableColumns columns = cttable.addNewTableColumns();
        columns.setCount(3);
        for (int i = 1; i <= 3; i++) {
            CTTableColumn column = columns.addNewTableColumn();
            column.setId(i);
            column.setName("Column" + i);
        }

        for (int r = 0; r < 2; r++) {
            XSSFRow row = sheet.createRow(r);
            for(int c = 0; c < 3; c++) {
                XSSFCell cell = row.createCell(c);
                if(r == 0) { //first row is for column headers
                    cell.setCellValue("Column"+ (c+1)); //content **must** be here for table column names
                } else {
                    cell.setCellValue("Data R"+ (r+1) + "C" + (c+1));
                }
            }
        }

        try (FileOutputStream outputStream = new FileOutputStream(targetFile)) {
            workbook.write(outputStream);
        }
        return targetFile;
    }
}
