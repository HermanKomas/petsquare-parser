package com.petsquare;

import com.sun.org.apache.bcel.internal.generic.INSTANCEOF;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import javax.swing.text.html.ObjectView;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.*;

public class FileProcessor {

    public FileProcessor(){};

    public FileProcessor(String path){
        this.path = path;
    }

    private String path;

    private Object[] headers;

    public XSSFWorkbook openExcelFile() throws IOException{

        File file = new File(path);
        FileInputStream fip = new FileInputStream(file);

        XSSFWorkbook xssfWorkbook = new XSSFWorkbook(fip);

        fip.close();

        if(file.exists() && file.isFile()){
            return xssfWorkbook;
        }else {
            throw new IOException("file is invalid");
        }
    }

    public XSSFWorkbook validateExcelFile(XSSFWorkbook workbook){

        XSSFSheet spreadsheet = workbook.getSheetAt(0);
        Iterator<Row> rowIterator = spreadsheet.iterator();

        Map<String, Object[]> invalidRecords = new TreeMap<String, Object[]>();
        Map<String, Object[]> validRecords = new TreeMap<String, Object[]>();

        setupHeaders(spreadsheet.getRow(0));

        for(Row row : spreadsheet) {
            if(row.getRowNum() == 0)
                continue;

            if(!isRowValid(row)){
                invalidRecords.put(String.valueOf(row.getRowNum()), convertRowToObjectArray(row));
            }else {
                validRecords.put(String.valueOf(row.getRowNum()), convertRowToObjectArray(row));
            }
        }

        workbook = attachInvalidRecordsTab(workbook, invalidRecords);
        workbook = attachValidRecordsTab(workbook, validRecords);

        return workbook;
    }

    public void saveExcelFile(XSSFWorkbook workbook) throws IOException {
        FileOutputStream out = new FileOutputStream(new File(generateExportFileName()));
        workbook.write(out);
        out.close();
    }

    private boolean isRowValid(Row row){

        if(row.getLastCellNum() < 10){
            return false;
        }

        for (int cn = 0; cn < row.getLastCellNum(); cn++) {
            Cell cell = row.getCell(cn, Row.CREATE_NULL_AS_BLANK);

            switch (cell.getColumnIndex()){
                case 0:
                    if(cell.getStringCellValue().equals(null) || cell.getStringCellValue().toUpperCase().contains("NOT GIVEN"))
                        return false;
                    break;
                case 1:
                    if(cell.getStringCellValue().equals(null) || cell.getStringCellValue().toUpperCase().contains("NOT GIVEN"))
                        return false;
                    break;
                case 2:
                    if(cell.getStringCellValue().equals(null) || cell.getStringCellValue().toUpperCase().contains("NOT GIVEN"))
                        return false;
                    break;
                /*case 3:
                    if(cell.getStringCellValue().equals(null))
                        return false;
                    break;
                case 4:
                    if(cell.getStringCellValue().equals(null))
                        return false;
                    break;
                case 5:
                    if(cell.getStringCellValue().equals(null))
                        return false;
                    break;*/
                case 6:
                    if(cell.getStringCellValue() == null || cell.getStringCellValue().toUpperCase().contains(("NOT GIVEN"))){

                        Cell facebookPageAddress = row.getCell(7, Row.CREATE_NULL_AS_BLANK);

                        if(facebookPageAddress.getStringCellValue().equals("Nog given")
                                || facebookPageAddress.getStringCellValue() == null){
                            return false;
                        }else{
                            return true;
                        }
                    }
                    break;
                case 8:
                    if(cell.getStringCellValue() == null || cell.getStringCellValue().equals("Not given"))
                        return false;
                    break;
                /*case 9:
                    if(cell.getStringCellValue().equals(null))
                        return false;
                    break;
                case 10:
                    if(cell.getStringCellValue().equals(null))
                        return false;
                    break;
                case 11:
                    if(cell.getStringCellValue().equals(null))
                        return false;
                    break;
                case 12:
                    if(cell.getStringCellValue().equals(null))
                        return false;
                    break;
                case 13:
                    if(cell.getStringCellValue().equals(null))
                        return false;
                    break;
                case 14:
                    if(cell.getStringCellValue().equals(null))
                        return false;
                    break;
                case 15:
                    if(cell.getStringCellValue().equals(null))
                        return false;
                    break;*/
                case 16:
                    if(cell.getStringCellValue() == null || cell.getStringCellValue().toUpperCase().contains(("NOT GIVEN"))) {
                        Cell operationalArea = row.getCell(17, Row.CREATE_NULL_AS_BLANK);
                        if (operationalArea.getStringCellValue().toUpperCase().contains(("NOT GIVEN"))
                                || cell.getStringCellValue() == null ){
                            return false;
                        }else{
                            return true;
                        }
                    }
                    break;
                /*case 17:
                    if(cell.getStringCellValue() == null || cell.getStringCellValue().toUpperCase().contains(("NOT GIVEN")))
                        return false;
                    break;*/
                /*case 18:
                    if(cell.getStringCellValue().equals(null))
                        return false;
                    break;
                case 19:
                    if(cell.getStringCellValue().equals(null))
                        return false;
                    break;
                case 20:
                    if(cell.getStringCellValue().equals(null))
                        return false;
                    break;
                case 21:
                    if(cell.getStringCellValue().equals(null))
                        return false;
                    break;*/
            }
        }
        return true;
    }

    private void setupHeaders(Row row) {
        List<String> headersList = new ArrayList<>();
        for (int cn = 0; cn < row.getLastCellNum(); cn++) {
            Cell cell = row.getCell(cn, Row.CREATE_NULL_AS_BLANK);
            headersList.add(cell.getStringCellValue());
        }
        this.headers = headersList.toArray();
    }

    private Object[] convertRowToObjectArray(Row row){

        List<Object> cells = new ArrayList<>();

        for (int cn = 0; cn < row.getLastCellNum(); cn++) {
            Cell cell = row.getCell(cn, Row.CREATE_NULL_AS_BLANK);

            switch (cell.getCellType()) {
                case Cell.CELL_TYPE_NUMERIC:
                    cells.add(cell.getNumericCellValue());
                    break;

                case Cell.CELL_TYPE_STRING:
                    cells.add(cell.getStringCellValue());
                    break;
            }
        }
        return cells.toArray();
    }

    private XSSFWorkbook attachInvalidRecordsTab
            (XSSFWorkbook workbook, Map<String, Object[]> invalidRecords){

        XSSFSheet spreadsheet = workbook.createSheet("INVALID");

        spreadsheet = attachRecordsToSpreadsheet(spreadsheet, invalidRecords);

        return workbook;
    }

    private XSSFWorkbook attachValidRecordsTab
            (XSSFWorkbook workbook, Map<String, Object[]> validRecords){

        XSSFSheet spreadsheet = workbook.createSheet("VALID");

        spreadsheet = attachRecordsToSpreadsheet(spreadsheet, validRecords);

        return workbook;
    }

    private XSSFSheet attachRecordsToSpreadsheet(XSSFSheet spreadsheet, Map<String, Object[]> recordsMap){

        CellStyle headerStyle = spreadsheet.getWorkbook().createCellStyle();
        headerStyle.setFillBackgroundColor(IndexedColors.LIGHT_GREEN.getIndex());

        XSSFRow headersRow = spreadsheet.createRow(0);
        int headerCellId = 0;

        for (Object header : headers) {
            Cell cell = headersRow.createCell(headerCellId++);
            cell.setCellStyle(headerStyle);
            cell.setCellValue((String)header);

            spreadsheet.setColumnWidth(headerCellId, 4000);
        }

        CellStyle regularCellStyle = spreadsheet.getWorkbook().createCellStyle();

        XSSFRow row;

        //Iterate over data and write to sheet
        Set < String > keyid = recordsMap.keySet();
        int rowid = 1;

        for (String key : keyid) {
            row = spreadsheet.createRow(rowid++);
            Object [] objectArr = recordsMap.get(key);
            int cellid = 0;

            for (Object obj : objectArr) {
                Cell cell = row.createCell(cellid++);
                cell.setCellStyle(regularCellStyle);
                cell.setCellValue(String.valueOf(obj));
            }
        }


        return spreadsheet;
    }

    private void printWorkbookOut(XSSFWorkbook workbook){

        XSSFSheet spreadsheet = workbook.getSheetAt(0);
        Iterator<Row> rowIterator = spreadsheet.iterator();
        XSSFRow row;

        while (rowIterator.hasNext()) {
            row = (XSSFRow) rowIterator.next();
            Iterator <Cell>  cellIterator = row.cellIterator();

            while ( cellIterator.hasNext()) {
                Cell cell = cellIterator.next();

                switch (cell.getCellType()) {
                    case Cell.CELL_TYPE_NUMERIC:
                        System.out.print(cell.getNumericCellValue() + " \t\t ");
                        break;

                    case Cell.CELL_TYPE_STRING:
                        System.out.print(
                                cell.getStringCellValue() + " \t\t ");
                        break;
                }
            }
            System.out.println();
        }
    }

    private String generateExportFileName(){
        return "VALIDATED_DOGWALKERS_"
                + (new Random()).nextInt(50) + 1 + ".xlsx";
    }
}
