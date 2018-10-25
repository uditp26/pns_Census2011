package com.dmg.extractor;

import com.google.gson.Gson;
import com.google.gson.JsonObject;
import com.google.gson.JsonParser;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.util.CellAddress;

import java.io.*;
import java.util.ArrayList;
import java.util.List;

public class Main_Backup {

    static File destFile;
    static FileInputStream dfs;
    static HSSFWorkbook dwb;
    static HSSFSheet dsheet;

    static int destSheetLastRowNo;


    public static void main(String[] args) throws IOException {

//        String dirPath = "/media/mohit/New Volume/IIITD/Semeseter/Subjects/DMG/Project/Data Mining Project/Extraction/";
//
//        String filePath = "/media/mohit/New Volume/IIITD/Semeseter/Subjects/DMG/Project/DataMiningProject/Extraction/2425-Surat/T39_2425.xls";
//
//        String addresses = "D16,E16,F16,G16,H16,I16,J16,K16";
//
//        String[] addrArray = addresses.split(",");
//
//        List<String> addressesToRead = getAllCellAddressesInRange(addrArray);
//
//        readExcelFile(filePath, 0, addressesToRead);
        String jsonFilePath = "/home/mohit/Downloads/excel_mappings.json";

        Gson gson = new Gson();

        BufferedReader br = new BufferedReader(new FileReader(jsonFilePath));
        JsonParser parser = new JsonParser();
        JsonObject object = parser.parse(br).getAsJsonObject();

        JSONMappingModel model = gson.fromJson(object, JSONMappingModel.class);

        String dirPath = model.getSourceFolder();
        String targetFilePath = model.getTargetFilePath();


        destFile = new File(targetFilePath);
        dfs = new FileInputStream(destFile);
        dwb = new HSSFWorkbook(dfs);
        dsheet = dwb.getSheetAt(0);
        destSheetLastRowNo = dsheet.getPhysicalNumberOfRows();

        //System.out.println("model.getTargetFilePath() "+model.getAllExcelMappings().get(0).getCellMappings().get(0).getDestinationCellName());
        iterateFolders(dirPath, model);

    }

    private static void iterateFolders(String dirPath, JSONMappingModel model) {
        File root = new File(dirPath);
        File[] list = root.listFiles();
        if (list == null) return;
        for (File f : list) {
            if (f.isDirectory()) {
                iterateFolders(f.getAbsolutePath(), model);
                System.out.println("Dir:" + f.getAbsoluteFile().getName());
            } else {
                //System.out.println("File:" + f.getAbsoluteFile().getName());

                AllExcelMapping allExcelMapping = new AllExcelMapping();
                allExcelMapping.setExcelName(f.getAbsoluteFile().getName());
                if (model.getAllExcelMappings().contains(allExcelMapping)) {
                    System.out.println("@@@@@Yes matched " + model.getAllExcelMappings().indexOf(allExcelMapping));
                    int matchedIDX = model.getAllExcelMappings().indexOf(allExcelMapping);
                    List<CellMapping> cellMappings = model.getAllExcelMappings().get(matchedIDX).getCellMappings();
                    readWriteExcelFile(f.getAbsolutePath(), model.getTargetFilePath(), 0, cellMappings);
                }
            }
        }

    }


    private static void readWriteExcelFile(String sourceFilePath, String destinationFilePath, int sheetIndex, List<CellMapping> cellReferences) {
        try {
            File file = new File(sourceFilePath);
            POIFSFileSystem fs = new POIFSFileSystem(new FileInputStream(file));
            HSSFWorkbook wb = new HSSFWorkbook(fs);
            HSSFSheet sheet = wb.getSheetAt(0);


//            File destFile = new File(destinationFilePath);
//            FileInputStream dfs = new FileInputStream(destFile);
//           // POIFSFileSystem dfs = new POIFSFileSystem(new FileInputStream(destFile));
//            HSSFWorkbook dwb = new HSSFWorkbook(dfs);
//            HSSFSheet dsheet = dwb.getSheetAt(0);
//            HSSFRow drow;
//            HSSFCell dcell;


            for (CellMapping cellMapping : cellReferences) {

                String sourceCellAddress = getCellAddressFromCellMapping(cellMapping.getSourceRowPosition(), cellMapping.getSourceCellName(), sheet.getPhysicalNumberOfRows());
                String destinationCellAddress = getCellAddressFromCellMapping(cellMapping.getDestinationRowPosition(), cellMapping.getDestinationCellName(), destSheetLastRowNo);

                System.out.println("sourceCellAddress " + sourceCellAddress);
                System.out.println("destinationCellAddress " + destinationCellAddress);

                String sourceCellRowCol = getColumnNumberFromCellAddress(sourceCellAddress);
                String destinationRowCol = getColumnNumberFromCellAddress(destinationCellAddress);

                String srcRowNo = sourceCellRowCol.split(",")[0];
                String srcColNo = sourceCellRowCol.split(",")[1];

                HSSFRow row = sheet.getRow(Integer.parseInt(srcRowNo));
                HSSFCell cell = row.getCell(Integer.parseInt(srcColNo));

                                String srcCellValue = "";
                if(cell.getCellType() == CellType.NUMERIC){
                    srcCellValue = String.valueOf(cell.getNumericCellValue());
                }else if(cell.getCellType() == CellType.FORMULA){
                    srcCellValue = String.valueOf(cell.getNumericCellValue());
                } else
                    {
                    srcCellValue = new DataFormatter().formatCellValue(cell);
                }
                  System.out.println("srcCellValue " + srcCellValue);

                String destRowNo = destinationRowCol.split(",")[0];
                String destColNo = destinationRowCol.split(",")[1];

                System.out.println("destRowNo "+destRowNo);
                System.out.println("destColNo "+destColNo);


                HSSFRow drow = dsheet.getRow(Integer.parseInt(destRowNo));

                if (drow == null) {
                    // System.out.println("drow is null "+drow);
                    drow = dsheet.createRow(Integer.parseInt(destRowNo));
                }
                HSSFCell dcell = drow.getCell(Integer.parseInt(destColNo));

                if (dcell == null) {
                    //  System.out.println("dcell is null "+drow);
                    dcell = drow.createCell(Integer.parseInt(destColNo));
                }

                dcell.setCellValue(srcCellValue);

            }

            dfs.close();

            FileOutputStream output_file = new FileOutputStream(destFile);
            dwb.write(output_file);
            output_file.close();


        } catch (Exception ioe) {
            ioe.printStackTrace();
        }
    }

    private static String getCellAddressFromCellMapping(String rowPosition, String cellName, int totalRowsInSheet) {

        String finalSrcRowNoStr = null;
        String sourceRowPosition = rowPosition;
        if (sourceRowPosition.startsWith("@/")) { //calculate row number based on number of rows in sheet
            int rows = totalRowsInSheet;

            //System.out.println("totalRowsInSheet "+totalRowsInSheet);

            //int idx = sourceRowPosition.indexOf("@/");
            String rowFromLast = sourceRowPosition.replace("@/", "");

            //System.out.println("rowFromLast "+rowFromLast);

            int finalSrcRowNo = Integer.parseInt(rowFromLast);

            finalSrcRowNo = rows + finalSrcRowNo;
            finalSrcRowNoStr = String.valueOf(finalSrcRowNo);

        } else {
            finalSrcRowNoStr = sourceRowPosition;
        }

        int idx = cellName.indexOf("@/");
        // System.out.println("idx "+idx);
        String colName = cellName.substring(0, idx);

        //System.out.println("colName+finalSrcRowNoStr "+ colName+finalSrcRowNoStr);

        return colName + finalSrcRowNoStr;
    }

//    private static int getLastRowOfExcel(HSSFSheet sheet){
//        int lastRowIndex = -1;
//        if( sheet.getPhysicalNumberOfRows() > 0 )
//        {
//            // getLastRowNum() actually returns an index, not a row number
//            lastRowIndex = sheet.getLastRowNum();
//
//            // now, start at end of spreadsheet and work our way backwards until we find a row having data
//            for( ; lastRowIndex >= 0; lastRowIndex-- ){
//                Row row = sheet.getRow( lastRowIndex );
//                if( row != null ){
//                    return row.getRowNum();
//                    //break;
//                }
//            }
//        }
//
//        return -1;
//    }

//
//    private static int determineRowCount(HSSFSheet sheet )
//    {
////        HSSFFormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
////        this.formatter = new DataFormatter( true );
//
//        int lastRowIndex = -1;
//        if( sheet.getPhysicalNumberOfRows() > 0 )
//        {
//            // getLastRowNum() actually returns an index, not a row number
//            lastRowIndex = sheet.getLastRowNum();
//
//            // now, start at end of spreadsheet and work our way backwards until we find a row having data
//            for( ; lastRowIndex >= 0; lastRowIndex-- )
//            {
//                Row row = sheet.getRow( lastRowIndex );
//                if( !checkIfRowIsEmpty( row ) )
//                {
//                    break;
//                }
//            }
//        }
//        return lastRowIndex;
//    }

//    /**
//     * Determine whether a row is effectively completely empty - i.e. all cells either contain an empty string or nothing.
//     */
//    private boolean isRowEmpty( Row row )
//    {
//        if( row == null ){
//            return true;
//        }
//
//        int cellCount = row.getLastCellNum() + 1;
//        for( int i = 0; i < cellCount; i++ ){
//            String cellValue = getCellValue( row, i );
//            if( cellValue != null && cellValue.length() > 0 ){
//                return false;
//            }
//        }
//        return true;
//    }
//
//    private static boolean checkIfRowIsEmpty(Row row) {
//        if (row == null) {
//            return true;
//        }
//        if (row.getLastCellNum() <= 0) {
//            return true;
//        }
//        for (int cellNum = row.getFirstCellNum(); cellNum < row.getLastCellNum(); cellNum++) {
//            Cell cell = row.getCell(cellNum);
//            if (cell != null && cell.getCellTypeEnum() != CellType.BLANK && StringUtils.isNotBlank(cell.toString())) {
//                return false;
//            }
//        }
//        return true;
//    }

//    /**
//     * Get the effective value of a cell, formatted according to the formatting of the cell.
//     * If the cell contains a formula, it is evaluated first, then the result is formatted.
//     *
//     * @param row the row
//     * @param columnIndex the cell's column index
//     * @return the cell's value
//     */
//    private String getCellValue( Row row, int columnIndex )
//    {
//        String cellValue;
//        Cell cell = row.getCell( columnIndex );
//        if( cell == null ){
//            // no data in this cell
//            cellValue = null;
//        }
//        else{
//            if( cell.getCellType() != Cell.CELL_TYPE_FORMULA ){
//                // cell has a value, so format it into a string
//                cellValue = this.formatter.formatCellValue( cell );
//            }
//            else {
//                // cell has a formula, so evaluate it
//                cellValue = this.formatter.formatCellValue( cell, this.evaluator );
//            }
//        }
//        return cellValue;
//    }

    private static void readExcelFile(String filePath, int sheetIndex, List<String> cellReferences) {
        try {
            File file = new File(filePath);
            POIFSFileSystem fs = new POIFSFileSystem(new FileInputStream(file));
            HSSFWorkbook wb = new HSSFWorkbook(fs);
            HSSFSheet sheet = wb.getSheetAt(sheetIndex);
            HSSFRow row;
            HSSFCell cell;

            for (String cellAddress : cellReferences) {
                String rowNo = cellAddress.split(",")[0];
                String colNo = cellAddress.split(",")[1];

                int rowIdx = -1;
                int colIdx;
                if (!rowNo.equals("@")) { //TODO not Last Row
                    rowIdx = Integer.parseInt(rowNo);
                }
                colIdx = Integer.parseInt(colNo);
                row = sheet.getRow(rowIdx);
                cell = row.getCell(colIdx);

                if (cell.getCellType() == CellType.NUMERIC)
                    System.out.println(cell.getNumericCellValue());
            }


            int rows; // No of rows
            rows = sheet.getPhysicalNumberOfRows();

            int cols = 0; // No of columns
            int tmp = 0;

            // This trick ensures that we get the data properly even if it doesn't start from first few rows
            for (int i = 0; i < 10 || i < rows; i++) {
                row = sheet.getRow(i);
                if (row != null) {
                    tmp = sheet.getRow(i).getPhysicalNumberOfCells();
                    if (tmp > cols) cols = tmp;
                }
            }

            for (int r = 0; r < rows; r++) {
                row = sheet.getRow(r);
                if (row != null) {
                    for (int c = 0; c < cols; c++) {
                        cell = row.getCell((short) c);
                        if (cell != null) {
                            // Your code here
                        }
                    }
                }
            }
        } catch (Exception ioe) {
            ioe.printStackTrace();
        }
    }

    private static List<String> getAllCellAddressesInRange(String[] allAddresses) {
        List<String> addressesToRead = new ArrayList<String>();

        for (int i = 0; i < allAddresses.length; i++) {
            String addressRange = allAddresses[i];
            addressesToRead.add(getColumnNumberFromCellAddress(addressRange));
        }

        return addressesToRead;

    }

//    private static int convertColStringToIndex(String colName){
//        int colIndex = CellReference.convertColStringToIndex(colName);
//        System.out.println("colIndex "+colIndex);
//        return colIndex;
//    }

    private static String getColumnNumberFromCellAddress(String address) {
        CellAddress cellAddress = new CellAddress(address);
        int columnNo = cellAddress.getColumn();
        int rowNo = cellAddress.getRow();
        String addressInRowColIdx = String.valueOf(rowNo) + "," + String.valueOf(columnNo);
        //System.out.println("addressInRowColIdx "+addressInRowColIdx);
        return addressInRowColIdx;
    }
}
