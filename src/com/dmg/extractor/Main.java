package com.dmg.extractor;

import com.google.gson.Gson;
import com.google.gson.JsonObject;
import com.google.gson.JsonParser;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;

import java.util.List;

public class Main {

    private static File destFile;
    private static FileInputStream dfs;
    private static HSSFWorkbook dwb;
    private static HSSFSheet dsheet;
    private static Sheet ssheet;

    private static int destSheetLastRowNo;

    private static JSONMappingModel model;

    private static FormulaEvaluator objFormulaEvaluator;

    private static String JSON_FILE_PATH = "./resources/excel_mappings.json";


    public static void main(String[] args) throws IOException {

        String jsonFilePath = JSON_FILE_PATH;

        Gson gson = new Gson();

        BufferedReader br = new BufferedReader(new FileReader(jsonFilePath));
        JsonParser parser = new JsonParser();
        JsonObject object = parser.parse(br).getAsJsonObject();

        model = gson.fromJson(object, JSONMappingModel.class);

        String dirPath = model.getSourceFolder();
        String targetFilePath = model.getTargetFilePath();


        destFile = new File(targetFilePath);
        dfs = new FileInputStream(destFile);
        dwb = new HSSFWorkbook(dfs);
        dsheet = dwb.getSheetAt(0);
        destSheetLastRowNo = dsheet.getPhysicalNumberOfRows();

        iterateFolders(dirPath);

    }

    private static void iterateFolders(String dirPath) {
        File root = new File(dirPath);
        File[] list = root.listFiles();
        if (list == null) return;
        for (File f : list) {
            if (f.isDirectory()) {
                iterateFolders(f.getAbsolutePath());
                System.out.println("Dir:" + f.getAbsoluteFile().getName());
                destSheetLastRowNo = destSheetLastRowNo + 1; //adding 1 row after processing a folder
            } else {
                //System.out.println("File:" + f.getAbsoluteFile().getName());

                AllExcelMapping allExcelMapping = new AllExcelMapping();
                allExcelMapping.setExcelName(f.getAbsoluteFile().getName());
                if (model.getAllExcelMappings().contains(allExcelMapping)) {
                    //System.out.println("@@@@@Yes matched " + model.getAllExcelMappings().indexOf(allExcelMapping));
                    int matchedIDX = model.getAllExcelMappings().indexOf(allExcelMapping);
                    List<CellMapping> cellMappings = model.getAllExcelMappings().get(matchedIDX).getCellMappings();

                    String parentFolderName = f.getParentFile().getName();

                    String stateName = f.getParentFile().getParentFile().getName();

                    int idx = parentFolderName.indexOf("-");

                    String districtName = parentFolderName.substring(idx + 1);

                    readWriteExcelFile(f.getAbsolutePath(), stateName, districtName, cellMappings);
                }
            }
        }

    }


    private static void readWriteExcelFile(String sourceFilePath, String stateName, String districtName, List<CellMapping> cellReferences) {
        try {
            File file = new File(sourceFilePath);
            if(sourceFilePath.endsWith(".xls")) {
                POIFSFileSystem fs = new POIFSFileSystem(new FileInputStream(file));
                HSSFWorkbook wb = new HSSFWorkbook(fs);
                ssheet = wb.getSheetAt(0); //TODO sheet number 0 hardcoded
            }else if (sourceFilePath.endsWith(".xlsx")){
                XSSFWorkbook wb = new XSSFWorkbook(new FileInputStream(file));
                ssheet = wb.getSheetAt(0); //TODO sheet number 0 hardcoded
            }

            writeCellAtRowColInExcel(String.valueOf(destSheetLastRowNo), "0", stateName); //TODO  state name column hardcoded
            writeCellAtRowColInExcel(String.valueOf(destSheetLastRowNo), "1", districtName); //TODO  district name column hardcoded

            objFormulaEvaluator = ssheet.getWorkbook().getCreationHelper().createFormulaEvaluator();

            for (CellMapping cellMapping : cellReferences) {

                String sourceCellAddress = getCellAddressFromCellMapping(cellMapping.getSourceRowPosition(), cellMapping.getSourceSearchString(), cellMapping.getSourceCellName(), ssheet.getPhysicalNumberOfRows());
                String destinationCellAddress = getCellAddressFromCellMapping(cellMapping.getDestinationRowPosition(), null, cellMapping.getDestinationCellName(), destSheetLastRowNo);

//                System.out.println("sourceCellAddress " + sourceCellAddress);
//                System.out.println("destinationCellAddress " + destinationCellAddress);

                String sourceCellRowCol = getColumnNumberFromCellAddress(sourceCellAddress);
                String destinationRowCol = getColumnNumberFromCellAddress(destinationCellAddress);

                String srcRowNo = sourceCellRowCol.split(",")[0];
                String srcColNo = sourceCellRowCol.split(",")[1];

                Row row = ssheet.getRow(Integer.parseInt(srcRowNo));
                Cell cell = row.getCell(Integer.parseInt(srcColNo));

                String srcCellValue = getCellValue(cell);

                String destRowNo = destinationRowCol.split(",")[0];
                String destColNo = destinationRowCol.split(",")[1];

//                System.out.println("destRowNo "+destRowNo);
//                System.out.println("destColNo "+destColNo);

                writeCellAtRowColInExcel(destRowNo, destColNo, srcCellValue);

            }

            dfs.close();

            FileOutputStream output_file = new FileOutputStream(destFile);
            dwb.write(output_file);
            output_file.close();


        } catch (Exception ioe) {
            ioe.printStackTrace();
        }
    }

    private static String getCellValue(Cell cell) {
        DataFormatter objDefaultFormat = new DataFormatter();
        objFormulaEvaluator.evaluate(cell); // This will evaluate the cell, And any type of cell will return string value
        String cellValueStr = objDefaultFormat.formatCellValue(cell, objFormulaEvaluator);
        return cellValueStr;
    }

    private static void writeCellAtRowColInExcel(String destRowNo, String destColNo, String cellValue) {
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

        dcell.setCellValue(cellValue);
    }

    private static String getCellAddressFromCellMapping(String rowPosition, String searchString, String cellName, int totalRowsInSheet) {
        String finalSrcRowNoStr = null;
        String sourceRowPosition = rowPosition;
        if (sourceRowPosition.startsWith("@/")) { //calculate row number based on number of rows in sheet
            int rows = totalRowsInSheet;

            String rowFromLast = sourceRowPosition.replace("@/", "");

            int finalSrcRowNo = Integer.parseInt(rowFromLast);

            finalSrcRowNo = rows + finalSrcRowNo;
            finalSrcRowNoStr = String.valueOf(finalSrcRowNo);

        } else {
            finalSrcRowNoStr = sourceRowPosition;
        }

        int idx = cellName.indexOf("@/");
        String colName = cellName.substring(0, idx);

        if (searchString != null && !searchString.trim().equals("")) {
            finalSrcRowNoStr = getRowNoForSearchString(finalSrcRowNoStr, colName);
           //  System.out.println("finalSrcRowNoStr " + finalSrcRowNoStr);
        }

        return colName + finalSrcRowNoStr;
    }

    private static String getRowNoForSearchString(String initialRow, String colName) {
       // System.out.println("initialRow "+initialRow + "colName "+colName);
        int lastRow = Integer.parseInt(initialRow);
        lastRow = lastRow - 1; //for 0 based index consideration in poi
        for (; lastRow >= 0; lastRow--) {
            //System.out.println("lastRow " + lastRow);
            Row drow = ssheet.getRow(lastRow);

            if(drow!=null) {
               // int totalCellsInRow = drow.getPhysicalNumberOfCells();


                String sourceCellRowCol = getColumnNumberFromCellAddress(colName + String.valueOf(lastRow));

                String srcRowNo = sourceCellRowCol.split(",")[0];
                String srcColNo = sourceCellRowCol.split(",")[1];

                Cell cell = drow.getCell(Integer.parseInt(srcColNo));

                //   System.out.println("totalCellsInRow " + totalCellsInRow);

                //for (int c = 0; c < totalCellsInRow; c++) {

                    //HSSFCell cell = drow.getCell(colName);

                    String cellValue = getCellValue(cell);

                    //     System.out.println("cellValue " + cellValue);

                    if (cellValue != null && !cellValue.trim().equals(""))
                        return String.valueOf(drow.getRowNum() + 1);
             //   }
            }
        }
        return initialRow;
    }

    private static String getColumnNumberFromCellAddress(String address) {
        CellAddress cellAddress = new CellAddress(address);
        int columnNo = cellAddress.getColumn();
        int rowNo = cellAddress.getRow();
        String addressInRowColIdx = String.valueOf(rowNo) + "," + String.valueOf(columnNo);
        return addressInRowColIdx;
    }
}
