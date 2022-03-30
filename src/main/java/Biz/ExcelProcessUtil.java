package Biz;

import VO.ExportItem;
import VO.Result;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.yaml.snakeyaml.Yaml;

import java.io.*;
import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Map;

public class ExcelProcessUtil {
    public static String getSourceFilePath() throws FileNotFoundException {
        Yaml yml = new Yaml();
        HashMap configuration = yml.load(new FileInputStream("src/main/java/Config/configuration.yml"));
        String path = configuration.get("sourceFilePath").toString();
        return path;
    }

    public static ArrayList<ExportItem> getDataFromDrillDown(String excelPath, int columnLength) throws IOException {
        ArrayList<ExportItem> itemList = new ArrayList<>();
        XSSFWorkbook workbook = new XSSFWorkbook(ExcelProcessUtil.getSourceFilePath() + "source1.xlsx");
        XSSFSheet sheet = workbook.getSheetAt(0);
        int maxRow = sheet.getLastRowNum();
        for(int i = 0 ; i < (maxRow+1) / columnLength ; i++){
            ExportItem item = new ExportItem();
            item.date = sheet.getRow(i * 6).getCell(0).toString();
            item.referenceID = sheet.getRow(i * columnLength + 1).getCell(0).toString();
            item.transactionType = sheet.getRow(i * columnLength + 2).getCell(0).toString();
            item.description = sheet.getRow(i * columnLength + 3).getCell(0).toString();
            item.debit = new BigDecimal(sheet.getRow(i * columnLength + 4).getCell(0).toString());
            item.credit = new BigDecimal(sheet.getRow(i * columnLength + 5).getCell(0).toString());
            itemList.add(item);
        }
        return itemList;
    }

    public static void exportByFormatForDrillDown(String excelPath) throws IOException {
        ArrayList<ExportItem> itemList = getDataFromDrillDown(excelPath, 6);
        XSSFWorkbook workbook=new XSSFWorkbook();
        XSSFSheet sheet=workbook.createSheet("Report");
        XSSFRow row=sheet.createRow(0);
        row.createCell(0).setCellValue("Date");
        row.createCell(1).setCellValue("Reference Number");
        row.createCell(2).setCellValue("Transaction Type");
        row.createCell(3).setCellValue("Description");
        row.createCell(4).setCellValue("Debit");
        row.createCell(5).setCellValue("Credit");
        for(int j = 1 ; j <= itemList.size() ; j++){
            XSSFRow row_j=sheet.createRow(j);
            row_j.createCell(0).setCellValue(itemList.get(j-1).date);
            row_j.createCell(1).setCellValue(itemList.get(j-1).referenceID);
            row_j.createCell(2).setCellValue(itemList.get(j-1).transactionType);
            row_j.createCell(3).setCellValue(itemList.get(j-1).description);
            row_j.createCell(4).setCellValue(itemList.get(j-1).debit.toString());
            row_j.createCell(5).setCellValue(itemList.get(j-1).credit.toString());
        }
        FileOutputStream out=new FileOutputStream(getSourceFilePath() + "Report_"+ System.currentTimeMillis() + ".xlsx");
        workbook.write(out);
    }

    public static Map<String, ArrayList<Result>> changeToMapFromExcel(String filePath) throws IOException {
        Map<String, ArrayList<Result>> map = new HashMap<>();
        XSSFWorkbook workbook = new XSSFWorkbook(ExcelProcessUtil.getSourceFilePath() + "result.xlsx");
        XSSFSheet sheet = workbook.getSheetAt(0);
        int maxRow = sheet.getLastRowNum();
        for(int i = 1 ; i <= maxRow ; i++){
            String referenceID = sheet.getRow(i).getCell(1).toString();
            if(map.containsKey(referenceID)){
                Result item = new Result();
                item.referenceID = referenceID;
                item.credit = new BigDecimal(sheet.getRow(i).getCell(5).toString());
                item.debit = new BigDecimal(sheet.getRow(i).getCell(4).toString());
                map.get(referenceID).add(item);
            }else{
                ArrayList<Result> list = new ArrayList<>();
                Result item = new Result();
                item.referenceID = referenceID;
                item.credit = new BigDecimal(sheet.getRow(i).getCell(5).toString());
                item.debit = new BigDecimal(sheet.getRow(i).getCell(4).toString());
                list.add(item);
                map.put(referenceID, list);
            }
        }
        return map;
    }

    public static void generateResultToExcel(Map<String, ArrayList<Result>> map) throws IOException {
        XSSFWorkbook workbook=new XSSFWorkbook();
        XSSFSheet sheet=workbook.createSheet("Report");
        XSSFRow row=sheet.createRow(0);
        row.createCell(0).setCellValue("Reference Number");
        row.createCell(1).setCellValue("Balance");
        row.createCell(2).setCellValue("IsBalanced");
        int i = 1;
        for(Map.Entry<String, ArrayList<Result>> entry : map.entrySet()){
            XSSFRow row_j=sheet.createRow(i);
            i++;
            String referenceID = entry.getKey();
            ArrayList<Result> list = entry.getValue();
            BigDecimal totalCredit = new BigDecimal(0);
            BigDecimal totalDebit = new BigDecimal(0);
            for(int j = 0; j < list.size() ; j++){
                Result item = list.get(j);
                totalCredit = totalCredit.add(item.credit);
                totalDebit = totalDebit.add(item.debit);
            }
            BigDecimal balance = totalCredit.subtract(totalDebit);
            row_j.createCell(0).setCellValue(referenceID);
            row_j.createCell(1).setCellValue(balance.toString());
            if(balance.compareTo(new BigDecimal(0)) == 0){
                row_j.createCell(2).setCellValue(true);
            }else{
                row_j.createCell(2).setCellValue(false);
            }
        }

        FileOutputStream out=new FileOutputStream(getSourceFilePath() + "Report_"+ System.currentTimeMillis() + ".xlsx");
        workbook.write(out);
    }
}
