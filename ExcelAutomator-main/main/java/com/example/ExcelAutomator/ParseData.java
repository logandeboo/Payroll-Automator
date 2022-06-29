package com.example.ExcelAutomator;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.IOException;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.time.LocalDate;
import java.util.ArrayList;
import java.util.HashSet;
import java.util.Set;
import java.util.logging.Logger;

public class ParseData {

    int rowCount;
    int columnCount;

    String cur_emp, prev_emp;
    String cur_POR, prev_POR;
    XSSFSheet curSheet;
    Set<String> sheetList;

    int goldRateFinal, platRateFinal;

    ArrayList<XSSFWorkbook> workbooks = new ArrayList<>();
    private static final Logger LOGGER = Logger.getLogger(SQLUtils.class.getName());
    public ArrayList<XSSFWorkbook> parseMaster(ResultSet qResult,LocalDate start, LocalDate end, int goldRate, int platRate) throws SQLException, IOException {

        // set gold & platinum rate
        goldRateFinal = goldRate;
        platRateFinal = platRate;

        sheetList = new HashSet<>();
        XSSFWorkbook workBookSBSH = new XSSFWorkbook();
        XSSFWorkbook workBookSBSM = new XSSFWorkbook();

        // (SBSH)
        // 1. Write data to master sheet (all employees)
        // 2. Reset qResult pointer to original position
        // 3. Write data to distinct employee sheets

        writeMaster(qResult,workBookSBSH,start,end,"SBSH");
        qResult.beforeFirst(); // reset result set pointer to original position
        writeEmployeeSheets(qResult,workBookSBSH,start,end,"SBSH");
        qResult.beforeFirst(); // reset result set pointer to original position

        // (SBSH)
        // 1. Write data to master sheet (all employees)
        // 2. Reset qResult pointer to original position
        // 3. Write data to distinct employee sheets

        writeMaster(qResult,workBookSBSM,start,end, "SBSM");
        qResult.beforeFirst(); // reset pointer to original position
        writeEmployeeSheets(qResult,workBookSBSM,start,end,"SBSM");

        workbooks.add(workBookSBSH);
        workbooks.add(workBookSBSM);


        return workbooks;
    }

    private void writeMaster(ResultSet rs, XSSFWorkbook workbook, LocalDate start, LocalDate end, String state) throws SQLException {

        //write column headers
        workbook.createSheet("Master");
        writeHeaderHelper(workbook.getSheetAt(0));

        while (rs.next()){


            if (rs.getString("Metal") == null){
                continue;
            }
            if (rs.getString("Metal").equals("SILVER")){
                continue;
            }
            //Filter employee code (SBSH or SBHSM)
            if (!rs.getString("Employee_Code").startsWith(state)){
                continue;
            }



            // filter dates from query results
            String temp = rs.getString("Return_Date");
            LocalDate tempDate = LocalDate.parse(temp);
            if (tempDate.isAfter(end) || tempDate.isBefore(start)){
                continue;
            }
            //write row data
            writeRowHelper(workbook, workbook.getSheetAt(0),rs,false);
        }
    }





    private void writeEmployeeSheets(ResultSet rs, XSSFWorkbook workbook, LocalDate start, LocalDate end, String state) throws SQLException {

        while (rs.next()){
            cur_emp = rs.getString("Employee_Code");
            cur_POR = rs.getString("POR");

            if (rs.getString("Metal") == null){
                continue;
            }
            if (rs.getString("Metal").equals("SILVER")){
                continue;
            }

            //Filter employee code (SBSH or SBHSM)
            if (!cur_emp.startsWith(state)){
                continue;
            }
            // filter dates from query results
            String temp = rs.getString("Return_Date");
            LocalDate tempDate = LocalDate.parse(temp);
            if (tempDate.isAfter(end) || tempDate.isBefore(start)){
                continue;
            }

            // if current employee matches previous employee --> keep current sheet & write new row
            // if current employee does not match previous employee --> create new sheet --> write new header --> write new row
            // sheetList maintains a unique set of all sheet names and is used to check if a sheet already exists for an employee
            // if employee sheet already exists --> add data to existing sheet ELSE create new sheet and write data
            if (cur_emp.equals(prev_emp)){
                writeRowHelper(workbook,curSheet,rs, getJobFlag());
            } else {
                if (sheetList.contains(cur_emp)){
                    curSheet = workbook.getSheet(cur_emp);
                    writeRowHelper(workbook,curSheet,rs,true);
                } else {
                    curSheet = workbook.createSheet(cur_emp);
                    sheetList.add(cur_emp);
                    writeHeaderHelper(curSheet);
                    writeRowHelper(workbook,curSheet,rs,false);
                }
            }
            prev_emp = cur_emp;
            prev_POR = cur_POR;
        }
    }

    public Boolean getJobFlag(){
        return !prev_POR.equals(cur_POR);
    }

    public void writeRowHelper(XSSFWorkbook workbook, XSSFSheet sheet, ResultSet rs, Boolean jobFlag) throws SQLException {
        CellStyle dateFormat = workbook.createCellStyle();
        CreationHelper creationHelper = workbook.getCreationHelper();
        dateFormat.setDataFormat(creationHelper.createDataFormat().getFormat("dd/MM/yyyy"));

        String empCode = rs.getString("Employee_Code");
        String metalType = rs.getString("Metal");
        String POR = rs.getString("POR");
        String transType = rs.getString("Trans_Type");
        String qualityCode = rs.getString("Quality_Code");
        String itemNo = rs.getString("Item_No");
        String itemType = calculate_item_type(rs);
        String issueDate = rs.getString("Issue_Date");
        String returnDate = rs.getString("Return_Date");

        String settingType = rs.getString("Setting_Type");
        String BRC = rs.getString("BRC");
        int metalRate = get_rate(rs, goldRateFinal, platRateFinal);
        int jobQty = rs.getInt("QTY");
        int issuedPieces = rs.getInt("Issued_Pcs");
        int thb = calculate_thb(rs,metalRate,issuedPieces);
        double issuedWeight = rs.getDouble("Issue_Weight");
        double grossLoss = calculate_gross_loss(rs);
        double returnedWeight = issuedWeight - grossLoss;
        double allowedLoss = calculate_allowed_loss(rs,(int)returnedWeight);
        double netLoss = grossLoss - allowedLoss;
        double metalLossValue = metalRate*netLoss*rs.getInt("Purity");
        double Final = thb - metalLossValue;

        if (jobFlag){
            rowCount = sheet.getLastRowNum() + 2;
        } else {
            rowCount = sheet.getLastRowNum() + 1;
        }
        Row row = sheet.createRow(rowCount++);
        columnCount = 0;

        Cell cell = row.createCell(columnCount++);
        cell.setCellValue(empCode);

        Cell cell1 = row.createCell(columnCount++);
        cell1.setCellValue(metalType);

        Cell cell2 = row.createCell(columnCount++);
        cell2.setCellValue(POR);

        Cell cell3 = row.createCell(columnCount++);
        cell3.setCellValue(transType);

        Cell cell4 = row.createCell(columnCount++);
        cell4.setCellValue(qualityCode);

        Cell cell5 = row.createCell(columnCount++);
        cell5.setCellValue(itemNo);

        Cell cell6 = row.createCell(columnCount++);
        cell6.setCellValue(itemType);

        Cell cell7 = row.createCell(columnCount++);
        cell7.setCellValue(issueDate);
        cell7.setCellStyle(dateFormat);

        Cell cell8 = row.createCell(columnCount++);
        cell8.setCellValue(returnDate);
        cell8.setCellStyle(dateFormat);

        Cell cell9 = row.createCell(columnCount++);
        cell9.setCellValue(settingType);

        Cell cell10 = row.createCell(columnCount++);
        cell10.setCellValue(BRC);

        Cell cell11 = row.createCell(columnCount++);
        cell11.setCellValue(metalRate);

        Cell cell12 = row.createCell(columnCount++);
        cell12.setCellValue(jobQty);

        Cell cell13 = row.createCell(columnCount++);
        cell13.setCellValue(issuedPieces);

        Cell cell14 = row.createCell(columnCount++);
        cell14.setCellValue(thb);

        Cell cell15 = row.createCell(columnCount++);
        cell15.setCellValue(issuedWeight);

        Cell cell16 = row.createCell(columnCount++);
        cell16.setCellValue(returnedWeight);

        Cell cell17 = row.createCell(columnCount++);
        cell17.setCellValue(grossLoss);

        Cell cell18 = row.createCell(columnCount++);
        cell18.setCellValue(allowedLoss);

        Cell cell19 = row.createCell(columnCount++);
        cell19.setCellValue(netLoss);

        Cell cell20 = row.createCell(columnCount++);
        cell20.setCellValue(metalLossValue);

        Cell cell21 = row.createCell(columnCount);
        cell21.setCellValue(Final);
    }


    private double calculate_allowed_loss(ResultSet rs, int returnWeight) throws SQLException {
        if (rs.getInt("Gross_Loss") == 0){
            return 0;
        } else if (rs.getString("Trans_Type").equals("REPAIR")){
            return 0;
        } else if (rs.getString("Metal").equals("PLATINUM")){
            return returnWeight * 0.02;
        } else {
            return returnWeight * 0.015;
        }

    }

    private void writeHeaderHelper(XSSFSheet sheet){

        XSSFCellStyle style = sheet.getWorkbook().createCellStyle();
        XSSFFont bold = sheet.getWorkbook().createFont();
        bold.setBold(true);
        style.setFont(bold);
        style.setBorderTop(BorderStyle.valueOf((short) 1));
        style.setBorderBottom(BorderStyle.valueOf((short) 1));

        Row headerRow = sheet.createRow(0);

        Cell headerCell = headerRow.createCell(0);
        headerCell.setCellValue("Employee Code");
        headerCell.setCellStyle(style);

        Cell headerCell1 = headerRow.createCell(1);
        headerCell1.setCellValue("Metal Type");
        headerCell1.setCellStyle(style);

        Cell headerCell2 = headerRow.createCell(2);
        headerCell2.setCellValue("POR");
        headerCell2.setCellStyle(style);

        Cell headerCell3 = headerRow.createCell(3);
        headerCell3.setCellValue("Trans Type");
        headerCell3.setCellStyle(style);

        Cell headerCell4 = headerRow.createCell(4);
        headerCell4.setCellValue("Quality Code");
        headerCell4.setCellStyle(style);

        Cell headerCell5 = headerRow.createCell(5);
        headerCell5.setCellValue("Item No.");
        headerCell5.setCellStyle(style);

        Cell headerCell6 = headerRow.createCell(6);
        headerCell6.setCellValue("Item Type");
        headerCell6.setCellStyle(style);

        Cell headerCell7 = headerRow.createCell(7);
        headerCell7.setCellValue("Issue Date");
        headerCell7.setCellStyle(style);

        Cell headerCell8 = headerRow.createCell(8);
        headerCell8.setCellValue("Return Date");
        headerCell8.setCellStyle(style);

        Cell headerCell9 = headerRow.createCell(9);
        headerCell9.setCellValue("Type of Setting");
        headerCell9.setCellStyle(style);

        Cell headerCell10 = headerRow.createCell(10);
        headerCell10.setCellValue("BRC");
        headerCell10.setCellStyle(style);

        Cell headerCell11 = headerRow.createCell(11);
        headerCell11.setCellValue("Rate");
        headerCell11.setCellStyle(style);

        Cell headerCell12 = headerRow.createCell(12);
        headerCell12.setCellValue("(Job) QTY");
        headerCell12.setCellStyle(style);

        Cell headerCell13 = headerRow.createCell(13);
        headerCell13.setCellValue("Issued Pieces");
        headerCell13.setCellStyle(style);

        Cell headerCell14 = headerRow.createCell(14);
        headerCell14.setCellValue("THB");
        headerCell14.setCellStyle(style);

        Cell headerCell15 = headerRow.createCell(15);
        headerCell15.setCellValue("Issue Weight");
        headerCell15.setCellStyle(style);

        Cell headerCell16 = headerRow.createCell(16);
        headerCell16.setCellValue("Returned Weight");
        headerCell16.setCellStyle(style);

        Cell headerCell17 = headerRow.createCell(17);
        headerCell17.setCellValue("Gross Loss");
        headerCell17.setCellStyle(style);

        Cell headerCell18 = headerRow.createCell(18);
        headerCell18.setCellValue("Allowed Loss");
        headerCell18.setCellStyle(style);

        Cell headerCell19 = headerRow.createCell(19);
        headerCell19.setCellValue("Net Loss");
        headerCell19.setCellStyle(style);

        Cell headerCell20 = headerRow.createCell(20);
        headerCell20.setCellValue("Metal Loss Value");
        headerCell20.setCellStyle(style);

        Cell headerCell21 = headerRow.createCell(21);
        headerCell21.setCellValue("Final");
        headerCell21.setCellStyle(style);

        Cell headerCell22 = headerRow.createCell(22);
        headerCell22.setCellValue("Broken-Missing");
        headerCell22.setCellStyle(style);

        Cell headerCell23 = headerRow.createCell(23);
        headerCell23.setCellValue("Next Payment");
        headerCell23.setCellStyle(style);


        sheet.setColumnWidth(0, 15*256);
        sheet.setColumnWidth(1, 12*256);
        sheet.setColumnWidth(2, 20*256);
        sheet.setColumnWidth(3, 13*256);
        sheet.setColumnWidth(4, 12*256);
        sheet.setColumnWidth(5, 13*256);
        sheet.setColumnWidth(6, 11*256);
        sheet.setColumnWidth(7, 12*256);
        sheet.setColumnWidth(8, 12*256);
        sheet.setColumnWidth(9, 26*256);
        sheet.setColumnWidth(10, 10*256);
        sheet.setColumnWidth(11, 10*256);
        sheet.setColumnWidth(12, 10*256);
        sheet.setColumnWidth(13, 13*256);
        sheet.setColumnWidth(14, 10*256);
        sheet.setColumnWidth(15, 12*256);
        sheet.setColumnWidth(16, 16*256);
        sheet.setColumnWidth(17, 10*256);
        sheet.setColumnWidth(18, 13*256);
        sheet.setColumnWidth(20, 16*256);
        sheet.setColumnWidth(21, 16*256);
        sheet.setColumnWidth(22, 20*256);
        headerRow.setHeight((short)650);
    }


    public String calculate_item_type(ResultSet rs) throws SQLException {
        if (rs.getString("Stone_Type").equals("0")){
            return "Alloy";
        } else {
            return rs.getString("Stone_Type");
        }
    }

    public int get_rate(ResultSet rs, int goldRate, int platRate) throws SQLException {

        if (rs.getString("Metal") != null) {

            if (rs.getString("Metal").equals("PLATINUM")) {
                return platRate;
            }
            if (rs.getString("Metal").equals("GOLD")) {
                return goldRate;
            }
        }
        System.out.println(rs.getRow() + rs.getString("Employee_Code") + rs.getString("POR"));
        return 0;
    }

    public int calculate_thb(ResultSet rs, int metalRate, int issuedPcs) throws SQLException {
        if (rs.getString("Trans_Type").equals("SET")){
            return metalRate * issuedPcs;
        } else {
            return rs.getInt("THB");
        }
    }



    public double calculate_gross_loss(ResultSet rs) throws SQLException {
        if (rs.getString("Trans_Type").equals("REPAIR") && rs.getInt("Gross_Loss") < 0){
            return 0;
        } else {
            return rs.getDouble("Gross_Loss");
        }
    }
}
