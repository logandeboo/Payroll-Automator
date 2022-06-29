package com.example.ExcelAutomator;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;

import java.awt.*;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.LocalDate;
import java.util.ArrayList;
import java.util.logging.Level;
import java.util.logging.Logger;

public class ExportSubIncentive {
    String filePathSBSH;
    String filePathSBSM;
    byte[] headerBlue = {91,(byte)155,(byte)213};
    byte[] subTotalPeach = {(byte)252,(byte)228,(byte)214};
    byte[] empTotalBlue = {(byte)155,(byte)194,(byte)230};
    byte[] empCodeBlue = {(byte)221,(byte)235,(byte)247};

    FileOutputStream outputStreamSH;
    FileOutputStream outputStreamSM;
    private static final Logger LOGGER = Logger.getLogger(SQLUtils.class.getName());


    public ArrayList<XSSFWorkbook> formatSubInc(ArrayList<XSSFWorkbook> workbooks){

        for (XSSFWorkbook book: workbooks){
            addFinalSubTotal(book);
            addSummaryRows(book);
            format_header(book);
            format_job_subtotals(book);
            format_metal_column(book);
            format_emp_column(book);
            subTotalSums(book);
            grandTotals(book);
            freeze_edges_helper(book);
        }
        return workbooks;
    }

    public void exportWorkbooks(ArrayList<XSSFWorkbook> workbooks, String path){


        filePathSBSH = path+ "/SubContractSetting-SBSH-" + LocalDate.now() + ".xlsx";
        filePathSBSM = path+ "/SubContractSetting-SBSM-" + LocalDate.now() + ".xlsx";


        try {
            outputStreamSH = new FileOutputStream(filePathSBSH);
            outputStreamSM = new FileOutputStream(filePathSBSM);

            workbooks.get(0).write(outputStreamSH);
            workbooks.get(1).write(outputStreamSM);

        } catch (IOException e) {
            LOGGER.log(Level.SEVERE, e.toString(), e);
            System.out.println(e);

        } finally {
            try{
                if (outputStreamSH != null && outputStreamSM != null){
                    outputStreamSH.close();
                    outputStreamSM.close();
                }
            } catch (IOException e) {
                LOGGER.log(Level.SEVERE, e.toString(), e);
                System.out.println(e);

            } finally {
                openFile(filePathSBSH);
                openFile(filePathSBSM);
            }

        }


    }

    public void addFinalSubTotal(XSSFWorkbook workbook){

        XSSFCellStyle peachSubTotal = workbook.createCellStyle();
        XSSFFont makeBold = workbook.createFont();
        makeBold.setBold(true);
        peachSubTotal.setFont(makeBold);
        peachSubTotal.setFillForegroundColor(new XSSFColor(subTotalPeach, new DefaultIndexedColorMap()));
        peachSubTotal.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        peachSubTotal.setBorderTop(BorderStyle.valueOf((short) 1));
        peachSubTotal.setBorderBottom(BorderStyle.valueOf((short) 1));

        for (int sheetNum = 1; sheetNum < workbook.getNumberOfSheets(); sheetNum++){
            XSSFSheet sheet = workbook.getSheetAt(sheetNum);
            Row myRow = sheet.createRow(sheet.getLastRowNum()+1);
            myRow.createCell(0).setCellValue(sheet.getRow(myRow.getRowNum()-1).getCell(0).getStringCellValue());
            myRow.createCell(1).setCellValue(sheet.getRow(myRow.getRowNum()-1).getCell(1).getStringCellValue());
            myRow.createCell(2).setCellValue(sheet.getRow(myRow.getRowNum()-1).getCell(2).getStringCellValue() + " Total");

            for (int cellNum = 2; cellNum < 24; cellNum++){
                if (cellNum > 2){
                    myRow.createCell(cellNum);
                }
                myRow.getCell(cellNum).setCellStyle(peachSubTotal);
            }
        }
    }
    public void addSummaryRows(XSSFWorkbook workbook){

        XSSFCellStyle empTotal = workbook.createCellStyle();
        XSSFFont makeBold = workbook.createFont();
        makeBold.setBold(true);
        empTotal.setAlignment(HorizontalAlignment.RIGHT);
        empTotal.setFont(makeBold);
        empTotal.setFillForegroundColor(new XSSFColor(empTotalBlue, new DefaultIndexedColorMap()));
        empTotal.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        empTotal.setBorderTop(BorderStyle.valueOf((short) 1));
        empTotal.setBorderBottom(BorderStyle.valueOf((short) 1));

        XSSFCellStyle boldFinalRow = workbook.createCellStyle();
        XSSFFont makeFinalBold = workbook.createFont();
        makeFinalBold.setBold(true);
        boldFinalRow.setAlignment(HorizontalAlignment.RIGHT);
        boldFinalRow.setFont(makeBold);
        boldFinalRow.setBorderTop(BorderStyle.valueOf((short) 1));
        boldFinalRow.setBorderBottom(BorderStyle.valueOf((short) 1));

        for (int sheetNum = 1; sheetNum < workbook.getNumberOfSheets(); sheetNum++){
            XSSFSheet sheet = workbook.getSheetAt(sheetNum);
            Row myRow = sheet.createRow(sheet.getLastRowNum()+1);
            myRow.createCell(0).setCellValue(sheet.getRow(myRow.getRowNum()-1).getCell(0).getStringCellValue() + " Total");
            for (int cellNum = 0; cellNum < 24; cellNum++){
                if (cellNum > 0){
                    myRow.createCell(cellNum);
                }
                myRow.getCell(cellNum).setCellStyle(empTotal);
            }
        }

        for (int sheetNum = 1; sheetNum < workbook.getNumberOfSheets(); sheetNum++){
            XSSFSheet sheet = workbook.getSheetAt(sheetNum);
            Row myRow = sheet.createRow(sheet.getLastRowNum()+1);
            myRow.createCell(0).setCellValue("Grand Total");

            for (int cellNum = 0; cellNum < 24; cellNum++){
                if (cellNum > 0){
                    myRow.createCell(cellNum);
                }
                myRow.getCell(cellNum).setCellStyle(boldFinalRow);
            }

        }
    }
    public void format_header(XSSFWorkbook workbook){

        XSSFCellStyle style = workbook.createCellStyle();
        XSSFFont makeWhite = workbook.createFont();
        makeWhite.setBold(true);
        makeWhite.setColor(IndexedColors.WHITE.index);
        style.setFont(makeWhite);
        style.setFillForegroundColor(new XSSFColor(headerBlue, new DefaultIndexedColorMap()));
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        for (int i = 0; i < workbook.getNumberOfSheets(); i++){
            XSSFSheet sheet = workbook.getSheetAt(i);
            Row myRow = sheet.getRow(0);

            for (int j = 0; j < 24; j++){
                myRow.getCell(j).setCellStyle(style);
            }
        }
    }

    public void format_job_subtotals(XSSFWorkbook workbook){

        XSSFCellStyle peachSubTotal = workbook.createCellStyle();
        XSSFFont makeBold = workbook.createFont();
        makeBold.setBold(true);
        peachSubTotal.setAlignment(HorizontalAlignment.RIGHT);
        peachSubTotal.setFont(makeBold);
        peachSubTotal.setFillForegroundColor(new XSSFColor(subTotalPeach, new DefaultIndexedColorMap()));
        peachSubTotal.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        peachSubTotal.setBorderTop(BorderStyle.valueOf((short) 1));
        peachSubTotal.setBorderBottom(BorderStyle.valueOf((short) 1));

        for (int i = 1; i < workbook.getNumberOfSheets(); i++){
            XSSFSheet sheet = workbook.getSheetAt(i);
            for (int rowNum = 1; rowNum < sheet.getLastRowNum()+1; rowNum++){
                if (sheet.getRow(rowNum) == null){
                    Row myRow = sheet.createRow(rowNum);
                    myRow.createCell(0).setCellValue(sheet.getRow(rowNum-1).getCell(0).getStringCellValue());
                    myRow.createCell(1).setCellValue(sheet.getRow(rowNum-1).getCell(1).getStringCellValue());
                    myRow.createCell(2).setCellValue(sheet.getRow(rowNum-1).getCell(2).getStringCellValue() + " Total");

                    for (int cellNum = 2; cellNum < 24; cellNum++){
                        if (cellNum > 2){
                            myRow.createCell(cellNum);
                        }
                        myRow.getCell(cellNum).setCellStyle(peachSubTotal);
                    }
                }
            }
        }
    }
    public void format_metal_column(XSSFWorkbook workbook){

        XSSFCellStyle metalCol = workbook.createCellStyle();
        XSSFFont bold = workbook.createFont();
        bold.setBold(true);
        metalCol.setFont(bold);


        for (int sheetNum = 0; sheetNum < workbook.getNumberOfSheets(); sheetNum++){
            XSSFSheet temp = workbook.getSheetAt(sheetNum);
            for (int rowNum = 1; rowNum < temp.getPhysicalNumberOfRows(); rowNum++){
                Row row = temp.getRow(rowNum);
                row.getCell(1).setCellStyle(metalCol);
            }
        }
    }

    public void format_emp_column(XSSFWorkbook workbook){

        XSSFCellStyle empCodeColumn = workbook.createCellStyle();
        XSSFFont bold = workbook.createFont();
        bold.setBold(true);
        empCodeColumn.setFont(bold);
        empCodeColumn.setFillForegroundColor(new XSSFColor(empCodeBlue, new DefaultIndexedColorMap()));
        empCodeColumn.setFillPattern(FillPatternType.SOLID_FOREGROUND);


        for (int sheetNum = 0; sheetNum < workbook.getNumberOfSheets(); sheetNum++){
            XSSFSheet temp = workbook.getSheetAt(sheetNum);
            for (int rowNum = 1; rowNum < temp.getLastRowNum()+1; rowNum++){
                Row row = temp.getRow(rowNum);
                row.getCell(0).setCellStyle(empCodeColumn);
            }
        }
    }

    public void subTotalSums(XSSFWorkbook workbook){
        double jobQty = 0;
        double issuePieces = 0;
        double thb = 0;
        double issueWeight = 0;
        double returnedWeight = 0;
        double grossLoss = 0;
        double allowedLoss = 0;
        double netLoss = 0;
        double metalLossVal = 0;
        double fin = 0;

        for (int i = 1; i < workbook.getNumberOfSheets(); i++){
            XSSFSheet sheet = workbook.getSheetAt(i);
            for (int rowNum = 1; rowNum < sheet.getLastRowNum()+1; rowNum++) {
                Row temp = sheet.getRow(rowNum);
                if (!temp.getCell(2).getStringCellValue().contains("Total")){
                    jobQty += temp.getCell(12).getNumericCellValue();
                    issuePieces += temp.getCell(13).getNumericCellValue();
                    thb += temp.getCell(14).getNumericCellValue();
                    issueWeight += temp.getCell(15).getNumericCellValue();
                    returnedWeight += temp.getCell(16).getNumericCellValue();
                    grossLoss += temp.getCell(17).getNumericCellValue();
                    allowedLoss += temp.getCell(18).getNumericCellValue();
                    netLoss += temp.getCell(19).getNumericCellValue();
                    metalLossVal += temp.getCell(20).getNumericCellValue();
                    fin += temp.getCell(21).getNumericCellValue();
                } else if (temp.getCell(2).getStringCellValue().contains("Total")){
                    temp.getCell(12).setCellValue(jobQty);
                    jobQty = 0;
                    temp.getCell(13).setCellValue(issuePieces);
                    issuePieces = 0;
                    temp.getCell(14).setCellValue(thb);
                    thb = 0;
                    temp.getCell(15).setCellValue(issueWeight);
                    issueWeight = 0;
                    temp.getCell(16).setCellValue(returnedWeight);
                    returnedWeight = 0;
                    temp.getCell(17).setCellValue(grossLoss);
                    grossLoss = 0;
                    temp.getCell(18).setCellValue(allowedLoss);
                    allowedLoss = 0;
                    temp.getCell(19).setCellValue(netLoss);
                    netLoss = 0;
                    temp.getCell(20).setCellValue(metalLossVal);
                    metalLossVal = 0;
                    temp.getCell(21).setCellValue(fin);
                    fin = 0;
                }
            }
        }
    }

    public void grandTotals(XSSFWorkbook workbook){


        for (int sheetNum = 1; sheetNum < workbook.getNumberOfSheets(); sheetNum++){
            XSSFSheet sheet = workbook.getSheetAt(sheetNum);
            Row employeeTotal = sheet.getRow(sheet.getLastRowNum()-1);
            Row grandTotal = sheet.getRow(sheet.getLastRowNum());
            double jobQty = 0;
            double issuePieces = 0;
            double thb = 0;
            double issueWeight = 0;
            double returnedWeight = 0;
            double grossLoss = 0;
            double allowedLoss = 0;
            double netLoss = 0;
            double metalLossVal = 0;
            double fin = 0;

            for (int rowNum = 1; rowNum < sheet.getLastRowNum(); rowNum++){
                Row temp = sheet.getRow(rowNum);
                if (temp.getCell(2).getStringCellValue().contains("Total")){
                    jobQty += temp.getCell(12).getNumericCellValue();
                    issuePieces += temp.getCell(13).getNumericCellValue();
                    thb += temp.getCell(14).getNumericCellValue();
                    issueWeight += temp.getCell(15).getNumericCellValue();
                    returnedWeight += temp.getCell(16).getNumericCellValue();
                    grossLoss += temp.getCell(17).getNumericCellValue();
                    allowedLoss += temp.getCell(18).getNumericCellValue();
                    netLoss += temp.getCell(19).getNumericCellValue();
                    metalLossVal += temp.getCell(20).getNumericCellValue();
                    fin += temp.getCell(21).getNumericCellValue();
                }
            }


            employeeTotal.getCell(12).setCellValue(jobQty);
            grandTotal.getCell(12).setCellValue(jobQty);

            employeeTotal.getCell(13).setCellValue(issuePieces);
            grandTotal.getCell(13).setCellValue(issuePieces);

            employeeTotal.getCell(14).setCellValue("฿"+thb);
            grandTotal.getCell(14).setCellValue("฿"+thb);

            employeeTotal.getCell(15).setCellValue(issueWeight);
            grandTotal.getCell(15).setCellValue(issueWeight);

            employeeTotal.getCell(16).setCellValue(returnedWeight);
            grandTotal.getCell(16).setCellValue(returnedWeight);

            employeeTotal.getCell(17).setCellValue(grossLoss);
            grandTotal.getCell(17).setCellValue(grossLoss);

            employeeTotal.getCell(18).setCellValue(allowedLoss);
            grandTotal.getCell(18).setCellValue(allowedLoss);

            employeeTotal.getCell(19).setCellValue(netLoss);
            grandTotal.getCell(19).setCellValue(netLoss);

            employeeTotal.getCell(20).setCellValue("฿"+metalLossVal);
            grandTotal.getCell(20).setCellValue("฿"+metalLossVal);

            employeeTotal.getCell(21).setCellValue("฿"+fin);
            grandTotal.getCell(21).setCellValue("฿"+fin);
        }
    }



    public void freeze_edges_helper(XSSFWorkbook workbook) {

        for (int sheetNum = 0; sheetNum < workbook.getNumberOfSheets(); sheetNum++){
            XSSFSheet temp = workbook.getSheetAt(sheetNum);
            temp.createFreezePane(0,1);
        }
    }

    public void openFile(String path) {

        try {

            File file = new File(path);

            if (!Desktop.isDesktopSupported()){
                System.out.println("System not supported");
                return;
            }
            Desktop myComputer = Desktop.getDesktop();
            if (file.exists()){
                myComputer.open(file);
            }
        } catch (IOException e) {
            LOGGER.log(Level.SEVERE, e.toString(), e);
            System.out.println(e);
        }
    }



}
