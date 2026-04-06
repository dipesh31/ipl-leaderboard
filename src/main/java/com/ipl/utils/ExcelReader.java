package com.ipl.utils;

import org.apache.poi.ss.usermodel.*;


import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class ExcelReader {
    private String path;

    public ExcelReader(String path) {
        this.path = path;
    }

    public Object[][] getSheetData(String sheetName) throws Exception {
        FileInputStream fis = new FileInputStream(path);
        Workbook workbook = WorkbookFactory.create(fis);
        Sheet sheet = workbook.getSheet(sheetName);
        DataFormatter formatter = new DataFormatter();

        int rowCount = sheet.getLastRowNum();
        int colCount = sheet.getRow(0).getLastCellNum();
        Object[][] data = new Object[rowCount][colCount];

        for (int i = 0; i < rowCount; i++) {
            Row row = sheet.getRow(i + 1);
            for (int j = 0; j < colCount; j++) {
                data[i][j] = formatter.formatCellValue(row.getCell(j));
            }
        }
        fis.close();
        return data;
    }

    /**
     * Writes ONLY player points to Sheet1 (Column E)
     * Applies C/VC multiplier
     */
    public void writePoints(String sheetName, String playerName, String rawPoints) throws Exception {
        FileInputStream fis = new FileInputStream(path);
        Workbook workbook = WorkbookFactory.create(fis);
        Sheet sheet = workbook.getSheet(sheetName);
        DataFormatter formatter = new DataFormatter();

        for (int i = 1; i <= sheet.getLastRowNum(); i++) { // start from row 1 (skip header)
            Row row = sheet.getRow(i);
            if (row == null) continue;

            String nameInSheet = formatter.formatCellValue(row.getCell(0)).trim();

            if (nameInSheet.equalsIgnoreCase(playerName.trim())) {

                double p = 0.0;
                try {
                    p = Double.parseDouble(rawPoints.replaceAll("[^0-9.-]", ""));
                } catch (Exception e) {
                    p = 0.0;
                }

                String role = formatter.formatCellValue(row.getCell(3)).trim(); // Column D

                // Apply multiplier
                if (role.equalsIgnoreCase("C")) {
                    p = p * 2.0;
                } else if (role.equalsIgnoreCase("VC")) {
                    p = p * 1.5;
                }

                Cell pointsCell = row.getCell(4); // Column E
                if (pointsCell == null) {
                    pointsCell = row.createCell(4);
                }

                pointsCell.setCellValue(p); // ✅ ONLY player points written
                break;
            }
        }

        fis.close();
        FileOutputStream fos = new FileOutputStream(path);
        workbook.write(fos);
        fos.close();
    }

    /**
     * NEW METHOD:
     * Calculates totals from Sheet1 and writes ONLY to Sheet2
     */
    public void updateAllTeamTotals(String sourceSheet, String targetSheet) throws Exception {

        FileInputStream fis = new FileInputStream(path);
        Workbook workbook = WorkbookFactory.create(fis);
        Sheet s1 = workbook.getSheet(sourceSheet);
        Sheet s2 = workbook.getSheet(targetSheet);
        DataFormatter formatter = new DataFormatter();

        Map<String, Double> teamTotals = new HashMap<>();

        // Step 1: Calculate totals from Sheet1 (UNCHANGED)
        for (int i = 1; i <= s1.getLastRowNum(); i++) {
            Row row = s1.getRow(i);
            if (row == null) continue;

            String team = formatter.formatCellValue(row.getCell(2)).trim();
            String pointsStr = formatter.formatCellValue(row.getCell(4)).trim();

            if (!team.isEmpty() && !pointsStr.isEmpty()) {
                try {
                    double points = Double.parseDouble(pointsStr);
                    teamTotals.put(team, teamTotals.getOrDefault(team, 0.0) + points);
                } catch (Exception ignored) {}
            }
        }

        // 🔥 NEW: Sort teams by points DESC
        List<Map.Entry<String, Double>> sortedList = new ArrayList<>(teamTotals.entrySet());
        sortedList.sort((a, b) -> Double.compare(b.getValue(), a.getValue()));

        // Step 2: Clear Sheet2 (UNCHANGED)
        int lastRow = s2.getLastRowNum();
        for (int i = lastRow; i > 0; i--) {
            Row row = s2.getRow(i);
            if (row != null) {
                s2.removeRow(row);
            }
        }

        // 🔥 NEW: Add Rank in header
        Row header = s2.getRow(0);
        if (header == null) header = s2.createRow(0);

        header.createCell(0).setCellValue("Rank");
        header.createCell(1).setCellValue("Team");
        header.createCell(2).setCellValue("Total");

        // 🔥 NEW: Write sorted data with rank
        int rowIndex = 1;
        int rank = 1;

        for (Map.Entry<String, Double> entry : sortedList) {
            Row row = s2.createRow(rowIndex++);

            row.createCell(0).setCellValue(rank++);           // Rank
            row.createCell(1).setCellValue(entry.getKey());   // Team
            row.createCell(2).setCellValue(entry.getValue()); // Points
        }

        fis.close();
        FileOutputStream fos = new FileOutputStream(path);
        workbook.write(fos);
        fos.close();
    }
    
    public void clearPointsColumn(String sheetName) throws Exception {
	    FileInputStream fis = new FileInputStream(path);
	    Workbook workbook = WorkbookFactory.create(fis);
	    Sheet sheet = workbook.getSheet(sheetName);

	    for (int i = 1; i <= sheet.getLastRowNum(); i++) {
	        Row row = sheet.getRow(i);
	        if (row != null) {
	            Cell cell = row.getCell(4); // Column E
	            if (cell != null) {
	                cell.setBlank(); // 🔥 clears old values
	            }
	        }
	    }

	    fis.close();
	    FileOutputStream fos = new FileOutputStream(path);
	    workbook.write(fos);
	    fos.close();
	}
}