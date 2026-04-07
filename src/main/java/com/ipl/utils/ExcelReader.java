package com.ipl.utils;

import org.apache.poi.ss.usermodel.*;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.*;

public class ExcelReader {
    private String path;

    // ✅ Column mapping (LOCKED)
    private static final int PLAYER_COL = 0;
    private static final int IPL_TEAM_COL = 1;
    private static final int FANTASY_TEAM_COL = 2;
    private static final int ROLE_COL = 3;
    private static final int POINTS_COL = 4;

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
     * ✅ FIXED: Safe write (prevents IPL Team overwrite)
     */
    public void writePoints(String sheetName, String playerName, String rawPoints) throws Exception {

        FileInputStream fis = new FileInputStream(path);
        Workbook workbook = WorkbookFactory.create(fis);
        Sheet sheet = workbook.getSheet(sheetName);
        DataFormatter formatter = new DataFormatter();

        for (int i = 1; i <= sheet.getLastRowNum(); i++) {
            Row row = sheet.getRow(i);
            if (row == null) continue;

            String nameInSheet = formatter.formatCellValue(row.getCell(PLAYER_COL)).trim();

            if (nameInSheet.equalsIgnoreCase(playerName.trim())) {

                // ✅ Ensure ALL cells exist (prevents row corruption)
                for (int c = 0; c <= POINTS_COL; c++) {
                    if (row.getCell(c) == null) {
                        row.createCell(c);
                    }
                }

                double p = 0.0;
                try {
                    p = Double.parseDouble(rawPoints.replaceAll("[^0-9.-]", ""));
                } catch (Exception e) {
                    p = 0.0;
                }

                String role = formatter.formatCellValue(row.getCell(ROLE_COL)).trim();

                // Apply multiplier
                if (role.equalsIgnoreCase("C")) {
                    p = p * 2.0;
                } else if (role.equalsIgnoreCase("VC")) {
                    p = p * 1.5;
                }

                // ✅ SAFELY overwrite ONLY points column
                Cell pointsCell = row.getCell(POINTS_COL);
                pointsCell.setBlank(); // remove any formula
                pointsCell.setCellValue(p);

                break;
            }
        }

        fis.close();
        FileOutputStream fos = new FileOutputStream(path);
        workbook.write(fos);
        fos.close();
    }

    /**
     * ✅ Unchanged logic + safe column usage
     */
    public void updateAllTeamTotals(String sourceSheet, String targetSheet) throws Exception {

        FileInputStream fis = new FileInputStream(path);
        Workbook workbook = WorkbookFactory.create(fis);
        Sheet s1 = workbook.getSheet(sourceSheet);
        Sheet s2 = workbook.getSheet(targetSheet);
        DataFormatter formatter = new DataFormatter();

        Map<String, Double> teamTotals = new HashMap<>();

        for (int i = 1; i <= s1.getLastRowNum(); i++) {
            Row row = s1.getRow(i);
            if (row == null) continue;

            String team = formatter.formatCellValue(row.getCell(FANTASY_TEAM_COL)).trim();
            String pointsStr = formatter.formatCellValue(row.getCell(POINTS_COL)).trim();

            if (!team.isEmpty() && !pointsStr.isEmpty()) {
                try {
                    double points = Double.parseDouble(pointsStr);
                    teamTotals.put(team, teamTotals.getOrDefault(team, 0.0) + points);
                } catch (Exception ignored) {}
            }
        }

        // Sorting
        List<Map.Entry<String, Double>> sortedList = new ArrayList<>(teamTotals.entrySet());
        sortedList.sort((a, b) -> Double.compare(b.getValue(), a.getValue()));

        // Clear Sheet2
        int lastRow = s2.getLastRowNum();
        for (int i = lastRow; i > 0; i--) {
            Row row = s2.getRow(i);
            if (row != null) s2.removeRow(row);
        }

        // Header
        Row header = s2.getRow(0);
        if (header == null) header = s2.createRow(0);

        header.createCell(0).setCellValue("Rank");
        header.createCell(1).setCellValue("Team");
        header.createCell(2).setCellValue("Total");

        int rowIndex = 1;
        int rank = 1;

        for (Map.Entry<String, Double> entry : sortedList) {
            Row row = s2.createRow(rowIndex++);
            row.createCell(0).setCellValue(rank++);
            row.createCell(1).setCellValue(entry.getKey());
            row.createCell(2).setCellValue(entry.getValue());
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
                Cell cell = row.getCell(POINTS_COL);
                if (cell != null) {
                    cell.setBlank();
                }
            }
        }

        fis.close();
        FileOutputStream fos = new FileOutputStream(path);
        workbook.write(fos);
        fos.close();
    }
}