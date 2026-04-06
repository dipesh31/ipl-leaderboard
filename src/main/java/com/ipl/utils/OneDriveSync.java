package com.ipl.utils;

import java.io.File;
import java.io.IOException;
import java.nio.file.*;

public class OneDriveSync {
    
    // This matches your screenshot exactly
    private static final String ONEDRIVE_PATH = System.getProperty("user.home") + "/OneDrive - Pearson PLC/Players.xlsx";
    private static final String LOCAL_PATH = "src/test/resources/Players.xlsx";

    public static void sync() {
        try {
            File localFile = new File(LOCAL_PATH);
            File remoteFile = new File(ONEDRIVE_PATH);

            // Debug: Print absolute paths so you can see where Java is looking
            System.out.println("DEBUG: Local Path: " + localFile.getAbsolutePath());
            System.out.println("DEBUG: Target Path: " + remoteFile.getAbsolutePath());

            if (!localFile.exists()) {
                System.err.println("Error: Local Players.xlsx not found in src/test/resources/");
                return;
            }

            // Ensure the destination directory exists
            if (remoteFile.getParentFile() != null && !remoteFile.getParentFile().exists()) {
                remoteFile.getParentFile().mkdirs();
            }

            // Perform the copy
            Files.copy(localFile.toPath(), remoteFile.toPath(), StandardCopyOption.REPLACE_EXISTING);
            System.out.println("OneDrive Sync Successful!");

        } catch (IOException e) {
            System.err.println("Sync Failed: " + e.getMessage());
            if (e.getMessage().contains("used by another process")) {
                System.err.println("ACTION REQUIRED: Close the Excel file on your Desktop/iPad.");
            }
            e.printStackTrace();
        }
    }
}