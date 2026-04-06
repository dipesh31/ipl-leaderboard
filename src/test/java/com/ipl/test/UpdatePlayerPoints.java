package com.ipl.test;

import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.annotations.AfterClass;
import org.testng.annotations.BeforeClass;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

import com.ipl.base.TestBase;
import com.ipl.utils.ExcelReader;

import java.time.Duration;
import java.util.HashSet;
import java.util.Set;

public class UpdatePlayerPoints extends TestBase {

	String excelPath = "src/test/resources/Players.xlsx";
	ExcelReader excel = new ExcelReader(excelPath);
	Set<String> uniqueTeams = new HashSet<>();

	// dependsOnMethods ensures the driver in TestBase3 is initialized first
	@BeforeClass(dependsOnMethods = "setup")
	public void setupStatsPage() {
		
		try {
	        excel.clearPointsColumn("Sheet1"); // 🔥 IMPORTANT FIX
	    } catch (Exception e) {
	        e.printStackTrace();
	    }
		driver.get("https://fantasy.iplt20.com/classic/stats");
		WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(30));
		wait.until(ExpectedConditions.visibilityOfElementLocated(By.className("m11c-tbl__body")));
		try {
			Thread.sleep(5000);
		} catch (Exception e) {
		}
	}

	@DataProvider(name = "IPLPlayerData")
	public Object[][] getPlayerData() throws Exception {
		return excel.getSheetData("Sheet1");
	}

	@Test(dataProvider = "IPLPlayerData")
	public void capturePoints(String playerName, String iplTeam, String fantasyTeam, String role, String currentPts) {
		WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(15));
		JavascriptExecutor js = (JavascriptExecutor) driver;

		if (fantasyTeam != null && !fantasyTeam.isEmpty())
			uniqueTeams.add(fantasyTeam.trim());

		try {
			String playerXpath = "//span[text()='" + playerName + "']";
			String ptsXpath = "//span[text()='" + playerName
					+ "']/ancestor::div[contains(@class,'m11c-tbl__row')]//div[contains(@class,'cell--amt')]/span";

			WebElement row = driver.findElement(By.xpath(playerXpath));
			js.executeScript("arguments[0].scrollIntoView({behavior: 'smooth', block: 'center'});", row);
			Thread.sleep(1500);

			WebElement ptsElem = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(ptsXpath)));
			String webValue = ptsElem.getText().trim();

			excel.writePoints("Sheet1", playerName, webValue);
			System.out.println("Processed " + playerName + " (" + role + "): Web=" + webValue);

		} catch (Exception e) {
			System.err.println("Skipped: " + playerName);
		}
	}

	@AfterClass(alwaysRun = true)
	public void finalizeTotals() {
		System.out.println("--- Calculating Team Leaderboard ---");
		try {
			excel.updateAllTeamTotals("Sheet1", "Sheet2"); // ✅ single call
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
	
	
}