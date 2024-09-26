package org.example;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.testng.annotations.AfterClass;
import org.testng.annotations.BeforeClass;
import org.testng.annotations.Test;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.*;
import java.util.concurrent.TimeUnit;

public class Main {
    public WebDriver driver;

    @BeforeClass
    public void setup() {
        System.setProperty("webdriver.chrome.driver", "C:\\Users\\Insha Praveen\\Desktop\\Collection\\Assisment\\chromedriver-win64\\chromedriver.exe");

        driver = new ChromeDriver();
        driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
        driver.manage().window().maximize();
    }

    @Test
    public void searchLgSoundbar() throws IOException {
        driver.get("https://www.amazon.in");

        WebElement searchBox = driver.findElement(By.id("twotabsearchtextbox"));
        searchBox.sendKeys("lg soundbar");
        searchBox.submit();

        driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);

        List<WebElement> productNames = driver.findElements(By.cssSelector("span.a-size-medium.a-color-base.a-text-normal"));
        List<WebElement> productPrices = driver.findElements(By.cssSelector("span.a-price-whole"));

        Map<String, Integer> productsMap = new HashMap<>();

        for (int i = 0; i < productNames.size(); i++) {
            String name = productNames.get(i).getText();
            int price = 0;

            if (i < productPrices.size()) {
                try {
                    price = Integer.parseInt(productPrices.get(i).getText().replace(",", ""));
                } catch (NumberFormatException e) {
                    price = 0;
                }
            }
            productsMap.put(name, price);
        }

        List<Map.Entry<String, Integer>> sortedProducts = new ArrayList<>(productsMap.entrySet());
        sortedProducts.sort(Map.Entry.comparingByValue());

        // Write data to an Excel file
        writeToExcel(sortedProducts);
    }

    public void writeToExcel(List<Map.Entry<String, Integer>> sortedProducts) throws IOException {
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Products");

        // Create header row
        Row headerRow = sheet.createRow(0);
        Cell headerCell1 = headerRow.createCell(0);
        headerCell1.setCellValue("Product Name");
        Cell headerCell2 = headerRow.createCell(1);
        headerCell2.setCellValue("Price");

        // Populate rows with product data
        int rowNum = 1;
        for (Map.Entry<String, Integer> entry : sortedProducts) {
            Row row = sheet.createRow(rowNum++);
            row.createCell(0).setCellValue(entry.getKey());
            row.createCell(1).setCellValue(entry.getValue());
        }

        // Write the Excel file to the file system
        try (FileOutputStream fileOut = new FileOutputStream("products.xlsx")) {
            workbook.write(fileOut);
        }

        workbook.close();
        System.out.println("Excel file created: products.xlsx");
    }

    @AfterClass
    public void teardown() {
        if (driver != null) {
            driver.quit();
        }
    }
}
