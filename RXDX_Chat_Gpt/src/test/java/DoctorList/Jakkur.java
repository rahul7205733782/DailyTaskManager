package DoctorList;

import java.time.Duration;
import java.util.*;
import org.openqa.selenium.*;
import org.openqa.selenium.edge.EdgeDriver;
import org.openqa.selenium.edge.EdgeOptions;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

// Excel Imports
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.util.CellRangeAddress;

import java.awt.Desktop;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;

public class Jakkur {

    public static void main(String[] args)
            throws InterruptedException, IOException {

        System.setProperty("org.apache.poi.util.POILogger", "org.apache.poi.util.NullLogger");

        String projectPath = System.getProperty("user.dir");
        String timestamp = new SimpleDateFormat("yyyyMMdd_HHmmss").format(new Date());

        String screenshotsPath = projectPath + "/RXDX_DoctorList_Jakkur/Screenshots/" + timestamp;
        String excelFolderPath = projectPath + "/RXDX_DoctorList_Jakkur/Test_Excel";
        String excelFilePath = excelFolderPath + "/RXDX_Jakkur_DoctorList_Results_" + timestamp + ".xlsx";

        new File(screenshotsPath).mkdirs();
        new File(excelFolderPath).mkdirs();

        System.out.println("\n📁 SCREENSHOTS will be saved in: " + screenshotsPath);
        System.out.println("📊 EXCEL FILE will be saved as: " + excelFilePath);

        // ================= DEPARTMENTS WITH DOCTOR LIST QUESTIONS =================
        Map<String, String> departmentQuestions = new LinkedHashMap<>();

        departmentQuestions.put("Adult Cardiology", "show me cardiology doctors");
        departmentQuestions.put("Diabetologist", "list of diabetologists");
        departmentQuestions.put("Family Medicine", "family medicine doctor list");
        departmentQuestions.put("Gastroenterologist", "need gastroenterologist list");
        departmentQuestions.put("General Medicine", "general medicine physicians");
        departmentQuestions.put("General Surgeon", "surgeons in general surgery");
        departmentQuestions.put("Nephrologist", "kidney specialist doctors");
        departmentQuestions.put("Obs & Gynae", "gynecologist list please");
        departmentQuestions.put("Ophthalmology", "eye doctors available");
        departmentQuestions.put("Orthopedician", "orthopedic surgeon list");
        departmentQuestions.put("Paediatric Endocrinologist", "child hormone specialists");
        departmentQuestions.put("Paediatrician", "pediatric doctors list");
        departmentQuestions.put("Physiotherapy", "physiotherapist details");
        departmentQuestions.put("Psychology", "psychologist list");
        departmentQuestions.put("Pulmonologist", "lung specialist doctors");
        departmentQuestions.put("Urologist", "urology doctor list");

        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("RXDX_Jakkur_DoctorList_Results");
        addProjectHeader(sheet, workbook, "RXDX JAKKUR - DOCTOR LIST VALIDATION TESTS");
        int headerRowIndex = addColumnHeaders(sheet, workbook);
        int excelRowNum = headerRowIndex + 1;

        EdgeOptions options = new EdgeOptions();
        options.addArguments("--user-data-dir=C:\\EdgeAutomationProfile");
        options.addArguments("--start-maximized");

        WebDriver driver = new EdgeDriver(options);
        WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(30));

        driver.manage().window().maximize();
        driver.get("https://web.whatsapp.com/");

        System.out.println("\n⚠️ PLEASE SCAN QR CODE...");
        Thread.sleep(20000);

        // Open RXDX chat
        driver.findElement(By.xpath("//span[@data-icon='new-chat-outline']")).click();
        driver.findElement(By.xpath("//input[@aria-label='Search name or number']")).sendKeys("916363415530");
        Thread.sleep(4000);
        driver.findElement(By.xpath("//span[@title='+91 63634 15530']")).click();
        Thread.sleep(4000);

        // Wait for chat input ready
        wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//footer//div[@contenteditable='true']")));
        System.out.println("✅ Chat ready.");

        int totalTests = departmentQuestions.size();
        List<Map<String, String>> failedTestsData = new ArrayList<>();

        System.out.println("\n" + "=".repeat(80));
        System.out.println("     DOCTOR LIST VALIDATION TESTS STARTING - TOTAL TEST CASES: " + totalTests);
        System.out.println("     BRANCH: Jakkur (Option 7)");
        System.out.println("=".repeat(80));

        int testNumber = 1;

        for (Map.Entry<String, String> entry : departmentQuestions.entrySet()) {
            String expectedDepartment = entry.getKey();
            String question = entry.getValue();

            String responseAfterHi = "";
            String responseAfterQuestion = "";
            String responseAfterBranch = "";
            String doctorListResponse = "";

            String status = "PASS";
            String failureType = "";
            String failureReason = "";

            String screenshotPath = "";
            boolean hasClosingText = false;
            boolean hasDoctors = false;
            int doctorCount = 0;

            System.out.println("\n" + "=".repeat(60));
            System.out.println("TEST #" + testNumber + " of " + totalTests);
            System.out.println("Department: " + expectedDepartment);
            System.out.println("Question: " + question);
            System.out.println("=".repeat(60));

            try {
                // STEP 1: Send "hi"
                System.out.println("\n📤 STEP 1: Sending 'hi'...");
                sendMessage(driver, "hi", wait);
                Thread.sleep(8000);
                responseAfterHi = getLastMessage(driver);
                System.out.println("✅ Bot responded to 'hi': " + responseAfterHi.substring(0, Math.min(100, responseAfterHi.length())));

                // STEP 2: Send doctor list question
                System.out.println("\n📤 STEP 2: Sending question: " + question);
                sendMessage(driver, question, wait);
                Thread.sleep(8000);
                responseAfterQuestion = getLastMessage(driver);
                System.out.println("✅ Branch menu received: " + responseAfterQuestion.substring(0, Math.min(100, responseAfterQuestion.length())));

                // STEP 3: Send branch number (7 for Jakkur)
                System.out.println("\n📤 STEP 3: Sending branch number 7 (RxDx Jakkur)...");
                sendMessage(driver, "7", wait);
                Thread.sleep(8000);
                responseAfterBranch = getLastMessage(driver);
                System.out.println("✅ Branch selected");

                // STEP 4: Send "1" to get doctor list
                System.out.println("\n📤 STEP 4: Sending '1' to get doctor list...");
                sendMessage(driver, "1", wait);
                Thread.sleep(8000);
                doctorListResponse = getLastMessage(driver);
                System.out.println("\n📋 DOCTOR LIST RESPONSE:\n" + doctorListResponse);

                // STEP 5: VALIDATION
                System.out.println("\n" + "=".repeat(50));
                System.out.println("STEP 5: VALIDATING DOCTOR LIST RESPONSE");
                System.out.println("=".repeat(50));

                // Validation 1: Closing text
                String expectedClosingText = "Please choose from the following list of doctors";
                hasClosingText = doctorListResponse.contains(expectedClosingText);
                if (!hasClosingText) {
                    status = "FAIL";
                    failureType = "MISSING_CLOSING_TEXT";
                    failureReason = "Missing closing text: '" + expectedClosingText + "'";
                    System.out.println("❌ FAIL: Missing closing text!");
                } else {
                    System.out.println("✅ PASS: Closing text found.");
                }

                // Validation 2: Doctor count
                String[] lines = doctorListResponse.split("\n");
                for (String line : lines) {
                    if (line.contains("Dr.") || (line.matches(".*\\d+\\..*") && line.contains("Dr"))) {
                        doctorCount++;
                    }
                }
                hasDoctors = doctorCount > 0;
                if (!hasDoctors) {
                    status = "FAIL";
                    failureType = (status.equals("FAIL") ? failureType + "_AND_NO_DOCTORS" : "NO_DOCTORS_FOUND");
                    failureReason = (failureReason.isEmpty() ? "" : failureReason + ". ") +
                                   "No doctors found in the response.";
                    System.out.println("❌ FAIL: No doctors found in response!");
                } else {
                    System.out.println("✅ PASS: Found " + doctorCount + " doctor(s) in the list.");
                }

                if (status.equals("PASS")) {
                    System.out.println("\n✅✅✅ TEST PASSED for: " + expectedDepartment);
                } else {
                    System.out.println("\n❌❌❌ TEST FAILED for: " + expectedDepartment);
                    System.out.println("   Reason: " + failureReason);
                    screenshotPath = takeScreenshot(driver, screenshotsPath, expectedDepartment);
                }

            } catch (Exception e) {
                status = "FAIL";
                failureType = "EXCEPTION";
                failureReason = "Exception: " + e.getMessage();
                screenshotPath = takeScreenshot(driver, screenshotsPath, expectedDepartment);
                e.printStackTrace();
            }

            addTestResultToExcel(sheet, excelRowNum++, testNumber, expectedDepartment, question,
                        responseAfterHi, responseAfterQuestion, responseAfterBranch, doctorListResponse,
                        status, failureType, failureReason, hasClosingText, hasDoctors, doctorCount, screenshotPath);

            System.out.println("\n📊 FINAL STATUS: " + status);

            if (status.equals("FAIL")) {
                Map<String, String> failureData = new HashMap<>();
                failureData.put("testNumber", String.valueOf(testNumber));
                failureData.put("departmentName", expectedDepartment);
                failureData.put("failureType", failureType);
                failureData.put("failureReason", failureReason);
                failureData.put("screenshotPath", screenshotPath);
                failedTestsData.add(failureData);
            }

            testNumber++;
            Thread.sleep(3000);
        }

        for (int i = 0; i <= 13; i++) sheet.autoSizeColumn(i);

        try (FileOutputStream fileOut = new FileOutputStream(excelFilePath)) {
            workbook.write(fileOut);
            System.out.println("\n✅ EXCEL SAVED: " + excelFilePath);
        }
        workbook.close();

        printFinalReport(totalTests, failedTestsData);

        File screenshotFolder = new File(screenshotsPath);
        File[] screenshotFiles = screenshotFolder.listFiles();
        if (screenshotFiles != null && screenshotFiles.length > 0) {
            System.out.println("\n📸 Total Screenshots: " + screenshotFiles.length);
            for (int i = 0; i < screenshotFiles.length; i++) {
                System.out.println((i + 1) + ". " + screenshotFiles[i].getAbsolutePath());
            }
        }

        try {
            File excelFile = new File(excelFilePath);
            if (Desktop.isDesktopSupported()) Desktop.getDesktop().open(excelFile);
            System.out.println("\n✅ Excel Opened Automatically");
        } catch (Exception e) {
            System.out.println("❌ Unable to open Excel automatically");
        }

        driver.quit();
    }

    // ========== IMPROVED sendMessage (reliable, clears input) ==========
    static void sendMessage(WebDriver driver, String text, WebDriverWait wait) throws InterruptedException {
        WebElement input = wait.until(ExpectedConditions.elementToBeClickable(
                By.xpath("//footer//div[@contenteditable='true']")
        ));
        input.click();
        input.sendKeys(Keys.CONTROL + "a");
        Thread.sleep(100);
        input.sendKeys(Keys.DELETE);
        Thread.sleep(200);
        for (char c : text.toCharArray()) {
            input.sendKeys(String.valueOf(c));
            Thread.sleep(50);
        }
        Thread.sleep(500);
        WebElement sendBtn = wait.until(ExpectedConditions.elementToBeClickable(
                By.xpath("//span[@data-icon='wds-ic-send-filled']")
        ));
        sendBtn.click();
        System.out.println("✅ Message sent: " + text);
        // No extra sleep – caller will sleep 8 seconds.
    }

    // ========== EXCEL AND HELPER METHODS ==========
    static void addProjectHeader(Sheet sheet, Workbook workbook, String projectTitle) {
        Row headerRow = sheet.createRow(0);
        Cell headerCell = headerRow.createCell(0);
        headerCell.setCellValue(projectTitle);
        CellStyle headerStyle = workbook.createCellStyle();
        Font headerFont = workbook.createFont();
        headerFont.setBold(true);
        headerFont.setFontHeightInPoints((short) 16);
        headerFont.setColor(IndexedColors.DARK_BLUE.getIndex());
        headerStyle.setFont(headerFont);
        headerStyle.setFillForegroundColor(IndexedColors.LIGHT_YELLOW.getIndex());
        headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        headerStyle.setAlignment(HorizontalAlignment.CENTER);
        headerStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        headerStyle.setBorderBottom(BorderStyle.MEDIUM);
        headerStyle.setBorderTop(BorderStyle.MEDIUM);
        headerStyle.setBorderLeft(BorderStyle.MEDIUM);
        headerStyle.setBorderRight(BorderStyle.MEDIUM);
        headerCell.setCellStyle(headerStyle);
        sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 13));

        Row infoRow = sheet.createRow(1);
        Cell infoCell = infoRow.createCell(0);
        infoCell.setCellValue("Project: RXDX Jakkur Doctor List Validation | Branch: Jakkur (Option 7) | Departments: 16 | Date: " +
                             new SimpleDateFormat("yyyy-MM-dd HH:mm:ss").format(new Date()));
        CellStyle infoStyle = workbook.createCellStyle();
        Font infoFont = workbook.createFont();
        infoFont.setBold(true);
        infoFont.setFontHeightInPoints((short) 11);
        infoFont.setColor(IndexedColors.WHITE.getIndex());
        infoStyle.setFont(infoFont);
        infoStyle.setFillForegroundColor(IndexedColors.GREY_50_PERCENT.getIndex());
        infoStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        infoStyle.setAlignment(HorizontalAlignment.CENTER);
        infoCell.setCellStyle(infoStyle);
        sheet.addMergedRegion(new CellRangeAddress(1, 1, 0, 13));
    }

    static int addColumnHeaders(Sheet sheet, Workbook workbook) {
        Row columnHeaderRow = sheet.createRow(2);
        String[] columns = {
            "Test No", "Department", "Question", "Response After Hi",
            "Response After Question", "Response After Branch", "Doctor List Response", "Status",
            "Failure Type", "Failure Reason", "Closing Text Found", "Doctors Found",
            "Doctor Count", "Screenshot", "Timestamp"
        };
        CellStyle headerStyle = workbook.createCellStyle();
        Font headerFont = workbook.createFont();
        headerFont.setBold(true);
        headerFont.setFontHeightInPoints((short) 10);
        headerFont.setColor(IndexedColors.WHITE.getIndex());
        headerStyle.setFont(headerFont);
        headerStyle.setFillForegroundColor(IndexedColors.DARK_GREEN.getIndex());
        headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        headerStyle.setAlignment(HorizontalAlignment.CENTER);
        headerStyle.setBorderBottom(BorderStyle.MEDIUM);
        for (int i = 0; i < columns.length; i++) {
            Cell cell = columnHeaderRow.createCell(i);
            cell.setCellValue(columns[i]);
            cell.setCellStyle(headerStyle);
        }
        return 2;
    }

    static void addTestResultToExcel(Sheet sheet, int rowNum, int testNumber, String expectedDepartment,
                                      String question, String responseAfterHi, String responseAfterQuestion,
                                      String responseAfterBranch, String doctorListResponse, String status,
                                      String failureType, String failureReason, boolean hasClosingText,
                                      boolean hasDoctors, int doctorCount, String screenshotPath) {
        Row row = sheet.createRow(rowNum);
        Workbook workbook = sheet.getWorkbook();
        CellStyle failStyle = workbook.createCellStyle();
        Font failFont = workbook.createFont();
        failFont.setColor(IndexedColors.RED.getIndex());
        failFont.setBold(true);
        failStyle.setFont(failFont);
        failStyle.setFillForegroundColor(IndexedColors.LIGHT_YELLOW.getIndex());
        failStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        CellStyle passStyle = workbook.createCellStyle();
        Font passFont = workbook.createFont();
        passFont.setColor(IndexedColors.DARK_GREEN.getIndex());
        passFont.setBold(true);
        passStyle.setFont(passFont);
        passStyle.setFillForegroundColor(IndexedColors.LIGHT_GREEN.getIndex());
        passStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        CellStyle wrapStyle = workbook.createCellStyle();
        wrapStyle.setWrapText(true);
        wrapStyle.setVerticalAlignment(VerticalAlignment.TOP);

        row.createCell(0).setCellValue(testNumber);
        row.createCell(1).setCellValue(expectedDepartment);
        row.createCell(2).setCellValue(question);
        Cell cell3 = row.createCell(3);
        cell3.setCellValue(responseAfterHi != null && !responseAfterHi.isEmpty() ? responseAfterHi : "No response");
        cell3.setCellStyle(wrapStyle);
        Cell cell4 = row.createCell(4);
        cell4.setCellValue(responseAfterQuestion != null && !responseAfterQuestion.isEmpty() ? responseAfterQuestion : "No response");
        cell4.setCellStyle(wrapStyle);
        Cell cell5 = row.createCell(5);
        cell5.setCellValue(responseAfterBranch != null && !responseAfterBranch.isEmpty() ? responseAfterBranch : "No response");
        cell5.setCellStyle(wrapStyle);
        Cell cell6 = row.createCell(6);
        cell6.setCellValue(doctorListResponse != null && !doctorListResponse.isEmpty() ? doctorListResponse : "No response");
        cell6.setCellStyle(wrapStyle);
        Cell statusCell = row.createCell(7);
        statusCell.setCellValue(status);
        if (status.equals("PASS")) statusCell.setCellStyle(passStyle);
        else statusCell.setCellStyle(failStyle);
        row.createCell(8).setCellValue(failureType != null ? failureType : "N/A");
        Cell cell9 = row.createCell(9);
        cell9.setCellValue(failureReason != null ? failureReason : "N/A");
        cell9.setCellStyle(wrapStyle);
        row.createCell(10).setCellValue(hasClosingText ? "YES (PASS)" : "NO (FAIL)");
        row.createCell(11).setCellValue(hasDoctors ? "YES (PASS)" : "NO (FAIL)");
        row.createCell(12).setCellValue(doctorCount);
        row.createCell(13).setCellValue(screenshotPath != null ? screenshotPath : "No screenshot");
        row.createCell(14).setCellValue(new SimpleDateFormat("yyyy-MM-dd HH:mm:ss").format(new Date()));
    }

    static String getLastMessage(WebDriver driver) {
        try {
            List<WebElement> messages = driver.findElements(By.xpath("//div[contains(@class,'message-in')]"));
            if (!messages.isEmpty()) return messages.get(messages.size() - 1).getText();
        } catch (Exception e) {}
        return "NO RESPONSE";
    }

    static String takeScreenshot(WebDriver driver, String folderPath, String fileName) {
        try {
            String cleanFileName = fileName.replaceAll("[^a-zA-Z0-9]", "_");
            String path = folderPath + "/" + cleanFileName + ".png";
            File src = ((TakesScreenshot) driver).getScreenshotAs(OutputType.FILE);
            File destFile = new File(path);
            src.renameTo(destFile);
            System.out.println("📸 Screenshot Saved: " + path);
            return path;
        } catch (Exception e) {
            return "Screenshot Failed: " + e.getMessage();
        }
    }

    static void printFinalReport(int totalTests, List<Map<String, String>> failedTestsData) {
        System.out.println("\n" + "=".repeat(80));
        System.out.println("                     FINAL TEST REPORT");
        System.out.println("=".repeat(80));
        if (failedTestsData.isEmpty()) {
            System.out.println("\n🎉 ALL TESTS PASSED!");
            System.out.println("✅ Total Tests: " + totalTests);
            System.out.println("✅ Success Rate: 100%");
        } else {
            System.out.println("\n❌ FAILURES: " + failedTestsData.size() + " of " + totalTests);
            Map<String, Integer> failuresByType = new HashMap<>();
            for (Map<String, String> failure : failedTestsData) {
                String type = failure.get("failureType");
                failuresByType.put(type, failuresByType.getOrDefault(type, 0) + 1);
            }
            System.out.println("\nFAILURES BY TYPE:");
            for (Map.Entry<String, Integer> entry : failuresByType.entrySet()) {
                System.out.println("• " + entry.getKey() + ": " + entry.getValue());
            }
            System.out.println("\n📊 SUMMARY:");
            System.out.println("Passed: " + (totalTests - failedTestsData.size()));
            System.out.println("Failed: " + failedTestsData.size());
            System.out.println("Success Rate: " +
                String.format("%.1f%%", ((totalTests - failedTestsData.size()) * 100.0 / totalTests)));
        }
    }
}