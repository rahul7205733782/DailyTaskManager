package RXDX_Kadugodi;

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

        System.setProperty(
                "org.apache.poi.util.POILogger",
                "org.apache.poi.util.NullLogger"
        );

        String projectPath = System.getProperty("user.dir");
        String timestamp = new SimpleDateFormat("yyyyMMdd_HHmmss").format(new Date());

        String screenshotsPath = projectPath + "/RXDX_Validation_Jakkur/Screenshots/" + timestamp;
        String excelFolderPath = projectPath + "/RXDX_Validation_Jakkur/Test_Excel";
        String excelFilePath = excelFolderPath + "/RXDX_Jakkur_Test_Results_" + timestamp + ".xlsx";

        new File(screenshotsPath).mkdirs();
        new File(excelFolderPath).mkdirs();

        System.out.println("\n📁 SCREENSHOTS will be saved in: " + screenshotsPath);
        System.out.println("📊 EXCEL FILE will be saved as: " + excelFilePath);

        Map<String, String> departmentQuestions = new LinkedHashMap<>();

        departmentQuestions.put("Adult Cardiology", "I get chest pain while walking fast");
        departmentQuestions.put("Diabetologist", "My sugar levels are very high recently");
        departmentQuestions.put("Family Medicine", "I need a family doctor for regular health checkup");
        departmentQuestions.put("Gastroenterologist", "I have severe acidity and stomach bloating");
        departmentQuestions.put("General Medicine", "I have fever weakness and body pain");
        departmentQuestions.put("General Surgeon", "I need consultation for gallbladder surgery");
        departmentQuestions.put("Nephrologist", "I have kidney stone and urine burning");
        departmentQuestions.put("Obs & Gynae", "I have irregular periods and abdominal cramps");
        departmentQuestions.put("Ophthalmology", "I have blurry vision and eye irritation");
        departmentQuestions.put("Orthopedician", "I have knee pain while climbing stairs");
        departmentQuestions.put("Paediatric Endocrinologist", "My child has hormonal growth problem");
        departmentQuestions.put("Paediatrician", "My child has fever cough and cold");
        departmentQuestions.put("Physiotherapy", "I need physiotherapy after sports injury");
        departmentQuestions.put("Psychology", "I need counselling for anxiety and stress");
        departmentQuestions.put("Pulmonologist", "I have breathing difficulty and chest congestion");
        departmentQuestions.put("Urologist", "I have urine infection and lower abdominal pain");
        
        Set<String> validDepartments = new HashSet<>(departmentQuestions.keySet());

        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("RXDX_Jakkur_Results");
        addProjectHeader(sheet, workbook, "RXDX JAKKUR CHATBOT - DEPARTMENT APPOINTMENT TESTS");
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

        int totalTests = departmentQuestions.size();
        List<Map<String, String>> failedTestsData = new ArrayList<>();

        System.out.println("\n" + "=".repeat(80));
        System.out.println("     DEPARTMENT-WISE APPOINTMENT TESTS STARTING - TOTAL TEST CASES: " + totalTests);
        System.out.println("=".repeat(80));

        int testNumber = 1;

        for (Map.Entry<String, String> entry : departmentQuestions.entrySet()) {
            String expectedDepartment = entry.getKey();
            String question = entry.getValue();
            String botResponse = "";
            String placeResponse = "";
            String status = "PASS";
            String failureType = "";
            String failureReason = "";
            String screenshotPath = "";
            String extraDepartments = "";
            String allDepartmentsShown = "";

            System.out.println("\n======================================");
            System.out.println("TEST #" + testNumber + " of " + totalTests);
            System.out.println("Department: " + expectedDepartment);
            System.out.println("Question: " + question);
            System.out.println("======================================");

            try {
                // STEP 1: Send "hi" and wait 8 seconds
                System.out.println("\n📤 STEP 1: Sending 'hi' to start conversation...");
                sendMessage(driver, "hi", wait);
                Thread.sleep(8000);
                botResponse = getLastMessage(driver);
                System.out.println("✅ Response after 'hi': " + botResponse.substring(0, Math.min(150, botResponse.length())));

                // STEP 2: Send department question and wait 8 seconds
                System.out.println("\n📤 STEP 2: Sending department question: " + question);
                sendMessage(driver, question, wait);
                Thread.sleep(8000);
                botResponse = getLastMessage(driver);
                System.out.println("\nBOT RESPONSE:\n" + botResponse);

                // Basic validation (check if response contains relevant keywords)
                boolean valid = validateResponse(expectedDepartment, botResponse);
                if (!valid) {
                    status = "FAIL";
                    failureType = "WRONG_GPT_RESPONSE";
                    failureReason = "Wrong department mapping - Expected department keywords not found in GPT response";
                    System.out.println("❌ WRONG GPT RESPONSE");
                }

                // STEP 3: Click "Book an Appointment" button
                System.out.println("\n🔍 STEP 3: Looking for 'Book an Appointment' button...");
                List<WebElement> buttons = driver.findElements(By.xpath("//span[contains(text(),'Book an Appointment')]"));
                if (!buttons.isEmpty()) {
                    buttons.get(buttons.size() - 1).click();
                    System.out.println("✅ Book Appointment Clicked");
                    Thread.sleep(8000);
                    String bookAppResponse = getLastMessage(driver);
                    System.out.println("✅ Response after Book Appointment:\n" + bookAppResponse);
                } else {
                    status = "FAIL";
                    failureType = "MISSING_BOOK_APPOINTMENT_BUTTON";
                    failureReason = (failureReason.isEmpty() ? "" : failureReason + ". ") + "Book Appointment button missing";
                    System.out.println("❌ BUTTON NOT FOUND");
                }

                // STEP 4: Send "7" for Jakkur location and wait 8 seconds
                System.out.println("\n📤 STEP 4: Sending '7' for Jakkur location selection...");
                sendMessage(driver, "7", wait);
                Thread.sleep(8000);
                placeResponse = getLastMessage(driver);
                System.out.println("\nLOCATION RESPONSE:\n" + placeResponse);

                // STEP 5: Validate filtering
                System.out.println("\n🔍 STEP 5: Validating filtering (Rajagiri style)...");
                allDepartmentsShown = extractAllDepartmentsFromResponse(placeResponse, validDepartments);
                extraDepartments = findExtraDepartments(allDepartmentsShown, expectedDepartment);
                
                if (!placeResponse.toLowerCase().contains(expectedDepartment.toLowerCase())) {
                    status = "FAIL";
                    failureType = "DEPARTMENT_NOT_FOUND";
                    failureReason = (failureReason.isEmpty() ? "" : failureReason + ". ") +
                                   "Expected department '" + expectedDepartment + "' not displayed in location response";
                } else if (!extraDepartments.isEmpty()) {
                    status = "FAIL";
                    failureType = "FILTERING_FAILED_EXTRA_DEPARTMENTS";
                    failureReason = (failureReason.isEmpty() ? "" : failureReason + ". ") +
                                   "FILTERING VALIDATION FAILED: Bot showed extra departments not requested. " +
                                   "Patient asked for: " + expectedDepartment + ". Extra departments shown: " + extraDepartments;
                    System.out.println("❌ FILTERING FAILED! Extra departments found: " + extraDepartments);
                } else {
                    System.out.println("✅ FILTERING VALIDATION PASSED! Only showing: " + expectedDepartment);
                }
                
                if (status.equals("PASS")) {
                    System.out.println("\n✅ All validations passed for department: " + expectedDepartment);
                }

                if (status.equals("FAIL")) {
                    screenshotPath = takeScreenshot(driver, screenshotsPath, expectedDepartment);
                }

            } catch (Exception e) {
                status = "FAIL";
                failureType = "EXCEPTION";
                failureReason = (failureReason.isEmpty() ? "" : failureReason + ". ") + "Exception: " + e.getMessage();
                screenshotPath = takeScreenshot(driver, screenshotsPath, expectedDepartment);
                System.out.println("❌ EXCEPTION : " + e.getMessage());
            }

            addTestResultToExcel(sheet, excelRowNum++, testNumber, expectedDepartment, question,
                         botResponse, placeResponse, failureType, failureReason, extraDepartments,
                         allDepartmentsShown, screenshotPath, status);

            System.out.println("\nSTATUS: " + status);
            
            if (status.equals("FAIL")) {
                Map<String, String> failureData = new HashMap<>();
                failureData.put("testNumber", String.valueOf(testNumber));
                failureData.put("departmentName", expectedDepartment);
                failureData.put("failureType", failureType);
                failureData.put("failureReason", failureReason);
                failureData.put("screenshotPath", screenshotPath);
                failedTestsData.add(failureData);
            }
            
            if (!extraDepartments.isEmpty()) {
                System.out.println("⚠️ EXTRA DEPARTMENTS FOUND: " + extraDepartments);
                System.out.println("📋 ALL DEPARTMENTS SHOWN: " + allDepartmentsShown);
            }

            testNumber++;
            Thread.sleep(5000);
        }

        for (int i = 0; i <= 11; i++) {
            sheet.autoSizeColumn(i);
        }

        try (FileOutputStream fileOut = new FileOutputStream(excelFilePath)) {
            workbook.write(fileOut);
            System.out.println("\n✅ EXCEL SAVED: " + excelFilePath);
        }
        workbook.close();
        
        printFinalReport(totalTests, failedTestsData);
        
        System.out.println("\n" + "=".repeat(80));
        System.out.println("                    ALL SCREENSHOT PATHS");
        System.out.println("=".repeat(80));
        
        File screenshotFolder = new File(screenshotsPath);
        File[] screenshotFiles = screenshotFolder.listFiles();
        if (screenshotFiles != null && screenshotFiles.length > 0) {
            System.out.println("📸 Total Screenshots: " + screenshotFiles.length);
            for (int i = 0; i < screenshotFiles.length; i++) {
                System.out.println((i + 1) + ". " + screenshotFiles[i].getAbsolutePath());
            }
        }

        try {
            File excelFile = new File(excelFilePath);
            if (Desktop.isDesktopSupported()) {
                Desktop.getDesktop().open(excelFile);
                System.out.println("\n✅ Excel Opened Automatically");
            }
        } catch (Exception e) {
            System.out.println("❌ Unable to open Excel automatically");
            e.printStackTrace();
        }

        driver.quit();
    }
    
    // ================= IMPROVED sendMessage WITH CHARACTER-BY-CHARACTER TYPING =================
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
        // No extra sleep – caller will add Thread.sleep(8000)
    }

    // ================= EXCEL METHODS (unchanged) =================
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
        sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 11));
        
        Row infoRow = sheet.createRow(1);
        Cell infoCell = infoRow.createCell(0);
        infoCell.setCellValue("Project: RXDX Jakkur Chatbot Testing | Location: Jakkur (Option 7) | Testing: Department-wise Appointment Booking | Total Departments: 16 | Wait: 8 Seconds per Step | Date: " + 
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
        sheet.addMergedRegion(new CellRangeAddress(1, 1, 0, 11));
    }
    
    static int addColumnHeaders(Sheet sheet, Workbook workbook) {
        Row columnHeaderRow = sheet.createRow(2);
        String[] columns = {
            "Test No", "Department", "Question", "Bot Response", "Location Response",
            "Status", "Failure Type", "Failure Reason", "Extra Departments Found",
            "All Departments Shown", "Screenshot", "Timestamp"
        };
        CellStyle headerStyle = workbook.createCellStyle();
        Font headerFont = workbook.createFont();
        headerFont.setBold(true);
        headerFont.setFontHeightInPoints((short) 12);
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
                                      String question, String botResponse, String placeResponse, 
                                      String failureType, String failureReason, String extraDepartments,
                                      String allDepartmentsShown, String screenshotPath, String status) {
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
        Cell botResponseCell = row.createCell(3);
        botResponseCell.setCellValue(botResponse != null && !botResponse.isEmpty() ? botResponse : "No response");
        botResponseCell.setCellStyle(wrapStyle);
        Cell placeResponseCell = row.createCell(4);
        placeResponseCell.setCellValue(placeResponse != null && !placeResponse.isEmpty() ? placeResponse : "No response");
        placeResponseCell.setCellStyle(wrapStyle);
        Cell statusCell = row.createCell(5);
        statusCell.setCellValue(status);
        if (status.equals("PASS")) statusCell.setCellStyle(passStyle);
        else statusCell.setCellStyle(failStyle);
        row.createCell(6).setCellValue(failureType != null ? failureType : "N/A");
        Cell failureCell = row.createCell(7);
        failureCell.setCellValue(failureReason != null ? failureReason : "N/A");
        failureCell.setCellStyle(wrapStyle);
        row.createCell(8).setCellValue(extraDepartments != null && !extraDepartments.isEmpty() ? extraDepartments : "None");
        row.createCell(9).setCellValue(allDepartmentsShown != null ? allDepartmentsShown : "None");
        row.createCell(10).setCellValue(screenshotPath != null ? screenshotPath : "No screenshot");
        row.createCell(11).setCellValue(new SimpleDateFormat("yyyy-MM-dd HH:mm:ss").format(new Date()));
    }

    static String getLastMessage(WebDriver driver) {
        try {
            List<WebElement> messages = driver.findElements(By.xpath("//div[contains(@class,'message-in')]"));
            if (!messages.isEmpty()) {
                return messages.get(messages.size() - 1).getText();
            }
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
    
    static String extractAllDepartmentsFromResponse(String response, Set<String> validDepartments) {
        List<String> foundDepartments = new ArrayList<>();
        String lowerResponse = response.toLowerCase();
        for (String department : validDepartments) {
            String lowerDept = department.toLowerCase();
            if (lowerResponse.contains(lowerDept)) {
                foundDepartments.add(department);
            } else {
                String[] deptParts = lowerDept.split("[ ,&]");
                for (String part : deptParts) {
                    if (part.length() > 3 && lowerResponse.contains(part)) {
                        foundDepartments.add(department);
                        break;
                    }
                }
            }
        }
        return foundDepartments.isEmpty() ? "None" : String.join(", ", foundDepartments);
    }
    
    static String findExtraDepartments(String allDepartments, String expectedDepartment) {
        if (allDepartments == null || allDepartments.isEmpty() || allDepartments.equals("None")) {
            return "";
        }
        String[] departments = allDepartments.split(", ");
        List<String> extraDepts = new ArrayList<>();
        Set<String> expectedSet = new HashSet<>(Arrays.asList(expectedDepartment.split(", ")));
        for (String dept : departments) {
            boolean isExpected = false;
            for (String expected : expectedSet) {
                if (dept.equalsIgnoreCase(expected) || 
                    dept.toLowerCase().contains(expected.toLowerCase()) ||
                    expected.toLowerCase().contains(dept.toLowerCase())) {
                    isExpected = true;
                    break;
                }
            }
            if (!isExpected && !dept.isEmpty() && !dept.equals("None")) {
                extraDepts.add(dept);
            }
        }
        return extraDepts.isEmpty() ? "" : String.join(", ", extraDepts);
    }

    static boolean validateResponse(String department, String response) {
        String lower = response.toLowerCase();
        if (department.toLowerCase().contains("cardiology") && lower.contains("heart")) return true;
        if (department.toLowerCase().contains("diabetologist") && lower.contains("sugar")) return true;
        if (department.toLowerCase().contains("family medicine") && lower.contains("family")) return true;
        if (department.toLowerCase().contains("gastro") && (lower.contains("stomach") || lower.contains("acidity"))) return true;
        if (department.toLowerCase().contains("general medicine") && lower.contains("fever")) return true;
        if (department.toLowerCase().contains("general surgeon") && lower.contains("surgery")) return true;
        if (department.toLowerCase().contains("nephrologist") && lower.contains("kidney")) return true;
        if (department.toLowerCase().contains("obs") && lower.contains("period")) return true;
        if (department.toLowerCase().contains("ophthalmology") && lower.contains("eye")) return true;
        if (department.toLowerCase().contains("orthopedician") && lower.contains("knee")) return true;
        if (department.toLowerCase().contains("paediatric endocrinologist") && (lower.contains("hormonal") || lower.contains("growth"))) return true;
        if (department.toLowerCase().contains("paediatrician") && lower.contains("child")) return true;
        if (department.toLowerCase().contains("physiotherapy") && lower.contains("physiotherapy")) return true;
        if (department.toLowerCase().contains("psychology") && lower.contains("stress")) return true;
        if (department.toLowerCase().contains("pulmonologist") && lower.contains("breathing")) return true;
        if (department.toLowerCase().contains("urologist") && lower.contains("urine")) return true;
        return false;
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