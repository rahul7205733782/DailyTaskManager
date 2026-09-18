package Kanglish_Language;

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

public class Kang_Kadugodi {

    public static void main(String[] args)
            throws InterruptedException, IOException {

        // ================= DISABLE LOGS =================

        System.setProperty(
                "org.apache.poi.util.POILogger",
                "org.apache.poi.util.NullLogger"
        );

        // ================= PROJECT PATH =================

        String projectPath =
                System.getProperty("user.dir");

        String timestamp =
                new SimpleDateFormat(
                        "yyyyMMdd_HHmmss"
                ).format(new Date());

        // ================= FOLDERS =================

        String screenshotsPath =
                projectPath
                        + "/RXDX_Kadugodi_Kanglish/Screenshots/"
                        + timestamp;

        String excelFolderPath =
                projectPath
                        + "/RXDX_Kadugodi_Kanglish/Test_Excel";

        String excelFilePath =
                excelFolderPath
                        + "/RXDX_Kadugodi_Kanglish_Results_"
                        + timestamp
                        + ".xlsx";

        new File(screenshotsPath).mkdirs();
        new File(excelFolderPath).mkdirs();

        System.out.println("\n📁 SCREENSHOTS will be saved in: " + screenshotsPath);
        System.out.println("📊 EXCEL FILE will be saved as: " + excelFilePath);

        // ================= DEPARTMENTS WITH KANGLISH QUESTIONS =================
        // Questions converted to Kanglish (Romanized Kannada)

        Map<String, String> departmentQuestions =
                new LinkedHashMap<>();

        // Kanglish questions (matching the format from your screenshot)
        departmentQuestions.put("Adult Cardiology", "Nange chest pain matte heart problem ide");
        departmentQuestions.put("Dermatology", "Nange skin rashes matte itching ide");
        departmentQuestions.put("ENT", "Nange ear pain matte throat infection ide");
        departmentQuestions.put("General Medicine", "Nange weak tired matte sick anstide");
        departmentQuestions.put("Homeopathy", "Nange homeopathy treatment bekku");
        departmentQuestions.put("Nephrologist", "Nange kidney stone problem ide");
        departmentQuestions.put("Obs & Gynae", "Nange pregnancy matte period problem ide");
        departmentQuestions.put("Ophthalmology", "Nange blurry vision matte eye pain ide");
        departmentQuestions.put("Orthopedician", "Nange knee joint matte bone pain ide");
        departmentQuestions.put("Paediatrician", "Nange makka ge fever matte cough ide");
        departmentQuestions.put("Physiotherapy", "Nange muscle pain ge physiotherapy bekku");
        departmentQuestions.put("Psychiatry", "Nange anxiety depression matte mental stress ide");

        // ================= KANGLISH KEYWORDS FOR VALIDATION =================
        String[] kanglishKeywords = {
            "bekku", "ide", "matte", "nange", "makka", "ge", "aagide", "madabeku",
            "nodabeku", "iddare", "aagi", "maadbeku", "anstide", "illa"
        };

        // Department-specific Kanglish keywords
        Map<String, List<String>> departmentKanglishKeywords = new LinkedHashMap<>();
        departmentKanglishKeywords.put("Adult Cardiology", Arrays.asList("chest pain", "heart", "problem ide"));
        departmentKanglishKeywords.put("Dermatology", Arrays.asList("skin rashes", "itching", "ide"));
        departmentKanglishKeywords.put("ENT", Arrays.asList("ear pain", "throat infection", "ide"));
        departmentKanglishKeywords.put("General Medicine", Arrays.asList("weak", "tired", "sick anstide"));
        departmentKanglishKeywords.put("Homeopathy", Arrays.asList("homeopathy", "treatment", "bekku"));
        departmentKanglishKeywords.put("Nephrologist", Arrays.asList("kidney stone", "problem ide"));
        departmentKanglishKeywords.put("Obs & Gynae", Arrays.asList("pregnancy", "period", "problem ide"));
        departmentKanglishKeywords.put("Ophthalmology", Arrays.asList("blurry vision", "eye pain", "ide"));
        departmentKanglishKeywords.put("Orthopedician", Arrays.asList("knee joint", "bone pain", "ide"));
        departmentKanglishKeywords.put("Paediatrician", Arrays.asList("makka ge", "fever", "cough", "ide"));
        departmentKanglishKeywords.put("Physiotherapy", Arrays.asList("muscle pain", "physiotherapy", "bekku"));
        departmentKanglishKeywords.put("Psychiatry", Arrays.asList("anxiety", "depression", "mental stress", "ide"));

        // ================= VALID DEPARTMENTS SET =================
        Set<String> validDepartments = new HashSet<>(departmentQuestions.keySet());

        // ================= CREATE EXCEL WORKBOOK WITH HEADERS =================

        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("RXDX_Kadugodi_Kanglish_Results");
        
        addProjectHeader(sheet, workbook, "RXDX KADUGODI - KANGLISH RESPONSE VALIDATION TESTS");
        
        int headerRowIndex = addColumnHeaders(sheet, workbook);
        int excelRowNum = headerRowIndex + 1;

        // ================= DRIVER =================

        EdgeOptions options = new EdgeOptions();
        options.addArguments("--user-data-dir=C:\\EdgeAutomationProfile");
        options.addArguments("--start-maximized");

        WebDriver driver = new EdgeDriver(options);
        WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(30));

        driver.manage().window().maximize();
        driver.get("https://web.whatsapp.com/");

        System.out.println("\n⚠️ PLEASE SCAN QR CODE...");
        Thread.sleep(20000);

        // ================= OPEN RXDX CHAT =================

        driver.findElement(By.xpath("//span[@data-icon='new-chat-outline']")).click();
        driver.findElement(By.xpath("//input[@aria-label='Search name or number']")).sendKeys("916363415530");
        Thread.sleep(4000);
        driver.findElement(By.xpath("//span[@title='+91 63634 15530']")).click();
        Thread.sleep(4000);

        // ================= CALCULATE TOTAL TESTS =================
        int totalTests = departmentQuestions.size();
        List<Map<String, String>> failedTestsData = new ArrayList<>();

        System.out.println("\n" + "=".repeat(80));
        System.out.println("     KANGLISH RESPONSE VALIDATION TESTS STARTING - TOTAL TEST CASES: " + totalTests);
        System.out.println("     BRANCH: Kadugodi (Option 2)");
        System.out.println("=".repeat(80));

        // ================= TEST LOOP =================

        int testNumber = 1;

        for (Map.Entry<String, String> entry : departmentQuestions.entrySet()) {

            String expectedDepartment = entry.getKey();
            String kanglishQuestion = entry.getValue();

            String botResponse = "";
            String placeResponse = "";

            String status = "PASS";
            String failureType = "";
            String failureReason = "";

            String screenshotPath = "";
            String extraDepartments = "";
            String allDepartmentsShown = "";
            
            boolean hasKanglish = false;
            boolean hasDepartmentKeywords = false;
            int kanglishWordCount = 0;

            System.out.println("\n" + "=".repeat(60));
            System.out.println("TEST #" + testNumber + " of " + totalTests);
            System.out.println("Department: " + expectedDepartment);
            System.out.println("Kanglish Question: " + kanglishQuestion);
            System.out.println("=".repeat(60));

            try {

                // ================= STEP 1: Send "hi" =================
                System.out.println("\n📤 STEP 1: Sending 'hi'...");
                String initialLastMessage = getLastMessage(driver);
                sendMessage(driver, "hi", wait);
                
                System.out.println("⏳ Waiting for response...");
                boolean responseReceived = waitForNewMessage(driver, initialLastMessage, 300000);
                
                if (!responseReceived) {
                    status = "FAIL";
                    failureType = "NO_RESPONSE_AFTER_HI";
                    botResponse = "NO RESPONSE FROM BOT";
                    failureReason = "Bot did not respond after sending 'hi' within 5 minutes";
                    screenshotPath = takeScreenshot(driver, screenshotsPath, expectedDepartment + "_NoResponse");
                    addTestResultToExcel(sheet, excelRowNum++, testNumber, expectedDepartment, kanglishQuestion,
                                botResponse, placeResponse, failureType, failureReason, extraDepartments,
                                allDepartmentsShown, screenshotPath, status, hasKanglish, hasDepartmentKeywords, kanglishWordCount);
                    testNumber++;
                    continue;
                }
                
                Thread.sleep(2000);
                botResponse = getLastMessage(driver);
                System.out.println("✅ Bot responded to 'hi'");

                // ================= STEP 2: Send Kanglish question =================
                System.out.println("\n📤 STEP 2: Sending Kanglish question: " + kanglishQuestion);
                initialLastMessage = getLastMessage(driver);
                sendMessage(driver, kanglishQuestion, wait);
                
                System.out.println("⏳ Waiting for Kanglish response...");
                responseReceived = waitForNewMessage(driver, initialLastMessage, 300000);
                
                if (!responseReceived) {
                    status = "FAIL";
                    failureType = "NO_RESPONSE_TO_QUESTION";
                    botResponse = "NO RESPONSE FROM BOT";
                    failureReason = "Bot did not respond to Kanglish question within 5 minutes";
                    screenshotPath = takeScreenshot(driver, screenshotsPath, expectedDepartment + "_NoResponse");
                    addTestResultToExcel(sheet, excelRowNum++, testNumber, expectedDepartment, kanglishQuestion,
                                botResponse, placeResponse, failureType, failureReason, extraDepartments,
                                allDepartmentsShown, screenshotPath, status, hasKanglish, hasDepartmentKeywords, kanglishWordCount);
                    testNumber++;
                    continue;
                }
                
                Thread.sleep(2000);
                botResponse = getLastMessage(driver);
                System.out.println("\nBOT RESPONSE:\n" + botResponse);

                // ================= VALIDATION 1: Check if response contains Kanglish words =================
                System.out.println("\n🔍 VALIDATION 1: Checking if bot responded in Kanglish...");
                
                String lowerResponse = botResponse.toLowerCase();
                for (String keyword : kanglishKeywords) {
                    if (lowerResponse.contains(keyword)) {
                        hasKanglish = true;
                        kanglishWordCount++;
                    }
                }
                
                if (!hasKanglish) {
                    status = "FAIL";
                    failureType = "NO_KANGLISH_RESPONSE";
                    failureReason = "Bot responded in English, not in Kanglish. Expected Kanglish (Romanized Kannada) response.";
                    System.out.println("❌ FAIL: No Kanglish words found in response!");
                } else {
                    System.out.println("✅ PASS: Bot responded in Kanglish (Found " + kanglishWordCount + " Kanglish word(s))");
                }

                // ================= VALIDATION 2: Check for department-specific keywords =================
                System.out.println("\n🔍 VALIDATION 2: Checking department-specific keywords...");
                
                List<String> deptKeywords = departmentKanglishKeywords.get(expectedDepartment);
                if (deptKeywords != null) {
                    for (String keyword : deptKeywords) {
                        if (lowerResponse.contains(keyword.toLowerCase())) {
                            hasDepartmentKeywords = true;
                            System.out.println("✅ Found keyword: '" + keyword + "'");
                            break;
                        }
                    }
                }
                
                if (!hasDepartmentKeywords && deptKeywords != null) {
                    if (status.equals("PASS")) {
                        status = "FAIL";
                        failureType = "DEPARTMENT_KEYWORDS_NOT_FOUND";
                        failureReason = "Department-specific Kanglish keywords not found in response.";
                    } else {
                        failureReason = failureReason + " Department-specific keywords not found.";
                    }
                    System.out.println("❌ FAIL: Department-specific keywords not found!");
                } else if (hasDepartmentKeywords) {
                    System.out.println("✅ PASS: Department-specific keywords found.");
                }

                // ================= STEP 3: Check for Book Appointment button =================
                System.out.println("\n🔍 STEP 3: Looking for 'Book an Appointment' button...");
                List<WebElement> buttons = driver.findElements(By.xpath("//span[contains(text(),'Book an Appointment')]"));

                if (!buttons.isEmpty()) {
                    initialLastMessage = getLastMessage(driver);
                    buttons.get(buttons.size() - 1).click();
                    System.out.println("✅ Book Appointment Clicked");
                    
                    responseReceived = waitForNewMessage(driver, initialLastMessage, 300000);
                    if (!responseReceived) {
                        status = "FAIL";
                        failureType = "NO_RESPONSE_AFTER_BOOK_APPOINTMENT";
                        failureReason = "Bot did not respond after clicking button";
                        screenshotPath = takeScreenshot(driver, screenshotsPath, expectedDepartment + "_NoBookAppResponse");
                        addTestResultToExcel(sheet, excelRowNum++, testNumber, expectedDepartment, kanglishQuestion,
                                    botResponse, placeResponse, failureType, failureReason, extraDepartments,
                                    allDepartmentsShown, screenshotPath, status, hasKanglish, hasDepartmentKeywords, kanglishWordCount);
                        testNumber++;
                        continue;
                    }
                    Thread.sleep(2000);
                } else {
                    status = "FAIL";
                    failureType = "MISSING_BOOK_APPOINTMENT_BUTTON";
                    failureReason = "Book Appointment button missing";
                }

                // ================= STEP 4: Send "2" for Kadugodi =================
                System.out.println("\n📤 STEP 4: Sending '2' for Kadugodi...");
                initialLastMessage = getLastMessage(driver);
                sendMessage(driver, "2", wait);
                
                responseReceived = waitForNewMessage(driver, initialLastMessage, 300000);
                if (!responseReceived) {
                    status = "FAIL";
                    failureType = "NO_RESPONSE_TO_LOCATION";
                    placeResponse = "NO RESPONSE";
                    failureReason = "Bot did not respond to location selection";
                    screenshotPath = takeScreenshot(driver, screenshotsPath, expectedDepartment + "_NoLocationResponse");
                    addTestResultToExcel(sheet, excelRowNum++, testNumber, expectedDepartment, kanglishQuestion,
                                botResponse, placeResponse, failureType, failureReason, extraDepartments,
                                allDepartmentsShown, screenshotPath, status, hasKanglish, hasDepartmentKeywords, kanglishWordCount);
                    testNumber++;
                    continue;
                }
                Thread.sleep(2000);
                placeResponse = getLastMessage(driver);
                System.out.println("\nLOCATION RESPONSE:\n" + placeResponse);

                // ================= VALIDATION 3: Department filtering =================
                allDepartmentsShown = extractAllDepartmentsFromResponse(placeResponse, validDepartments);
                extraDepartments = findExtraDepartments(allDepartmentsShown, expectedDepartment);
                
                if (!placeResponse.toLowerCase().contains(expectedDepartment.toLowerCase())) {
                    status = "FAIL";
                    failureType = "DEPARTMENT_NOT_FOUND";
                    failureReason = "Expected department not found in response";
                } else if (!extraDepartments.isEmpty()) {
                    status = "FAIL";
                    failureType = "EXTRA_DEPARTMENTS_FOUND";
                    failureReason = "Extra departments found: " + extraDepartments;
                }

                if (status.equals("PASS")) {
                    System.out.println("\n✅✅✅ ALL VALIDATIONS PASSED for: " + expectedDepartment);
                    System.out.println("   ✓ Bot responded in Kanglish");
                    System.out.println("   ✓ Department keywords found");
                    System.out.println("   ✓ Correct department filtered");
                } else {
                    System.out.println("\n❌❌❌ TEST FAILED: " + failureReason);
                }

                if (status.equals("FAIL")) {
                    screenshotPath = takeScreenshot(driver, screenshotsPath, expectedDepartment);
                }

            } catch (Exception e) {
                status = "FAIL";
                failureType = "EXCEPTION";
                failureReason = e.getMessage();
                screenshotPath = takeScreenshot(driver, screenshotsPath, expectedDepartment);
                e.printStackTrace();
            }

            addTestResultToExcel(sheet, excelRowNum++, testNumber, expectedDepartment, kanglishQuestion,
                        botResponse, placeResponse, failureType, failureReason, extraDepartments,
                        allDepartmentsShown, screenshotPath, status, hasKanglish, hasDepartmentKeywords, kanglishWordCount);

            System.out.println("\n📊 STATUS: " + status);
            
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
            Thread.sleep(5000);
        }

        for (int i = 0; i <= 13; i++) {
            sheet.autoSizeColumn(i);
        }

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
            if (Desktop.isDesktopSupported()) {
                Desktop.getDesktop().open(excelFile);
                System.out.println("\n✅ Excel Opened Automatically");
            }
        } catch (Exception e) {
            System.out.println("❌ Unable to open Excel automatically");
        }

        driver.quit();
    }
    
    static boolean waitForNewMessage(WebDriver driver, String previousLastMessage, long timeoutMillis) throws InterruptedException {
        long startTime = System.currentTimeMillis();
        int secondsWaited = 0;
        
        System.out.println();
        System.out.println("⏰ Maximum wait time: " + (timeoutMillis / 1000) + " seconds (5 minutes)");
        System.out.print("⏳ Waiting for response");
        
        while (System.currentTimeMillis() - startTime < timeoutMillis) {
            Thread.sleep(1000);
            secondsWaited++;
            String currentMessage = getLastMessage(driver);
            
            System.out.print(".");
            if (secondsWaited % 30 == 0) {
                System.out.println();
                System.out.print("   Still waiting... " + secondsWaited + " seconds elapsed");
            }
            
            if (!currentMessage.equals(previousLastMessage) && 
                !currentMessage.equals("NO RESPONSE") &&
                currentMessage.length() > 5) {
                System.out.println("\n✅ New message received after " + secondsWaited + " seconds");
                return true;
            }
        }
        
        System.out.println("\n⚠️ Timeout after " + (timeoutMillis / 1000) + " seconds - no new message");
        return false;
    }
    
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
        infoCell.setCellValue("Project: RXDX Kadugodi Kanglish Response Validation | Location: Kadugodi (Option 2) | Departments: 12 | Max Wait: 5 Minutes | Validating Kanglish Responses | Date: " + 
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
            "Test No", "Department", "Kanglish Question", "Bot Response", "Location Response",
            "Status", "Failure Type", "Failure Reason", "Extra Departments Found",
            "All Departments Shown", "Kanglish Response", "Dept Keywords Found",
            "Kanglish Word Count", "Screenshot", "Timestamp"
        };
        
        CellStyle headerStyle = workbook.createCellStyle();
        Font headerFont = workbook.createFont();
        headerFont.setBold(true);
        headerFont.setFontHeightInPoints((short) 9);
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
                                      String allDepartmentsShown, String screenshotPath, String status,
                                      boolean hasKanglish, boolean hasDepartmentKeywords, int kanglishWordCount) {
        
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
        if (status.equals("PASS")) {
            statusCell.setCellStyle(passStyle);
        } else {
            statusCell.setCellStyle(failStyle);
        }
        
        row.createCell(6).setCellValue(failureType != null ? failureType : "N/A");
        
        Cell failureCell = row.createCell(7);
        failureCell.setCellValue(failureReason != null ? failureReason : "N/A");
        failureCell.setCellStyle(wrapStyle);
        
        row.createCell(8).setCellValue(extraDepartments != null && !extraDepartments.isEmpty() ? extraDepartments : "None");
        row.createCell(9).setCellValue(allDepartmentsShown != null ? allDepartmentsShown : "None");
        row.createCell(10).setCellValue(hasKanglish ? "YES (PASS)" : "NO (FAIL)");
        row.createCell(11).setCellValue(hasDepartmentKeywords ? "YES (PASS)" : "NO (FAIL)");
        row.createCell(12).setCellValue(kanglishWordCount);
        row.createCell(13).setCellValue(screenshotPath != null ? screenshotPath : "No screenshot");
        row.createCell(14).setCellValue(new SimpleDateFormat("yyyy-MM-dd HH:mm:ss").format(new Date()));
    }

    static void sendMessage(WebDriver driver, String text, WebDriverWait wait) throws InterruptedException {
        WebElement input = wait.until(
                ExpectedConditions.elementToBeClickable(
                        By.xpath("//div[@aria-placeholder='Type a message']")
                )
        );
        input.clear();
        input.sendKeys(text);
        Thread.sleep(1000);
        
        WebElement sendBtn = wait.until(
                ExpectedConditions.elementToBeClickable(
                        By.xpath("//span[@data-icon='wds-ic-send-filled']")
                )
        );
        sendBtn.click();
        Thread.sleep(2000);
    }

    static String getLastMessage(WebDriver driver) {
        try {
            List<WebElement> messages = driver.findElements(
                    By.xpath("//div[contains(@class,'message-in')]")
            );
            if (!messages.isEmpty()) {
                return messages.get(messages.size() - 1).getText();
            }
        } catch (Exception e) {
            return "NO RESPONSE";
        }
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
        
        for (String dept : departments) {
            if (!dept.equalsIgnoreCase(expectedDepartment)) {
                extraDepts.add(dept);
            }
        }
        
        return extraDepts.isEmpty() ? "" : String.join(", ", extraDepts);
    }

    static void printFinalReport(int totalTests, List<Map<String, String>> failedTestsData) {
        System.out.println("\n" + "=".repeat(80));
        System.out.println("                     FINAL TEST REPORT");
        System.out.println("=".repeat(80));
        
        if (failedTestsData.isEmpty()) {
            System.out.println("\n🎉 ALL TESTS PASSED!");
            System.out.println("✅ Total Tests: " + totalTests);
            System.out.println("✅ Success Rate: 100%");
            System.out.println("✅ Bot is responding in Kanglish correctly!");
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