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

public class RXDXWhitefieldValidation {

    // Kanglish keywords for validation
    static String[] kanglishKeywords = {
        "bekku", "ide", "matte", "nange", "makka", "ge", "aagide", "madabeku",
        "nodabeku", "illa", "iddare", "aagi", "maadbeku", "anstide"
    };

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
                        + "/RXDX_Validation_Whitefield_Kanglish/Screenshots/"
                        + timestamp;

        String excelFolderPath =
                projectPath
                        + "/RXDX_Validation_Whitefield_Kanglish/Test_Excel";

        String excelFilePath =
                excelFolderPath
                        + "/RXDX_Whitefield_Kanglish_Results_"
                        + timestamp
                        + ".xlsx";

        new File(screenshotsPath).mkdirs();
        new File(excelFolderPath).mkdirs();

        System.out.println("\n📁 SCREENSHOTS will be saved in: " + screenshotsPath);
        System.out.println("📊 EXCEL FILE will be saved as: " + excelFilePath);

        // ================= DEPARTMENTS WITH KANGLISH QUESTIONS =================

        Map<String, String> departmentQuestions =
                new LinkedHashMap<>();

        departmentQuestions.put("Acupuncture", "ನನಗೆ ದೇಹದ ನೋವಿಗೆ ಆಕ್ಯುಪಂಕ್ಚರ್ ಥೆರಪಿ ಬೇಕು");
        departmentQuestions.put("Adult Cardiology", "ನನಗೆ ಎದೆನೋವು ಮತ್ತು ಹೃದಯ ಸಮಸ್ಯೆ ಇದೆ");
        departmentQuestions.put("Clinical Nutrition", "ನನಗೆ ತೂಕ ಇಳಿಸುವ ಡಯಟ್ ಸಲಹೆ ಬೇಕು");
        departmentQuestions.put("Dental", "ನನಗೆ ತೀವ್ರ ಹಲ್ಲಿನ ನೋವು ಇದೆ");
        departmentQuestions.put("Dermatology", "ನನಗೆ ಚರ್ಮದ ರ್ಯಾಶ್ ಮತ್ತು ಉರಿ ಇದೆ");
        departmentQuestions.put("Diabetologist", "ನನಗೆ ಮಧುಮೇಹ ಮತ್ತು ಸಕ್ಕರೆ ಮಟ್ಟ ಹೆಚ್ಚು ಇದೆ");
        departmentQuestions.put("Endocrinologist", "ನನಗೆ ಥೈರಾಯ್ಡ್ ಮತ್ತು ಹಾರ್ಮೋನ್ ಸಮಸ್ಯೆ ಇದೆ");
        departmentQuestions.put("Endodentist", "ನನಗೆ ರೂಟ್ ಕ್ಯಾನಲ್ ಚಿಕಿತ್ಸೆ ಬೇಕು");
        departmentQuestions.put("Endodontist", "ನನಗೆ ರೂಟ್ ಕ್ಯಾನಲ್ ಚಿಕಿತ್ಸೆ ಬೇಕು");
        departmentQuestions.put("Endovascular Neurosurgeon, Vascular Surgeon", "ನನಗೆ ರಕ್ತನಾಳ ಶಸ್ತ್ರಚಿಕಿತ್ಸೆ ಚಿಕಿತ್ಸೆ ಬೇಕು");
        departmentQuestions.put("ENT", "ನನಗೆ ಕಿವಿನೋವು ಮತ್ತು ಗಂಟಲು ಸೋಂಕು ಇದೆ");
        departmentQuestions.put("Gastroenterologist, Hepatologist", "ನನಗೆ ಹೊಟ್ಟೆನೋವು ಮತ್ತು ಯಕೃತ್ ಸಮಸ್ಯೆ ಇದೆ");
        departmentQuestions.put("General Medicine", "ನನಗೆ ದುರ್ಬಲತೆ, ಆಯಾಸ ಮತ್ತು ಅನಾರೋಗ್ಯದ ಅನುಭವವಾಗಿದೆ");
        departmentQuestions.put("General Surgeon", "ನನಗೆ ಶಸ್ತ್ರಚಿಕಿತ್ಸೆ ಸಲಹೆ ಬೇಕು");
        departmentQuestions.put("Homeopathy", "ನನಗೆ ಹೋಮಿಯೋಪತಿ ಚಿಕಿತ್ಸೆ ಬೇಕು");
        departmentQuestions.put("Internal Medicine", "ನನಗೆ ಜ್ವರ, ದುರ್ಬಲತೆ ಮತ್ತು ದೇಹದ ನೋವು ಇದೆ");
        departmentQuestions.put("Nephrologist", "ನನಗೆ ಮೂತ್ರಪಿಂಡದ ಕಲ್ಲಿನ ಸಮಸ್ಯೆ ಇದೆ");
        departmentQuestions.put("Neurologist", "ನನಗೆ ಮೈಗ್ರೇನ್ ಮತ್ತು ತೀವ್ರ ತಲೆನೋವು ಇದೆ");
        departmentQuestions.put("Neurosurgery", "ನನಗೆ ಮೆದುಳು ಮತ್ತು ಬೆನ್ನುಹುರಿ ಸಮಸ್ಯೆ ಇದೆ");
        departmentQuestions.put("Obs & Gynae", "ನನಗೆ ಗರ್ಭಧಾರಣೆ ಮತ್ತು ಮಾಸಿಕ ಧರ್ಮದ ಸಮಸ್ಯೆ ಇದೆ");
        departmentQuestions.put("Occupational Therapy", "ನನಗೆ ಆಕ್ಯುಪೇಷನಲ್ ಥೆರಪಿ ಸಹಾಯ ಬೇಕು");
        departmentQuestions.put("Ophthalmology", "ನನಗೆ ಮಸುಕಾದ ದೃಷ್ಟಿ ಮತ್ತು ಕಣ್ಣಿನ ನೋವು ಇದೆ");
        departmentQuestions.put("Oral Surgeon", "ನನಗೆ ಬಾಯಿಯ ಶಸ್ತ್ರಚಿಕಿತ್ಸೆ ಚಿಕಿತ್ಸೆ ಬೇಕು");
        departmentQuestions.put("Orthodentists", "ನನಗೆ ಬ್ರೇಸಸ್ ಮತ್ತು ಹಲ್ಲಿನ ಸರಿಹೊಂದಿಸುವ ಚಿಕಿತ್ಸೆ ಬೇಕು");
        departmentQuestions.put("Orthopedic Surgeon", "ನನಗೆ ಮೊಣಕಾಲು, ಸಂಧಿ ಮತ್ತು ಎಲುಬಿನ ನೋವು ಇದೆ");
        departmentQuestions.put("Orthotist", "ನನಗೆ ನಡೆಯಲು ಆರ್ಥೋಟಿಕ್ ಸಹಾಯ ಬೇಕು");
        departmentQuestions.put("Paediatric and Fetal Cardiologist", "ನನ್ನ ಮಗುವಿಗೆ ಹೃದಯ ಸಮಸ್ಯೆ ಇದೆ");
        departmentQuestions.put("Paediatric Cardiologist", "ನನ್ನ ಮಗುವಿಗೆ ಹೃದಯ ಸಮಸ್ಯೆ ಇದೆ");
        departmentQuestions.put("Paediatric Neurologist", "ನನ್ನ ಮಗುವಿಗೆ ಸೈಜರ್ ಮತ್ತು ನರ ಸಮಸ್ಯೆ ಇದೆ");
        departmentQuestions.put("Paediatric Surgeon", "ನನ್ನ ಮಗುವಿಗೆ ಶಸ್ತ್ರಚಿಕಿತ್ಸೆ ಬೇಕು");
        departmentQuestions.put("Paediatrician", "ನನ್ನ ಮಗುವಿಗೆ ಜ್ವರ ಮತ್ತು ಕೆಮ್ಮು ಇದೆ");
        departmentQuestions.put("Pain medicine", "ನನಗೆ ದೀರ್ಘಕಾಲದ ದೇಹದ ನೋವು ಇದೆ");
        departmentQuestions.put("Pedodentists", "ನನ್ನ ಮಗುವಿಗೆ ಹಲ್ಲಿನ ನೋವು ಇದೆ");
        departmentQuestions.put("Periodontist", "ನನಗೆ ಹಲ್ಲಿನ ಹಸಿ ಭಾಗದಲ್ಲಿ ಊತ ಮತ್ತು ರಕ್ತಸ್ರಾವ ಇದೆ");
        departmentQuestions.put("Physiotherapy", "ನನಗೆ ಸ್ನಾಯು ನೋವಿಗೆ ಫಿಸಿಯೋಥೆರಪಿ ಬೇಕು");
        departmentQuestions.put("Plastic Surgery", "ನನಗೆ ಕಾಸ್ಮೆಟಿಕ್ ಶಸ್ತ್ರಚಿಕಿತ್ಸೆ ಸಲಹೆ ಬೇಕು");
        departmentQuestions.put("Podiatric Surgeon", "ನನಗೆ ಪಾದದ ನೋವು ಮತ್ತು ನಡೆಯಲು ತೊಂದರೆ ಇದೆ");
        departmentQuestions.put("Psychiatry", "ನನಗೆ ಆತಂಕ, ಖಿನ್ನತೆ ಮತ್ತು ಮಾನಸಿಕ ಒತ್ತಡ ಇದೆ");
        departmentQuestions.put("Psychology", "ನನಗೆ ಮಾನಸಿಕ ಆರೋಗ್ಯ ಕೌನ್ಸೆಲಿಂಗ್ ಬೇಕು");
        departmentQuestions.put("Pulmonologist", "ನನಗೆ ಉಸಿರಾಟ ಮತ್ತು ಶ್ವಾಸಕೋಶದ ಸಮಸ್ಯೆ ಇದೆ");
        departmentQuestions.put("Radiology", "ನನಗೆ ಎಂಆರ್‌ಐ, ಸಿಟಿ ಸ್ಕ್ಯಾನ್ ಅಥವಾ ಎಕ್ಸ್‌ರೇ ಬೇಕು");
        departmentQuestions.put("Rheumatology", "ನನಗೆ ಆರ್ಥ್ರೈಟಿಸ್ ಮತ್ತು ಸಂಧಿ ನೋವು ಇದೆ");
        departmentQuestions.put("Speech Therapy", "ನನಗೆ ಮಾತನಾಡುವ ಮತ್ತು ಸಂವಹನದ ಸಮಸ್ಯೆ ಇದೆ");
        departmentQuestions.put("Urologist", "ನನಗೆ ಮೂತ್ರದ ಸೋಂಕು ಮತ್ತು ಮೂತ್ರಾಶಯದ ನೋವು ಇದೆ");
        // ================= VALID DEPARTMENTS SET =================
        Set<String> validDepartments = new HashSet<>(departmentQuestions.keySet());

        // ================= CREATE EXCEL WORKBOOK WITH HEADERS =================
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("RXDX_Whitefield_Kanglish_Results");
        
        addProjectHeader(sheet, workbook, "RXDX WHITEFIELD HOSPITAL - KANGLISH VALIDATION TESTS");
        
        int headerRowIndex = addColumnHeaders(sheet, workbook);

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
        System.out.println("     KANGLISH VALIDATION TESTS STARTING - TOTAL TEST CASES: " + totalTests);
        System.out.println("     Using 5 seconds wait time between messages");
        System.out.println("=".repeat(80));

        // ================= TEST LOOP =================

        int testNumber = 1;
        int excelRowNum = headerRowIndex + 1;

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
            int kanglishWordCount = 0;

            System.out.println("\n" + "=".repeat(60));
            System.out.println("TEST #" + testNumber + " of " + totalTests);
            System.out.println("Department: " + expectedDepartment);
            System.out.println("Kanglish Question: " + kanglishQuestion);
            System.out.println("=".repeat(60));

            try {

                // ================= STEP 1: Send "hi" =================
                System.out.println("\n📤 STEP 1: Sending 'hi' to start...");
                sendMessage(driver, "hi", wait);
                System.out.println("⏳ Waiting 5 seconds for response to load...");
                Thread.sleep(5000);
                
                botResponse = getLastMessageReliable(driver);
                System.out.println("✅ Response received after 'hi'");

                // ================= STEP 2: Send Kanglish department question =================
                System.out.println("\n📤 STEP 2: Sending Kanglish question: " + kanglishQuestion);
                sendMessage(driver, kanglishQuestion, wait);
                System.out.println("⏳ Waiting 5 seconds for response to load...");
                Thread.sleep(5000);
                
                botResponse = getLastMessageReliable(driver);
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

                // ================= STEP 3: Click Book Appointment button =================
                System.out.println("\n🔍 STEP 3: Looking for 'Book an Appointment' button...");
                
                List<WebElement> buttons = driver.findElements(By.xpath("//span[contains(text(),'Book an Appointment')]"));

                if (!buttons.isEmpty()) {
                    buttons.get(buttons.size() - 1).click();
                    System.out.println("✅ Book Appointment Clicked");
                    System.out.println("⏳ Waiting 5 seconds for response to load...");
                    Thread.sleep(5000);
                    
                } else {
                    status = "FAIL";
                    failureType = "MISSING_BOOK_APPOINTMENT_BUTTON";
                    failureReason = "Book Appointment button missing";
                    System.out.println("❌ Book Appointment button not found!");
                }

                // ================= STEP 4: Send "1" for place selection =================
                System.out.println("\n📤 STEP 4: Sending '1' for Whitefield location selection...");
                sendMessage(driver, "1", wait);
                System.out.println("⏳ Waiting 5 seconds for response to load...");
                Thread.sleep(5000);
                
                placeResponse = getLastMessageReliable(driver);
                System.out.println("\nPLACE RESPONSE:\n" + placeResponse);

                // ================= VALIDATION 2: Department filtering =================
                System.out.println("\n🔍 VALIDATION 2: Validating department filtering...");
                
                allDepartmentsShown = extractAllDepartmentsFromResponse(placeResponse, validDepartments);
                extraDepartments = findExtraDepartments(allDepartmentsShown, expectedDepartment);
                
                if (!placeResponse.toLowerCase().contains(expectedDepartment.toLowerCase())) {
                    status = "FAIL";
                    failureType = "DEPARTMENT_NOT_FOUND";
                    failureReason = "Expected department '" + expectedDepartment + "' not displayed in place response";
                } 
                else if (!extraDepartments.isEmpty()) {
                    status = "FAIL";
                    failureType = "FILTERING_FAILED_EXTRA_DEPARTMENTS";
                    failureReason = "FILTERING VALIDATION FAILED: Bot showed extra departments not requested. " +
                                   "Patient asked for: " + expectedDepartment + ". Extra departments shown: " + extraDepartments;
                    System.out.println("❌ FILTERING FAILED! Extra departments found: " + extraDepartments);
                } else {
                    System.out.println("✅ FILTERING VALIDATION PASSED! Only showing: " + expectedDepartment);
                }
                
                if (status.equals("PASS")) {
                    System.out.println("\n✅ All validations passed for department: " + expectedDepartment);
                    System.out.println("   ✓ Kanglish response detected");
                    System.out.println("   ✓ Correct department filtered");
                }

                if (status.equals("FAIL")) {
                    screenshotPath = takeScreenshot(driver, screenshotsPath, expectedDepartment);
                }

            } catch (Exception e) {
                status = "FAIL";
                failureType = "EXCEPTION";
                failureReason = "Exception: " + e.getMessage();
                e.printStackTrace();
                screenshotPath = takeScreenshot(driver, screenshotsPath, expectedDepartment);
            }

            addTestResultToExcel(sheet, excelRowNum++, testNumber, expectedDepartment, kanglishQuestion,
                                 botResponse, placeResponse, failureType, failureReason, extraDepartments,
                                 allDepartmentsShown, screenshotPath, status, hasKanglish, kanglishWordCount);

            System.out.println("\nSTATUS: " + status);
            System.out.println("   Kanglish Response: " + (hasKanglish ? "YES" : "NO"));
            System.out.println("   Kanglish Word Count: " + kanglishWordCount);
            
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
            Thread.sleep(3000);
        }

        for (int i = 0; i <= 12; i++) {
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
    
    // ================= RELIABLE GET LAST MESSAGE WITH MULTIPLE XPATHS =================
    
    static String getLastMessageReliable(WebDriver driver) {
        String[] xpaths = {
            "//div[contains(@class,'message-in')]//div[contains(@class,'copyable-text')]//span[@dir='ltr']",
            "//div[contains(@class,'message-in')]//div[contains(@class,'selectable-text')]",
            "//div[contains(@class,'message-in')]",
            "//div[@data-testid='msg-container']",
            "//div[contains(@class,'_akbu')]",
            "//div[contains(@class,'focusable-list-item')]//div[contains(@class,'copyable-text')]",
            "//div[contains(@class,'web')]//div[contains(@class,'message')]"
        };
        
        for (String xpath : xpaths) {
            try {
                List<WebElement> messages = driver.findElements(By.xpath(xpath));
                if (!messages.isEmpty()) {
                    String text = messages.get(messages.size() - 1).getText();
                    if (text != null && !text.isEmpty() && !text.equals("NO RESPONSE")) {
                        return text;
                    }
                }
            } catch (Exception e) {
                // Try next xpath
            }
        }
        
        return "NO RESPONSE";
    }
    
    // ================= EXCEL METHODS =================
    
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
        sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 12));
        
        Row infoRow = sheet.createRow(1);
        Cell infoCell = infoRow.createCell(0);
        infoCell.setCellValue("Project: RXDX Whitefield Kanglish Validation | Testing: Department-wise Appointment Booking | Total Departments: 45 | Wait Time: 5 Seconds | Language: Kanglish | Date: " + 
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
        sheet.addMergedRegion(new CellRangeAddress(1, 1, 0, 12));
    }
    
    static int addColumnHeaders(Sheet sheet, Workbook workbook) {
        Row columnHeaderRow = sheet.createRow(2);
        
        String[] columns = {
            "Test No", "Department", "Kanglish Question", "Bot Response", "Place Response",
            "Status", "Failure Type", "Failure Reason", "Extra Departments Found",
            "All Departments Shown", "Kanglish Response", "Kanglish Word Count", "Screenshot", "Timestamp"
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
                                      String question, String botResponse, String placeResponse, 
                                      String failureType, String failureReason, String extraDepartments,
                                      String allDepartmentsShown, String screenshotPath, String status,
                                      boolean hasKanglish, int kanglishWordCount) {
        
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
        row.createCell(11).setCellValue(kanglishWordCount);
        row.createCell(12).setCellValue(screenshotPath != null ? screenshotPath : "No screenshot");
        row.createCell(13).setCellValue(new SimpleDateFormat("yyyy-MM-dd HH:mm:ss").format(new Date()));
    }

    // ================= SEND MESSAGE =================

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

    // ================= SCREENSHOT =================

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
    
    // ================= RAJAGIRI STYLE: EXTRACT ALL DEPARTMENTS FROM RESPONSE =================
    
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
    
    // ================= RAJAGIRI STYLE: FIND EXTRA DEPARTMENTS =================
    
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
    
    // ================= PRINT FINAL REPORT =================
    
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