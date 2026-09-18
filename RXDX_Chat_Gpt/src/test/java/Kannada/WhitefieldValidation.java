package Kannada;

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
import java.util.regex.Pattern;

public class WhitefieldValidation {

    // Kannada Unicode range: \u0C80-\u0CFF
    static final String KANNADA_UNICODE_RANGE = "[\\u0C80-\\u0CFF]";
    static final Pattern KANNADA_PATTERN = Pattern.compile(".*" + KANNADA_UNICODE_RANGE + ".*");

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
                        + "/RXDX_Validation_Whitefield_Kannada/Screenshots/"
                        + timestamp;

        String excelFolderPath =
                projectPath
                        + "/RXDX_Validation_Whitefield_Kannada/Test_Excel";

        String excelFilePath =
                excelFolderPath
                        + "/RXDX_Whitefield_Kannada_Results_"
                        + timestamp
                        + ".xlsx";

        new File(screenshotsPath).mkdirs();
        new File(excelFolderPath).mkdirs();

        System.out.println("\n📁 SCREENSHOTS will be saved in: " + screenshotsPath);
        System.out.println("📊 EXCEL FILE will be saved as: " + excelFilePath);

        // ================= DEPARTMENTS WITH KANNADA QUESTIONS =================
        // Questions in Kannada script

        Map<String, String> departmentQuestions =
                new LinkedHashMap<>();

        // Kannada questions
        departmentQuestions.put("Acupuncture", "ನನಗೆ ದೇಹದ ನೋವಿಗೆ ಅಕ್ಯುಪಂಕ್ಚರ್ ಚಿಕಿತ್ಸೆ ಬೇಕು");
        departmentQuestions.put("Adult Cardiology", "ನನಗೆ ಎದೆ ನೋವು ಮತ್ತು ಹೃದಯ ಸಮಸ್ಯೆ ಇದೆ");
        departmentQuestions.put("Allergy Specialist", "ನನಗೆ ತೀವ್ರ ಅಲರ್ಜಿ ಮತ್ತು ಚರ್ಮದ ಪ್ರತಿಕ್ರಿಯೆ ಇದೆ");
        departmentQuestions.put("Clinical Nutrition", "ನನಗೆ ತೂಕ ಇಳಿಕೆ ಆಹಾರ ಸಮಾಲೋಚನೆ ಬೇಕು");
        departmentQuestions.put("Dental", "ನನಗೆ ತೀವ್ರ ಹಲ್ಲು ನೋವು ಇದೆ");
        departmentQuestions.put("Dermatology", "ನನಗೆ ಚರ್ಮದ ದದ್ದು ಮತ್ತು ತುರಿಕೆ ಇದೆ");
        departmentQuestions.put("Diabetologist", "ನನಗೆ ಮಧುಮೇಹ ಸಕ್ಕರೆ ಮಟ್ಟ ಹೆಚ್ಚಿದೆ");
        departmentQuestions.put("Endocrinologist", "ನನಗೆ ಥೈರಾಯ್ಡ್ ಹಾರ್ಮೋನ್ ಸಮಸ್ಯೆ ಇದೆ");
        departmentQuestions.put("Endodentist", "ನನಗೆ ರೂಟ್ ಕೆನಾಲ್ ಚಿಕಿತ್ಸೆ ಬೇಕು");
        departmentQuestions.put("Endodontist", "ನನಗೆ ರೂಟ್ ಕೆನಾಲ್ ಚಿಕಿತ್ಸೆ ಬೇಕು");
        departmentQuestions.put("Endovascular Neurosurgeon, Vascular Surgeon", "ನನಗೆ ರಕ್ತನಾಳ ಶಸ್ತ್ರಚಿಕಿತ್ಸೆ ಚಿಕಿತ್ಸೆ ಬೇಕು");
        departmentQuestions.put("ENT", "ನನಗೆ ಕಿವಿ ನೋವು ಮತ್ತು ಗಂಟಲು ಸೋಂಕು ಇದೆ");
        departmentQuestions.put("Gastroenterologist, Hepatologist", "ನನಗೆ ಹೊಟ್ಟೆ ನೋವು ಮತ್ತು ಯಕೃತ್ತಿನ ಸಮಸ್ಯೆ ಇದೆ");
        departmentQuestions.put("General Medicine", "ನನಗೆ ದುರ್ಬಲ, ಆಯಾಸ ಮತ್ತು ಅನಾರೋಗ್ಯ ಅನಿಸುತ್ತಿದೆ");
        departmentQuestions.put("General Surgeon", "ನನಗೆ ಶಸ್ತ್ರಚಿಕಿತ್ಸೆ ಸಮಾಲೋಚನೆ ಬೇಕು");
        departmentQuestions.put("Homeopathy", "ನನಗೆ ಹೋಮಿಯೋಪತಿ ಚಿಕಿತ್ಸೆ ಬೇಕು");
        departmentQuestions.put("Internal Medicine", "ನನಗೆ ಜ್ವರ, ದೌರ್ಬಲ್ಯ ಮತ್ತು ದೇಹದ ನೋವು ಇದೆ");
        departmentQuestions.put("Nephrologist", "ನನಗೆ ಮೂತ್ರಪಿಂಡದ ಕಲ್ಲಿನ ಸಮಸ್ಯೆ ಇದೆ");
        departmentQuestions.put("Neurologist", "ನನಗೆ ಮೈಗ್ರೇನ್ ಮತ್ತು ತೀವ್ರ ತಲೆನೋವು ಇದೆ");
        departmentQuestions.put("Neurosurgery", "ನನಗೆ ಬೆನ್ನುಮೂಳೆ ಮತ್ತು ಮೆದುಳಿನ ಸಮಸ್ಯೆ ಇದೆ");
        departmentQuestions.put("Obs & Gynae", "ನನಗೆ ಗರ್ಭಧಾರಣೆ ಮತ್ತು ಋತುಚಕ್ರದ ಸಮಸ್ಯೆ ಇದೆ");
        departmentQuestions.put("Occupational Therapy", "ನನಗೆ ಔದ್ಯೋಗಿಕ ಚಿಕಿತ್ಸೆ ಬೆಂಬಲ ಬೇಕು");
        departmentQuestions.put("Ophthalmology", "ನನಗೆ ಮಸುಕಾದ ದೃಷ್ಟಿ ಮತ್ತು ಕಣ್ಣಿನ ನೋವು ಇದೆ");
        departmentQuestions.put("Oral Surgeon", "ನನಗೆ ಮೌಖಿಕ ಶಸ್ತ್ರಚಿಕಿತ್ಸೆ ಚಿಕಿತ್ಸೆ ಬೇಕು");
        departmentQuestions.put("Orthodentists", "ನನಗೆ ಬ್ರೇಸ್ ಮತ್ತು ಹಲ್ಲು ಜೋಡಣೆ ಚಿಕಿತ್ಸೆ ಬೇಕು");
        departmentQuestions.put("Orthopedic Surgeon", "ನನಗೆ ಮೊಣಕಾಲು ಕೀಲು ಮತ್ತು ಮೂಳೆ ನೋವು ಇದೆ");
        departmentQuestions.put("Orthotist", "ನನಗೆ ನಡಿಗೆಗೆ ಆರ್ಥೋಟಿಕ್ ಬೆಂಬಲ ಬೇಕು");
        departmentQuestions.put("Paediatric and Fetal Cardiologist", "ನನ್ನ ಮಗುವಿಗೆ ಹೃದಯ ಸಮಸ್ಯೆ ಇದೆ");
        departmentQuestions.put("Paediatric Cardiologist", "ನನ್ನ ಮಗುವಿಗೆ ಎದೆ ಮತ್ತು ಹೃದಯ ಸಮಸ್ಯೆ ಇದೆ");
        departmentQuestions.put("Paediatric Neurologist", "ನನ್ನ ಮಗುವಿಗೆ ಅಪಸ್ಮಾರ ಮತ್ತು ನರಗಳ ಸಮಸ್ಯೆ ಇದೆ");
        departmentQuestions.put("Paediatric Surgeon", "ನನ್ನ ಮಗುವಿಗೆ ಶಸ್ತ್ರಚಿಕಿತ್ಸೆ ಬೇಕು");
        departmentQuestions.put("Paediatrician", "ನನ್ನ ಮಗುವಿಗೆ ಜ್ವರ ಮತ್ತು ಕೆಮ್ಮು ಇದೆ");
        departmentQuestions.put("Pain medicine", "ನನಗೆ ದೀರ್ಘಕಾಲದ ದೇಹದ ನೋವು ಇದೆ");
        departmentQuestions.put("Pedodentists", "ನನ್ನ ಮಗುವಿಗೆ ಹಲ್ಲು ನೋವು ಇದೆ");
        departmentQuestions.put("Periodontist", "ನನಗೆ ಒಸಡು ಊತ ಮತ್ತು ರಕ್ತಸ್ರಾವ ಇದೆ");
        departmentQuestions.put("Physiotherapy", "ನನಗೆ ಸ್ನಾಯು ನೋವಿಗೆ ಫಿಸಿಯೋಥೆರಪಿ ಬೇಕು");
        departmentQuestions.put("Plastic Surgery", "ನನಗೆ ಕಾಸ್ಮೆಟಿಕ್ ಸರ್ಜರಿ ಸಮಾಲೋಚನೆ ಬೇಕು");
        departmentQuestions.put("Podiatric Surgeon", "ನನಗೆ ಪಾದ ನೋವು ಮತ್ತು ನಡಿಗೆ ಸಮಸ್ಯೆ ಇದೆ");
        departmentQuestions.put("Psychiatry", "ನನಗೆ ಆತಂಕ, ಖಿನ್ನತೆ ಮತ್ತು ಮಾನಸಿಕ ಒತ್ತಡ ಇದೆ");
        departmentQuestions.put("Psychology", "ನನಗೆ ಮಾನಸಿಕ ಆರೋಗ್ಯ ಸಮಾಲೋಚನೆ ಬೇಕು");
        departmentQuestions.put("Pulmonologist", "ನನಗೆ ಉಸಿರಾಟ ಮತ್ತು ಶ್ವಾಸಕೋಶದ ಸಮಸ್ಯೆ ಇದೆ");
        departmentQuestions.put("Radiology", "ನನಗೆ MRI CT ಸ್ಕ್ಯಾನ್ ಮತ್ತು ಎಕ್ಸರೆ ಬೇಕು");
        departmentQuestions.put("Rheumatology", "ನನಗೆ ಸಂಧಿವಾತ ಮತ್ತು ಕೀಲು ನೋವು ಇದೆ");
        departmentQuestions.put("Speech Therapy", "ನನಗೆ ಮಾತನಾಡುವ ಮತ್ತು ಸಂವಹನ ತೊಂದರೆ ಇದೆ");
        departmentQuestions.put("Urologist", "ನನಗೆ ಮೂತ್ರ ಸೋಂಕು ಮತ್ತು ಮೂತ್ರಕೋಶದ ನೋವು ಇದೆ");
        
        // ================= VALID DEPARTMENTS SET =================
        Set<String> validDepartments = new HashSet<>(departmentQuestions.keySet());

        // ================= CREATE EXCEL WORKBOOK WITH HEADERS =================
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("RXDX_Whitefield_Kannada_Results");
        
        addProjectHeader(sheet, workbook, "RXDX WHITEFIELD HOSPITAL - KANNADA LANGUAGE VALIDATION TESTS");
        
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
        System.out.println("     KANNADA LANGUAGE VALIDATION TESTS STARTING - TOTAL TEST CASES: " + totalTests);
        System.out.println("     Validating ChatGPT responses in Kannada script");
        System.out.println("=".repeat(80));

        // ================= TEST LOOP =================

        int testNumber = 1;
        int excelRowNum = headerRowIndex + 1;

        for (Map.Entry<String, String> entry : departmentQuestions.entrySet()) {

            String expectedDepartment = entry.getKey();
            String kannadaQuestion = entry.getValue();

            String botResponse = "";
            String placeResponse = "";
            String status = "PASS";
            String failureType = "";
            String failureReason = "";
            String screenshotPath = "";
            String extraDepartments = "";
            String allDepartmentsShown = "";
            boolean isKannadaResponse = false;
            int kannadaCharCount = 0;

            System.out.println("\n" + "=".repeat(60));
            System.out.println("TEST #" + testNumber + " of " + totalTests);
            System.out.println("Department: " + expectedDepartment);
            System.out.println("Kannada Question: " + kannadaQuestion);
            System.out.println("=".repeat(60));

            try {

                // ================= STEP 1: Send "hi" =================
                System.out.println("\n📤 STEP 1: Sending 'hi' to start...");
                sendMessage(driver, "hi", wait);
                System.out.println("⏳ Waiting 5 seconds for response to load...");
                Thread.sleep(5000);
                
                botResponse = getLastMessageReliable(driver);
                System.out.println("✅ Response received after 'hi'");

                // ================= STEP 2: Send Kannada department question =================
                System.out.println("\n📤 STEP 2: Sending Kannada question: " + kannadaQuestion);
                sendMessage(driver, kannadaQuestion, wait);
                System.out.println("⏳ Waiting 5 seconds for response to load...");
                Thread.sleep(5000);
                
                botResponse = getLastMessageReliable(driver);
                System.out.println("\nBOT RESPONSE:\n" + botResponse);

                // ================= VALIDATION 1: Check if response is in Kannada =================
                System.out.println("\n🔍 VALIDATION 1: Checking if ChatGPT responded in Kannada...");
                
                // Count Kannada characters in response
                for (char c : botResponse.toCharArray()) {
                    if (c >= '\u0C80' && c <= '\u0CFF') {
                        isKannadaResponse = true;
                        kannadaCharCount++;
                    }
                }
                
                // Also check for GPT footer in Kannada or English
                boolean hasGPTFooter = botResponse.contains("This Response was generated using ChatGPT.") ||
                                       botResponse.contains("ಈ ಪ್ರತಿಕ್ರಿಯೆಯನ್ನು ChatGPT ಬಳಸಿ ರಚಿಸಲಾಗಿದೆ.");
                
                if (!isKannadaResponse && kannadaCharCount == 0) {
                    status = "FAIL";
                    failureType = "NOT_KANNADA_RESPONSE";
                    failureReason = "ChatGPT responded in English/Roman script, not in Kannada script";
                    System.out.println("❌ FAIL: Response is NOT in Kannada script!");
                } else {
                    System.out.println("✅ PASS: Response is in Kannada script! (Found " + kannadaCharCount + " Kannada characters)");
                }
                
                System.out.println("   Kannada characters count: " + kannadaCharCount);

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

                // ================= VALIDATION 2: Check if place response is in Kannada =================
                System.out.println("\n🔍 VALIDATION 2: Checking if place response is in Kannada...");
                
                int placeKannadaCount = 0;
                for (char c : placeResponse.toCharArray()) {
                    if (c >= '\u0C80' && c <= '\uCFFF') {
                        placeKannadaCount++;
                    }
                }
                
                if (placeKannadaCount == 0) {
                    System.out.println("⚠️ WARNING: Place response is not in Kannada script");
                } else {
                    System.out.println("✅ Place response has " + placeKannadaCount + " Kannada characters");
                }

                // ================= VALIDATION 3: Department filtering =================
                System.out.println("\n🔍 VALIDATION 3: Validating department filtering...");
                
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
                    System.out.println("   ✓ ChatGPT responded in Kannada script");
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

            addTestResultToExcel(sheet, excelRowNum++, testNumber, expectedDepartment, kannadaQuestion,
                                 botResponse, placeResponse, failureType, failureReason, extraDepartments,
                                 allDepartmentsShown, screenshotPath, status, isKannadaResponse, kannadaCharCount);

            System.out.println("\nSTATUS: " + status);
            System.out.println("   Kannada Response: " + (isKannadaResponse ? "YES" : "NO"));
            System.out.println("   Kannada Char Count: " + kannadaCharCount);
            
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
        infoCell.setCellValue("Project: RXDX Whitefield Kannada Validation | Testing: Department-wise Appointment Booking | Total Departments: 45 | Wait Time: 5 Seconds | Language: Kannada | Date: " + 
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
            "Test No", "Department", "Kannada Question", "Bot Response", "Place Response",
            "Status", "Failure Type", "Failure Reason", "Extra Departments Found",
            "All Departments Shown", "Kannada Response", "Kannada Char Count", "Screenshot", "Timestamp"
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
                                      boolean isKannadaResponse, int kannadaCharCount) {
        
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
        row.createCell(10).setCellValue(isKannadaResponse ? "YES (PASS)" : "NO (FAIL)");
        row.createCell(11).setCellValue(kannadaCharCount);
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
            System.out.println("✅ ChatGPT is responding in Kannada script correctly!");
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