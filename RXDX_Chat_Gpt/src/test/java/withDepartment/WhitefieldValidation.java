package withDepartment;

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

public class WhitefieldValidation {

    public static void main(String[] args)
            throws InterruptedException, IOException {

        System.setProperty("org.apache.poi.util.POILogger", "org.apache.poi.util.NullLogger");

        String projectPath = System.getProperty("user.dir");
        String timestamp = new SimpleDateFormat("yyyyMMdd_HHmmss").format(new Date());

        String screenshotsPath = projectPath + "/RXDX_Validation_Whitefield/Screenshots/" + timestamp;
        String excelFolderPath = projectPath + "/RXDX_Validation_Whitefield/Test_Excel";
        String excelFilePath = excelFolderPath + "/RXDX_Whitefield_Test_Results_" + timestamp + ".xlsx";

        new File(screenshotsPath).mkdirs();
        new File(excelFolderPath).mkdirs();

        System.out.println("\n📁 SCREENSHOTS will be saved in: " + screenshotsPath);
        System.out.println("📊 EXCEL FILE will be saved as: " + excelFilePath);

        // ================= DEPARTMENTS (EXACT 45 DEPARTMENTS FROM RXDX WHITEFIELD) =================

        Map<String, String> departmentQuestions = new LinkedHashMap<>();

        departmentQuestions.put("Acupuncture", "Acupuncture doctor jothe appointment book maadi");
        departmentQuestions.put("Adult Cardiology", "Cardiology doctor jothe appointment book maadi");
        departmentQuestions.put("Allergy Specialist", "Allergy specialist doctor jothe appointment book maadi");
        departmentQuestions.put("Clinical Nutrition", "Nutrition specialist appointment beku");
        departmentQuestions.put("Dental", "Dental doctor jothe appointment book maadi");
        departmentQuestions.put("Dermatology", "Dermatology doctor jothe appointment book maadi");
        departmentQuestions.put("Diabetologist", "Diabetologist doctor appointment beku");
        departmentQuestions.put("Endocrinologist", "Endocrinologist doctor jothe appointment book maadi");
        departmentQuestions.put("Endodentist", "Endodentist doctor consultation beku");
        departmentQuestions.put("Endodontist", "Endodontist doctor jothe appointment book maadi");
        departmentQuestions.put("Endovascular Neurosurgeon, Vascular Surgeon", "Vascular Surgeon doctor jothe appointment book maadi");
        departmentQuestions.put("ENT", "ENT doctor jothe appointment book maadi");
        departmentQuestions.put("Gastroenterologist, Hepatologist", "Hepatologist doctor jothe appointment book maadi");
        departmentQuestions.put("General Medicine", "General Medicine doctor appointment beku");
        departmentQuestions.put("General Surgeon", "General Surgeon doctor jothe appointment book maadi");
        departmentQuestions.put("Homeopathy", "Homeopathy doctor jothe appointment book maadi");
        departmentQuestions.put("Internal Medicine", "Internal Medicine doctor consultation beku");
        departmentQuestions.put("Nephrologist", "Nephrologist doctor jothe appointment book maadi");
        departmentQuestions.put("Neurologist", "Neurologist doctor jothe appointment book maadi");
        departmentQuestions.put("Neurosurgery", "Neurosurgery consultation beku");
        departmentQuestions.put("Obs & Gynae", "Gynecology doctor jothe appointment book maadi");
        departmentQuestions.put("Occupational Therapy", "Occupational Therapy appointment beku");
        departmentQuestions.put("Ophthalmology", "Kannina doctor jothe appointment book maadi");
        departmentQuestions.put("Oral Surgeon", "Oral Surgeon doctor jothe appointment book maadi");
        departmentQuestions.put("Orthodentists", "Orthodontist doctor appointment beku");
        departmentQuestions.put("Orthopedic Surgeon", "Orthopedic Surgeon doctor jothe appointment book maadi");
        departmentQuestions.put("Orthotist", "Orthotist consultation beku");
        departmentQuestions.put("Paediatric and Fetal Cardiologist", "Makkala heart doctor jothe appointment book maadi");
        departmentQuestions.put("Paediatric Cardiologist", "Pediatric Cardiology doctor jothe appointment book maadi");
        departmentQuestions.put("Paediatric Neurologist", "Makkala Neurologist doctor jothe appointment book maadi");
        departmentQuestions.put("Paediatric Surgeon", "Pediatric Surgeon doctor jothe appointment book maadi");
        departmentQuestions.put("Paediatrician", "Makkala doctor appointment beku");
        departmentQuestions.put("Pain medicine", "Pain Medicine doctor jothe appointment book maadi");
        departmentQuestions.put("Pedodentists", "Makkala Dental doctor consultation beku");
        departmentQuestions.put("Periodontist", "Periodontist doctor jothe appointment book maadi");
        departmentQuestions.put("Physiotherapy", "Physiotherapy appointment beku");
        departmentQuestions.put("Plastic Surgery", "Plastic Surgery doctor appointment beku");
        departmentQuestions.put("Podiatric Surgeon", "Podiatric Surgeon doctor jothe appointment book maadi");
        departmentQuestions.put("Psychiatry", "Psychiatry doctor jothe appointment book maadi");
        departmentQuestions.put("Psychology", "Psychology doctor jothe appointment book maadi");
        departmentQuestions.put("Pulmonologist", "Pulmonologist doctor appointment beku");
        departmentQuestions.put("Radiology", "Radiology appointment beku");
        departmentQuestions.put("Rheumatology", "Rheumatology consultation beku");
        departmentQuestions.put("Speech Therapy", "Speech Therapy appointment beku");
        departmentQuestions.put("Urologist", "Urologist doctor jothe appointment book maadi");

        Set<String> validDepartments = new HashSet<>(departmentQuestions.keySet());

        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("RXDX_Whitefield_Results");
        addProjectHeader(sheet, workbook, "RXDX WHITEFIELD HOSPITAL CHATBOT - DEPARTMENT APPOINTMENT TESTS");
        int headerRowIndex = addColumnHeaders(sheet, workbook);

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
        System.out.println("     DEPARTMENT-WISE APPOINTMENT TESTS STARTING - TOTAL TEST CASES: " + totalTests);
        System.out.println("=".repeat(80));

        int testNumber = 1;
        int excelRowNum = headerRowIndex + 1;

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
                botResponse = getLastMessage(driver);
                System.out.println("✅ Response after 'hi': " + botResponse.substring(0, Math.min(150, botResponse.length())));

                // STEP 2: Send department question
                System.out.println("\n📤 STEP 2: Sending department question: " + question);
                sendMessage(driver, question, wait);
                Thread.sleep(8000);
                botResponse = getLastMessage(driver);
                System.out.println("\nBOT RESPONSE:\n" + botResponse);

                // STEP 3: Click Book Appointment button
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
                    failureReason = "Book Appointment button missing";
                }

                // STEP 4: Send "1" for Whitefield location
                System.out.println("\n📤 STEP 4: Sending '1' for Whitefield location...");
                sendMessage(driver, "1", wait);
                Thread.sleep(8000);
                placeResponse = getLastMessage(driver);
                System.out.println("\nPLACE RESPONSE:\n" + placeResponse);

                // STEP 5: Validate filtering
                System.out.println("\n🔍 STEP 5: Validating filtering...");
                allDepartmentsShown = extractAllDepartmentsFromResponse(placeResponse, validDepartments);
                extraDepartments = findExtraDepartments(allDepartmentsShown, expectedDepartment);

                if (!placeResponse.toLowerCase().contains(expectedDepartment.toLowerCase())) {
                    status = "FAIL";
                    failureType = "DEPARTMENT_NOT_FOUND";
                    failureReason = "Expected department '" + expectedDepartment + "' not displayed";
                } else if (!extraDepartments.isEmpty()) {
                    status = "FAIL";
                    failureType = "FILTERING_FAILED_EXTRA_DEPARTMENTS";
                    failureReason = "Extra departments shown: " + extraDepartments;
                    System.out.println("❌ FILTERING FAILED! Extra departments found: " + extraDepartments);
                } else {
                    System.out.println("✅ FILTERING VALIDATION PASSED! Only showing: " + expectedDepartment);
                }

                if (status.equals("PASS")) {
                    System.out.println("\n✅ All validations passed for department: " + expectedDepartment);
                } else {
                    screenshotPath = takeScreenshot(driver, screenshotsPath, expectedDepartment);
                }

            } catch (Exception e) {
                status = "FAIL";
                failureType = "EXCEPTION";
                failureReason = "Exception: " + e.getMessage();
                e.printStackTrace();
                screenshotPath = takeScreenshot(driver, screenshotsPath, expectedDepartment);
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

            testNumber++;
            Thread.sleep(3000);
        }

        for (int i = 0; i <= 11; i++) sheet.autoSizeColumn(i);
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

    // ========== IMPROVED sendMessage (clears input, types character by character) ==========
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
        // Caller will sleep 8 seconds.
    }

    // ========== EXCEL AND HELPER METHODS (unchanged) ==========
    static void addProjectHeader(Sheet sheet, Workbook workbook, String projectTitle) { /* same as before */ }
    static int addColumnHeaders(Sheet sheet, Workbook workbook) { /* same */ }
    static void addTestResultToExcel(Sheet sheet, int rowNum, int testNumber, String expectedDepartment,
                                      String question, String botResponse, String placeResponse, 
                                      String failureType, String failureReason, String extraDepartments,
                                      String allDepartmentsShown, String screenshotPath, String status) { /* same */ }
    static String getLastMessage(WebDriver driver) { /* same */ }
    static String takeScreenshot(WebDriver driver, String folderPath, String fileName) { /* same */ }
    static String extractAllDepartmentsFromResponse(String response, Set<String> validDepartments) { /* same */ }
    static String findExtraDepartments(String allDepartments, String expectedDepartment) { /* same */ }
    static void printFinalReport(int totalTests, List<Map<String, String>> failedTestsData) { /* same */ }
}