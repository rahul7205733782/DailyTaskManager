package RXDX_Kadugodi;

import java.time.Duration;
import java.util.*;
import org.openqa.selenium.*;
import org.openqa.selenium.edge.EdgeDriver;
import org.openqa.selenium.edge.EdgeOptions;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.util.CellRangeAddress;

import java.awt.Desktop;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;

public class Sarjapur_Rd_Dommasandra {

    // ================= DEPARTMENT SYNONYMS (from Whitefield code) =================
    static Map<String, Set<String>> acceptableDepartments = new HashMap<>();

    static {
        acceptableDepartments.put("ENT", new HashSet<>(Arrays.asList("ENT")));
        acceptableDepartments.put("Family Medicine", new HashSet<>(Arrays.asList("Family Medicine", "General Medicine")));
        acceptableDepartments.put("Gastroenterologist", new HashSet<>(Arrays.asList("Gastroenterologist", "Gastroenterologist, Hepatologist")));
        acceptableDepartments.put("General Medicine", new HashSet<>(Arrays.asList("General Medicine", "Internal Medicine")));
        acceptableDepartments.put("General Surgeon", new HashSet<>(Arrays.asList("General Surgeon", "Paediatric Surgeon", "Orthopedic Surgeon")));
        acceptableDepartments.put("Homeopathy", new HashSet<>(Arrays.asList("Homeopathy")));
        acceptableDepartments.put("Nephrologist", new HashSet<>(Arrays.asList("Nephrologist", "Urologist")));
        acceptableDepartments.put("Neurologist", new HashSet<>(Arrays.asList("Neurologist", "Paediatric Neurologist")));
        acceptableDepartments.put("Ophthalmology", new HashSet<>(Arrays.asList("Ophthalmology")));
        acceptableDepartments.put("Orthopedician", new HashSet<>(Arrays.asList("Orthopedic Surgeon", "Orthopedician", "Orthopaedic")));
        acceptableDepartments.put("Paediatrician", new HashSet<>(Arrays.asList("Paediatrician", "Pediatrician")));
        acceptableDepartments.put("Physiotherapy", new HashSet<>(Arrays.asList("Physiotherapy", "Physical Therapy")));
        acceptableDepartments.put("Pulmonologist", new HashSet<>(Arrays.asList("Pulmonologist")));
    }

    static final String GPT_FOOTER = "This Response was generated using ChatGPT.";

    public static void main(String[] args) throws InterruptedException, IOException {
        System.setProperty("org.apache.poi.util.POILogger", "org.apache.poi.util.NullLogger");

        String projectPath = System.getProperty("user.dir");
        String timestamp = new SimpleDateFormat("yyyyMMdd_HHmmss").format(new Date());

        String screenshotsPath = projectPath + "/RXDX_Validation_Sarjapur/Screenshots/" + timestamp;
        String excelFolderPath = projectPath + "/RXDX_Validation_Sarjapur/Test_Excel";
        String excelFilePath = excelFolderPath + "/Sarjapur_Road_Dommasandra_Results_" + timestamp + ".xlsx";

        new File(screenshotsPath).mkdirs();
        new File(excelFolderPath).mkdirs();

        System.out.println("\n📁 SCREENSHOTS: " + screenshotsPath);
        System.out.println("📊 EXCEL FILE: " + excelFilePath);

        // ================= DEPARTMENT QUESTIONS (from Sarjapur_Rd_Dommasandra) =================
        Map<String, String> departmentQuestions = new LinkedHashMap<>();
        departmentQuestions.put("ENT", "I have ear pain and throat infection");
        departmentQuestions.put("Family Medicine", "I need family doctor consultation");
        departmentQuestions.put("Gastroenterologist", "I have stomach pain and digestion issue");
        departmentQuestions.put("General Medicine", "I feel weak tired and sick");
        departmentQuestions.put("General Surgeon", "I need surgery consultation");
        departmentQuestions.put("Homeopathy", "I need homeopathy treatment");
        departmentQuestions.put("Nephrologist", "I have kidney stone problem");
        departmentQuestions.put("Neurologist", "I have migraine and severe headache");
        departmentQuestions.put("Ophthalmology", "I have blurry vision and eye pain");
        departmentQuestions.put("Orthopedician", "I have knee joint and bone pain");
        departmentQuestions.put("Paediatrician", "My child has fever and cough");
        departmentQuestions.put("Physiotherapy", "I need physiotherapy for muscle pain");
        departmentQuestions.put("Pulmonologist", "I have breathing and lung problem");

        Set<String> validDepartments = new HashSet<>(departmentQuestions.keySet());

        // ================= EXCEL SETUP =================
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Sarjapur_Results");
        addProjectHeader(sheet, workbook);
        int headerRowIndex = addColumnHeaders(sheet, workbook);
        int excelRowNum = headerRowIndex + 1;

        // ================= WEBDRIVER =================
        EdgeOptions options = new EdgeOptions();
        options.addArguments("--user-data-dir=C:\\EdgeAutomationProfile");
        options.addArguments("--start-maximized");
        WebDriver driver = new EdgeDriver(options);
        WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(120));

        driver.manage().window().maximize();
        driver.get("https://web.whatsapp.com/");
        System.out.println("\n⚠️ PLEASE SCAN QR CODE...");
        Thread.sleep(20000);

        // Open RXDX chat
        driver.findElement(By.xpath("//span[@data-icon='new-chat-outline']")).click();
        Thread.sleep(2000);
        driver.findElement(By.xpath("//input[@aria-label='Search name or number']")).sendKeys("916363415530");
        Thread.sleep(2000);
        driver.findElement(By.xpath("//span[@title='+91 63634 15530']")).click();
        Thread.sleep(2000);

        int totalTests = departmentQuestions.size();
        List<Map<String, String>> failedTestsData = new ArrayList<>();

        System.out.println("\n" + "=".repeat(80));
        System.out.println("   SARJAPUR ROAD DOMMASANDRA APPOINTMENT TESTS - TOTAL: " + totalTests);
        System.out.println("=".repeat(80));

        int testNumber = 1;

        for (Map.Entry<String, String> entry : departmentQuestions.entrySet()) {
            String expectedDepartment = entry.getKey();
            String question = entry.getValue();
            String botResponse = "";
            String placeResponse = "";
            String status = "PASS";
            String mappingStatus = "PASS";
            String filteringStatus = "PASS";
            String failureType = "";
            String failureReason = "";
            String screenshotPath = "";
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
                Thread.sleep(2000);
                System.out.println("⏳ Waiting for main menu...");
                wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//div[contains(text(),'Main Menu')]")));
                System.out.println("✅ Main menu detected!");
                Thread.sleep(2000);
                botResponse = getLastMessage(driver);
                Thread.sleep(2000);
                System.out.println("✅ Bot response received");

                // STEP 2: Send department question
                System.out.println("\n📤 STEP 2: Sending department question: " + question);
                sendMessage(driver, question, wait);
                Thread.sleep(2000);
                System.out.println("⏳ Waiting for GPT response (max 2 min)...");
                wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//em[contains(text(),'This Response was generated using ChatGPT.')]")));
                System.out.println("✅ GPT response detected!");
                Thread.sleep(2000);
                botResponse = getLastMessage(driver);
                System.out.println("\nGPT RESPONSE:\n" + botResponse.substring(0, Math.min(300, botResponse.length())));

                // Validation 1: GPT Footer
                if (!botResponse.contains(GPT_FOOTER)) {
                    status = "FAIL";
                    failureType = "GPT_RESPONSE_NOT_COMPLETED";
                    failureReason = "GPT footer not found";
                    screenshotPath = takeScreenshot(driver, screenshotsPath, expectedDepartment + "_GPTFooterMissing");
                    addTestResultToExcel(sheet, excelRowNum++, testNumber, expectedDepartment, question,
                            botResponse, placeResponse, failureType, failureReason, allDepartmentsShown,
                            screenshotPath, status, mappingStatus, filteringStatus);
                    testNumber++;
                    continue;
                }
                System.out.println("✅ GPT footer found");

                // Validation 2: Department mapping using synonym map
                boolean valid = validateResponseWithSynonyms(expectedDepartment, botResponse);
                if (!valid) {
                    status = "FAIL";
                    mappingStatus = "FAIL";
                    failureType = "WRONG_GPT_RESPONSE";
                    failureReason = "Department keywords not found (even with synonyms)";
                } else {
                    System.out.println("✅ Department mapping correct");
                }

                // STEP 3: Click Book an Appointment button - MULTIPLE STRATEGIES
                System.out.println("\n🔍 STEP 3: Looking for 'Book an Appointment' button...");
                boolean buttonClicked = false;

                List<WebElement> exactSpans = driver.findElements(By.xpath("(//span[contains(text(),'Book an Appointment')])[last()]"));
                for (WebElement span : exactSpans) {
                    if (span.isDisplayed() && span.isEnabled()) {
                        span.click();
                        buttonClicked = true;
                        System.out.println("   ✅ Clicked exact span");
                        Thread.sleep(2000);
                        break;
                    }
                }

                if (!buttonClicked) {
                    List<WebElement> anyElements = driver.findElements(By.xpath("//*[contains(text(),'Book an Appointment')]"));
                    for (WebElement elem : anyElements) {
                        if (elem.isDisplayed() && elem.isEnabled()) {
                            elem.click();
                            buttonClicked = true;
                            System.out.println("   ✅ Clicked element containing text");
                            Thread.sleep(2000);
                            break;
                        }
                    }
                }

                if (!buttonClicked) {
                    try {
                        List<WebElement> msgContainers = driver.findElements(By.xpath("//div[@data-testid='msg-container']"));
                        if (!msgContainers.isEmpty()) {
                            WebElement lastMsg = msgContainers.get(msgContainers.size() - 1);
                            List<WebElement> btns = lastMsg.findElements(By.xpath(".//span[contains(text(),'Book')]"));
                            for (WebElement btn : btns) {
                                if (btn.isDisplayed() && btn.isEnabled()) {
                                    btn.click();
                                    buttonClicked = true;
                                    System.out.println("   ✅ Clicked button from message container");
                                    Thread.sleep(2000);
                                    break;
                                }
                            }
                        }
                    } catch (Exception e) {}
                }

                if (!buttonClicked) {
                    try {
                        WebElement jsButton = driver.findElement(By.xpath("//span[contains(text(),'Book an Appointment')]"));
                        ((JavascriptExecutor) driver).executeScript("arguments[0].click();", jsButton);
                        buttonClicked = true;
                        System.out.println("   ✅ Clicked using JavaScript");
                        Thread.sleep(2000);
                    } catch (Exception e) {}
                }

                if (!buttonClicked) {
                    try {
                        WebElement actionsButton = driver.findElement(By.xpath("//span[contains(text(),'Book an Appointment')]"));
                        org.openqa.selenium.interactions.Actions actions = new org.openqa.selenium.interactions.Actions(driver);
                        actions.moveToElement(actionsButton).click().perform();
                        buttonClicked = true;
                        System.out.println("   ✅ Clicked using Actions");
                        Thread.sleep(2000);
                    } catch (Exception e) {}
                }

                if (!buttonClicked) {
                    status = "FAIL";
                    failureType = "BOOK_APPOINTMENT_BUTTON_MISSING";
                    failureReason = "Book Appointment button not found with any strategy";
                    screenshotPath = takeScreenshot(driver, screenshotsPath, expectedDepartment + "_NoBookButton");
                    addTestResultToExcel(sheet, excelRowNum++, testNumber, expectedDepartment, question,
                            botResponse, placeResponse, failureType, failureReason, allDepartmentsShown,
                            screenshotPath, status, mappingStatus, filteringStatus);
                    testNumber++;
                    continue;
                }

                System.out.println("⏳ Waiting for location options after click...");
                Thread.sleep(4000);

                // STEP 4: Select Sarjapur Road Dommasandra (option 3)
                System.out.println("\n📤 STEP 4: Sending '3' for Sarjapur Road Dommasandra...");
                sendMessage(driver, "3", wait);
                Thread.sleep(2000);
                Thread.sleep(5000);
                placeResponse = getLastMessage(driver);
                Thread.sleep(2000);
                System.out.println("\nLOCATION RESPONSE:\n" + placeResponse);

                // STEP 5: Validate filtering – count departments shown (max 2 allowed)
                System.out.println("\n🔍 STEP 5: Validating filtering (max 2 departments allowed)...");
                allDepartmentsShown = extractAllDepartmentsFromResponse(placeResponse, validDepartments);
                int departmentsCount = allDepartmentsShown.equals("None") ? 0 : allDepartmentsShown.split(", ").length;
                System.out.println("   Departments shown (" + departmentsCount + "): " + allDepartmentsShown);

                if (departmentsCount > 2) {
                    status = "FAIL";
                    filteringStatus = "FAIL";
                    failureType = "TOO_MANY_DEPARTMENTS";
                    failureReason = "Expected at most 2 departments, but found " + departmentsCount + ": " + allDepartmentsShown;
                } else if (!containsDepartmentWithSynonyms(placeResponse, expectedDepartment)) {
                    status = "FAIL";
                    mappingStatus = "FAIL";
                    failureType = "DEPARTMENT_NOT_FOUND";
                    failureReason = "Expected department '" + expectedDepartment + "' not found in location response";
                } else {
                    System.out.println("✅ Department found and departments count ≤ 2");
                }

                if (status.equals("PASS")) {
                    System.out.println("\n✅✅✅ TEST PASSED for: " + expectedDepartment);
                }

                if (status.equals("FAIL")) {
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
                    botResponse, placeResponse, failureType, failureReason, allDepartmentsShown,
                    screenshotPath, status, mappingStatus, filteringStatus);

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

        // Auto-size columns & save Excel
        for (int i = 0; i <= 11; i++) sheet.autoSizeColumn(i);
        try (FileOutputStream fileOut = new FileOutputStream(excelFilePath)) {
            workbook.write(fileOut);
            System.out.println("\n✅ EXCEL SAVED: " + excelFilePath);
        }
        workbook.close();

        printFinalReport(totalTests, failedTestsData);

        try {
            File excelFile = new File(excelFilePath);
            if (Desktop.isDesktopSupported()) Desktop.getDesktop().open(excelFile);
            System.out.println("\n✅ Excel opened automatically");
        } catch (Exception e) {}

        driver.quit();
    }

    // ================= HELPER METHODS (same as Whitefield version) =================

    static void addProjectHeader(Sheet sheet, Workbook workbook) {
        Row headerRow = sheet.createRow(0);
        Cell headerCell = headerRow.createCell(0);
        headerCell.setCellValue("SARJAPUR ROAD DOMMASANDRA CHATBOT - DEPARTMENT APPOINTMENT TESTS");
        CellStyle style = workbook.createCellStyle();
        Font font = workbook.createFont();
        font.setBold(true);
        font.setFontHeightInPoints((short) 16);
        style.setFont(font);
        headerCell.setCellStyle(style);
        sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 11));

        Row infoRow = sheet.createRow(1);
        Cell infoCell = infoRow.createCell(0);
        infoCell.setCellValue("Date: " + new SimpleDateFormat("yyyy-MM-dd HH:mm:ss").format(new Date()));
        sheet.addMergedRegion(new CellRangeAddress(1, 1, 0, 11));
    }

    static int addColumnHeaders(Sheet sheet, Workbook workbook) {
        Row row = sheet.createRow(2);
        String[] cols = {"Test No","Department","Question","Bot Response","Location Response",
                "Status","Mapping Status","Filtering Status","Failure Type","Failure Reason",
                "Departments Shown","Screenshot"};
        for (int i=0; i<cols.length; i++) row.createCell(i).setCellValue(cols[i]);
        return 2;
    }

    static void addTestResultToExcel(Sheet sheet, int rowNum, int testNumber, String dept, String q,
                                     String bot, String place, String failType, String failReason,
                                     String deptsShown, String screenshot, String status,
                                     String mapStatus, String filterStatus) {
        Row row = sheet.createRow(rowNum);
        row.createCell(0).setCellValue(testNumber);
        row.createCell(1).setCellValue(dept);
        row.createCell(2).setCellValue(q);
        row.createCell(3).setCellValue(bot);
        row.createCell(4).setCellValue(place);
        row.createCell(5).setCellValue(status);
        row.createCell(6).setCellValue(mapStatus);
        row.createCell(7).setCellValue(filterStatus);
        row.createCell(8).setCellValue(failType);
        row.createCell(9).setCellValue(failReason);
        row.createCell(10).setCellValue(deptsShown);
        row.createCell(11).setCellValue(screenshot);
    }

    static void sendMessage(WebDriver driver, String text, WebDriverWait wait) throws InterruptedException {
        WebElement input = wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//div[@aria-placeholder='Type a message']")));
        input.clear();
        input.sendKeys(text);
        Thread.sleep(1000);
        input.sendKeys(Keys.ENTER);
        Thread.sleep(2000);
    }

    static String getLastMessage(WebDriver driver) {
        String[] xpaths = {
                "//div[@data-testid='msg-container']",
                "//div[contains(@class,'message-in')]",
                "//div[contains(@class,'_akbu')]",
                "//div[contains(@class,'copyable-text')]"
        };
        for (String xp : xpaths) {
            try {
                List<WebElement> msgs = driver.findElements(By.xpath(xp));
                if (!msgs.isEmpty()) {
                    String txt = msgs.get(msgs.size()-1).getText();
                    if (txt != null && txt.length() > 10) return txt;
                }
            } catch (Exception e) {}
        }
        return "NO RESPONSE";
    }

    static String takeScreenshot(WebDriver driver, String folder, String name) {
        try {
            String clean = name.replaceAll("[^a-zA-Z0-9]", "_");
            String path = folder + "/" + clean + ".png";
            File src = ((TakesScreenshot)driver).getScreenshotAs(OutputType.FILE);
            src.renameTo(new File(path));
            System.out.println("📸 Screenshot: " + path);
            return path;
        } catch (Exception e) { return "Screenshot failed"; }
    }

    static String extractAllDepartmentsFromResponse(String resp, Set<String> valid) {
        List<String> found = new ArrayList<>();
        for (String d : valid) {
            if (resp.toLowerCase().contains(d.toLowerCase())) {
                found.add(d);
            }
        }
        return found.isEmpty() ? "None" : String.join(", ", found);
    }

    static boolean validateResponseWithSynonyms(String expectedDept, String botResp) {
        String lowerResp = botResp.toLowerCase();
        Set<String> synonyms = acceptableDepartments.get(expectedDept);
        if (synonyms == null) {
            // fallback: just check if department name appears
            return lowerResp.contains(expectedDept.toLowerCase());
        }
        for (String syn : synonyms) {
            if (lowerResp.contains(syn.toLowerCase())) {
                return true;
            }
        }
        return false;
    }

    static boolean containsDepartmentWithSynonyms(String response, String expectedDept) {
        String lowerResp = response.toLowerCase();
        Set<String> synonyms = acceptableDepartments.get(expectedDept);
        if (synonyms == null) {
            return lowerResp.contains(expectedDept.toLowerCase());
        }
        for (String syn : synonyms) {
            if (lowerResp.contains(syn.toLowerCase())) {
                return true;
            }
        }
        return false;
    }

    static void printFinalReport(int total, List<Map<String,String>> fails) {
        System.out.println("\n" + "=".repeat(80) + "\nFINAL TEST REPORT\n" + "=".repeat(80));
        if (fails.isEmpty()) System.out.println("\n🎉 ALL TESTS PASSED! (" + total + "/" + total + ")");
        else System.out.println("\n❌ FAILURES: " + fails.size() + " of " + total);
    }
}