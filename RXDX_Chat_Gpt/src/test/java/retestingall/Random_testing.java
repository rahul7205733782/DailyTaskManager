package retestingall;

import java.time.Duration;
import java.util.*;
import java.awt.Desktop;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;

import org.openqa.selenium.*;
import org.openqa.selenium.edge.EdgeDriver;
import org.openqa.selenium.edge.EdgeOptions;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.util.CellRangeAddress;

public class Random_testing {

    // ================================================================
    // ACCEPTABLE DEPARTMENT MAPPING
    // ================================================================

    static Map<String, Set<String>> acceptableDepartments = new HashMap<>();

    static {

        acceptableDepartments.put("Nephrologist",
                new HashSet<>(Arrays.asList(
                        "Nephrologist",
                        "Urologist"
                )));

        acceptableDepartments.put("Neurologist",
                new HashSet<>(Arrays.asList(
                        "Neurologist",
                        "Paediatric Neurologist"
                )));

        acceptableDepartments.put("General Medicine",
                new HashSet<>(Arrays.asList(
                        "General Medicine",
                        "Internal Medicine"
                )));

        acceptableDepartments.put("Internal Medicine",
                new HashSet<>(Arrays.asList(
                        "Internal Medicine",
                        "General Medicine"
                )));

        acceptableDepartments.put("General Surgeon",
                new HashSet<>(Arrays.asList(
                        "General Surgeon",
                        "Paediatric Surgeon",
                        "Orthopedic Surgeon",
                        "Oral Surgeon",
                        "Plastic Surgery",
                        "Podiatric Surgeon"
                )));

        acceptableDepartments.put("Neurosurgery",
                new HashSet<>(Arrays.asList(
                        "Neurosurgery",
                        "Neurologist"
                )));

        acceptableDepartments.put("Paediatrician",
                new HashSet<>(Arrays.asList(
                        "Paediatrician",
                        "Paediatric Cardiologist",
                        "Paediatric Neurologist",
                        "Paediatric Surgeon"
                )));

        acceptableDepartments.put("Paediatric Cardiologist",
                new HashSet<>(Arrays.asList(
                        "Paediatric Cardiologist",
                        "Paediatric and Fetal Cardiologist",
                        "Adult Cardiology"
                )));

        acceptableDepartments.put("Paediatric Neurologist",
                new HashSet<>(Arrays.asList(
                        "Paediatric Neurologist",
                        "Neurologist"
                )));

        acceptableDepartments.put("Pain medicine",
                new HashSet<>(Arrays.asList(
                        "Pain medicine",
                        "Physiotherapy",
                        "Neurologist",
                        "Orthopedic Surgeon"
                )));

        acceptableDepartments.put("Dental",
                new HashSet<>(Arrays.asList(
                        "Dental",
                        "Endodentist",
                        "Endodontist",
                        "Oral Surgeon",
                        "Orthodentists",
                        "Pedodentists",
                        "Periodontist"
                )));

        acceptableDepartments.put("Endodentist",
                new HashSet<>(Arrays.asList(
                        "Endodentist",
                        "Endodontist",
                        "Dental"
                )));

        acceptableDepartments.put("Endodontist",
                new HashSet<>(Arrays.asList(
                        "Endodontist",
                        "Endodentist",
                        "Dental"
                )));

        acceptableDepartments.put("Adult Cardiology",
                new HashSet<>(Arrays.asList(
                        "Adult Cardiology",
                        "Paediatric Cardiologist",
                        "Paediatric and Fetal Cardiologist"
                )));

        acceptableDepartments.put("Dermatology",
                new HashSet<>(Arrays.asList(
                        "Dermatology",
                        "Allergy Specialist"
                )));

        acceptableDepartments.put("Ophthalmology",
                new HashSet<>(Arrays.asList(
                        "Ophthalmology"
                )));

        acceptableDepartments.put("ENT",
                new HashSet<>(Arrays.asList(
                        "ENT"
                )));

        acceptableDepartments.put("Diabetologist",
                new HashSet<>(Arrays.asList(
                        "Diabetologist",
                        "Endocrinologist"
                )));

        acceptableDepartments.put("Endocrinologist",
                new HashSet<>(Arrays.asList(
                        "Endocrinologist",
                        "Diabetologist"
                )));

        acceptableDepartments.put("Gastroenterologist, Hepatologist",
                new HashSet<>(Arrays.asList(
                        "Gastroenterologist, Hepatologist"
                )));

        acceptableDepartments.put("Homeopathy",
                new HashSet<>(Arrays.asList(
                        "Homeopathy"
                )));

        acceptableDepartments.put("Occupational Therapy",
                new HashSet<>(Arrays.asList(
                        "Occupational Therapy",
                        "Physiotherapy"
                )));

        acceptableDepartments.put("Orthopedic Surgeon",
                new HashSet<>(Arrays.asList(
                        "Orthopedic Surgeon",
                        "Physiotherapy",
                        "Pain medicine"
                )));

        acceptableDepartments.put("Physiotherapy",
                new HashSet<>(Arrays.asList(
                        "Physiotherapy",
                        "Pain medicine",
                        "Orthopedic Surgeon"
                )));

        acceptableDepartments.put("Psychiatry",
                new HashSet<>(Arrays.asList(
                        "Psychiatry",
                        "Psychology"
                )));

        acceptableDepartments.put("Psychology",
                new HashSet<>(Arrays.asList(
                        "Psychology",
                        "Psychiatry"
                )));

        acceptableDepartments.put("Pulmonologist",
                new HashSet<>(Arrays.asList(
                        "Pulmonologist"
                )));

        acceptableDepartments.put("Urologist",
                new HashSet<>(Arrays.asList(
                        "Urologist",
                        "Nephrologist"
                )));

        acceptableDepartments.put("Radiology",
                new HashSet<>(Arrays.asList(
                        "Radiology"
                )));

        acceptableDepartments.put("Rheumatology",
                new HashSet<>(Arrays.asList(
                        "Rheumatology",
                        "Orthopedic Surgeon"
                )));

        acceptableDepartments.put("Speech Therapy",
                new HashSet<>(Arrays.asList(
                        "Speech Therapy"
                )));

        acceptableDepartments.put("Acupuncture",
                new HashSet<>(Arrays.asList(
                        "Acupuncture",
                        "Pain medicine"
                )));

        acceptableDepartments.put("Clinical Nutrition",
                new HashSet<>(Arrays.asList(
                        "Clinical Nutrition"
                )));

        acceptableDepartments.put(
                "Endovascular Neurosurgeon, Vascular Surgeon",
                new HashSet<>(Arrays.asList(
                        "Endovascular Neurosurgeon, Vascular Surgeon",
                        "General Surgeon"
                ))
        );

        acceptableDepartments.put("Plastic Surgery",
                new HashSet<>(Arrays.asList(
                        "Plastic Surgery",
                        "General Surgeon"
                )));

        acceptableDepartments.put("Podiatric Surgeon",
                new HashSet<>(Arrays.asList(
                        "Podiatric Surgeon",
                        "Orthopedic Surgeon"
                )));

        acceptableDepartments.put("Orthotist",
                new HashSet<>(Arrays.asList(
                        "Orthotist",
                        "Physiotherapy"
                )));

        acceptableDepartments.put("Pedodentists",
                new HashSet<>(Arrays.asList(
                        "Pedodentists",
                        "Dental"
                )));

        acceptableDepartments.put("Periodontist",
                new HashSet<>(Arrays.asList(
                        "Periodontist",
                        "Dental"
                )));

        acceptableDepartments.put("Orthodentists",
                new HashSet<>(Arrays.asList(
                        "Orthodentists",
                        "Dental"
                )));

        acceptableDepartments.put("Oral Surgeon",
                new HashSet<>(Arrays.asList(
                        "Oral Surgeon",
                        "Dental",
                        "General Surgeon"
                )));

        acceptableDepartments.put("Obs & Gynae",
                new HashSet<>(Arrays.asList(
                        "Obs & Gynae"
                )));

        acceptableDepartments.put(
                "Paediatric and Fetal Cardiologist",
                new HashSet<>(Arrays.asList(
                        "Paediatric and Fetal Cardiologist",
                        "Paediatric Cardiologist",
                        "Adult Cardiology"
                ))
        );

        acceptableDepartments.put("Paediatric Surgeon",
                new HashSet<>(Arrays.asList(
                        "Paediatric Surgeon",
                        "General Surgeon"
                )));
    }

    static final String GPT_FOOTER =
            "This Response was generated using ChatGPT.";

    // ================================================================
    // MAIN
    // ================================================================

    public static void main(String[] args)
            throws InterruptedException, IOException {

        System.setProperty(
                "org.apache.poi.util.POILogger",
                "org.apache.poi.util.NullLogger"
        );

        String projectPath =
                System.getProperty("user.dir");

        String timestamp =
                new SimpleDateFormat(
                        "yyyyMMdd_HHmmss"
                ).format(new Date());

        String screenshotsPath =
                projectPath
                        + "/RXDX_Validation_Whitefield/Screenshots/"
                        + timestamp;

        String excelFolderPath =
                projectPath
                        + "/RXDX_Validation_Whitefield/Test_Excel";

        String excelFilePath =
                excelFolderPath
                        + "/RXDX_Whitefield_Test_Results_"
                        + timestamp
                        + ".xlsx";

        new File(screenshotsPath).mkdirs();
        new File(excelFolderPath).mkdirs();

        System.out.println(
                "\n📁 SCREENSHOTS: "
                        + screenshotsPath
        );

        System.out.println(
                "📊 EXCEL FILE: "
                        + excelFilePath
        );

        // ================================================================
        // DEPARTMENT QUESTIONS
        // 2 QUESTIONS PER DEPARTMENT
        // ================================================================

        Map<String, List<String>> departmentQuestions =
                new LinkedHashMap<>();

        addQuestions(
                departmentQuestions,
                "Acupuncture",
                "I have back pain. Can acupuncture help?",
                "I have neck pain. Would acupuncture be useful?"
        );

        addQuestions(
                departmentQuestions,
                "Adult Cardiology",
                "I’m having chest pain sometimes. Who should I see?",
                "I get short of breath when I walk. Could it be a heart problem?"
        );

        addQuestions(
                departmentQuestions,
                "Allergy Specialist",
                "I keep getting skin rashes. Could it be an allergy?",
                "I sneeze a lot and my eyes are itchy. Is this an allergy?"
        );

        addQuestions(
                departmentQuestions,
                "Clinical Nutrition",
                "I want help with my diet. Who should I see?",
                "I want to lose weight. Can a nutritionist help me?"
        );

        addQuestions(
                departmentQuestions,
                "Dental",
                "My tooth is hurting badly. Can I see a dentist?",
                "I have a cavity and tooth pain. Which doctor should I book?"
        );

        addQuestions(
                departmentQuestions,
                "Dermatology",
                "I have red patches on my skin. Which doctor should I see?",
                "I’m losing a lot of hair lately. Can a skin doctor help?"
        );

        addQuestions(
                departmentQuestions,
                "Diabetologist",
                "My sugar level is high. Which doctor should I see?",
                "I have diabetes and my sugar isn’t under control. Who should I consult?"
        );

        addQuestions(
                departmentQuestions,
                "Endocrinologist",
                "I think I have a hormone problem. Who should I see?",
                "I’m gaining weight for no reason. Could it be a hormone issue?"
        );

        addQuestions(
                departmentQuestions,
                "Endodentist",
                "I have a lot of pain inside my tooth. Where should I go?",
                "My tooth hurts when I bite. Do I need special dental treatment?"
        );

        addQuestions(
                departmentQuestions,
                "Endodontist",
                "My tooth is hurting badly. Do I need a root canal?",
                "I have pain deep inside my tooth. Can an endodontist help?"
        );

        addQuestions(
                departmentQuestions,
                "Endovascular Neurosurgeon, Vascular Surgeon",
                "I have a blood vessel problem. Which doctor can treat it?",
                "I was told I have a blocked blood vessel. Who should I consult?"
        );

        addQuestions(
                departmentQuestions,
                "ENT",
                "My ear is blocked and I can’t hear properly.",
                "I have a sore throat and trouble swallowing. Which doctor should I see?"
        );

        addQuestions(
                departmentQuestions,
                "Gastroenterologist, Hepatologist",
                "I’m having stomach and digestion problems. Which doctor should I see?",
                "I have a liver problem. Who should I consult?"
        );

        addQuestions(
                departmentQuestions,
                "General Medicine",
                "I’m not feeling well. I want to see a general physician.",
                "I have fever and body pain. Which doctor should I see?"
        );

        addQuestions(
                departmentQuestions,
                "General Surgeon",
                "I’ve been told I might need surgery. Who should I consult?",
                "I have a lump that may need surgery. Which doctor should I see?"
        );

        addQuestions(
                departmentQuestions,
                "Homeopathy",
                "I want to try homeopathic treatment. Is a doctor available?",
                "Can I consult a doctor for homeopathy treatment?"
        );

        addQuestions(
                departmentQuestions,
                "Internal Medicine",
                "I have fever and weakness. Which doctor should I see?",
                "I’ve been feeling tired and unwell for a few days. Who should I consult?"
        );

        addQuestions(
                departmentQuestions,
                "Nephrologist",
                "I’m having pain near my kidneys. Can I see a kidney specialist?",
                "I have a kidney problem. Which doctor should I book with?"
        );

        addQuestions(
                departmentQuestions,
                "Neurologist",
                "I’m having nerve pain. Which doctor should I see?",
                "I’ve been getting headaches and dizziness. Who should I consult?"
        );

        addQuestions(
                departmentQuestions,
                "Neurosurgery",
                "I have a brain problem and may need surgery. Who should I see?",
                "I was told I might need surgery for a spine problem. Which doctor should I consult?"
        );

        addQuestions(
                departmentQuestions,
                "Obs & Gynae",
                "I have irregular periods. Which doctor should I see?",
                "I’m having some problems during my pregnancy. Who should I consult?"
        );

        addQuestions(
                departmentQuestions,
                "Occupational Therapy",
                "I’m having trouble doing my daily activities. Can therapy help?",
                "I’m struggling to use my hand after an injury. Can occupational therapy help?"
        );

        addQuestions(
                departmentQuestions,
                "Ophthalmology",
                "My eyesight is blurry. I need to see an eye doctor.",
                "My eyes have been hurting and my vision is getting worse. Who should I see?"
        );

        addQuestions(
                departmentQuestions,
                "Oral Surgeon",
                "I have a problem with my jaw. Do I need an oral surgeon?",
                "My wisdom tooth is causing a lot of trouble. Who should I consult?"
        );

        addQuestions(
                departmentQuestions,
                "Orthodentists",
                "My teeth are not straight. I want to get them aligned.",
                "I have gaps between my teeth. Can I get braces?"
        );

        addQuestions(
                departmentQuestions,
                "Orthopedic Surgeon",
                "My knee is hurting and I’m having trouble walking.",
                "I injured my shoulder and it’s still painful. Which doctor should I see?"
        );

        addQuestions(
                departmentQuestions,
                "Orthotist",
                "I need a support or brace for my leg. Who should I see?",
                "I need a brace for my ankle. Can someone help me with that?"
        );

        addQuestions(
                departmentQuestions,
                "Paediatric and Fetal Cardiologist",
                "The scan showed a heart problem in my baby. Who should I consult?",
                "My unborn baby may have a heart issue. Which specialist should I see?"
        );

        addQuestions(
                departmentQuestions,
                "Paediatric Cardiologist",
                "My child has a heart problem. Which doctor should I book with?",
                "My child gets tired easily while playing. Could it be a heart problem?"
        );

        addQuestions(
                departmentQuestions,
                "Paediatric Neurologist",
                "My child is having unusual movements. Who should I see?",
                "My child has frequent headaches. Should I see a child neurologist?"
        );

        addQuestions(
                departmentQuestions,
                "Paediatric Surgeon",
                "My child may need surgery. Which doctor should I consult?",
                "My child has a condition that might require an operation. Who should I see?"
        );

        addQuestions(
                departmentQuestions,
                "Paediatrician",
                "My child has fever and cough. Can I book a pediatrician?",
                "My child isn’t eating properly and seems weak. Which doctor should I see?"
        );

        addQuestions(
                departmentQuestions,
                "Pain medicine",
                "I have pain that isn’t going away. Which doctor can help?",
                "I’ve had severe pain for a long time. Who should I consult?"
        );

        addQuestions(
                departmentQuestions,
                "Pedodentists",
                "My child’s tooth is hurting. Can I see a pediatric dentist?",
                "My child has a cavity. Which dentist should I book with?"
        );

        addQuestions(
                departmentQuestions,
                "Periodontist",
                "My gums are bleeding when I brush. Who should I see?",
                "My gums are swollen and painful. Can a specialist help?"
        );

        addQuestions(
                departmentQuestions,
                "Physiotherapy",
                "I have muscle pain. Can I get physiotherapy?",
                "My back hurts after an injury. Can physiotherapy help me?"
        );

        addQuestions(
                departmentQuestions,
                "Plastic Surgery",
                "I want to improve my appearance. Which doctor should I consult?",
                "I have a scar that I want to get treated. Who should I see?"
        );

        addQuestions(
                departmentQuestions,
                "Podiatric Surgeon",
                "My foot hurts when I walk. Which specialist should I see?",
                "I have a problem with my toe that may need surgery. Who should I consult?"
        );

        addQuestions(
                departmentQuestions,
                "Psychiatry",
                "I’ve been feeling very stressed lately. Who can I talk to?",
                "I’ve been having trouble sleeping because of my thoughts. Which doctor should I see?"
        );

        addQuestions(
                departmentQuestions,
                "Psychology",
                "I’m having some personal and emotional problems. Can I see a psychologist?",
                "I’ve been feeling anxious lately and need someone to talk to."
        );

        addQuestions(
                departmentQuestions,
                "Pulmonologist",
                "I’m having trouble breathing. Which doctor should I see?",
                "I’ve had a cough for a long time. Should I see a lung specialist?"
        );

        addQuestions(
                departmentQuestions,
                "Radiology",
                "My doctor asked me to get a scan. Where should I go?",
                "I need to get an MRI. Which department should I book with?"
        );

        addQuestions(
                departmentQuestions,
                "Rheumatology",
                "My joints have been hurting and feeling stiff. Who should I consult?",
                "I have swelling and pain in several joints. Which doctor should I see?"
        );

        addQuestions(
                departmentQuestions,
                "Speech Therapy",
                "I’m having trouble speaking clearly. Can speech therapy help?",
                "My child is not speaking properly for their age. Who should I consult?"
        );

        addQuestions(
                departmentQuestions,
                "Urologist",
                "It hurts when I pee. Which doctor should I see?",
                "I’m having a problem with frequent urination. Who should I consult?"
        );

        // ================================================================
        // VALIDATE QUESTION COUNT
        // ================================================================

        for (Map.Entry<String, List<String>> entry :
                departmentQuestions.entrySet()) {

            if (entry.getValue().size() != 2) {

                throw new IllegalStateException(
                        "Department must have exactly 2 questions: "
                                + entry.getKey()
                );
            }
        }

        int totalTests =
                departmentQuestions.size() * 2;

        System.out.println(
                "\n📋 TOTAL DEPARTMENTS: "
                        + departmentQuestions.size()
        );

        System.out.println(
                "🧪 TOTAL TEST CASES: "
                        + totalTests
        );

        // ================================================================
        // EXCEL
        // ================================================================

        Workbook workbook =
                new XSSFWorkbook();

        Sheet sheet =
                workbook.createSheet(
                        "RXDX_Whitefield_Results"
                );

        addProjectHeader(
                sheet,
                workbook
        );

        int headerRowIndex =
                addColumnHeaders(
                        sheet,
                        workbook
                );

        int excelRowNum =
                headerRowIndex + 1;

        // ================================================================
        // EDGE
        // ================================================================

        EdgeOptions options =
                new EdgeOptions();

        options.addArguments(
                "--user-data-dir=C:\\EdgeAutomationProfile"
        );

        options.addArguments(
                "--start-maximized"
        );

        WebDriver driver =
                new EdgeDriver(options);

        WebDriverWait wait =
                new WebDriverWait(
                        driver,
                        Duration.ofSeconds(120)
                );

        driver.manage()
                .window()
                .maximize();

        driver.get(
                "https://web.whatsapp.com/"
        );

        System.out.println(
                "\n⚠️ PLEASE SCAN QR CODE..."
        );

        Thread.sleep(20000);

        // ================================================================
        // OPEN RXDX CHAT
        // ================================================================

        WebElement searchBox =
                wait.until(
                        ExpectedConditions
                                .elementToBeClickable(
                                        By.xpath(
                                                "//input[@placeholder='Search or start a new chat']"
                                        )
                                )
                );

        searchBox.click();

        Thread.sleep(2000);

        searchBox.sendKeys(
                "916363415530"
        );

        Thread.sleep(2000);

        wait.until(
                ExpectedConditions
                        .elementToBeClickable(
                                By.xpath(
                                        "(//span[normalize-space()='RxDx Healthcare'])[1]"
                                )
                        )
        ).click();

        Thread.sleep(3000);

        // ================================================================
        // FAILED TEST STORAGE
        // ================================================================

        List<Map<String, String>> failedTestsData =
                new ArrayList<>();

        int testNumber = 1;

        // ================================================================
        // MAIN LOOP
        // ================================================================

        for (Map.Entry<String, List<String>> entry :
                departmentQuestions.entrySet()) {

            String expectedDepartment =
                    entry.getKey();

            List<String> questions =
                    entry.getValue();

            for (int questionNumber = 0;
                 questionNumber < questions.size();
                 questionNumber++) {

                String question =
                        questions.get(questionNumber);

                String botResponse = "";

                String locationResponse = "";

                String departmentsShown = "None";

                String status = "PASS";

                String mappingStatus = "PASS";

                String filteringStatus = "PASS";

                String failureType = "";

                String failureReason = "";

                String screenshotPath = "";

                String testCase =
                        "Q" + (questionNumber + 1);

                System.out.println(
                        "\n"
                                + "=".repeat(90)
                );

                System.out.println(
                        "TEST #"
                                + testNumber
                                + " / "
                                + totalTests
                );

                System.out.println(
                        "Department: "
                                + expectedDepartment
                );

                System.out.println(
                        "Question: "
                                + testCase
                                + " - "
                                + question
                );

                System.out.println(
                        "=".repeat(90)
                );

                try {

                    // ====================================================
                    // STEP 1 - START FLOW
                    // ====================================================

                    System.out.println(
                            "\n📤 STEP 1: Sending HI"
                    );

                    sendMessage(
                            driver,
                            "hi",
                            wait
                    );

                    Thread.sleep(5000);

                    waitForMainMenu(
                            driver,
                            wait
                    );

                    System.out.println(
                            "✅ Main Menu received"
                    );

                    Thread.sleep(5000);

                    // ====================================================
                    // STEP 2 - SEND QUESTION
                    // ====================================================

                    System.out.println(
                            "\n📤 STEP 2: Sending "
                                    + testCase
                                    + ": "
                                    + question
                    );

                    sendMessage(
                            driver,
                            question,
                            wait
                    );

                    System.out.println(
                            "⏳ Waiting for GPT response..."
                    );

                    boolean gptReceived =
                            waitForGPTResponse(
                                    driver,
                                    wait
                            );

                    if (!gptReceived) {

                        status = "FAIL";

                        mappingStatus = "FAIL";

                        failureType =
                                "GPT_RESPONSE_TIMEOUT";

                        failureReason =
                                "GPT response was not received within 120 seconds.";

                        screenshotPath =
                                takeScreenshot(
                                        driver,
                                        screenshotsPath,
                                        expectedDepartment
                                                + "_"
                                                + testCase
                                                + "_GPTTimeout"
                                );

                        saveResult(
                                sheet,
                                excelRowNum++,
                                testNumber,
                                expectedDepartment,
                                testCase,
                                question,
                                botResponse,
                                locationResponse,
                                status,
                                mappingStatus,
                                filteringStatus,
                                failureType,
                                failureReason,
                                departmentsShown,
                                screenshotPath
                        );

                        addFailure(
                                failedTestsData,
                                testNumber,
                                expectedDepartment,
                                testCase,
                                failureType,
                                failureReason,
                                screenshotPath
                        );

                        testNumber++;

                        continue;
                    }

                    Thread.sleep(5000);

                    botResponse =
                            getLastMessage(
                                    driver
                            );

                    System.out.println(
                            "\n🤖 GPT RESPONSE:\n"
                                    + botResponse
                    );

                    // ====================================================
                    // VALIDATION 1
                    // GPT RESPONSE CONTENT
                    // ====================================================

                    ValidationResult gptValidation =
                            validateGPTResponse(
                                    botResponse
                            );

                    if (!gptValidation.valid) {

                        status = "FAIL";

                        mappingStatus = "FAIL";

                        failureType =
                                gptValidation.failureType;

                        failureReason =
                                gptValidation.reason;

                        screenshotPath =
                                takeScreenshot(
                                        driver,
                                        screenshotsPath,
                                        expectedDepartment
                                                + "_"
                                                + testCase
                                                + "_GPTValidation"
                                );

                        saveResult(
                                sheet,
                                excelRowNum++,
                                testNumber,
                                expectedDepartment,
                                testCase,
                                question,
                                botResponse,
                                locationResponse,
                                status,
                                mappingStatus,
                                filteringStatus,
                                failureType,
                                failureReason,
                                departmentsShown,
                                screenshotPath
                        );

                        addFailure(
                                failedTestsData,
                                testNumber,
                                expectedDepartment,
                                testCase,
                                failureType,
                                failureReason,
                                screenshotPath
                        );

                        testNumber++;

                        continue;
                    }

                    System.out.println(
                            "✅ GPT response validation PASS"
                    );

                    // ====================================================
                    // STEP 3 - BOOK APPOINTMENT
                    // ====================================================

                    System.out.println(
                            "\n🔍 STEP 3: Looking for Book an Appointment"
                    );

                    boolean buttonClicked =
                            clickBookAppointment(
                                    driver
                            );

                    if (!buttonClicked) {

                        status = "FAIL";

                        failureType =
                                "BOOK_APPOINTMENT_BUTTON_MISSING";

                        failureReason =
                                "Book an Appointment button was not found.";

                        screenshotPath =
                                takeScreenshot(
                                        driver,
                                        screenshotsPath,
                                        expectedDepartment
                                                + "_"
                                                + testCase
                                                + "_NoBookButton"
                                );

                        saveResult(
                                sheet,
                                excelRowNum++,
                                testNumber,
                                expectedDepartment,
                                testCase,
                                question,
                                botResponse,
                                locationResponse,
                                status,
                                mappingStatus,
                                filteringStatus,
                                failureType,
                                failureReason,
                                departmentsShown,
                                screenshotPath
                        );

                        addFailure(
                                failedTestsData,
                                testNumber,
                                expectedDepartment,
                                testCase,
                                failureType,
                                failureReason,
                                screenshotPath
                        );

                        testNumber++;

                        continue;
                    }

                    System.out.println(
                            "✅ Book an Appointment clicked"
                    );

                    Thread.sleep(5000);

                    // ====================================================
                    // STEP 4 - WHITEFIELD
                    // ====================================================

                    System.out.println(
                            "\n📤 STEP 4: Selecting Whitefield using 1"
                    );

                    sendMessage(
                            driver,
                            "1",
                            wait
                    );

                    Thread.sleep(5000);

                    locationResponse =
                            getLastMessage(
                                    driver
                            );

                    System.out.println(
                            "\n📍 LOCATION RESPONSE:\n"
                                    + locationResponse
                    );

                    // ====================================================
                    // VALIDATION 2
                    // EXTRACT DEPARTMENTS
                    // ====================================================

                    Set<String> detectedDepartments =
                            extractDepartments(
                                    locationResponse
                            );

                    departmentsShown =
                            detectedDepartments.isEmpty()
                                    ? "None"
                                    : String.join(
                                            ", ",
                                            detectedDepartments
                                    );

                    System.out.println(
                            "\n🔎 Departments detected: "
                                    + departmentsShown
                    );

                    // ====================================================
                    // VALIDATION 3
                    // MAX 2 DEPARTMENTS
                    // ====================================================

                    if (detectedDepartments.size() > 2) {

                        status = "FAIL";

                        filteringStatus = "FAIL";

                        failureType =
                                "TOO_MANY_DEPARTMENTS";

                        failureReason =
                                "Expected maximum 2 departments, but "
                                        + detectedDepartments.size()
                                        + " were found: "
                                        + departmentsShown;
                    }

                    // ====================================================
                    // VALIDATION 4
                    // ACCEPTABLE MAPPING
                    // ====================================================

                    if (status.equals("PASS")) {

                        Set<String> allowedDepartments =
                                acceptableDepartments.get(
                                        expectedDepartment
                                );

                        if (allowedDepartments == null) {

                            allowedDepartments =
                                    new HashSet<>();

                            allowedDepartments.add(
                                    expectedDepartment
                            );
                        }

                        boolean mappingPassed =
                                false;

                        for (String detected :
                                detectedDepartments) {

                            if (allowedDepartments
                                    .contains(detected)) {

                                mappingPassed = true;

                                break;
                            }
                        }

                        if (!mappingPassed) {

                            status = "FAIL";

                            mappingStatus = "FAIL";

                            failureType =
                                    "DEPARTMENT_MAPPING_FAILED";

                            failureReason =
                                    "Expected department '"
                                            + expectedDepartment
                                            + "' or an accepted mapped department "
                                            + allowedDepartments
                                            + " but found: "
                                            + departmentsShown;
                        }
                    }

                    // ====================================================
                    // VALIDATION 5
                    // RESPONSE NOT EMPTY
                    // ====================================================

                    if (status.equals("PASS")
                            && !hasMeaningfulText(
                                    locationResponse
                            )) {

                        status = "FAIL";

                        filteringStatus = "FAIL";

                        failureType =
                                "EMPTY_LOCATION_RESPONSE";

                        failureReason =
                                "Location response was empty or invalid.";
                    }

                    // ====================================================
                    // FINAL STATUS
                    // ====================================================

                    if (status.equals("PASS")) {

                        System.out.println(
                                "\n🎉 TEST PASSED"
                        );

                        System.out.println(
                                "Department: "
                                        + expectedDepartment
                        );

                        System.out.println(
                                "Question: "
                                        + testCase
                        );

                        System.out.println(
                                "Departments shown: "
                                        + departmentsShown
                        );

                    } else {

                        System.out.println(
                                "\n❌ TEST FAILED"
                        );

                        System.out.println(
                                "Failure Type: "
                                        + failureType
                        );

                        System.out.println(
                                "Reason: "
                                        + failureReason
                        );

                        screenshotPath =
                                takeScreenshot(
                                        driver,
                                        screenshotsPath,
                                        expectedDepartment
                                                + "_"
                                                + testCase
                                );
                    }

                } catch (Exception e) {

                    status = "FAIL";

                    failureType =
                            "AUTOMATION_EXCEPTION";

                    failureReason =
                            e.getClass().getSimpleName()
                                    + ": "
                                    + e.getMessage();

                    System.out.println(
                            "\n❌ EXCEPTION: "
                                    + failureReason
                    );

                    screenshotPath =
                            takeScreenshot(
                                    driver,
                                    screenshotsPath,
                                    expectedDepartment
                                            + "_"
                                            + testCase
                                            + "_Exception"
                            );
                }

                // ========================================================
                // SAVE RESULT
                // ========================================================

                saveResult(
                        sheet,
                        excelRowNum++,
                        testNumber,
                        expectedDepartment,
                        testCase,
                        question,
                        botResponse,
                        locationResponse,
                        status,
                        mappingStatus,
                        filteringStatus,
                        failureType,
                        failureReason,
                        departmentsShown,
                        screenshotPath
                );

                // ========================================================
                // STORE FAILURE
                // ========================================================

                if (status.equals("FAIL")) {

                    addFailure(
                            failedTestsData,
                            testNumber,
                            expectedDepartment,
                            testCase,
                            failureType,
                            failureReason,
                            screenshotPath
                    );
                }

                testNumber++;

                Thread.sleep(5000);
            }
        }

        // ================================================================
        // EXCEL FORMATTING
        // ================================================================

        for (int i = 0; i <= 13; i++) {

            sheet.autoSizeColumn(i);
        }

        // ================================================================
        // SAVE EXCEL
        // ================================================================

        try (
                FileOutputStream fileOut =
                        new FileOutputStream(
                                excelFilePath
                        )
        ) {

            workbook.write(fileOut);

            System.out.println(
                    "\n✅ EXCEL SAVED:"
                            + "\n"
                            + excelFilePath
            );
        }

        workbook.close();

        // ================================================================
        // FINAL REPORT
        // ================================================================

        printFinalReport(
                totalTests,
                failedTestsData
        );

        // ================================================================
        // OPEN EXCEL
        // ================================================================

        try {

            File excelFile =
                    new File(
                            excelFilePath
                    );

            if (Desktop.isDesktopSupported()) {

                Desktop
                        .getDesktop()
                        .open(excelFile);
            }

        } catch (Exception e) {

            System.out.println(
                    "⚠️ Could not open Excel automatically."
            );
        }

        // ================================================================
        // CLOSE DRIVER
        // ================================================================

        driver.quit();

        System.out.println(
                "\n🏁 AUTOMATION COMPLETED"
        );
    }

    // ================================================================
    // ADD TWO QUESTIONS
    // ================================================================

    static void addQuestions(
            Map<String, List<String>> map,
            String department,
            String question1,
            String question2) {

        map.put(
                department,
                Arrays.asList(
                        question1,
                        question2
                )
        );
    }

    // ================================================================
    // WAIT FOR MAIN MENU
    // ================================================================

    static void waitForMainMenu(
            WebDriver driver,
            WebDriverWait wait) {

        wait.until(
                ExpectedConditions
                        .presenceOfElementLocated(
                                By.xpath(
                                        "//*[contains(text(),'Main Menu')]"
                                )
                        )
        );
    }

    // ================================================================
    // WAIT FOR GPT RESPONSE
    // ================================================================

    static boolean waitForGPTResponse(
            WebDriver driver,
            WebDriverWait wait) {

        try {

            wait.until(
                    ExpectedConditions
                            .presenceOfElementLocated(
                                    By.xpath(
                                            "//em[contains(text(),'This Response was generated using ChatGPT.')]"
                                    )
                            )
            );

            return true;

        } catch (TimeoutException e) {

            return false;
        }
    }

    // ================================================================
    // GPT VALIDATION
    // ================================================================

    static ValidationResult validateGPTResponse(
            String response) {

        if (response == null
                || response.trim().isEmpty()) {

            return new ValidationResult(
                    false,
                    "GPT_RESPONSE_EMPTY",
                    "GPT response is empty."
            );
        }

        if (response.equalsIgnoreCase(
                "NO RESPONSE"
        )) {

            return new ValidationResult(
                    false,
                    "GPT_RESPONSE_MISSING",
                    "No bot response was captured."
            );
        }

        if (!response.contains(
                GPT_FOOTER
        )) {

            return new ValidationResult(
                    false,
                    "GPT_FOOTER_MISSING",
                    "GPT footer was not found."
            );
        }

        if (response.trim().length() < 50) {

            return new ValidationResult(
                    false,
                    "GPT_RESPONSE_TOO_SHORT",
                    "GPT response contains insufficient content."
            );
        }

        return new ValidationResult(
                true,
                "",
                ""
        );
    }

    // ================================================================
    // CLICK BOOK APPOINTMENT
    // ================================================================

    static boolean clickBookAppointment(
            WebDriver driver) {

        // Strategy 1
        try {

            List<WebElement> elements =
                    driver.findElements(
                            By.xpath(
                                    "(//span[contains(normalize-space(),'Book an Appointment')])[last()]"
                            )
                    );

            for (WebElement element :
                    elements) {

                if (element.isDisplayed()
                        && element.isEnabled()) {

                    element.click();

                    return true;
                }
            }

        } catch (Exception e) {
        }

        // Strategy 2
        try {

            List<WebElement> elements =
                    driver.findElements(
                            By.xpath(
                                    "//*[contains(normalize-space(),'Book an Appointment')]"
                            )
                    );

            for (WebElement element :
                    elements) {

                if (element.isDisplayed()
                        && element.isEnabled()) {

                    element.click();

                    return true;
                }
            }

        } catch (Exception e) {
        }

        // Strategy 3 - JavaScript
        try {

            WebElement element =
                    driver.findElement(
                            By.xpath(
                                    "//*[contains(normalize-space(),'Book an Appointment')]"
                            )
                    );

            ((JavascriptExecutor) driver)
                    .executeScript(
                            "arguments[0].click();",
                            element
                    );

            return true;

        } catch (Exception e) {
        }

        // Strategy 4 - Actions
        try {

            WebElement element =
                    driver.findElement(
                            By.xpath(
                                    "//*[contains(normalize-space(),'Book an Appointment')]"
                            )
                    );

            org.openqa.selenium.interactions.Actions actions =
                    new org.openqa.selenium.interactions.Actions(
                            driver
                    );

            actions
                    .moveToElement(element)
                    .click()
                    .perform();

            return true;

        } catch (Exception e) {
        }

        return false;
    }

    // ================================================================
    // SEND MESSAGE
    // ================================================================

    static void sendMessage(
            WebDriver driver,
            String text,
            WebDriverWait wait)
            throws InterruptedException {

        WebElement input =
                wait.until(
                        ExpectedConditions
                                .elementToBeClickable(
                                        By.xpath(
                                                "//div[@aria-placeholder='Type a message']"
                                        )
                                )
                );

        input.click();

        input.sendKeys(
                text
        );

        Thread.sleep(1000);

        input.sendKeys(
                Keys.ENTER
        );

        Thread.sleep(2000);
    }

    // ================================================================
    // GET LAST MESSAGE
    // ================================================================

    static String getLastMessage(
            WebDriver driver) {

        String[] xpaths = {

                "//div[@data-testid='msg-container']",

                "//div[contains(@class,'message-in')]",

                "//div[contains(@class,'_akbu')]",

                "//div[contains(@class,'copyable-text')]"
        };

        for (String xpath :
                xpaths) {

            try {

                List<WebElement> messages =
                        driver.findElements(
                                By.xpath(xpath)
                        );

                if (!messages.isEmpty()) {

                    WebElement last =
                            messages.get(
                                    messages.size() - 1
                            );

                    String text =
                            last.getText();

                    if (text != null
                            && !text.trim().isEmpty()) {

                        return text.trim();
                    }
                }

            } catch (Exception e) {
            }
        }

        return "NO RESPONSE";
    }

    // ================================================================
    // EXTRACT DEPARTMENTS
    // ================================================================

    static Set<String> extractDepartments(
            String response) {

        Set<String> found =
                new LinkedHashSet<>();

        if (response == null) {

            return found;
        }

        String lowerResponse =
                response.toLowerCase();

        // Sort longest first so that
        // "Paediatric and Fetal Cardiologist"
        // is checked before "Adult Cardiology", etc.

        List<String> departments =
                new ArrayList<>(
                        getAllDepartments()
                );

        departments.sort(
                Comparator.comparingInt(
                        String::length
                ).reversed()
        );

        for (String department :
                departments) {

            if (lowerResponse.contains(
                    department.toLowerCase()
            )) {

                found.add(
                        department
                );
            }
        }

        return found;
    }

    // ================================================================
    // ALL DEPARTMENTS
    // ================================================================

    static Set<String> getAllDepartments() {

        Set<String> departments =
                new LinkedHashSet<>();

        departments.add("Acupuncture");
        departments.add("Adult Cardiology");
        departments.add("Allergy Specialist");
        departments.add("Clinical Nutrition");
        departments.add("Dental");
        departments.add("Dermatology");
        departments.add("Diabetologist");
        departments.add("Endocrinologist");
        departments.add("Endodentist");
        departments.add("Endodontist");
        departments.add(
                "Endovascular Neurosurgeon, Vascular Surgeon"
        );
        departments.add("ENT");
        departments.add(
                "Gastroenterologist, Hepatologist"
        );
        departments.add("General Medicine");
        departments.add("General Surgeon");
        departments.add("Homeopathy");
        departments.add("Internal Medicine");
        departments.add("Nephrologist");
        departments.add("Neurologist");
        departments.add("Neurosurgery");
        departments.add("Obs & Gynae");
        departments.add("Occupational Therapy");
        departments.add("Ophthalmology");
        departments.add("Oral Surgeon");
        departments.add("Orthodentists");
        departments.add("Orthopedic Surgeon");
        departments.add("Orthotist");
        departments.add(
                "Paediatric and Fetal Cardiologist"
        );
        departments.add("Paediatric Cardiologist");
        departments.add("Paediatric Neurologist");
        departments.add("Paediatric Surgeon");
        departments.add("Paediatrician");
        departments.add("Pain medicine");
        departments.add("Pedodentists");
        departments.add("Periodontist");
        departments.add("Physiotherapy");
        departments.add("Plastic Surgery");
        departments.add("Podiatric Surgeon");
        departments.add("Psychiatry");
        departments.add("Psychology");
        departments.add("Pulmonologist");
        departments.add("Radiology");
        departments.add("Rheumatology");
        departments.add("Speech Therapy");
        departments.add("Urologist");

        return departments;
    }

    // ================================================================
    // MEANINGFUL TEXT VALIDATION
    // ================================================================

    static boolean hasMeaningfulText(
            String text) {

        if (text == null) {
            return false;
        }

        String cleaned =
                text.replaceAll(
                        "\\s+",
                        " "
                ).trim();

        return cleaned.length() >= 10;
    }

    // ================================================================
    // SCREENSHOT
    // ================================================================

    static String takeScreenshot(
            WebDriver driver,
            String folder,
            String name) {

        try {

            String cleanName =
                    name.replaceAll(
                            "[^a-zA-Z0-9_-]",
                            "_"
                    );

            String path =
                    folder
                            + "/"
                            + cleanName
                            + "_"
                            + System.currentTimeMillis()
                            + ".png";

            File source =
                    ((TakesScreenshot) driver)
                            .getScreenshotAs(
                                    OutputType.FILE
                            );

            File destination =
                    new File(path);

            if (source.renameTo(
                    destination
            )) {

                System.out.println(
                        "📸 Screenshot: "
                                + path
                );

                return path;
            }

            return "Screenshot rename failed";

        } catch (Exception e) {

            return "Screenshot failed: "
                    + e.getMessage();
        }
    }

    // ================================================================
    // SAVE EXCEL RESULT
    // ================================================================

    static void saveResult(
            Sheet sheet,
            int rowNumber,
            int testNumber,
            String department,
            String testCase,
            String question,
            String botResponse,
            String locationResponse,
            String status,
            String mappingStatus,
            String filteringStatus,
            String failureType,
            String failureReason,
            String departmentsShown,
            String screenshot) {

        Row row =
                sheet.createRow(
                        rowNumber
                );

        row.createCell(0)
                .setCellValue(
                        testNumber
                );

        row.createCell(1)
                .setCellValue(
                        department
                );

        row.createCell(2)
                .setCellValue(
                        testCase
                );

        row.createCell(3)
                .setCellValue(
                        question
                );

        row.createCell(4)
                .setCellValue(
                        botResponse
                );

        row.createCell(5)
                .setCellValue(
                        locationResponse
                );

        row.createCell(6)
                .setCellValue(
                        status
                );

        row.createCell(7)
                .setCellValue(
                        mappingStatus
                );

        row.createCell(8)
                .setCellValue(
                        filteringStatus
                );

        row.createCell(9)
                .setCellValue(
                        failureType
                );

        row.createCell(10)
                .setCellValue(
                        failureReason
                );

        row.createCell(11)
                .setCellValue(
                        departmentsShown
                );

        row.createCell(12)
                .setCellValue(
                        screenshot
                );
    }

    // ================================================================
    // ADD FAILURE
    // ================================================================

    static void addFailure(
            List<Map<String, String>> failures,
            int testNumber,
            String department,
            String testCase,
            String failureType,
            String reason,
            String screenshot) {

        Map<String, String> data =
                new LinkedHashMap<>();

        data.put(
                "Test",
                String.valueOf(
                        testNumber
                )
        );

        data.put(
                "Department",
                department
        );

        data.put(
                "Question",
                testCase
        );

        data.put(
                "Failure Type",
                failureType
        );

        data.put(
                "Reason",
                reason
        );

        data.put(
                "Screenshot",
                screenshot
        );

        failures.add(
                data
        );
    }

    // ================================================================
    // EXCEL HEADER
    // ================================================================

    static void addProjectHeader(
            Sheet sheet,
            Workbook workbook) {

        Row row =
                sheet.createRow(0);

        Cell cell =
                row.createCell(0);

        cell.setCellValue(
                "RXDX WHITEFIELD HOSPITAL CHATBOT - DEPARTMENT APPOINTMENT VALIDATION"
        );

        CellStyle style =
                workbook.createCellStyle();

        Font font =
                workbook.createFont();

        font.setBold(true);

        font.setFontHeightInPoints(
                (short) 16
        );

        style.setFont(
                font
        );

        cell.setCellStyle(
                style
        );

        sheet.addMergedRegion(
                new CellRangeAddress(
                        0,
                        0,
                        0,
                        12
                )
        );

        Row info =
                sheet.createRow(1);

        info.createCell(0)
                .setCellValue(
                        "Execution Date: "
                                + new SimpleDateFormat(
                                        "yyyy-MM-dd HH:mm:ss"
                                ).format(
                                        new Date()
                                )
                );

        sheet.addMergedRegion(
                new CellRangeAddress(
                        1,
                        1,
                        0,
                        12
                )
        );
    }

    // ================================================================
    // EXCEL COLUMN HEADERS
    // ================================================================

    static int addColumnHeaders(
            Sheet sheet,
            Workbook workbook) {

        Row row =
                sheet.createRow(2);

        String[] headers = {

                "Test No",

                "Department",

                "Test Case",

                "Question",

                "Bot Response",

                "Location Response",

                "Status",

                "Mapping Status",

                "Filtering Status",

                "Failure Type",

                "Failure Reason",

                "Departments Shown",

                "Screenshot"
        };

        for (int i = 0;
             i < headers.length;
             i++) {

            Cell cell =
                    row.createCell(i);

            cell.setCellValue(
                    headers[i]
            );

            CellStyle style =
                    workbook.createCellStyle();

            Font font =
                    workbook.createFont();

            font.setBold(true);

            style.setFont(
                    font
            );

            cell.setCellStyle(
                    style
            );
        }

        return 2;
    }

    // ================================================================
    // FINAL REPORT
    // ================================================================

    static void printFinalReport(
            int totalTests,
            List<Map<String, String>> failures) {

        int failed =
                failures.size();

        int passed =
                totalTests - failed;

        double passPercentage =
                totalTests == 0
                        ? 0
                        : (passed * 100.0)
                        / totalTests;

        System.out.println(
                "\n"
                        + "=".repeat(90)
        );

        System.out.println(
                "                    FINAL TEST REPORT"
        );

        System.out.println(
                "=".repeat(90)
        );

        System.out.println(
                "Total Tests     : "
                        + totalTests
        );

        System.out.println(
                "Passed          : "
                        + passed
        );

        System.out.println(
                "Failed          : "
                        + failed
        );

        System.out.println(
                "Pass Percentage : "
                        + String.format(
                                "%.2f",
                                passPercentage
                        )
                        + "%"
        );

        System.out.println(
                "=".repeat(90)
        );

        if (failures.isEmpty()) {

            System.out.println(
                    "\n🎉 ALL 90 TEST CASES PASSED!"
            );

        } else {

            System.out.println(
                    "\n❌ FAILED TEST CASES:"
            );

            for (Map<String, String> failure :
                    failures) {

                System.out.println(
                        "\nTest #"
                                + failure.get("Test")
                );

                System.out.println(
                        "Department: "
                                + failure.get(
                                        "Department"
                                )
                );

                System.out.println(
                        "Question: "
                                + failure.get(
                                        "Question"
                                )
                );

                System.out.println(
                        "Failure: "
                                + failure.get(
                                        "Failure Type"
                                )
                );

                System.out.println(
                        "Reason: "
                                + failure.get(
                                        "Reason"
                                )
                );

                System.out.println(
                        "Screenshot: "
                                + failure.get(
                                        "Screenshot"
                                )
                );
            }
        }

        System.out.println(
                "\n"
                        + "=".repeat(90)
        );
    }

    // ================================================================
    // VALIDATION RESULT CLASS
    // ================================================================

    static class ValidationResult {

        boolean valid;

        String failureType;

        String reason;

        ValidationResult(
                boolean valid,
                String failureType,
                String reason) {

            this.valid =
                    valid;

            this.failureType =
                    failureType;

            this.reason =
                    reason;
        }
    }
}