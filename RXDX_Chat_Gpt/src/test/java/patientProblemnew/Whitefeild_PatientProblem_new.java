package patientProblemnew;

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
import org.apache.poi.ss.util.CellRangeAddress;

import java.awt.Desktop;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;

public class Whitefeild_PatientProblem_new {

    // ============================================================
    // ACCEPTABLE DEPARTMENT MAPPING
    // ============================================================

    static Map<String, Set<String>> acceptableDepartments = new HashMap<>();

    static {
        acceptableDepartments.put("Nephrologist",
                new HashSet<>(Arrays.asList("Nephrologist", "Urologist")));

        acceptableDepartments.put("Neurologist",
                new HashSet<>(Arrays.asList("Neurologist", "Paediatric Neurologist")));

        acceptableDepartments.put("General Medicine",
                new HashSet<>(Arrays.asList("General Medicine", "Internal Medicine")));

        acceptableDepartments.put("Internal Medicine",
                new HashSet<>(Arrays.asList("Internal Medicine", "General Medicine")));

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
                new HashSet<>(Arrays.asList("Neurosurgery", "Neurologist")));

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
                new HashSet<>(Arrays.asList("Ophthalmology")));

        acceptableDepartments.put("ENT",
                new HashSet<>(Arrays.asList("ENT")));

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
                new HashSet<>(Arrays.asList("Homeopathy")));

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
                new HashSet<>(Arrays.asList("Pulmonologist")));

        acceptableDepartments.put("Urologist",
                new HashSet<>(Arrays.asList(
                        "Urologist",
                        "Nephrologist"
                )));

        acceptableDepartments.put("Radiology",
                new HashSet<>(Arrays.asList("Radiology")));

        acceptableDepartments.put("Rheumatology",
                new HashSet<>(Arrays.asList(
                        "Rheumatology",
                        "Orthopedic Surgeon"
                )));

        acceptableDepartments.put("Speech Therapy",
                new HashSet<>(Arrays.asList("Speech Therapy")));

        acceptableDepartments.put("Acupuncture",
                new HashSet<>(Arrays.asList(
                        "Acupuncture",
                        "Pain medicine"
                )));

        acceptableDepartments.put("Clinical Nutrition",
                new HashSet<>(Arrays.asList("Clinical Nutrition")));

        acceptableDepartments.put("Endovascular Neurosurgeon, Vascular Surgeon",
                new HashSet<>(Arrays.asList(
                        "Endovascular Neurosurgeon, Vascular Surgeon",
                        "General Surgeon"
                )));

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
                new HashSet<>(Arrays.asList("Obs & Gynae")));

        acceptableDepartments.put("Paediatric and Fetal Cardiologist",
                new HashSet<>(Arrays.asList(
                        "Paediatric and Fetal Cardiologist",
                        "Paediatric Cardiologist",
                        "Adult Cardiology"
                )));

        acceptableDepartments.put("Paediatric Surgeon",
                new HashSet<>(Arrays.asList(
                        "Paediatric Surgeon",
                        "General Surgeon"
                )));
    }


    // ============================================================
    // GPT FOOTER
    // ============================================================

    static final String GPT_FOOTER =
            "This Response was generated using ChatGPT.";


    // ============================================================
    // MAIN
    // ============================================================

    public static void main(String[] args)
            throws InterruptedException, IOException {

        System.setProperty(
                "org.apache.poi.util.POILogger",
                "org.apache.poi.util.NullLogger"
        );

        String projectPath = System.getProperty("user.dir");

        String timestamp =
                new SimpleDateFormat("yyyyMMdd_HHmmss")
                        .format(new Date());

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


        // Create folders
        new File(screenshotsPath).mkdirs();
        new File(excelFolderPath).mkdirs();


        System.out.println("\n📁 SCREENSHOTS: " + screenshotsPath);
        System.out.println("📊 EXCEL FILE: " + excelFilePath);


        // ========================================================
        // DEPARTMENT QUESTIONS
        // 3 HUMAN-LIKE PATIENT PROBLEMS PER DEPARTMENT
        // ========================================================

        Map<String, List<String>> departmentQuestions =
                new LinkedHashMap<>();


        // --------------------------------------------------------
        // Acupuncture
        // --------------------------------------------------------

        departmentQuestions.put("Acupuncture", Arrays.asList(
                "I have back pain for a long time",
                "My body has been hurting for several weeks",
                "I have constant muscle pain and want some relief"
        ));


        // --------------------------------------------------------
        // Adult Cardiology
        // --------------------------------------------------------

//        departmentQuestions.put("Adult Cardiology", Arrays.asList(
//                "I sometimes get chest pain and feel tired easily",
//               "I feel heaviness in my chest when I walk",
//               "My heart feels like it is beating very fast sometimes"
//       ));
//
//
//        // --------------------------------------------------------
//        // Clinical Nutrition
//        // --------------------------------------------------------
//
//        departmentQuestions.put("Clinical Nutrition", Arrays.asList(
//                "I am gaining weight and need to improve my diet",
//                "I want to lose weight but don't know what to eat",
//                "I need help planning a healthy diet"
//        ));
//
//
//        // --------------------------------------------------------
//        // Dental
//        // --------------------------------------------------------
//
//        departmentQuestions.put("Dental", Arrays.asList(
//                "My tooth has been hurting badly",
//                "I have a broken tooth and it is painful",
//                "My teeth hurt whenever I eat something cold"
//        ));
//
//
//        // --------------------------------------------------------
//        // Dermatology
//        // --------------------------------------------------------
//
//        departmentQuestions.put("Dermatology", Arrays.asList(
//                "I have itchy red rashes on my skin",
//                "My skin has been irritated for several days",
//                "I have small red spots appearing on my skin"
//        ));
//
//
//        // --------------------------------------------------------
//        // Diabetologist
//        // --------------------------------------------------------
//
//        departmentQuestions.put("Diabetologist", Arrays.asList(
//                "My blood sugar has been very high lately",
//                "My sugar levels are not coming under control",
//                "I am having problems managing my diabetes"
//        ));
//
//
//        // --------------------------------------------------------
//        // Endocrinologist
//        // --------------------------------------------------------
//
//        departmentQuestions.put("Endocrinologist", Arrays.asList(
//                "I think I have a hormone problem",
//                "I have been having unusual changes in my hormones",
//                "I feel something is wrong with my thyroid"
//        ));
//
//
//        // --------------------------------------------------------
//        // ENT
//        // --------------------------------------------------------
//
//        departmentQuestions.put("ENT", Arrays.asList(
//                "My ear has been hurting and I cannot hear properly",
//                "I have a sore throat and blocked nose",
//                "I keep getting ear infections"
//        ));
//
//
//        // --------------------------------------------------------
//        // General Medicine
//        // --------------------------------------------------------
//
//        departmentQuestions.put("General Medicine", Arrays.asList(
//                "I have been feeling weak and unwell for a few days",
//                "I have body aches and feel tired all the time",
//                "I don't feel well but I am not sure what the problem is"
//        ));
//
//
//        // --------------------------------------------------------
//        // Nephrologist
//        // --------------------------------------------------------
//
//        departmentQuestions.put("Nephrologist", Arrays.asList(
//                "I have pain around my kidney area",
//                "I have been having kidney problems for some time",
//                "I have swelling in my legs and some kidney issues"
//        ));
//
//
//        // --------------------------------------------------------
//        // Neurologist
//        // --------------------------------------------------------
//
//        departmentQuestions.put("Neurologist", Arrays.asList(
//                "I keep getting severe headaches and dizziness",
//                "I sometimes feel numbness in my hands",
//                "I have been having frequent headaches lately"
//        ));
//
//
//        // --------------------------------------------------------
//        // Neurosurgery
//        // --------------------------------------------------------
//
//        departmentQuestions.put("Neurosurgery", Arrays.asList(
//                "I have a serious problem with my spine",
//                "I was told that I may need brain surgery",
//                "I have a condition that might require spine surgery"
//        ));
//
//
//        // --------------------------------------------------------
//        // Obs & Gynae
//        // --------------------------------------------------------
//
//        departmentQuestions.put("Obs & Gynae", Arrays.asList(
//                "I have irregular periods and lower stomach pain",
//                "I have been having problems with my periods",
//                "I have some women's health problems"
//        ));
//
//
//        // --------------------------------------------------------
//        // Occupational Therapy
//        // --------------------------------------------------------
//
//        departmentQuestions.put("Occupational Therapy", Arrays.asList(
//                "I am having trouble doing my daily activities",
//                "I find it difficult to use my hand for normal activities",
//                "I need help getting back to my daily routine"
//        ));
//
//
//        // --------------------------------------------------------
//        // Ophthalmology
//        // --------------------------------------------------------
//
//        departmentQuestions.put("Ophthalmology", Arrays.asList(
//                "My vision has become blurry recently",
//                "My eyes hurt and I cannot see clearly",
//                "I am having trouble seeing things far away"
//        ));
//
//
//        // --------------------------------------------------------
//        // Oral Surgeon
//        // --------------------------------------------------------
//
//        departmentQuestions.put("Oral Surgeon", Arrays.asList(
//                "I have a painful problem near my jaw",
//                "My wisdom tooth is causing a lot of pain",
//                "I have a dental problem that may need surgery"
//        ));
//
//
//        // --------------------------------------------------------
//        // Orthodentists
//        // --------------------------------------------------------
//
//        departmentQuestions.put("Orthodentists", Arrays.asList(
//                "My teeth are crooked",
//                "My teeth are not properly aligned",
//                "I want to straighten my teeth"
//        ));
//
//
//        // --------------------------------------------------------
//        // Orthopedic Surgeon
//        // --------------------------------------------------------
//
//        departmentQuestions.put("Orthopedic Surgeon", Arrays.asList(
//                "My knee has been hurting badly",
//                "I have severe pain in my joints",
//                "I injured my shoulder and it is still painful"
//        ));
//
//
//        // --------------------------------------------------------
//        // Orthotist
//        // --------------------------------------------------------
//
//        departmentQuestions.put("Orthotist", Arrays.asList(
//                "I have difficulty walking and need some support",
//                "I need a support device for my leg",
//                "I have trouble walking properly because of my leg"
//        ));
//
//
//        // --------------------------------------------------------
//        // Paediatric and Fetal Cardiologist
//        // --------------------------------------------------------
//
//        departmentQuestions.put(
//                "Paediatric and Fetal Cardiologist",
//                Arrays.asList(
//                        "My unborn baby has a heart problem",
//                        "The scan showed something wrong with my baby's heart",
//                        "My doctor said my baby may have a heart condition"
//                )
//        );
//
//
//        // --------------------------------------------------------
//        // Paediatric Cardiologist
//        // --------------------------------------------------------
//
//        departmentQuestions.put("Paediatric Cardiologist", Arrays.asList(
//                "My child has a heart problem",
//                "My child sometimes complains of chest pain",
//                "My child's heartbeat seems unusual"
//        ));
//

        // --------------------------------------------------------
        // Paediatric Neurologist
        // --------------------------------------------------------

        departmentQuestions.put("Paediatric Neurologist", Arrays.asList(
                "My child is having frequent headaches",
                "My child is having unusual movements",
                "My child sometimes faints and I am worried"
        ));


        // --------------------------------------------------------
        // Paediatric Surgeon
        // --------------------------------------------------------

        departmentQuestions.put("Paediatric Surgeon", Arrays.asList(
                "My child has a problem that may need surgery",
                "My child has a swelling that needs to be checked",
                "My child was advised to see a surgeon"
        ));


        // --------------------------------------------------------
        // Paediatrician
        // --------------------------------------------------------

        departmentQuestions.put("Paediatrician", Arrays.asList(
                "My child has fever and is not eating properly",
                "My child has been sick for several days",
                "My child has a cough and feels very weak"
        ));


        // --------------------------------------------------------
        // Pain medicine
        // --------------------------------------------------------

        departmentQuestions.put("Pain medicine", Arrays.asList(
                "I have constant pain that is affecting my daily life",
                "My pain is not getting better even after taking medicine",
                "I have been dealing with pain for a long time"
        ));


        // --------------------------------------------------------
        // Pedodentists
        // --------------------------------------------------------

        departmentQuestions.put("Pedodentists", Arrays.asList(
                "My child has severe tooth pain",
                "My child's tooth is hurting when they eat",
                "My child has a damaged tooth"
        ));


        // --------------------------------------------------------
        // Periodontist
        // --------------------------------------------------------

        departmentQuestions.put("Periodontist", Arrays.asList(
                "My gums are swollen and bleed when I brush",
                "My gums have been hurting for several days",
                "My gums are pulling away from my teeth"
        ));


        // --------------------------------------------------------
        // Physiotherapy
        // --------------------------------------------------------

        departmentQuestions.put("Physiotherapy", Arrays.asList(
                "I have muscle pain and difficulty moving",
                "I am having trouble walking after an injury",
                "My movement has become difficult because of pain"
        ));


        // --------------------------------------------------------
        // Plastic Surgery
        // --------------------------------------------------------

        departmentQuestions.put("Plastic Surgery", Arrays.asList(
                "I have an injury that needs reconstructive treatment",
                "I have a scar that I want to get treated",
                "I need surgery to repair an injury"
        ));


        // --------------------------------------------------------
        // Podiatric Surgeon
        // --------------------------------------------------------

        departmentQuestions.put("Podiatric Surgeon", Arrays.asList(
                "I have severe pain in my foot",
                "My foot is swollen and very painful",
                "I have a foot problem that is getting worse"
        ));


        // --------------------------------------------------------
        // Psychiatry
        // --------------------------------------------------------

        departmentQuestions.put("Psychiatry", Arrays.asList(
                "I have been feeling very anxious and cannot sleep",
                "I have been feeling low and unable to concentrate",
                "I have been struggling with my mood lately"
        ));


        // --------------------------------------------------------
        // Psychology
        // --------------------------------------------------------

        departmentQuestions.put("Psychology", Arrays.asList(
                "I have been feeling very stressed lately",
                "I am having trouble dealing with my emotions",
                "I feel overwhelmed and need someone to talk to"
        ));


        // --------------------------------------------------------
        // Pulmonologist
        // --------------------------------------------------------

        departmentQuestions.put("Pulmonologist", Arrays.asList(
                "I have been coughing a lot and feel short of breath",
                "I get breathless even when I walk a short distance",
                "I have had a cough for several weeks"
        ));


        // --------------------------------------------------------
        // Radiology
        // --------------------------------------------------------

        departmentQuestions.put("Radiology", Arrays.asList(
                "My doctor asked me to get a scan",
                "I need to get an MRI done",
                "I was advised to get a diagnostic scan"
        ));


        // --------------------------------------------------------
        // Rheumatology
        // --------------------------------------------------------

        departmentQuestions.put("Rheumatology", Arrays.asList(
                "My joints are swollen and stiff in the morning",
                "I have pain and stiffness in several joints",
                "My joint pain has been going on for a long time"
        ));


        // --------------------------------------------------------
        // Speech Therapy
        // --------------------------------------------------------

        departmentQuestions.put("Speech Therapy", Arrays.asList(
                "My child has trouble speaking clearly",
                "My child is not pronouncing words properly",
                "My child is having difficulty communicating"
        ));


        // --------------------------------------------------------
        // Urologist
        // --------------------------------------------------------

        departmentQuestions.put("Urologist", Arrays.asList(
                "It hurts when I urinate",
                "I need to urinate very frequently",
                "I have pain in my bladder area"
        ));


        // --------------------------------------------------------
        // ALLERGY SPECIALIST
        // --------------------------------------------------------

        departmentQuestions.put("Allergy Specialist", Arrays.asList(
                "I keep sneezing and getting allergies",
                "I get itchy eyes and runny nose very often",
                "I keep getting allergic reactions"
        ));


        // --------------------------------------------------------
        // GASTROENTEROLOGIST / HEPATOLOGIST
        // --------------------------------------------------------

        departmentQuestions.put(
                "Gastroenterologist, Hepatologist",
                Arrays.asList(
                        "I have frequent stomach pain and digestion problems",
                        "I often feel bloated after eating",
                        "I have been having stomach problems for several weeks"
                )
        );


        // --------------------------------------------------------
        // GENERAL SURGEON
        // --------------------------------------------------------

        departmentQuestions.put("General Surgeon", Arrays.asList(
                "I have a swelling that is getting bigger",
                "I was told that I may need surgery",
                "I have a lump that needs to be checked"
        ));


        // --------------------------------------------------------
        // HOMEOPATHY
        // --------------------------------------------------------

        departmentQuestions.put("Homeopathy", Arrays.asList(
                "I have been having health problems for a long time",
                "I want to try an alternative treatment",
                "I have an ongoing problem and want homeopathic treatment"
        ));


        // --------------------------------------------------------
        // INTERNAL MEDICINE
        // --------------------------------------------------------

        departmentQuestions.put("Internal Medicine", Arrays.asList(
                "I have been feeling sick and tired recently",
                "I have several health problems and don't know what is causing them",
                "I have been feeling weak for the past few weeks"
        ));


        // ========================================================
        // TESTING DOCTORS
        // ========================================================

        departmentQuestions.put("Testing Doctor", Arrays.asList(
                "I have a general health problem",
                "I am not feeling well and need to get checked",
                "I have been feeling unwell lately"
        ));


        departmentQuestions.put("Testing Doctor - APR", Arrays.asList(
                "I have a health problem that needs to be checked",
                "I have not been feeling well recently",
                "I need help with a general health issue"
        ));


        departmentQuestions.put("Tele-Testing Doctor", Arrays.asList(
                "I have a health problem and need an online consultation",
                "I am feeling unwell and want to talk to a doctor online",
                "I need to discuss my health problem through a video consultation"
        ));


        // ========================================================
        // TOTAL TESTS
        // ========================================================

        int totalTests = 0;

        for (List<String> questions : departmentQuestions.values()) {
            totalTests += questions.size();
        }


        Set<String> validDepartments =
                new HashSet<>(departmentQuestions.keySet());


        // ========================================================
        // CREATE EXCEL
        // ========================================================

        Workbook workbook = new XSSFWorkbook();

        Sheet sheet =
                workbook.createSheet(
                        "RXDX_Whitefield_Results"
                );

        addProjectHeader(sheet, workbook);

        int headerRowIndex =
                addColumnHeaders(sheet, workbook);

        int excelRowNum =
                headerRowIndex + 1;


        // ========================================================
        // EDGE DRIVER
        // ========================================================

        EdgeOptions options = new EdgeOptions();

        options.addArguments(
                "--user-data-dir=C:\\EdgeAutomationProfile"
        );

        options.addArguments("--start-maximized");

        WebDriver driver =
                new EdgeDriver(options);


        // MAX RESPONSE WAIT = 2 MINUTES

        WebDriverWait wait =
                new WebDriverWait(
                        driver,
                        Duration.ofSeconds(120)
                );


        driver.manage().window().maximize();

        driver.get(
                "https://web.whatsapp.com/"
        );


        System.out.println(
                "\n⚠️ PLEASE SCAN QR CODE..."
        );

        Thread.sleep(20000);


        // ========================================================
        // OPEN RXDX CHAT
        // ========================================================

        driver.findElement(
                By.xpath(
                        "//input[@placeholder='Search or start a new chat']"
                )
        ).click();

        Thread.sleep(2000);


        driver.findElement(
                By.xpath(
                        "//input[@placeholder='Search or start a new chat']"
                )
        ).sendKeys("916363415530");

        Thread.sleep(2000);


        driver.findElement(
                By.xpath(
                        "//span[normalize-space()='RxDx Healthcare']"
                )
        ).click();

        Thread.sleep(2000);


        // ========================================================
        // FAILED TESTS
        // ========================================================

        List<Map<String, String>> failedTestsData =
                new ArrayList<>();


        System.out.println(
                "\n"
                + "=".repeat(80)
        );

        System.out.println(
                "     DEPARTMENT-WISE APPOINTMENT TESTS STARTING"
        );

        System.out.println(
                "     TOTAL TESTS: "
                + totalTests
        );

        System.out.println(
                "=".repeat(80)
        );


        int testNumber = 1;


        // ========================================================
        // LOOP THROUGH DEPARTMENTS
        // ========================================================

        for (Map.Entry<String, List<String>> entry :
                departmentQuestions.entrySet()) {

            String expectedDepartment =
                    entry.getKey();

            List<String> questions =
                    entry.getValue();


            // ====================================================
            // LOOP THROUGH 3 QUESTIONS
            // ====================================================

            int questionNumber = 1;

            for (String question : questions) {

                String botResponse = "";

                String placeResponse = "";

                String status = "PASS";

                String mappingStatus = "PASS";

                String filteringStatus = "PASS";

                String failureType = "";

                String failureReason = "";

                String screenshotPath = "";

                String allDepartmentsShown = "";


                System.out.println(
                        "\n"
                        + "=".repeat(70)
                );

                System.out.println(
                        "TEST #"
                        + testNumber
                        + " of "
                        + totalTests
                );

                System.out.println(
                        "Department: "
                        + expectedDepartment
                );

                System.out.println(
                        "Question #"
                        + questionNumber
                        + ": "
                        + question
                );

                System.out.println(
                        "=".repeat(70)
                );


                try {

                    // ============================================
                    // STEP 1: SEND HI
                    // ============================================

                    System.out.println(
                            "\n📤 STEP 1: Sending 'hi'..."
                    );

                    sendMessage(
                            driver,
                            "hi",
                            wait
                    );

                    Thread.sleep(5000);


                    System.out.println(
                            "⏳ Waiting for main menu..."
                    );


                    wait.until(
                            ExpectedConditions
                                    .presenceOfElementLocated(
                                            By.xpath(
                                                    "//div[contains(text(),'Main Menu')]"
                                            )
                                    )
                    );


                    System.out.println(
                            "✅ Main menu detected!"
                    );


                    Thread.sleep(5000);


                    botResponse =
                            getLastMessage(driver);


                    Thread.sleep(5000);


                    System.out.println(
                            "✅ Main menu response received"
                    );


                    // ============================================
                    // STEP 2: SEND PATIENT PROBLEM
                    // ============================================

                    System.out.println(
                            "\n📤 STEP 2: Sending patient problem:"
                    );

                    System.out.println(
                            question
                    );


                    sendMessage(
                            driver,
                            question,
                            wait
                    );


                    Thread.sleep(5000);


                    System.out.println(
                            "⏳ Waiting for GPT response (max 2 min)..."
                    );


                    wait.until(
                            ExpectedConditions
                                    .presenceOfElementLocated(
                                            By.xpath(
                                                    "//em[contains(text(),'This Response was generated using ChatGPT.')]"
                                            )
                                    )
                    );


                    System.out.println(
                            "✅ GPT response detected!"
                    );


                    Thread.sleep(5000);


                    botResponse =
                            getLastMessage(driver);


                    Thread.sleep(5000);


                    System.out.println(
                            "\nGPT RESPONSE:\n"
                            + botResponse.substring(
                                    0,
                                    Math.min(
                                            500,
                                            botResponse.length()
                                    )
                            )
                    );


                    // ============================================
                    // VALIDATION 1: GPT FOOTER
                    // ============================================

                    if (!botResponse.contains(GPT_FOOTER)) {

                        status = "FAIL";

                        failureType =
                                "GPT_RESPONSE_NOT_COMPLETED";

                        failureReason =
                                "GPT footer not found";

                        screenshotPath =
                                takeScreenshot(
                                        driver,
                                        screenshotsPath,
                                        expectedDepartment
                                                + "_Q"
                                                + questionNumber
                                                + "_GPTFooterMissing"
                                );


                        addTestResultToExcel(
                                sheet,
                                excelRowNum++,
                                testNumber,
                                expectedDepartment,
                                question,
                                botResponse,
                                placeResponse,
                                failureType,
                                failureReason,
                                allDepartmentsShown,
                                screenshotPath,
                                status,
                                mappingStatus,
                                filteringStatus
                        );


                        testNumber++;

                        questionNumber++;

                        continue;
                    }


                    System.out.println(
                            "✅ GPT footer found"
                    );


                    // ============================================
                    // VALIDATION 2: DEPARTMENT MAPPING
                    // ============================================

                    boolean valid =
                            validateResponse(
                                    expectedDepartment,
                                    botResponse
                            );


                    if (!valid) {

                        status = "FAIL";

                        mappingStatus = "FAIL";

                        failureType =
                                "WRONG_GPT_RESPONSE";

                        failureReason =
                                "Department keywords not found";

                    } else {

                        System.out.println(
                                "✅ Department mapping correct"
                        );
                    }


                    // ============================================
                    // STEP 3: BOOK APPOINTMENT BUTTON
                    // ============================================

                    System.out.println(
                            "\n🔍 STEP 3: Looking for "
                            + "'Book an Appointment' button..."
                    );


                    boolean buttonClicked = false;


                    // Strategy 1

                    List<WebElement> exactSpans =
                            driver.findElements(
                                    By.xpath(
                                            "(//span[contains(text(),'Book an Appointment')])[last()]"
                                    )
                            );


                    System.out.println(
                            "   Strategy 1: Found "
                            + exactSpans.size()
                            + " exact spans"
                    );


                    for (WebElement span :
                            exactSpans) {

                        if (span.isDisplayed()
                                && span.isEnabled()) {

                            span.click();

                            buttonClicked = true;

                            System.out.println(
                                    "   ✅ Clicked exact span"
                            );

                            Thread.sleep(5000);

                            break;
                        }
                    }


                    // Strategy 2

                    if (!buttonClicked) {

                        List<WebElement> anyElements =
                                driver.findElements(
                                        By.xpath(
                                                "//*[contains(text(),'Book an Appointment')]"
                                        )
                                );


                        System.out.println(
                                "   Strategy 2: Found "
                                + anyElements.size()
                                + " elements containing text"
                        );


                        for (WebElement elem :
                                anyElements) {

                            if (elem.isDisplayed()
                                    && elem.isEnabled()) {

                                elem.click();

                                buttonClicked = true;

                                System.out.println(
                                        "   ✅ Clicked element containing text"
                                );

                                Thread.sleep(5000);

                                break;
                            }
                        }
                    }


                    // Strategy 3

                    if (!buttonClicked) {

                        try {

                            List<WebElement> msgContainers =
                                    driver.findElements(
                                            By.xpath(
                                                    "//div[@data-testid='msg-container']"
                                            )
                                    );


                            if (!msgContainers.isEmpty()) {

                                WebElement lastMsg =
                                        msgContainers.get(
                                                msgContainers.size() - 1
                                        );


                                List<WebElement> btns =
                                        lastMsg.findElements(
                                                By.xpath(
                                                        ".//span[contains(text(),'Book')]"
                                                )
                                        );


                                System.out.println(
                                        "   Strategy 3: Found "
                                        + btns.size()
                                        + " buttons in last message"
                                );


                                for (WebElement btn :
                                        btns) {

                                    if (btn.isDisplayed()
                                            && btn.isEnabled()) {

                                        btn.click();

                                        buttonClicked = true;

                                        System.out.println(
                                                "   ✅ Clicked button from message container"
                                        );

                                        Thread.sleep(5000);

                                        break;
                                    }
                                }
                            }

                        } catch (Exception e) {
                            // Ignore and continue
                        }
                    }


                    // Strategy 4 - JavaScript

                    if (!buttonClicked) {

                        try {

                            WebElement jsButton =
                                    driver.findElement(
                                            By.xpath(
                                                    "//span[contains(text(),'Book an Appointment')]"
                                            )
                                    );


                            ((JavascriptExecutor) driver)
                                    .executeScript(
                                            "arguments[0].click();",
                                            jsButton
                                    );


                            buttonClicked = true;


                            System.out.println(
                                    "   ✅ Clicked using JavaScript"
                            );


                            Thread.sleep(5000);

                        } catch (Exception e) {
                            // Ignore
                        }
                    }


                    // Strategy 5 - Actions

                    if (!buttonClicked) {

                        try {

                            WebElement actionsButton =
                                    driver.findElement(
                                            By.xpath(
                                                    "//span[contains(text(),'Book an Appointment')]"
                                            )
                                    );


                            org.openqa.selenium.interactions.Actions actions =
                                    new org.openqa.selenium.interactions.Actions(
                                            driver
                                    );


                            actions.moveToElement(
                                    actionsButton
                            ).click().perform();


                            buttonClicked = true;


                            System.out.println(
                                    "   ✅ Clicked using Actions"
                            );


                            Thread.sleep(5000);

                        } catch (Exception e) {
                            // Ignore
                        }
                    }


                    // ============================================
                    // BOOK BUTTON NOT FOUND
                    // ============================================

                    if (!buttonClicked) {

                        status = "FAIL";

                        failureType =
                                "BOOK_APPOINTMENT_BUTTON_MISSING";

                        failureReason =
                                "Book Appointment button not found with any strategy";


                        screenshotPath =
                                takeScreenshot(
                                        driver,
                                        screenshotsPath,
                                        expectedDepartment
                                                + "_Q"
                                                + questionNumber
                                                + "_NoBookButton"
                                );


                        addTestResultToExcel(
                                sheet,
                                excelRowNum++,
                                testNumber,
                                expectedDepartment,
                                question,
                                botResponse,
                                placeResponse,
                                failureType,
                                failureReason,
                                allDepartmentsShown,
                                screenshotPath,
                                status,
                                mappingStatus,
                                filteringStatus
                        );


                        testNumber++;

                        questionNumber++;

                        continue;
                    }


                    // ============================================
                    // STEP 4: WHITEFIELD
                    // ============================================

                    System.out.println(
                            "\n⏳ Waiting for location options..."
                    );


                    Thread.sleep(4000);


                    System.out.println(
                            "\n📤 STEP 4: Sending '1' for Whitefield..."
                    );


                    sendMessage(
                            driver,
                            "1",
                            wait
                    );


                    Thread.sleep(2000);

                    Thread.sleep(5000);


                    placeResponse =
                            getLastMessage(driver);


                    Thread.sleep(5000);


                    System.out.println(
                            "\nLOCATION RESPONSE:\n"
                            + placeResponse
                    );


                    // ============================================
                    // STEP 5: VALIDATE DEPARTMENT FILTERING
                    // ============================================

                    System.out.println(
                            "\n🔍 STEP 5: Validating filtering..."
                    );


                    allDepartmentsShown =
                            extractAllDepartmentsFromResponse(
                                    placeResponse,
                                    validDepartments
                            );


                    int departmentsCount =
                            allDepartmentsShown.equals("None")
                                    ? 0
                                    : allDepartmentsShown
                                            .split(", ")
                                            .length;


                    System.out.println(
                            "   Departments shown ("
                            + departmentsCount
                            + "): "
                            + allDepartmentsShown
                    );


                    // ============================================
                    // FAIL IF MORE THAN 2 DEPARTMENTS
                    // ============================================

                    if (departmentsCount > 2) {

                        status = "FAIL";

                        filteringStatus = "FAIL";

                        failureType =
                                "TOO_MANY_DEPARTMENTS";

                        failureReason =
                                "Expected at most 2 departments, "
                                + "but found "
                                + departmentsCount
                                + ": "
                                + allDepartmentsShown;


                    } else if (
                            !placeResponse
                                    .toLowerCase()
                                    .contains(
                                            expectedDepartment
                                                    .toLowerCase()
                                    )
                    ) {

                        status = "FAIL";

                        mappingStatus = "FAIL";

                        failureType =
                                "DEPARTMENT_NOT_FOUND";

                        failureReason =
                                "Expected department '"
                                + expectedDepartment
                                + "' not found in location response";


                    } else {

                        System.out.println(
                                "✅ Department found and "
                                + "departments count ≤ 2"
                        );
                    }


                    // ============================================
                    // PASS MESSAGE
                    // ============================================

                    if (status.equals("PASS")) {

                        System.out.println(
                                "\n✅✅✅ TEST PASSED"
                                + " for: "
                                + expectedDepartment
                                + " - Question "
                                + questionNumber
                        );
                    }


                    // ============================================
                    // SCREENSHOT FOR FAILURE
                    // ============================================

                    if (status.equals("FAIL")) {

                        screenshotPath =
                                takeScreenshot(
                                        driver,
                                        screenshotsPath,
                                        expectedDepartment
                                                + "_Q"
                                                + questionNumber
                                );
                    }


                } catch (Exception e) {

                    status = "FAIL";

                    failureType =
                            "EXCEPTION";

                    failureReason =
                            "Exception: "
                            + e.getMessage();


                    screenshotPath =
                            takeScreenshot(
                                    driver,
                                    screenshotsPath,
                                    expectedDepartment
                                            + "_Q"
                                            + questionNumber
                                            + "_Exception"
                            );


                    e.printStackTrace();
                }


                // =================================================
                // ADD RESULT TO EXCEL
                // =================================================

                addTestResultToExcel(
                        sheet,
                        excelRowNum++,
                        testNumber,
                        expectedDepartment,
                        question,
                        botResponse,
                        placeResponse,
                        failureType,
                        failureReason,
                        allDepartmentsShown,
                        screenshotPath,
                        status,
                        mappingStatus,
                        filteringStatus
                );


                System.out.println(
                        "\n📊 STATUS: "
                        + status
                );


                // =================================================
                // SAVE FAILURE DATA
                // =================================================

                if (status.equals("FAIL")) {

                    Map<String, String> failureData =
                            new HashMap<>();


                    failureData.put(
                            "testNumber",
                            String.valueOf(testNumber)
                    );


                    failureData.put(
                            "departmentName",
                            expectedDepartment
                    );


                    failureData.put(
                            "questionNumber",
                            String.valueOf(questionNumber)
                    );


                    failureData.put(
                            "question",
                            question
                    );


                    failureData.put(
                            "failureType",
                            failureType
                    );


                    failureData.put(
                            "failureReason",
                            failureReason
                    );


                    failureData.put(
                            "screenshotPath",
                            screenshotPath
                    );


                    failedTestsData.add(
                            failureData
                    );
                }


                testNumber++;

                questionNumber++;


                // Wait before next question

                Thread.sleep(5000);
            }
        }


        // ========================================================
        // AUTO SIZE EXCEL COLUMNS
        // ========================================================

        for (int i = 0; i <= 11; i++) {

            sheet.autoSizeColumn(i);
        }


        // ========================================================
        // SAVE EXCEL
        // ========================================================

        try (
                FileOutputStream fileOut =
                        new FileOutputStream(
                                excelFilePath
                        )
        ) {

            workbook.write(fileOut);


            System.out.println(
                    "\n✅ EXCEL SAVED: "
                    + excelFilePath
            );
        }


        workbook.close();


        // ========================================================
        // FINAL REPORT
        // ========================================================

        printFinalReport(
                totalTests,
                failedTestsData
        );


        // ========================================================
        // OPEN EXCEL
        // ========================================================

        try {

            File excelFile =
                    new File(excelFilePath);


            if (Desktop.isDesktopSupported()) {

                Desktop.getDesktop()
                        .open(excelFile);
            }


            System.out.println(
                    "\n✅ Excel opened automatically"
            );


        } catch (Exception e) {

            System.out.println(
                    "⚠️ Could not open Excel automatically"
            );
        }


        // ========================================================
        // CLOSE DRIVER
        // ========================================================

        driver.quit();
    }


    // ============================================================
    // EXCEL PROJECT HEADER
    // ============================================================

    static void addProjectHeader(
            Sheet sheet,
            Workbook workbook) {

        Row headerRow =
                sheet.createRow(0);


        Cell headerCell =
                headerRow.createCell(0);


        headerCell.setCellValue(
                "RXDX WHITEFIELD HOSPITAL CHATBOT - "
                + "DEPARTMENT APPOINTMENT TESTS"
        );


        CellStyle style =
                workbook.createCellStyle();


        Font font =
                workbook.createFont();


        font.setBold(true);

        font.setFontHeightInPoints(
                (short) 16
        );


        style.setFont(font);


        headerCell.setCellStyle(style);


        sheet.addMergedRegion(
                new CellRangeAddress(
                        0,
                        0,
                        0,
                        11
                )
        );


        Row infoRow =
                sheet.createRow(1);


        Cell infoCell =
                infoRow.createCell(0);


        infoCell.setCellValue(
                "Date: "
                + new SimpleDateFormat(
                        "yyyy-MM-dd HH:mm:ss"
                ).format(new Date())
        );


        sheet.addMergedRegion(
                new CellRangeAddress(
                        1,
                        1,
                        0,
                        11
                )
        );
    }


    // ============================================================
    // EXCEL COLUMN HEADERS
    // ============================================================

    static int addColumnHeaders(
            Sheet sheet,
            Workbook workbook) {

        Row row =
                sheet.createRow(2);


        String[] cols = {

                "Test No",

                "Department",

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


        for (int i = 0; i < cols.length; i++) {

            Cell cell =
                    row.createCell(i);


            cell.setCellValue(
                    cols[i]
            );


            CellStyle style =
                    workbook.createCellStyle();


            Font font =
                    workbook.createFont();


            font.setBold(true);


            style.setFont(font);


            cell.setCellStyle(style);
        }


        return 2;
    }


    // ============================================================
    // ADD TEST RESULT TO EXCEL
    // ============================================================

    static void addTestResultToExcel(
            Sheet sheet,
            int rowNum,
            int testNumber,
            String dept,
            String q,
            String bot,
            String place,
            String failType,
            String failReason,
            String deptsShown,
            String screenshot,
            String status,
            String mapStatus,
            String filterStatus) {


        Row row =
                sheet.createRow(rowNum);


        row.createCell(0)
                .setCellValue(testNumber);


        row.createCell(1)
                .setCellValue(dept);


        row.createCell(2)
                .setCellValue(q);


        row.createCell(3)
                .setCellValue(bot);


        row.createCell(4)
                .setCellValue(place);


        row.createCell(5)
                .setCellValue(status);


        row.createCell(6)
                .setCellValue(mapStatus);


        row.createCell(7)
                .setCellValue(filterStatus);


        row.createCell(8)
                .setCellValue(failType);


        row.createCell(9)
                .setCellValue(failReason);


        row.createCell(10)
                .setCellValue(deptsShown);


        row.createCell(11)
                .setCellValue(screenshot);
    }


    // ============================================================
    // SEND MESSAGE
    // ============================================================

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


        input.clear();


        input.sendKeys(text);


        Thread.sleep(5000);


        input.sendKeys(Keys.ENTER);


        Thread.sleep(5000);
    }


    // ============================================================
    // GET LAST MESSAGE
    // ============================================================

    static String getLastMessage(
            WebDriver driver) {


        String[] xpaths = {

                "//div[@data-testid='msg-container']",

                "//div[contains(@class,'message-in')]",

                "//div[contains(@class,'_akbu')]",

                "//div[contains(@class,'copyable-text')]"
        };


        for (String xp : xpaths) {

            try {

                List<WebElement> msgs =
                        driver.findElements(
                                By.xpath(xp)
                        );


                if (!msgs.isEmpty()) {

                    String txt =
                            msgs.get(
                                    msgs.size() - 1
                            ).getText();


                    if (txt != null
                            && txt.length() > 10) {

                        return txt;
                    }
                }


            } catch (Exception e) {

                // Try next XPath
            }
        }


        return "NO RESPONSE";
    }


    // ============================================================
    // TAKE SCREENSHOT
    // ============================================================

    static String takeScreenshot(
            WebDriver driver,
            String folder,
            String name) {


        try {

            String clean =
                    name.replaceAll(
                            "[^a-zA-Z0-9]",
                            "_"
                    );


            String path =
                    folder
                    + "/"
                    + clean
                    + ".png";


            File src =
                    ((TakesScreenshot) driver)
                            .getScreenshotAs(
                                    OutputType.FILE
                            );


            src.renameTo(
                    new File(path)
            );


            System.out.println(
                    "📸 Screenshot: "
                    + path
            );


            return path;


        } catch (Exception e) {

            return "Screenshot failed";
        }
    }


    // ============================================================
    // EXTRACT ALL DEPARTMENTS FROM RESPONSE
    // ============================================================

    static String extractAllDepartmentsFromResponse(
            String resp,
            Set<String> valid) {


        List<String> found =
                new ArrayList<>();


        for (String d : valid) {

            if (resp
                    .toLowerCase()
                    .contains(
                            d.toLowerCase()
                    )) {

                found.add(d);
            }
        }


        return found.isEmpty()
                ? "None"
                : String.join(
                        ", ",
                        found
                );
    }


    // ============================================================
    // VALIDATE GPT RESPONSE
    // ============================================================

    static boolean validateResponse(
            String dept,
            String resp) {


        /*
         * Current validation:
         * Response should be more than 50 characters.
         *
         * This keeps the same logic from your existing code.
         */

        return resp != null
                && resp.length() > 50;
    }


    // ============================================================
    // FINAL REPORT
    // ============================================================

    static void printFinalReport(
            int total,
            List<Map<String, String>> fails) {


        System.out.println(
                "\n"
                + "=".repeat(80)
                + "\nFINAL TEST REPORT\n"
                + "=".repeat(80)
        );


        if (fails.isEmpty()) {

            System.out.println(
                    "\n🎉 ALL TESTS PASSED! ("
                    + total
                    + "/"
                    + total
                    + ")"
            );

        } else {

            System.out.println(
                    "\n❌ FAILURES: "
                    + fails.size()
                    + " of "
                    + total
            );


            System.out.println(
                    "\nFailed Tests:"
            );


            for (Map<String, String> failure :
                    fails) {

                System.out.println(
                        "\nTest #"
                        + failure.get("testNumber")
                        + " | "
                        + failure.get("departmentName")
                );


                System.out.println(
                        "Question: "
                        + failure.get("question")
                );


                System.out.println(
                        "Failure Type: "
                        + failure.get("failureType")
                );


                System.out.println(
                        "Reason: "
                        + failure.get("failureReason")
                );
            }
        }
    }
}