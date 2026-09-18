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

public class Kanglish_DoctorList {

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
                        + "/RXDX_DoctorList_Whitefield/Screenshots/"
                        + timestamp;

        String excelFolderPath =
                projectPath
                        + "/RXDX_DoctorList_Whitefield/Test_Excel";

        String excelFilePath =
                excelFolderPath
                        + "/RXDX_Whitefield_DoctorList_Results_"
                        + timestamp
                        + ".xlsx";

        new File(screenshotsPath).mkdirs();
        new File(excelFolderPath).mkdirs();

        System.out.println("\n📁 SCREENSHOTS will be saved in: " + screenshotsPath);
        System.out.println("📊 EXCEL FILE will be saved as: " + excelFilePath);

        // ================= DEPARTMENTS WITH PROPER KANNADA QUESTIONS =================

        Map<String, String> departmentQuestions =
                new LinkedHashMap<>();

        // Proper Kannada text for each department
        departmentQuestions.put("Acupuncture", "ಆಕ್ಯುಪಂಕ್ಚರ್ ವೈದ್ಯರ ಪಟ್ಟಿಯನ್ನು ತೋರಿಸಿ");
        departmentQuestions.put("Adult Cardiology", "ಹೃದಯ ತಜ್ಞ ವೈದ್ಯರ ಪಟ್ಟಿಯನ್ನು ತೋರಿಸಿ");
        departmentQuestions.put("Allergy Specialist", "ಅಲರ್ಜಿ ತಜ್ಞ ವೈದ್ಯರ ಪಟ್ಟಿಯನ್ನು ತೋರಿಸಿ");
        departmentQuestions.put("Clinical Nutrition", "ಪೌಷ್ಟಿಕಾಂಶ ತಜ್ಞ ವೈದ್ಯರ ಪಟ್ಟಿಯನ್ನು ತೋರಿಸಿ");
        departmentQuestions.put("Dental", "ದಂತ ವೈದ್ಯರ ಪಟ್ಟಿಯನ್ನು ತೋರಿಸಿ");
        departmentQuestions.put("Dermatology", "ಚರ್ಮರೋಗ ತಜ್ಞ ವೈದ್ಯರ ಪಟ್ಟಿಯನ್ನು ತೋರಿಸಿ");
        departmentQuestions.put("Diabetologist", "ಮಧುಮೇಹ ತಜ್ಞ ವೈದ್ಯರ ಪಟ್ಟಿಯನ್ನು ತೋರಿಸಿ");
        departmentQuestions.put("Endocrinologist", "ಅಂತಃಸ್ರಾವಶಾಸ್ತ್ರ ತಜ್ಞ ವೈದ್ಯರ ಪಟ್ಟಿಯನ್ನು ತೋರಿಸಿ");
        departmentQuestions.put("Endodentist", "ಎಂಡೋಡಾಂಟಿಕ್ಸ್ ವೈದ್ಯರ ಪಟ್ಟಿಯನ್ನು ತೋರಿಸಿ");
        departmentQuestions.put("Endodontist", "ಎಂಡೋಡಾಂಟಿಕ್ಸ್ ವೈದ್ಯರ ಪಟ್ಟಿಯನ್ನು ತೋರಿಸಿ");
        departmentQuestions.put("Endovascular Neurosurgeon, Vascular Surgeon", "ರಕ್ತನಾಳ ಶಸ್ತ್ರಚಿಕಿತ್ಸಾ ತಜ್ಞರ ಪಟ್ಟಿಯನ್ನು ತೋರಿಸಿ");
        departmentQuestions.put("ENT", "ಕಿವಿ, ಮೂಗು ಮತ್ತು ಗಂಟಲು ತಜ್ಞರ ಪಟ್ಟಿಯನ್ನು ತೋರಿಸಿ");
        departmentQuestions.put("Gastroenterologist, Hepatologist", "ಜೀರ್ಣಾಂಗ ಮತ್ತು ಯಕೃತ್ ತಜ್ಞರ ಪಟ್ಟಿಯನ್ನು ತೋರಿಸಿ");
        departmentQuestions.put("General Medicine", "ಸಾಮಾನ್ಯ ವೈದ್ಯಕೀಯ ತಜ್ಞರ ಪಟ್ಟಿಯನ್ನು ತೋರಿಸಿ");
        departmentQuestions.put("General Surgeon", "ಸಾಮಾನ್ಯ ಶಸ್ತ್ರಚಿಕಿತ್ಸಾ ತಜ್ಞರ ಪಟ್ಟಿಯನ್ನು ತೋರಿಸಿ");
        departmentQuestions.put("Homeopathy", "ಹೋಮಿಯೋಪತಿ ವೈದ್ಯರ ಪಟ್ಟಿಯನ್ನು ತೋರಿಸಿ");
        departmentQuestions.put("Internal Medicine", "ಆಂತರಿಕ ವೈದ್ಯಕೀಯ ತಜ್ಞರ ಪಟ್ಟಿಯನ್ನು ತೋರಿಸಿ");
        departmentQuestions.put("Nephrologist", "ಮೂತ್ರಪಿಂಡ ತಜ್ಞರ ಪಟ್ಟಿಯನ್ನು ತೋರಿಸಿ");
        departmentQuestions.put("Neurologist", "ನರರೋಗ ತಜ್ಞರ ಪಟ್ಟಿಯನ್ನು ತೋರಿಸಿ");
        departmentQuestions.put("Neurosurgery", "ನ್ಯೂರೋ ಶಸ್ತ್ರಚಿಕಿತ್ಸಾ ತಜ್ಞರ ಪಟ್ಟಿಯನ್ನು ತೋರಿಸಿ");
        departmentQuestions.put("Obs & Gynae", "ಸ್ತ್ರೀರೋಗ ಮತ್ತು ಪ್ರಸೂತಿ ತಜ್ಞರ ಪಟ್ಟಿಯನ್ನು ತೋರಿಸಿ");
        departmentQuestions.put("Occupational Therapy", "ಆಕ್ಯುಪೇಷನಲ್ ಥೆರಪಿ ತಜ್ಞರ ಪಟ್ಟಿಯನ್ನು ತೋರಿಸಿ");
        departmentQuestions.put("Ophthalmology", "ನೇತ್ರ ತಜ್ಞರ ಪಟ್ಟಿಯನ್ನು ತೋರಿಸಿ");
        departmentQuestions.put("Oral Surgeon", "ಬಾಯಿ ಶಸ್ತ್ರಚಿಕಿತ್ಸಾ ತಜ್ಞರ ಪಟ್ಟಿಯನ್ನು ತೋರಿಸಿ");
        departmentQuestions.put("Orthodentists", "ಆರ್ಥೋಡಾಂಟಿಕ್ಸ್ ತಜ್ಞರ ಪಟ್ಟಿಯನ್ನು ತೋರಿಸಿ");
        departmentQuestions.put("Orthopedic Surgeon", "ಅಸ್ಥಿರೋಗ ತಜ್ಞರ ಪಟ್ಟಿಯನ್ನು ತೋರಿಸಿ");
        departmentQuestions.put("Orthotist", "ಆರ್ಥೋಟಿಕ್ಸ್ ತಜ್ಞರ ಪಟ್ಟಿಯನ್ನು ತೋರಿಸಿ");
        departmentQuestions.put("Paediatric and Fetal Cardiologist", "ಮಕ್ಕಳ ಹೃದಯ ತಜ್ಞರ ಪಟ್ಟಿಯನ್ನು ತೋರಿಸಿ");
        departmentQuestions.put("Paediatric Cardiologist", "ಮಕ್ಕಳ ಹೃದಯ ತಜ್ಞರ ಪಟ್ಟಿಯನ್ನು ತೋರಿಸಿ");
        departmentQuestions.put("Paediatric Neurologist", "ಮಕ್ಕಳ ನರರೋಗ ತಜ್ಞರ ಪಟ್ಟಿಯನ್ನು ತೋರಿಸಿ");
        departmentQuestions.put("Paediatric Surgeon", "ಮಕ್ಕಳ ಶಸ್ತ್ರಚಿಕಿತ್ಸಾ ತಜ್ಞರ ಪಟ್ಟಿಯನ್ನು ತೋರಿಸಿ");
        departmentQuestions.put("Paediatrician", "ಮಕ್ಕಳ ವೈದ್ಯರ ಪಟ್ಟಿಯನ್ನು ತೋರಿಸಿ");
        departmentQuestions.put("Pain medicine", "ನೋವು ನಿರ್ವಹಣಾ ತಜ್ಞರ ಪಟ್ಟಿಯನ್ನು ತೋರಿಸಿ");
        departmentQuestions.put("Pedodentists", "ಮಕ್ಕಳ ದಂತ ವೈದ್ಯರ ಪಟ್ಟಿಯನ್ನು ತೋರಿಸಿ");
        departmentQuestions.put("Periodontist", "ಪೆರಿಯೊಡಾಂಟಿಕ್ಸ್ ತಜ್ಞರ ಪಟ್ಟಿಯನ್ನು ತೋರಿಸಿ");
        departmentQuestions.put("Physiotherapy", "ಫಿಸಿಯೋಥೆರಪಿ ತಜ್ಞರ ಪಟ್ಟಿಯನ್ನು ತೋರಿಸಿ");
        departmentQuestions.put("Plastic Surgery", "ಪ್ಲಾಸ್ಟಿಕ್ ಶಸ್ತ್ರಚಿಕಿತ್ಸಾ ತಜ್ಞರ ಪಟ್ಟಿಯನ್ನು ತೋರಿಸಿ");
        departmentQuestions.put("Podiatric Surgeon", "ಪಾದ ಶಸ್ತ್ರಚಿಕಿತ್ಸಾ ತಜ್ಞರ ಪಟ್ಟಿಯನ್ನು ತೋರಿಸಿ");
        departmentQuestions.put("Psychiatry", "ಮನೋವೈದ್ಯರ ಪಟ್ಟಿಯನ್ನು ತೋರಿಸಿ");
        departmentQuestions.put("Psychology", "ಮನೋವಿಜ್ಞಾನ ತಜ್ಞರ ಪಟ್ಟಿಯನ್ನು ತೋರಿಸಿ");
        departmentQuestions.put("Pulmonologist", "ಶ್ವಾಸಕೋಶ ತಜ್ಞರ ಪಟ್ಟಿಯನ್ನು ತೋರಿಸಿ");
        departmentQuestions.put("Radiology", "ರೇಡಿಯಾಲಜಿ ತಜ್ಞರ ಪಟ್ಟಿಯನ್ನು ತೋರಿಸಿ");
        departmentQuestions.put("Rheumatology", "ಸಂಧಿವಾತ ತಜ್ಞರ ಪಟ್ಟಿಯನ್ನು ತೋರಿಸಿ");
        departmentQuestions.put("Speech Therapy", "ಮಾತಿನ ಚಿಕಿತ್ಸಾ ತಜ್ಞರ ಪಟ್ಟಿಯನ್ನು ತೋರಿಸಿ");
        departmentQuestions.put("Urologist", "ಮೂತ್ರರೋಗ ತಜ್ಞರ ಪಟ್ಟಿಯನ್ನು ತೋರಿಸಿ");

        // ================= ENGLISH MEANINGS MAP (For Excel) =================
        
        Map<String, String> englishMeaning = new HashMap<>();
//        englishMeaning.put("ಆಕ್ಯುಪಂಕ್ಚರ್ ವೈದ್ಯರ ಪಟ್ಟಿಯನ್ನು ತೋರಿಸಿ", "Show doctor list of Acupuncture");
//        englishMeaning.put("ಹೃದಯ ತಜ್ಞ ವೈದ್ಯರ ಪಟ್ಟಿಯನ್ನು ತೋರಿಸಿ", "Show doctor list of Adult Cardiology");
//        englishMeaning.put("ಅಲರ್ಜಿ ತಜ್ಞ ವೈದ್ಯರ ಪಟ್ಟಿಯನ್ನು ತೋರಿಸಿ", "Show doctor list of Allergy Specialist");
//        englishMeaning.put("ಪೌಷ್ಟಿಕಾಂಶ ತಜ್ಞ ವೈದ್ಯರ ಪಟ್ಟಿಯನ್ನು ತೋರಿಸಿ", "Show doctor list of Clinical Nutrition");
//        englishMeaning.put("ದಂತ ವೈದ್ಯರ ಪಟ್ಟಿಯನ್ನು ತೋರಿಸಿ", "Show doctor list of Dental");
//        englishMeaning.put("ಚರ್ಮರೋಗ ತಜ್ಞ ವೈದ್ಯರ ಪಟ್ಟಿಯನ್ನು ತೋರಿಸಿ", "Show doctor list of Dermatology");
//        englishMeaning.put("ಮಧುಮೇಹ ತಜ್ಞ ವೈದ್ಯರ ಪಟ್ಟಿಯನ್ನು ತೋರಿಸಿ", "Show doctor list of Diabetologist");
//        englishMeaning.put("ಅಂತಃಸ್ರಾವಶಾಸ್ತ್ರ ತಜ್ಞ ವೈದ್ಯರ ಪಟ್ಟಿಯನ್ನು ತೋರಿಸಿ", "Show doctor list of Endocrinologist");
//        englishMeaning.put("ಎಂಡೋಡಾಂಟಿಕ್ಸ್ ವೈದ್ಯರ ಪಟ್ಟಿಯನ್ನು ತೋರಿಸಿ", "Show doctor list of Endodentist/Endodontist");
//        englishMeaning.put("ರಕ್ತನಾಳ ಶಸ್ತ್ರಚಿಕಿತ್ಸಾ ತಜ್ಞರ ಪಟ್ಟಿಯನ್ನು ತೋರಿಸಿ", "Show doctor list of Endovascular Neurosurgeon, Vascular Surgeon");
//        englishMeaning.put("ಕಿವಿ, ಮೂಗು ಮತ್ತು ಗಂಟಲು ತಜ್ಞರ ಪಟ್ಟಿಯನ್ನು ತೋರಿಸಿ", "Show doctor list of ENT");
//        englishMeaning.put("ಜೀರ್ಣಾಂಗ ಮತ್ತು ಯಕೃತ್ ತಜ್ಞರ ಪಟ್ಟಿಯನ್ನು ತೋರಿಸಿ", "Show doctor list of Gastroenterologist, Hepatologist");
//        englishMeaning.put("ಸಾಮಾನ್ಯ ವೈದ್ಯಕೀಯ ತಜ್ಞರ ಪಟ್ಟಿಯನ್ನು ತೋರಿಸಿ", "Show doctor list of General Medicine");
//        englishMeaning.put("ಸಾಮಾನ್ಯ ಶಸ್ತ್ರಚಿಕಿತ್ಸಾ ತಜ್ಞರ ಪಟ್ಟಿಯನ್ನು ತೋರಿಸಿ", "Show doctor list of General Surgeon");
//        englishMeaning.put("ಹೋಮಿಯೋಪತಿ ವೈದ್ಯರ ಪಟ್ಟಿಯನ್ನು ತೋರಿಸಿ", "Show doctor list of Homeopathy");
        englishMeaning.put("ಆಂತರಿಕ ವೈದ್ಯಕೀಯ ತಜ್ಞರ ಪಟ್ಟಿಯನ್ನು ತೋರಿಸಿ", "Show doctor list of Internal Medicine");
        englishMeaning.put("ಮೂತ್ರಪಿಂಡ ತಜ್ಞರ ಪಟ್ಟಿಯನ್ನು ತೋರಿಸಿ", "Show doctor list of Nephrologist");
        englishMeaning.put("ನರರೋಗ ತಜ್ಞರ ಪಟ್ಟಿಯನ್ನು ತೋರಿಸಿ", "Show doctor list of Neurologist");
        englishMeaning.put("ನ್ಯೂರೋ ಶಸ್ತ್ರಚಿಕಿತ್ಸಾ ತಜ್ಞರ ಪಟ್ಟಿಯನ್ನು ತೋರಿಸಿ", "Show doctor list of Neurosurgery");
        englishMeaning.put("ಸ್ತ್ರೀರೋಗ ಮತ್ತು ಪ್ರಸೂತಿ ತಜ್ಞರ ಪಟ್ಟಿಯನ್ನು ತೋರಿಸಿ", "Show doctor list of Obs & Gynae");
        englishMeaning.put("ಆಕ್ಯುಪೇಷನಲ್ ಥೆರಪಿ ತಜ್ಞರ ಪಟ್ಟಿಯನ್ನು ತೋರಿಸಿ", "Show doctor list of Occupational Therapy");
        englishMeaning.put("ನೇತ್ರ ತಜ್ಞರ ಪಟ್ಟಿಯನ್ನು ತೋರಿಸಿ", "Show doctor list of Ophthalmology");
        englishMeaning.put("ಬಾಯಿ ಶಸ್ತ್ರಚಿಕಿತ್ಸಾ ತಜ್ಞರ ಪಟ್ಟಿಯನ್ನು ತೋರಿಸಿ", "Show doctor list of Oral Surgeon");
        englishMeaning.put("ಆರ್ಥೋಡಾಂಟಿಕ್ಸ್ ತಜ್ಞರ ಪಟ್ಟಿಯನ್ನು ತೋರಿಸಿ", "Show doctor list of Orthodentists");
        englishMeaning.put("ಅಸ್ಥಿರೋಗ ತಜ್ಞರ ಪಟ್ಟಿಯನ್ನು ತೋರಿಸಿ", "Show doctor list of Orthopedic Surgeon");
        englishMeaning.put("ಆರ್ಥೋಟಿಕ್ಸ್ ತಜ್ಞರ ಪಟ್ಟಿಯನ್ನು ತೋರಿಸಿ", "Show doctor list of Orthotist");
        englishMeaning.put("ಮಕ್ಕಳ ಹೃದಯ ತಜ್ಞರ ಪಟ್ಟಿಯನ್ನು ತೋರಿಸಿ", "Show doctor list of Paediatric and Fetal Cardiologist");
        englishMeaning.put("ಮಕ್ಕಳ ನರರೋಗ ತಜ್ಞರ ಪಟ್ಟಿಯನ್ನು ತೋರಿಸಿ", "Show doctor list of Paediatric Neurologist");
        englishMeaning.put("ಮಕ್ಕಳ ಶಸ್ತ್ರಚಿಕಿತ್ಸಾ ತಜ್ಞರ ಪಟ್ಟಿಯನ್ನು ತೋರಿಸಿ", "Show doctor list of Paediatric Surgeon");
        englishMeaning.put("ಮಕ್ಕಳ ವೈದ್ಯರ ಪಟ್ಟಿಯನ್ನು ತೋರಿಸಿ", "Show doctor list of Paediatrician");
        englishMeaning.put("ನೋವು ನಿರ್ವಹಣಾ ತಜ್ಞರ ಪಟ್ಟಿಯನ್ನು ತೋರಿಸಿ", "Show doctor list of Pain medicine");
        englishMeaning.put("ಮಕ್ಕಳ ದಂತ ವೈದ್ಯರ ಪಟ್ಟಿಯನ್ನು ತೋರಿಸಿ", "Show doctor list of Pedodentists");
        englishMeaning.put("ಪೆರಿಯೊಡಾಂಟಿಕ್ಸ್ ತಜ್ಞರ ಪಟ್ಟಿಯನ್ನು ತೋರಿಸಿ", "Show doctor list of Periodontist");
        englishMeaning.put("ಫಿಸಿಯೋಥೆರಪಿ ತಜ್ಞರ ಪಟ್ಟಿಯನ್ನು ತೋರಿಸಿ", "Show doctor list of Physiotherapy");
        englishMeaning.put("ಪ್ಲಾಸ್ಟಿಕ್ ಶಸ್ತ್ರಚಿಕಿತ್ಸಾ ತಜ್ಞರ ಪಟ್ಟಿಯನ್ನು ತೋರಿಸಿ", "Show doctor list of Plastic Surgery");
        englishMeaning.put("ಪಾದ ಶಸ್ತ್ರಚಿಕಿತ್ಸಾ ತಜ್ಞರ ಪಟ್ಟಿಯನ್ನು ತೋರಿಸಿ", "Show doctor list of Podiatric Surgeon");
        englishMeaning.put("ಮನೋವೈದ್ಯರ ಪಟ್ಟಿಯನ್ನು ತೋರಿಸಿ", "Show doctor list of Psychiatry");
        englishMeaning.put("ಮನೋವಿಜ್ಞಾನ ತಜ್ಞರ ಪಟ್ಟಿಯನ್ನು ತೋರಿಸಿ", "Show doctor list of Psychology");
        englishMeaning.put("ಶ್ವಾಸಕೋಶ ತಜ್ಞರ ಪಟ್ಟಿಯನ್ನು ತೋರಿಸಿ", "Show doctor list of Pulmonologist");
        englishMeaning.put("ರೇಡಿಯಾಲಜಿ ತಜ್ಞರ ಪಟ್ಟಿಯನ್ನು ತೋರಿಸಿ", "Show doctor list of Radiology");
        englishMeaning.put("ಸಂಧಿವಾತ ತಜ್ಞರ ಪಟ್ಟಿಯನ್ನು ತೋರಿಸಿ", "Show doctor list of Rheumatology");
        englishMeaning.put("ಮಾತಿನ ಚಿಕಿತ್ಸಾ ತಜ್ಞರ ಪಟ್ಟಿಯನ್ನು ತೋರಿಸಿ", "Show doctor list of Speech Therapy");
        englishMeaning.put("ಮೂತ್ರರೋಗ ತಜ್ಞರ ಪಟ್ಟಿಯನ್ನು ತೋರಿಸಿ", "Show doctor list of Urologist");
        
        // ================= CREATE EXCEL WORKBOOK WITH HEADERS =================

        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("RXDX_Whitefield_DoctorList_Results");
        
        addProjectHeader(sheet, workbook, "RXDX WHITEFIELD - KANNADA LANGUAGE DOCTOR LIST VALIDATION TESTS");
        
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
        System.out.println("     KANNADA LANGUAGE DOCTOR LIST VALIDATION TESTS STARTING");
        System.out.println("     TOTAL TEST CASES: " + totalTests);
        System.out.println("     BRANCH: Whitefield (Option 1)");
        System.out.println("     Using 10 seconds wait time between messages");
        System.out.println("=".repeat(80));

        // ================= TEST LOOP =================

        int testNumber = 1;

        for (Map.Entry<String, String> entry : departmentQuestions.entrySet()) {

            String expectedDepartment = entry.getKey();
            String kannadaQuestion = entry.getValue();
            String englishQuestion = englishMeaning.getOrDefault(kannadaQuestion, "Unknown");

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
            System.out.println("Question (Kannada): " + kannadaQuestion);
            System.out.println("Question (English): " + englishQuestion);
            System.out.println("=".repeat(60));

            try {

                // ================= STEP 1: Send "hi" =================
                System.out.println("\n📤 STEP 1: Sending 'hi'...");
                sendMessage(driver, "hi", wait);
                System.out.println("⏳ Waiting 10 seconds for response to load...");
                Thread.sleep(10000);
                
                responseAfterHi = getLastMessageReliable(driver);
                System.out.println("✅ Bot responded to 'hi'");

                // ================= STEP 2: Send Kannada doctor list question =================
                System.out.println("\n📤 STEP 2: Sending Kannada question: " + kannadaQuestion);
                sendMessage(driver, kannadaQuestion, wait);
                System.out.println("⏳ Waiting 10 seconds for response to load...");
                Thread.sleep(10000);
                
                responseAfterQuestion = getLastMessageReliable(driver);
                System.out.println("✅ Branch menu received (or unsupported response)");

                // ================= STEP 3: Send branch number (1 for Whitefield) =================
                System.out.println("\n📤 STEP 3: Sending branch number 1 (RxDx Whitefield)...");
                sendMessage(driver, "1", wait);
                System.out.println("⏳ Waiting 10 seconds for response to load...");
                Thread.sleep(10000);
                
                responseAfterBranch = getLastMessageReliable(driver);
                System.out.println("✅ Branch selected");

                // ================= STEP 4: Send "1" to get doctor list =================
                System.out.println("\n📤 STEP 4: Sending '1' to get doctor list...");
                sendMessage(driver, "1", wait);
                System.out.println("⏳ Waiting 10 seconds for response to load...");
                Thread.sleep(10000);
                
                doctorListResponse = getLastMessageReliable(driver);
                System.out.println("\n📋 DOCTOR LIST RESPONSE:\n" + doctorListResponse);

                // ================= STEP 5: VALIDATION WITH UNSUPPORTED DEPARTMENT CHECK =================
                System.out.println("\n" + "=".repeat(50));
                System.out.println("STEP 5: VALIDATING DOCTOR LIST RESPONSE");
                System.out.println("=".repeat(50));
                
                // ================= UNSUPPORTED DEPARTMENT CHECK =================
                if (doctorListResponse.toLowerCase()
                        .contains("sorry, i can assist only for the supported departments")) {

                    status = "FAIL";
                    failureType = "UNSUPPORTED_DEPARTMENT";
                    failureReason = "Bot does not support Kannada language for this department";

                    System.out.println("\n❌ FAIL: Unsupported Department - Bot doesn't understand Kannada");

                } else {

                    // Validation 1: Check closing text
                    String expectedClosingText = "Please choose from the following list of doctors";
                    hasClosingText = doctorListResponse.contains(expectedClosingText);

                    if (!hasClosingText) {
                        status = "FAIL";
                        failureType = "MISSING_CLOSING_TEXT";
                        failureReason = "Missing closing text: " + expectedClosingText;
                        System.out.println("❌ FAIL: Missing closing text!");
                    } else {
                        System.out.println("✅ PASS: Closing text found.");
                    }

                    // Validation 2: Check doctors using regex pattern
                    System.out.println("\n🔍 VALIDATION 2: Checking if doctors are listed...");
                    
                    String[] lines = doctorListResponse.split("\n");
                    doctorCount = 0;

                    for (String line : lines) {
                        String trimmed = line.trim();
                        if (trimmed.matches("^\\d+\\.\\s*Dr\\..*")) {
                            doctorCount++;
                        }
                    }

                    hasDoctors = doctorCount > 0;

                    if (!hasDoctors) {
                        status = "FAIL";
                        if (failureType.isEmpty()) {
                            failureType = "NO_DOCTORS_FOUND";
                        }
                        failureReason = failureReason + " No doctors found in response.";
                        System.out.println("❌ FAIL: No doctors found in response!");
                    } else {
                        System.out.println("✅ PASS: Found " + doctorCount + " doctor(s) in the list.");
                    }
                }
                
                // Final result
                if (status.equals("PASS")) {
                    System.out.println("\n✅✅✅ STEP 5: ALL VALIDATIONS PASSED for: " + expectedDepartment);
                    System.out.println("   ✓ Kannada language supported");
                    System.out.println("   ✓ Closing text present");
                    System.out.println("   ✓ " + doctorCount + " doctor(s) listed");
                } else {
                    System.out.println("\n❌❌❌ STEP 5: TEST FAILED for: " + expectedDepartment);
                    System.out.println("   Reason: " + failureReason);
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

            addTestResultToExcel(sheet, excelRowNum++, testNumber, expectedDepartment, kannadaQuestion, englishQuestion,
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
            Thread.sleep(5000);
        }

        for (int i = 0; i <= 15; i++) {
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
        sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 15));
        
        Row infoRow = sheet.createRow(1);
        Cell infoCell = infoRow.createCell(0);
        infoCell.setCellValue("Project: RXDX Whitefield Doctor List Validation | Language: Kannada | Branch: Whitefield (Option 1) | Departments: 45 | Wait Time: 10 Seconds | Date: " + 
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
        sheet.addMergedRegion(new CellRangeAddress(1, 1, 0, 15));
    }
    
    static int addColumnHeaders(Sheet sheet, Workbook workbook) {
        Row columnHeaderRow = sheet.createRow(2);
        
        String[] columns = {
            "Test No", "Department", "Question (Kannada)", "English Meaning",
            "Response After Hi", "Response After Question", "Response After Branch", 
            "Doctor List Response", "Status", "Failure Type", "Failure Reason", 
            "Closing Text Found", "Doctors Found", "Doctor Count", "Screenshot", "Timestamp"
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
                                      String kannadaQuestion, String englishQuestion, String responseAfterHi, 
                                      String responseAfterQuestion, String responseAfterBranch, 
                                      String doctorListResponse, String status, String failureType, 
                                      String failureReason, boolean hasClosingText, boolean hasDoctors, 
                                      int doctorCount, String screenshotPath) {
        
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
        row.createCell(2).setCellValue(kannadaQuestion);
        row.createCell(3).setCellValue(englishQuestion);
        
        Cell cell4 = row.createCell(4);
        cell4.setCellValue(responseAfterHi != null && !responseAfterHi.isEmpty() ? responseAfterHi : "No response");
        cell4.setCellStyle(wrapStyle);
        
        Cell cell5 = row.createCell(5);
        cell5.setCellValue(responseAfterQuestion != null && !responseAfterQuestion.isEmpty() ? responseAfterQuestion : "No response");
        cell5.setCellStyle(wrapStyle);
        
        Cell cell6 = row.createCell(6);
        cell6.setCellValue(responseAfterBranch != null && !responseAfterBranch.isEmpty() ? responseAfterBranch : "No response");
        cell6.setCellStyle(wrapStyle);
        
        Cell cell7 = row.createCell(7);
        cell7.setCellValue(doctorListResponse != null && !doctorListResponse.isEmpty() ? doctorListResponse : "No response");
        cell7.setCellStyle(wrapStyle);
        
        Cell statusCell = row.createCell(8);
        statusCell.setCellValue(status);
        if (status.equals("PASS")) {
            statusCell.setCellStyle(passStyle);
        } else {
            statusCell.setCellStyle(failStyle);
        }
        
        row.createCell(9).setCellValue(failureType != null ? failureType : "N/A");
        
        Cell cell10 = row.createCell(10);
        cell10.setCellValue(failureReason != null ? failureReason : "N/A");
        cell10.setCellStyle(wrapStyle);
        
        row.createCell(11).setCellValue(hasClosingText ? "YES (PASS)" : "NO (FAIL)");
        row.createCell(12).setCellValue(hasDoctors ? "YES (PASS)" : "NO (FAIL)");
        row.createCell(13).setCellValue(doctorCount);
        row.createCell(14).setCellValue(screenshotPath != null ? screenshotPath : "No screenshot");
        row.createCell(15).setCellValue(new SimpleDateFormat("yyyy-MM-dd HH:mm:ss").format(new Date()));
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