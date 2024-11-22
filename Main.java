import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

public class Main {
    public static void main(String[] args) {
        String filePath = "C:\\Users\\diedy\\IdeaProjects\\Student\\src\\Students.xlsx";
        List<Student> students = new ArrayList<>();

        // Чтение данных из Excel
        try (FileInputStream fis = new FileInputStream(filePath);
             Workbook workbook = new XSSFWorkbook(fis)) {

            Sheet sheet = workbook.getSheetAt(0);
            for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);

                String name = row.getCell(0).getStringCellValue();
                double currentScholarship = row.getCell(1).getNumericCellValue();
                double newScholarship = row.getCell(2).getNumericCellValue();

                students.add(new Student(name, currentScholarship, newScholarship));
            }

        } catch (IOException e) {
            e.printStackTrace();
        }

        for (Student student : students) {
            System.out.println(
                    "Имя: " + student.getName() +
                            " | Текущая стипендия: " + student.getCurrentScholarship() +
                            " | Новая стипендия: " + student.getNewScholarship() +
                            " | Увеличение: " + (student.getNewScholarship() - student.getCurrentScholarship())
            );
        }

    }
}