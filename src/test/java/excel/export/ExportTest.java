package excel.export;

import org.apache.commons.io.FileUtils;
import org.apache.commons.lang3.RandomStringUtils;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.junit.Test;

import java.util.ArrayList;
import java.util.Date;
import java.util.List;

public class ExportTest {

    private static final Logger logger = LogManager.getLogger(ExportTest.class);

    @Test
    public void export() {
        final int ROW_LENGTH = 10000;
        List<Employee> employees = new ArrayList<>(ROW_LENGTH);
        final int NAME_LENGTH = 5;
        final String CHARS = "abcdefahijklmnopqrstuvwxyz";
        final String filename = FileUtils.getTempDirectoryPath() + "employee.xls";

        ExportExcel exportExcel = new ExportExcel(filename);
        for (int i = 1; i <= 10; i++) {
            for (int j = 1; j <= ROW_LENGTH; j++) {
                employees.add(new Employee(j, RandomStringUtils.random(NAME_LENGTH, CHARS), true, new Date()));
            }
            exportExcel.addSheet("Sheet" + i, employees, Employee.class);
            employees.clear();
        }
        exportExcel.write();

        logger.info(filename);
    }
}
