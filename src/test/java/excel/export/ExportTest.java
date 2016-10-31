package excel.export;

import java.util.ArrayList;
import java.util.Date;
import java.util.List;

import org.apache.commons.io.FileUtils;
import org.apache.commons.lang3.RandomStringUtils;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.junit.Test;

public class ExportTest {

	private static final Logger logger = LogManager.getLogger(ExportTest.class);

	@Test
	public void export() {
		final int ROW_LENGTH = 10000;
		List<Employee> employees = new ArrayList<>(ROW_LENGTH);
		final int NAME_LENGTH = 5;
		final String CHARS = "abcdefahijklmnopqrstuvwxyz";
		for (int i = 1; i <= ROW_LENGTH; i++) {
			employees.add(new Employee(i, RandomStringUtils.random(NAME_LENGTH, CHARS), true, new Date()));
		}
		final String filename = FileUtils.getTempDirectoryPath() + "employee.xls";
		boolean result = new ExportExcel<Employee>(employees, filename, Employee.class).export();
		logger.info(filename);
		logger.info(result);
	}
}
