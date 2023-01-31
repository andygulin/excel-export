package excel.export;

import jxl.SheetSettings;
import jxl.Workbook;
import jxl.format.Alignment;
import jxl.format.Border;
import jxl.format.BorderLineStyle;
import jxl.format.VerticalAlignment;
import jxl.write.Number;
import jxl.write.*;
import org.apache.commons.beanutils.PropertyUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.commons.lang3.math.NumberUtils;
import org.apache.commons.lang3.time.DateFormatUtils;

import java.io.File;
import java.io.IOException;
import java.lang.reflect.Field;
import java.util.Calendar;
import java.util.Date;
import java.util.LinkedList;
import java.util.List;

public class ExportExcel {

    private static final String DATE_FORMAT_STR = "yyyy-MM-dd HH:mm:ss";

    private String filename;
    private WritableWorkbook workbook = null;
    private int sheetIdx = 0;

    public ExportExcel(String filename) {
        super();
        this.filename = filename;
        try {
            workbook = Workbook.createWorkbook(new File(filename));
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public String getFilename() {
        return filename;
    }

    public void setFilename(String filename) {
        this.filename = filename;
    }

    public <T> void addSheet(String sheetName, List<T> list, Class<T> clazz) {
        sheetIdx++;
        try {
            WritableSheet sheet = workbook.createSheet(sheetName, sheetIdx);
            SheetSettings sheetSettings = sheet.getSettings();
            sheetSettings.setProtected(false);
            WritableFont NormalFont = new WritableFont(WritableFont.COURIER, 12);
            WritableFont BoldFont = new WritableFont(WritableFont.COURIER, 12, WritableFont.NO_BOLD);

            WritableCellFormat wcf_center = new WritableCellFormat(BoldFont);
            wcf_center.setBorder(Border.NONE, BorderLineStyle.NONE);
            wcf_center.setVerticalAlignment(VerticalAlignment.CENTRE);
            wcf_center.setAlignment(Alignment.CENTRE);
            wcf_center.setWrap(false);

            WritableCellFormat wcf_left = new WritableCellFormat(NormalFont);
            wcf_left.setBorder(Border.NONE, BorderLineStyle.NONE);
            wcf_left.setVerticalAlignment(VerticalAlignment.CENTRE);
            wcf_left.setAlignment(Alignment.LEFT);
            wcf_left.setWrap(false);

            List<String> headers = getHeaders(clazz);
            List<String> fields = getFields(clazz);

            for (int i = 0; i < headers.size(); i++) {
                sheet.addCell(new Label(i, 0, headers.get(i), wcf_center));
            }

            int i = 1;
            for (T obj : list) {
                int j = 0;
                for (String field : fields) {
                    Object val = PropertyUtils.getProperty(obj, field);
                    if (val instanceof String) {
                        sheet.addCell(new Label(j, i, (String) val, wcf_left));
                    } else if (val instanceof Date) {
                        String dateStr = DateFormatUtils.format((Date) val, DATE_FORMAT_STR);
                        sheet.addCell(new Label(j, i, dateStr, wcf_left));
                    } else if (val instanceof Calendar) {
                        String dateStr = DateFormatUtils.format((Calendar) val, DATE_FORMAT_STR);
                        sheet.addCell(new Label(j, i, dateStr, wcf_left));
                    } else if (NumberUtils.isCreatable(String.valueOf(val))) {
                        String tmpValue = String.valueOf(val);
                        if (tmpValue.contains(".")) {
                            sheet.addCell(new Number(j, i, Double.parseDouble(tmpValue), wcf_left));
                        } else {
                            sheet.addCell(new Number(j, i, Long.parseLong(tmpValue), wcf_left));
                        }
                    } else {
                        sheet.addCell(new Label(j, i, String.valueOf(val), wcf_left));
                    }
                    j++;
                }
                i++;
            }
        } catch (WriteException | ReflectiveOperationException e) {
            e.printStackTrace();
        }
    }

    public void write() {
        try {
            workbook.write();
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            if (workbook != null) {
                try {
                    workbook.close();
                } catch (IOException | WriteException e) {
                    e.printStackTrace();
                }
            }
        }
    }

    private List<String> getHeaders(Class<?> clazz) {
        List<String> headers = new LinkedList<>();
        Field[] fields = clazz.getDeclaredFields();
        for (Field field : fields) {
            ExportField export = field.getAnnotation(ExportField.class);
            if (export != null) {
                if (StringUtils.isEmpty(export.name())) {
                    headers.add(field.getName());
                } else {
                    headers.add(export.name());
                }
            }
        }
        return headers;
    }

    private List<String> getFields(Class<?> clazz) {
        List<String> headers = new LinkedList<>();
        Field[] fields = clazz.getDeclaredFields();
        for (Field field : fields) {
            ExportField export = field.getAnnotation(ExportField.class);
            if (export != null) {
                headers.add(field.getName());
            }
        }
        return headers;
    }
}