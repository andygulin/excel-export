package excel.export;

import jxl.SheetSettings;
import jxl.Workbook;
import jxl.format.Alignment;
import jxl.format.Border;
import jxl.format.BorderLineStyle;
import jxl.format.VerticalAlignment;
import jxl.write.*;
import jxl.write.Number;
import org.apache.commons.beanutils.PropertyUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.commons.lang3.math.NumberUtils;
import org.apache.commons.lang3.time.DateFormatUtils;

import java.io.File;
import java.io.IOException;
import java.lang.reflect.Field;
import java.sql.Timestamp;
import java.util.Calendar;
import java.util.Date;
import java.util.LinkedList;
import java.util.List;

public class ExportExcel<T> {

    private static final String DATE_FORMAT_STR = "yyyy-MM-dd HH:mm:ss";

    private List<T> list;
    private String filename;
    private Class<T> clazz;

    public ExportExcel() {
    }

    public ExportExcel(List<T> list, String filename, Class<T> clazz) {
        super();
        this.list = list;
        this.filename = filename;
        this.clazz = clazz;
    }

    public List<T> getList() {
        return list;
    }

    public void setList(List<T> list) {
        this.list = list;
    }

    public String getFilename() {
        return filename;
    }

    public void setFilename(String filename) {
        this.filename = filename;
    }

    public Class<T> getClazz() {
        return clazz;
    }

    public void setClazz(Class<T> clazz) {
        this.clazz = clazz;
    }

    public boolean export(String sheetName) {
        boolean result = true;
        WritableWorkbook workbook = null;
        try {
            workbook = Workbook.createWorkbook(new File(filename));
            WritableSheet sheet = workbook.createSheet(sheetName, 0);
            SheetSettings sheetset = sheet.getSettings();
            sheetset.setProtected(false);
            WritableFont NormalFont = new WritableFont(WritableFont.ARIAL, 10);
            WritableFont BoldFont = new WritableFont(WritableFont.ARIAL, 10, WritableFont.BOLD);

            WritableCellFormat wcf_center = new WritableCellFormat(BoldFont);
            wcf_center.setBorder(Border.ALL, BorderLineStyle.THIN);
            wcf_center.setVerticalAlignment(VerticalAlignment.CENTRE);
            wcf_center.setAlignment(Alignment.CENTRE);
            wcf_center.setWrap(false);

            WritableCellFormat wcf_left = new WritableCellFormat(NormalFont);
            wcf_left.setBorder(Border.NONE, BorderLineStyle.THIN);
            wcf_left.setVerticalAlignment(VerticalAlignment.CENTRE);
            wcf_left.setAlignment(Alignment.LEFT);
            wcf_left.setWrap(false);

            List<String> headers = getHanders();
            List<String> fields = getFields();

            for (int i = 0; i < headers.size(); i++) {
                sheet.addCell(new Label(i, 0, headers.get(i), wcf_center));
            }

            int i = 1;
            for (T obj : list) {
                int j = 0;
                for (int k = 0; k < fields.size(); k++) {
                    Object val = PropertyUtils.getProperty(obj, fields.get(k));
                    if (val instanceof String) {
                        sheet.addCell(new Label(j, i, (String) val, wcf_left));
                    } else if (val instanceof Date) {
                        String dateStr = DateFormatUtils.format((Date) val, DATE_FORMAT_STR);
                        sheet.addCell(new Label(j, i, dateStr, wcf_left));
                    } else if (val instanceof java.sql.Date) {
                        String dateStr = DateFormatUtils.format((java.sql.Date) val, DATE_FORMAT_STR);
                        sheet.addCell(new Label(j, i, dateStr, wcf_left));
                    } else if (val instanceof Timestamp) {
                        String dateStr = DateFormatUtils.format((Timestamp) val, DATE_FORMAT_STR);
                        sheet.addCell(new Label(j, i, dateStr, wcf_left));
                    } else if (val instanceof Calendar) {
                        String dateStr = DateFormatUtils.format((Calendar) val, DATE_FORMAT_STR);
                        sheet.addCell(new Label(j, i, dateStr, wcf_left));
                    } else if (NumberUtils.isCreatable(String.valueOf(val))) {
                        String tmpValue = String.valueOf(val);
                        if (tmpValue.indexOf(".") != -1) {
                            sheet.addCell(new Number(j, i, Double.valueOf(tmpValue), wcf_left));
                        } else {
                            sheet.addCell(new Number(j, i, Long.valueOf(tmpValue), wcf_left));
                        }
                    } else {
                        sheet.addCell(new Label(j, i, String.valueOf(val), wcf_left));
                    }
                    j++;
                }
                i++;
            }
            workbook.write();
        } catch (IOException | WriteException | ReflectiveOperationException e) {
            e.printStackTrace();
            result = false;
        } finally {
            if (workbook != null) {
                try {
                    workbook.close();
                } catch (IOException e) {
                    result = false;
                } catch (WriteException e) {
                    result = false;
                }
            }
        }
        return result;
    }

    public boolean export() {
        return export("Sheet");
    }

    private List<String> getHanders() {
        List<String> headers = new LinkedList<>();
        Field[] fields = clazz.getDeclaredFields();
        for (Field field : fields) {
            Export export = field.getAnnotation(Export.class);
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

    private List<String> getFields() {
        List<String> headers = new LinkedList<>();
        Field[] fields = clazz.getDeclaredFields();
        for (Field field : fields) {
            Export export = field.getAnnotation(Export.class);
            if (export != null) {
                headers.add(field.getName());
            }
        }
        return headers;
    }
}