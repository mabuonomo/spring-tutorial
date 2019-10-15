package it.mabuonomo.javatoexcel.component;

import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.Map;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import org.apache.poi.ss.usermodel.*;
import org.springframework.context.annotation.ComponentScan;
import org.springframework.stereotype.Component;
import org.springframework.web.servlet.view.document.AbstractXlsView;

@Component("forexView")
@ComponentScan
public class ExcelView extends AbstractXlsView {

  @Override
  protected void buildExcelDocument(Map<String, Object> model, Workbook workbook, HttpServletRequest request,
      HttpServletResponse response) throws Exception {

    Sheet sheet = workbook.createSheet("Reports");
    sheet.setFitToPage(true);

    List<Object> objs = (List<Object>) model.get("reports");

    List<String> fieldNames = getFieldNamesForClass(objs.get(0).getClass());
    int rowCount = 0;
    int columnCount = 0;
    Row row = sheet.createRow(rowCount++);

    CellStyle style = workbook.createCellStyle();
    Font font = workbook.createFont();
    font.setBoldweight(Font.BOLDWEIGHT_BOLD);
    style.setFont(font);

    // HEADERS
    for (String fieldName : fieldNames) {
      Cell cell = row.createCell(columnCount++);
      cell.setCellValue(fieldName.toUpperCase());
      cell.setCellStyle(style);
    }

    // ROWS
    for (Object t : objs) {
      Class<Object> classz = (Class<Object>) t.getClass();
      row = sheet.createRow(rowCount++);
      columnCount = 0;
      for (String fieldName : fieldNames) {
        Cell cell = row.createCell(columnCount);
        Method method = null;
        try {
          method = classz.getMethod("get" + capitalize(fieldName));
        } catch (NoSuchMethodException nme) {
          method = classz.getMethod("get" + fieldName);
        }
        Object value = method.invoke(t, (Object[]) null);
        if (value != null) {
          if (value instanceof Long) {
            cell.setCellValue((Long) value);
          } else if (value instanceof Integer) {
            cell.setCellValue((Integer) value);
          } else if (value instanceof Double) {
            cell.setCellValue((Double) value);
          } else if (value instanceof Date) {
            SimpleDateFormat formatter = new SimpleDateFormat("dd/MM/yyyy HH:mm:ss");
            cell.setCellValue((formatter.format(value)));
          } else {
            cell.setCellValue((String) value);
          }
        } else {
          cell.setCellValue("");
        }
        columnCount++;
      }
    }
    response.setHeader("Content-Disposition", "attachment; filename=report.xls");
  }

  private static List<String> getFieldNamesForClass(Class<?> clazz) throws Exception {
    List<String> fieldNames = new ArrayList<String>();
    Field[] fields = clazz.getDeclaredFields();
    for (Field field : fields) {
      fieldNames.add(field.getName());
    }
    return fieldNames;
  }

  // capitalize the first letter of the field name for retriving value of the
  // field later
  private static String capitalize(String s) {
    if (s.length() == 0)
      return s;
    return s.substring(0, 1).toUpperCase() + s.substring(1);
  }
}
