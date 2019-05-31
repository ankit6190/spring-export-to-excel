
package com.example.demo.config;

import com.example.demo.domain.Student;
import lombok.Setter;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.*;
import org.springframework.stereotype.Component;
import org.springframework.web.servlet.view.document.AbstractXlsxView;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.beans.PropertyDescriptor;
import java.util.List;
import java.util.Map;

@Component
@Setter
public class ExportExcelView extends AbstractXlsxView {


    private String fileName;
    private String beanName;
    private String sheetName;

    @Override
    @SuppressWarnings("unchecked")
    public void buildExcelDocument(Map<String, Object> model, Workbook workbook, HttpServletRequest request, HttpServletResponse response) throws Exception {
        String[] headerList = (String[]) model.get("headerList");
        List<Object> dataList = (List<Object>) model.get("dataList");
        String[] beanMapColumn=(String[]) model.get("beanMapColumn");
        response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        response.setHeader("Content-Disposition", "attachment; filename="+fileName+".xlsx");

        // create a new Excel sheet
        //HSSFSheet sheet = workbook.createSheet();
        if(workbook==null){
             workbook = createWorkbook(model, request);
        }

            Sheet sheet = workbook.createSheet(sheetName);
            sheet.setDefaultColumnWidth(15);

            // create style for header cells
            CellStyle style = workbook.createCellStyle();
            Font font = workbook.createFont();
            font.setFontName("Arial");
            font.setBold(true);
            style.setFont(font);

            // create header row
            Row header = sheet.createRow(0);

            for (int i = 0; i < headerList.length; i++) {
                header.createCell(i).setCellValue(headerList[i]);
                header.getCell(i).setCellStyle(style);
            }

            int rowCount = 1;
            if (dataList != null) {
                for (int i = 0; i < dataList.size(); i++) {
                    Row aRow = sheet.createRow(rowCount++);
                    for (int j = 0; j < beanMapColumn.length; j++) {
                        Object value = new PropertyDescriptor(beanMapColumn[j], Class.forName(beanName)).getReadMethod().invoke(dataList.get(i));
                        aRow.createCell(j).setCellValue(value == null ? "" : value.toString());
                    }

                }
            }

    }
}
