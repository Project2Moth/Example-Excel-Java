package example.service.impl;

import example.dto.ExcelExampleDto;
import example.service.ExcelService;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.Arrays;
import java.util.List;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.core.io.ByteArrayResource;
import org.springframework.core.io.Resource;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Transactional;
import org.springframework.web.multipart.MultipartFile;


@Service
@Transactional
public class ExcelServiceImpl implements ExcelService {

  @Override
  public Resource exportExcel(List<ExcelExampleDto> excelExampleDtos) {
    // Define columns for file excel
    List<String> headers = Arrays
        .asList("Thời gian ", "Tên hiển thị ", "Số tài khoản ", "Tên cửa hàng ",
            "Doanh thu chưa  gồm chi phí", "Chi phí", "Số tiền cần thanh toán", "Thực nhận");
    XSSFWorkbook workbook = new XSSFWorkbook();
    XSSFSheet sheet = workbook.createSheet("Revenue");
    // Merge column and row to use for tittle
    sheet.addMergedRegion(new CellRangeAddress(0, 1, 0, 7));
    // Setup properties for tittle
    XSSFRow titleRow = sheet.createRow(0);
    XSSFFont titleFont = workbook.createFont();
    titleFont.setBold(true);
    titleFont.setColor(IndexedColors.DARK_BLUE.index);
    titleFont.setFontHeight(25);
    XSSFCell titleCell = titleRow.createCell(0);
    titleCell.setCellValue("Bảng lương trong thời gian " + excelExampleDtos.get(0).getTime());
    // Set up for tittle in center
    XSSFCellStyle xssfCellStyle = workbook.createCellStyle();
    xssfCellStyle.setFont(titleFont);
    xssfCellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
    titleCell.setCellStyle(xssfCellStyle);
    XSSFFont headerFont = workbook.createFont();
    headerFont.setBold(true);
    headerFont.setFontHeightInPoints((short) 10);
    headerFont.setColor(IndexedColors.DARK_BLUE.index);

    XSSFCellStyle headerCellStyle = workbook.createCellStyle();
    headerCellStyle.setAlignment(HorizontalAlignment.CENTER);
    headerCellStyle.setFont(headerFont);

    // Set style for cell excel and add value for header.
    XSSFRow headerRow = sheet.createRow(2);
    headers.forEach(str -> {
      XSSFCell cell = headerRow.createCell(headers.indexOf(str));
      cell.setCellStyle(headerCellStyle);
      cell.setCellValue(str);
    });

    // Insert value to cell
    int rowId = 3;
    for (ExcelExampleDto excelExampleDto : excelExampleDtos) {
      XSSFRow row = sheet.createRow(rowId++);
      row.createCell(0).setCellValue(excelExampleDto.getTime());
      row.createCell(1).setCellValue(excelExampleDto.getDisplayName());
      row.createCell(2).setCellValue(excelExampleDto.getAccountNumber());
      row.createCell(3).setCellValue(excelExampleDto.getStoreName());
      row.createCell(4).setCellValue(excelExampleDto.getTurnoverValue().doubleValue());
      row.createCell(5).setCellValue(excelExampleDto.getFee().doubleValue());
      row.createCell(6).setCellValue(excelExampleDto.getAmountToPay().doubleValue());
      row.createCell(7).setCellValue(excelExampleDto.getAmountToPay().doubleValue());
    }

    ByteArrayOutputStream result = new ByteArrayOutputStream();
    try {
      workbook.write(result);
      workbook.close();
    } catch (IOException e) {
      e.printStackTrace();
    }
    return new ByteArrayResource(result.toByteArray());
  }

  @Override
  public String importExcel(MultipartFile multipartFile) {
    return null;
  }
}
