package example.service.impl;

import example.dto.ExcelExampleDto;
import example.entity.ExcelExample;
import example.repository.ExcelExampleRepository;
import example.service.ExcelService;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.math.BigDecimal;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.LocalTime;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
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
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.core.io.ByteArrayResource;
import org.springframework.core.io.Resource;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Transactional;
import org.springframework.web.multipart.MultipartFile;


@Service
@Transactional
public class ExcelServiceImpl implements ExcelService {

  @Autowired
  private ExcelExampleRepository excelExampleRepository;

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
    try {
      XSSFWorkbook workbook = new XSSFWorkbook(multipartFile.getInputStream());
      XSSFSheet worksheet = workbook.getSheetAt(0);
      List<ExcelExample> excelExamples = new ArrayList<>();
      for (int i = 4; i < worksheet.getPhysicalNumberOfRows(); i++) {
        ExcelExample excelExample = new ExcelExample();

        XSSFRow row = worksheet.getRow(i);

        String str = row.getCell(0).getStringCellValue();
        DateTimeFormatter formatter = DateTimeFormatter.ofPattern("yyyy/MM/dd");
        LocalDate date = LocalDate.parse(str, formatter);
        LocalDateTime dateTime = LocalDateTime.of(date, LocalDateTime.MIN.toLocalTime());
        excelExample.setTime(dateTime);
        excelExample.setDisplayName(row.getCell(1).getStringCellValue());
        excelExample.setAccountNumber(String.valueOf(row.getCell(2).getStringCellValue()));
        excelExample.setStoreName(row.getCell(3).getStringCellValue());
        excelExample.setTurnoverValue(BigDecimal.valueOf(row.getCell(4).getNumericCellValue()));
        excelExample.setFee(BigDecimal.valueOf(row.getCell(5).getNumericCellValue()));
        excelExample.setAmountToPay(BigDecimal.valueOf(row.getCell(6).getNumericCellValue()));
        excelExample.setActuallyReceived(BigDecimal.valueOf(row.getCell(7).getNumericCellValue()));
        excelExamples.add(excelExample);
      }
      excelExampleRepository.saveAll(excelExamples);
    } catch (IOException e) {
      e.printStackTrace();
    }
    return "Successfully";
  }

}
