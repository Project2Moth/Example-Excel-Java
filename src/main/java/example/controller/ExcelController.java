package example.controller;

import example.dto.ExcelExampleDto;
import example.service.ExcelService;
import java.io.IOException;
import java.util.List;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.core.io.Resource;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

@RestController
@RequestMapping("api/v1/excel")
public class ExcelController {

  @Autowired
  private ExcelService excelService;

  @PostMapping
  public ResponseEntity<Resource> exportExcel(@RequestBody List<ExcelExampleDto> excelExampleDtos) {
    Resource resource = excelService.exportExcel(excelExampleDtos);
    HttpHeaders headers = new HttpHeaders();
    headers.add("Content-Disposition",
        "attachment; filename=doanhthutrongthoigian :" + excelExampleDtos.get(0).getTime()
            + ".xlsx");
    headers.add("Content-Type", "text/html");
    try {
      return ResponseEntity.ok().headers(headers).contentLength(resource.contentLength()).contentType(
          MediaType.parseMediaType("application/vnd.ms-excel")).body(resource);
    } catch (IOException e) {
      e.printStackTrace();
      return null;
    }
  }

}
