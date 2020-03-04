package example.service;

import example.dto.ExcelExampleDto;
import java.util.List;
import org.springframework.core.io.Resource;
import org.springframework.http.ResponseEntity;
import org.springframework.web.multipart.MultipartFile;

public interface ExcelService {

  Resource exportExcel(List<ExcelExampleDto> excelExampleDtos);

  String importExcel(MultipartFile multipartFile);
}
