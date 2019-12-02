package example.service;

import example.dto.ExcelExampleDto;
import java.util.List;
import org.springframework.core.io.Resource;

public interface ExcelService {

  Resource exportExcel(List<ExcelExampleDto> excelExampleDtos);
}
