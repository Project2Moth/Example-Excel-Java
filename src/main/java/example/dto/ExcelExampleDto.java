package example.dto;

import java.math.BigDecimal;
import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;

@Setter
@Getter
@NoArgsConstructor
@AllArgsConstructor
@Builder
public class ExcelExampleDto {

  private String time;
  private String displayName;
  private String accountNumber;
  private String storeName;
  private BigDecimal turnoverValue;
  private BigDecimal fee;
  private BigDecimal amountToPay;
  private BigDecimal actuallyReceived;

}
