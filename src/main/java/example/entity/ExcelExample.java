package example.entity;

import java.math.BigDecimal;
import javax.persistence.Entity;
import javax.persistence.GeneratedValue;
import javax.persistence.GenerationType;
import javax.persistence.Id;
import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.NoArgsConstructor;

@Entity
@NoArgsConstructor
@AllArgsConstructor
@Builder
public class ExcelExample {

  @Id
  @GeneratedValue(strategy = GenerationType.AUTO)
  private Long id;
  private String time;
  private String displayName;
  private String accountNumber;
  private String storeName;
  private BigDecimal turnoverValue;
  private BigDecimal fee;
  private BigDecimal amountToPay;
  private BigDecimal actuallyReceived;

}
