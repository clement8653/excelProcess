package VO;

import lombok.Getter;
import lombok.Setter;

import java.math.BigDecimal;

public class GenerateExcelResult {
    @Setter
    @Getter
    public String referenceID;

    @Setter
    @Getter
    public BigDecimal balance;

    @Setter
    @Getter
    public boolean isBalanced;
}
