package VO;

import lombok.Getter;
import lombok.Setter;

import java.math.BigDecimal;

public class Result {
    @Getter
    @Setter
    public String referenceID;

    @Getter
    @Setter
    public BigDecimal debit;

    @Getter
    @Setter
    public BigDecimal credit;
}
