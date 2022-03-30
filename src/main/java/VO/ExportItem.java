package VO;

import lombok.Getter;
import lombok.Setter;

import java.math.BigDecimal;
import java.util.Date;

public class ExportItem {
    @Getter
    @Setter
    public String date;

    @Getter
    @Setter
    public String referenceID;

    @Getter
    @Setter
    public String transactionType;

    @Getter
    @Setter
    public String description;

    @Getter
    @Setter
    public BigDecimal debit;

    @Getter
    @Setter
    public BigDecimal credit;
}
