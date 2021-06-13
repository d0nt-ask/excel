package io.velog.youmakemesmile.excel;

import io.velog.youmakemesmile.excel.config.*;
import lombok.AllArgsConstructor;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;
import org.apache.poi.ss.usermodel.HorizontalAlignment;

import java.math.BigDecimal;
import java.time.LocalDateTime;

@Setter
@Getter
@NoArgsConstructor
public class Sample1 {
    @ExcelHeader(headerName = "반", colIndex = 0, rowIndex = 0, rowSpan = 1, headerStyle = @HeaderStyle(background = @Background("#ECEFF3")))
    @ExcelBody(rowIndex = 0,colIndex = 0,rowGroup = true, bodyStyle = @BodyStyle(horizontalAlignment = HorizontalAlignment.CENTER))
    private String seq;

    @ExcelHeader(headerName = "이름", colIndex = 1, rowIndex = 0, rowSpan = 1, headerStyle = @HeaderStyle(background = @Background("#ECEFF3")))
    @ExcelBody(rowIndex = 0, colIndex = 1, rowSpan = 1)
    private String name;

    @ExcelHeader(headerName = "국어/수학", colIndex = 2, colSpan = 1, rowIndex = 0, headerStyle = @HeaderStyle(background = @Background("#ECEFF3")))
    private String koreanMathHeader;

    @ExcelBody(rowIndex = 0, colIndex = 2)
    private String korean;

    @ExcelBody(rowIndex = 0, colIndex = 3)
    private String math;

    @ExcelHeader(headerName = "영어", colIndex = 2, rowIndex = 1, headerStyle = @HeaderStyle(background = @Background("#ECEFF3")))
    @ExcelBody(rowIndex = 1, colIndex = 2)
    private String english;

    @ExcelHeader(headerName = "역사", colIndex = 3, rowIndex = 1, headerStyle = @HeaderStyle(background = @Background("#ECEFF3")))
    @ExcelBody(rowIndex = 1, colIndex = 3)
    private String history;

    @ExcelHeader(headerName = "생일", colIndex = 4, rowIndex = 0, headerStyle = @HeaderStyle(background = @Background("#ECEFF3")))
    @ExcelBody(rowIndex = 0, colIndex = 4, bodyStyle = @BodyStyle(dateFormat = "YYYY-MM-DD"), width = 12)
    private LocalDateTime birthDay;

    @ExcelHeader(headerName = "금액", colIndex = 4, rowIndex = 1, headerStyle = @HeaderStyle(background = @Background("#ECEFF3")))
    @ExcelBody(rowIndex = 1, colIndex = 4, bodyStyle = @BodyStyle(numberFormat = "#,000원"))
    private BigDecimal money;

    public Sample1(String seq, String name, String korean, String math, String english, String history, LocalDateTime birthDay, BigDecimal money) {
        this.seq = seq;
        this.name = name;
        this.korean = korean;
        this.math = math;
        this.english = english;
        this.history = history;
        this.birthDay = birthDay;
        this.money = money;
    }
}
