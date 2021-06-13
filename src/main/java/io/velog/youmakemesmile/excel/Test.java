package io.velog.youmakemesmile.excel;

import org.springframework.core.io.ByteArrayResource;
import org.springframework.core.io.Resource;
import org.springframework.http.ResponseEntity;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.math.BigDecimal;
import java.time.LocalDateTime;
import java.util.Arrays;
import java.util.List;

public class Test {
    public static void main(String[] args) throws IOException, IllegalAccessException {
        List<Sample1> sample1List = Arrays.asList(
                new Sample1("1반","박백아","100","90", "80", "100", LocalDateTime.of(1994, 1, 1,0,0,0), BigDecimal.valueOf(1000)),
                new Sample1("1반","김마리","100","90", "80", "100", LocalDateTime.of(1995, 2, 1,0,0,0), BigDecimal.valueOf(2000000)),
                new Sample1("3반","김철수","100","90", "80", "100", LocalDateTime.of(1996, 3, 1,0,0,0), BigDecimal.valueOf(300)),
                new Sample1("1반","김영희","100","90", "80", "100", LocalDateTime.of(1997, 4, 1,0,0,0), null)

        );
        ResponseEntity<Resource> responseEntity = ExcelUtil.export("test",Sample1.class, sample1List);
        FileOutputStream fileOutputStream = new FileOutputStream("test.xlsx");
        fileOutputStream.write(((ByteArrayResource)responseEntity.getBody()).getByteArray());
    }
}
