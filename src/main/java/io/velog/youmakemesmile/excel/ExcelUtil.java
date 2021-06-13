package io.velog.youmakemesmile.excel;

import io.velog.youmakemesmile.excel.config.ExcelBody;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.RegionUtil;
import org.apache.poi.xssf.usermodel.*;
import org.springframework.core.io.ByteArrayResource;
import org.springframework.core.io.Resource;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import io.velog.youmakemesmile.excel.config.ExcelHeader;

import java.awt.*;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.lang.reflect.Field;
import java.net.URLEncoder;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.*;
import java.util.List;
import java.util.stream.Collectors;

public class ExcelUtil {

    public static <T> ResponseEntity<Resource> export(String fileName, Class<T> excelClass, List<T> data) throws IllegalAccessException, IOException {

        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet(fileName);
        Map<Integer, List<ExcelHeader>> headerMap = Arrays.stream(excelClass.getDeclaredFields())
                .filter(field -> field.isAnnotationPresent(ExcelHeader.class))
                .map(field -> field.getDeclaredAnnotation(ExcelHeader.class))
                .sorted(Comparator.comparing(ExcelHeader::colIndex))
                .collect(Collectors.groupingBy(ExcelHeader::rowIndex));

        int index = 0;
        for (Integer key : headerMap.keySet()) {
            XSSFRow row = sheet.createRow(index++);
            for (ExcelHeader excelHeader : headerMap.get(key)) {
                XSSFCell cell = row.createCell(excelHeader.colIndex());
                XSSFCellStyle cellStyle = workbook.createCellStyle();
                cell.setCellValue(excelHeader.headerName());
                if (excelHeader.headerName().contains("\n")) {
                    cellStyle.setWrapText(true);
                }
                cellStyle.setAlignment(excelHeader.headerStyle().horizontalAlignment());
                cellStyle.setVerticalAlignment(excelHeader.headerStyle().verticalAlignment());
                if (isHex(excelHeader.headerStyle().background().value())) {
                    cellStyle.setFillForegroundColor(new XSSFColor(Color.decode(excelHeader.headerStyle().background().value()), new DefaultIndexedColorMap()));
                    cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                }
                XSSFFont font = workbook.createFont();
                font.setFontHeightInPoints((short) excelHeader.headerStyle().fontSize());
                cellStyle.setFont(font);
                cellStyle.setBorderBottom(BorderStyle.THIN);
                cellStyle.setBorderLeft(BorderStyle.THIN);
                cellStyle.setBorderRight(BorderStyle.THIN);
                cellStyle.setBorderTop(BorderStyle.THIN);
                cell.setCellStyle(cellStyle);
                if (excelHeader.colSpan() > 0 || excelHeader.rowSpan() > 0) {
                    CellRangeAddress cellAddresses = new CellRangeAddress(cell.getAddress().getRow(), cell.getAddress().getRow() + excelHeader.rowSpan(), cell.getAddress().getColumn(), cell.getAddress().getColumn() + excelHeader.colSpan());
                    sheet.addMergedRegion(cellAddresses);
                }
            }
        }

        Map<Integer, List<Field>> fieldMap = Arrays.stream(excelClass.getDeclaredFields())
                .filter(field -> field.isAnnotationPresent(ExcelBody.class))
                .map(field -> {
                    field.setAccessible(true);
                    return field;
                })
                .sorted(Comparator.comparing(field -> field.getDeclaredAnnotation(ExcelBody.class).colIndex()))
                .collect(Collectors.groupingBy(field -> field.getDeclaredAnnotation(ExcelBody.class).rowIndex()));

        for (T t : data) {
            for (Integer key : fieldMap.keySet()) {
                XSSFRow row = sheet.createRow(index++);
                for (Field field : fieldMap.get(key)) {
                    ExcelBody excelBody = field.getDeclaredAnnotation(ExcelBody.class);
                    Object o = field.get(t);
                    XSSFCell cell = row.createCell(excelBody.colIndex());
                    XSSFCellStyle cellStyle = workbook.createCellStyle();
                    XSSFDataFormat dataFormat = workbook.createDataFormat();

                    cellStyle.setAlignment(excelBody.bodyStyle().horizontalAlignment());
                    cellStyle.setVerticalAlignment(excelBody.bodyStyle().verticalAlignment());
                    if (isHex(excelBody.bodyStyle().background().value())) {
                        cellStyle.setFillForegroundColor(new XSSFColor(Color.decode(excelBody.bodyStyle().background().value()), new DefaultIndexedColorMap()));
                        cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                    }
                    cellStyle.setBorderBottom(BorderStyle.THIN);
                    cellStyle.setBorderLeft(BorderStyle.THIN);
                    cellStyle.setBorderRight(BorderStyle.THIN);
                    cellStyle.setBorderTop(BorderStyle.THIN);

                    if (o instanceof Number) {
                        if (StringUtils.isNoneBlank(excelBody.bodyStyle().numberFormat())) {
                            cellStyle.setDataFormat(dataFormat.getFormat(excelBody.bodyStyle().numberFormat()));
                        }
                        cell.setCellValue(((Number) o).doubleValue());
                    } else if (o instanceof String) {
                        cell.setCellValue((String) o);
                    } else if (o instanceof Date) {
                        cellStyle.setDataFormat(dataFormat.getFormat(excelBody.bodyStyle().dateFormat()));
                        cell.setCellValue((Date) o);
                    } else if (o instanceof LocalDateTime) {
                        cellStyle.setDataFormat(dataFormat.getFormat(excelBody.bodyStyle().dateFormat()));
                        cell.setCellValue((LocalDateTime) o);
                    } else if (o instanceof LocalDate) {
                        cellStyle.setDataFormat(dataFormat.getFormat(excelBody.bodyStyle().dateFormat()));
                        cell.setCellValue((LocalDate) o);
                    }
                    cell.setCellStyle(cellStyle);
                    if (excelBody.colSpan() > 0 || excelBody.rowSpan() > 0) {
                        CellRangeAddress cellAddresses = new CellRangeAddress(cell.getAddress().getRow(), cell.getAddress().getRow() + excelBody.rowSpan(), cell.getAddress().getColumn(), cell.getAddress().getColumn() + excelBody.colSpan());
                        sheet.addMergedRegion(cellAddresses);
                    }
                    if ((excelBody.width() > 0 && excelBody.width() != 8) && sheet.getColumnWidth(excelBody.colIndex()) == 2048) {
                        sheet.setColumnWidth(excelBody.colIndex(), excelBody.width() * 256);
                    }
                }
            }
        }
        List<Field> groupField = Arrays.stream(excelClass.getDeclaredFields())
                .filter(field -> field.isAnnotationPresent(ExcelBody.class) && field.getDeclaredAnnotation(ExcelBody.class).rowGroup())
                .map(field -> {
                    field.setAccessible(true);
                    return field;
                })
                .sorted(Comparator.comparing(field -> field.getDeclaredAnnotation(ExcelBody.class).colIndex()))
                .collect(Collectors.toList());

        Map<Field, List<Integer>> groupMap = new HashMap<>();
        for (Field field : groupField){
            groupMap.put(field, new ArrayList<>());
            for(int i=0; i< data.size(); i++){
                Object o1 = field.get(data.get(i));

                for(int j = i+1;j < data.size(); j++){
                    Object o2 = field.get(data.get(j));
                    if(!o1.equals(o2)){
                        groupMap.get(field).add((j)* headerMap.size()+headerMap.keySet().size()-1);
                        i = j-1;
                        break;
                    }
                }
            }
            groupMap.get(field).add(sheet.getLastRowNum());
        }

        for(Field field: groupMap.keySet()){
            int dataRowIndex = headerMap.keySet().size();
            for(int i=0; i<groupMap.get(field).size(); i++){
                XSSFRow row = sheet.getRow(dataRowIndex);
                XSSFCell cell = row.getCell(field.getDeclaredAnnotation(ExcelBody.class).colIndex());
                if(!(dataRowIndex == groupMap.get(field).get(i))){
                    CellRangeAddress cellAddresses = new CellRangeAddress(dataRowIndex, groupMap.get(field).get(i), cell.getColumnIndex(), cell.getColumnIndex());
                    sheet.addMergedRegion(cellAddresses);
                }
                dataRowIndex = groupMap.get(field).get(i)+1;

            }
        }
        List<CellRangeAddress> mergedRegions = sheet.getMergedRegions();
        for(CellRangeAddress rangeAddress : mergedRegions) {
            RegionUtil.setBorderBottom(BorderStyle.THIN, rangeAddress, sheet);
            RegionUtil.setBorderLeft(BorderStyle.THIN, rangeAddress, sheet);
            RegionUtil.setBorderRight(BorderStyle.THIN, rangeAddress, sheet);
            RegionUtil.setBorderTop(BorderStyle.THIN, rangeAddress, sheet);
        }
        ByteArrayOutputStream byteArrayOutputStream = new ByteArrayOutputStream();
        workbook.write(byteArrayOutputStream);
        return ResponseEntity
                .ok()
                .header("Content-Transfer-Encoding", "binary")
                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=" + URLEncoder.encode(fileName, "UTF-8")+".xlsx")
                .contentType(MediaType.APPLICATION_OCTET_STREAM)
                .contentLength(byteArrayOutputStream.size())
                .body(new ByteArrayResource(byteArrayOutputStream.toByteArray()));
    }


    private static boolean isHex(String hexCode) {
        if (StringUtils.startsWith(hexCode, "#")) {
            for (Character c : hexCode.substring(1).toCharArray()) {
                switch (c) {
                    case '0':
                    case '1':
                    case '2':
                    case '3':
                    case '4':
                    case '5':
                    case '6':
                    case '7':
                    case '8':
                    case '9':
                    case 'a':
                    case 'b':
                    case 'c':
                    case 'd':
                    case 'e':
                    case 'f':
                    case 'A':
                    case 'B':
                    case 'C':
                    case 'D':
                    case 'E':
                    case 'F':
                        break;
                    default:
                        return false;
                }
            }
        } else {
            return false;
        }
        return true;
    }
}
