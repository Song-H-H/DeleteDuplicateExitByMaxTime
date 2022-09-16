import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.lang.reflect.Field;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.stream.Collectors;


public class DeleteDuplicateExit {

    static DateTimeFormatter timeFormatter = DateTimeFormatter.ofPattern("HH시mm분ss초");
    static DateTimeFormatter dateFormatter = DateTimeFormatter.ofPattern("yyyy년MM월dd일");

    public static void main(String[] args) throws Exception {
        ArrayList<ExcelHeader> excelHeaderDataList = new ArrayList<ExcelHeader>();
        readExcelFile(excelHeaderDataList);
        List<List<ExcelHeader>> groupingDate = processor(excelHeaderDataList);
        writeExcelFile(groupingDate);
    }

    private static void writeExcelFile(List<List<ExcelHeader>> groupingDate) throws IllegalAccessException {
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("DeleteDuplicateExitByMaxTime");
        sheet.setDefaultColumnWidth(15);

        Class dataClass = ExcelHeader.class;
        Field[] fields = dataClass.getDeclaredFields();
        int rownum = 0;
        int cellIndex = 0;
        for(List<ExcelHeader> groupingCard : groupingDate){
            for (ExcelHeader data : groupingCard) {
                Row row = sheet.createRow(rownum++);
                cellIndex = 0;
                for (Field field : fields) {
                    field.setAccessible(true);
                    Object object = field.get(data);
                    Cell cell = row.createCell(cellIndex++);
                    cell.setCellValue(object.toString());
                }
            }
        }

        try (FileOutputStream out = new FileOutputStream(new File("./", "OUPUT.xlsx"))) {
            workbook.write(out);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static List<List<ExcelHeader>> processor(ArrayList<ExcelHeader> excelHeaderDataList) {
        List<List<ExcelHeader>> groupingDate = excelHeaderDataList.stream()
                .collect(Collectors.groupingBy(ExcelHeader::getIssueDate, Collectors.groupingBy(ExcelHeader::getCardNumber)))
                .entrySet()
                .stream()
                .sorted((a, b) -> a.getKey().compareTo(b.getKey()))
                .map(x -> x.getValue()
                        .entrySet()
                        .stream()
                        .map(z -> z.getValue())
                        .map(y -> y.stream()
                                .reduce((a, b) -> (a.getIssueTime().isAfter(b.getIssueTime())) ? a : b)
                                .get())
                        .sorted((a, b) -> a.getIssueTime().compareTo(b.getIssueTime()))
                        .collect(Collectors.toList()))
                .collect(Collectors.toList());
        return groupingDate;
    }

    private static void readExcelFile(ArrayList<ExcelHeader> excelHeaderDataList) throws Exception {
        try {
            //파일 read
            File readFilePath = new File("./INPUT.xlsx");
            Workbook workbook = new XSSFWorkbook(new FileInputStream(readFilePath));

            Sheet worksheet = workbook.getSheetAt(0);
            Iterator<Row> rowIterator = worksheet.iterator();

            //첫번쨰 행 Skip
            rowIterator.next();
            while (rowIterator.hasNext()) {
                Row readRow = rowIterator.next();
                Iterator<Cell> cellIterator = readRow.cellIterator();
                ExcelHeader excelHeaderData = new ExcelHeader();

                List<String> values = new ArrayList<>();
                List<String> colums = new ArrayList<>();
                String cellData;
                while (cellIterator.hasNext()) {
                    Cell readCell = cellIterator.next();

                    //컬럼 type 분기처리
                    switch (readCell.getCellType()) {
                        case STRING:
                            cellData = readCell.getRichStringCellValue().getString();
                            break;
                        case NUMERIC:
                            if (DateUtil.isCellDateFormatted(readCell)) {
                                cellData = readCell.getDateCellValue().toString();
                            } else {
                                Long roundVal = Math.round(readCell.getNumericCellValue());
                                Double doubleVal = readCell.getNumericCellValue();
                                if (doubleVal.equals(roundVal.doubleValue())) {
                                    cellData = String.valueOf(roundVal);
                                } else {
                                    cellData = String.valueOf(doubleVal);
                                }
                            }
                            break;
                        case BOOLEAN:
                            cellData = String.valueOf(readCell.getBooleanCellValue());
                            break;
                        case FORMULA:
                            cellData = readCell.getCellFormula();
                            break;
                        default:
                            cellData = "";
                    }

                    int cellIndex = readCell.getColumnIndex();
                    excelHeaderData.setHeaderIndexValue(cellIndex, cellData);
                }

                excelHeaderDataList.add(excelHeaderData);
            }
            workbook.close();

        }  catch (Exception e) {
            throw new Exception(e.getMessage());
        }
    }
}
