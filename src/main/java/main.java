import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class main {

    public static void main(String[] args) throws InvalidFormatException,
            IOException {

        String inpFn = "C:\\Users\\shaco\\Downloads\\GeozoneReport_13_05_2022.xls";
        String outFn = "C:\\Users\\shaco\\Downloads\\GeozoneReport_13_05_2022-style.xlsx";

        FileInputStream in = new FileInputStream(inpFn);
        try {
            Workbook wbIn = new HSSFWorkbook(in);
            File outF = new File(outFn);
            if (outF.exists())
                outF.delete();

            Workbook wbOut = new XSSFWorkbook();
            int sheetCnt = wbIn.getNumberOfSheets();
            for (int i = 0; i < sheetCnt; i++) {
                Sheet sIn = wbIn.getSheetAt(i);
                Sheet sOut = wbOut.createSheet(sIn.getSheetName());
                Iterator<Row> rowIt = sIn.rowIterator();
                while (rowIt.hasNext()) {
                    Row rowIn = rowIt.next();
                    Row rowOut = sOut.createRow(rowIn.getRowNum());

                    Iterator<Cell> cellIt = rowIn.cellIterator();
                    while (cellIt.hasNext()) {
                        Cell cellIn = cellIt.next();
                        Cell cellOut = rowOut.createCell(
                                cellIn.getColumnIndex(), cellIn.getCellType());

                        switch (cellIn.getCellType()) {
                            case Cell.CELL_TYPE_BLANK:
                                break;

                            case Cell.CELL_TYPE_BOOLEAN:
                                cellOut.setCellValue(cellIn.getBooleanCellValue());
                                break;

                            case Cell.CELL_TYPE_ERROR:
                                cellOut.setCellValue(cellIn.getErrorCellValue());
                                break;

                            case Cell.CELL_TYPE_FORMULA:
                                cellOut.setCellFormula(cellIn.getCellFormula());
                                break;

                            case Cell.CELL_TYPE_NUMERIC:
                                cellOut.setCellValue(cellIn.getNumericCellValue());
                                break;

                            case Cell.CELL_TYPE_STRING:
                                cellOut.setCellValue(cellIn.getStringCellValue());
                                break;
                        }

                        CellStyle styleIn = cellIn.getCellStyle();
                        CellStyle styleOut = cellOut.getCellStyle();
                        styleOut.setDataFormat(styleIn.getDataFormat());

                        styleIn.setBorderTop((short) 1);
                        styleIn.setBorderRight((short) 1);
                        styleIn.setBorderLeft((short) 1);
                        styleIn.setBorderBottom((short) 1);
                        styleIn.setFillForegroundColor(HSSFColor.LIGHT_TURQUOISE.index);
                        styleIn.setFillPattern(XSSFCellStyle.SOLID_FOREGROUND);
                        styleIn.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
                        styleIn.setAlignment(XSSFCellStyle.ALIGN_CENTER);
                        styleIn.setWrapText(true);
                        Font f = wbOut.createFont();
                        f.setFontHeightInPoints((short) 10);
                        f.setBoldweight((short) f.BOLDWEIGHT_BOLD);
//                            styleIn.setFont(f);
                        cellOut.setCellStyle(styleOut);
                        cellOut.setCellComment(cellIn.getCellComment());
                    }
                }
            }
            FileOutputStream out = new FileOutputStream(outF);
            try {
                wbOut.write(out);
            } finally {
                out.close();
            }
        } finally {
            in.close();
        }
    }
}
