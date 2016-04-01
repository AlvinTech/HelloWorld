package poi;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.HSSFCellUtil;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.util.CellRangeAddress;

import java.io.File;
import java.io.FileOutputStream;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.TimeZone;

/**
 * Created by Alvin on 16/4/1.
 */
public class ExportXls {


    public File exportTopGuests()
    {
       long fromDate = new Date().getTime();
       long toDate = new Date().getTime();

        HSSFWorkbook workbook = new HSSFWorkbook();
        HSSFSheet sheet = workbook.createSheet("TopGuestReport");

        HSSFCellStyle titleStyle = workbook.createCellStyle();

        HSSFFont titleFont = workbook.createFont();
        titleFont.setFontName(HSSFFont.FONT_ARIAL);
        titleFont.setFontHeightInPoints((short) 10);
        titleFont.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
        titleFont.setColor(HSSFColor.WHITE.index);

        titleStyle.setAlignment(HSSFCellStyle.ALIGN_LEFT);
        titleStyle.setBorderLeft(HSSFCellStyle.BORDER_THIN);
        titleStyle.setBorderRight(HSSFCellStyle.BORDER_THIN);
        titleStyle.setBorderTop(HSSFCellStyle.BORDER_THIN);
        titleStyle.setBorderBottom(HSSFCellStyle.BORDER_THIN);
        titleStyle.setFillForegroundColor(HSSFColor.BLUE_GREY.index);
        titleStyle.setFillBackgroundColor(HSSFColor.BLUE_GREY.index);
        titleStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
        titleStyle.setFont(titleFont);

        HSSFCellStyle headStyle = workbook.createCellStyle();

        HSSFFont headFont = workbook.createFont();
        headFont.setFontName(HSSFFont.FONT_ARIAL);
        headFont.setFontHeightInPoints((short) 10);
        headFont.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
        headFont.setColor(HSSFColor.GREY_50_PERCENT.index);

        headStyle.setAlignment(HSSFCellStyle.ALIGN_RIGHT);
        headStyle.setBorderLeft(HSSFCellStyle.BORDER_THIN);
        headStyle.setBorderRight(HSSFCellStyle.BORDER_THIN);
        headStyle.setBorderTop(HSSFCellStyle.BORDER_THIN);
        headStyle.setBorderBottom(HSSFCellStyle.BORDER_THIN);
        headStyle.setFillForegroundColor(HSSFColor.LIGHT_CORNFLOWER_BLUE.index);
        headStyle.setFillBackgroundColor(HSSFColor.LIGHT_CORNFLOWER_BLUE.index);
        headStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
        headStyle.setFont(headFont);

        HSSFCellStyle cellStyle = workbook.createCellStyle();

        HSSFFont cellFont = workbook.createFont();
        cellFont.setFontName(HSSFFont.FONT_ARIAL);
        cellFont.setFontHeightInPoints((short) 10);

        cellStyle.setAlignment(HSSFCellStyle.ALIGN_RIGHT);
        cellStyle.setBorderLeft(HSSFCellStyle.BORDER_THIN);
        cellStyle.setBorderRight(HSSFCellStyle.BORDER_THIN);
        cellStyle.setBorderTop(HSSFCellStyle.BORDER_THIN);
        cellStyle.setBorderBottom(HSSFCellStyle.BORDER_THIN);
        cellStyle.setFont(cellFont);

        HSSFRow queryTitleRow = sheet.createRow(0);
        HSSFCell queryTitleCell = queryTitleRow.createCell(0);

        CellRangeAddress region = new CellRangeAddress(0, 1, 0, 1);
        sheet.addMergedRegion(region);
        setRegionStyle(sheet, region, titleStyle);

        String arg1 = "Top 5 Guests";
        String arg2 = "Authentication Requests";






        queryTitleCell.setCellValue("Top 5 Guests by Authentication Requests");
        queryTitleCell.setCellStyle(titleStyle);

        region = new CellRangeAddress(0, 0, 2, 3);
        sheet.addMergedRegion(region);
        setRegionStyle(sheet, region, titleStyle);

        queryTitleCell = queryTitleRow.createCell(2);
        String arg = "Time Period";
        queryTitleCell.setCellValue(arg);
        queryTitleCell.setCellStyle(titleStyle);

        queryTitleRow = sheet.createRow(1);
        region = new CellRangeAddress(1, 1, 2, 3);
        sheet.addMergedRegion(region);
        setRegionStyle(sheet, region, titleStyle);

        queryTitleCell = queryTitleRow.createCell(2);
        SimpleDateFormat format = new SimpleDateFormat("yyyyy.MMMMM.dd GGG hh:mm aaa");
        format.setTimeZone(TimeZone.getDefault());
        arg = format.format(fromDate) + " ~ " + format.format(toDate);
        queryTitleCell.setCellValue(arg);
        queryTitleCell.setCellStyle(titleStyle);

        HSSFRow header = sheet.createRow(2);
        for (ETopGuestsColumnInfo columnInfo : ETopGuestsColumnInfo.values())
        {
            HSSFCell cell = header.createCell(columnInfo.index());
            cell.setCellType(columnInfo.cellType());
            String title = "";
            switch (columnInfo.index())
            {
            case 0:
                title = "Rank";
                break;
            case 1:
                title = "Guest Name";
                break;
            case 2:
                title = "Login Name";
                break;
            case 3:
                title = "Total Authentication Requests"+ arg2;
                break;
            }
            cell.setCellValue(new HSSFRichTextString(title));
            cell.setCellStyle(headStyle);
            sheet.setColumnWidth(columnInfo.index(), columnInfo.width());
        }

        List<IdmUser> allTopGuests = initUser();
        int rowCount = 3;

        for (int i = 0; i < allTopGuests.size(); i++)
        {
            IdmUser everyGuests = allTopGuests.get(i);
            HSSFRow contentRow = sheet.createRow(rowCount);
            HSSFCell cell = null;
            cell = contentRow.createCell(ETopGuestsColumnInfo.RANK.index());
            cell.setCellValue(i + 1);
            cell.setCellStyle(cellStyle);

            cell = contentRow.createCell(ETopGuestsColumnInfo.GUEST_NAME.index());
            cell.setCellValue(everyGuests.getGuestName());
            cell.setCellStyle(cellStyle);

            cell = contentRow.createCell(ETopGuestsColumnInfo.LOGIN_NAME.index());
            cell.setCellValue(everyGuests.getLoginNameForDisplay());
            cell.setCellStyle(cellStyle);

            cell = contentRow.createCell(ETopGuestsColumnInfo.TOTAL.index());
            cell.setCellValue(everyGuests.getTotalRequestsStr());
            cell.setCellStyle(cellStyle);

            rowCount++;
        }
        File fileToWrite = new File(System.getProperty("user.dir"), "TopGuestReport_" + System.currentTimeMillis() + ".xls");
        try
        {
            fileToWrite.createNewFile();
            FileOutputStream out = new FileOutputStream(fileToWrite);
            workbook.write(out);
            out.flush();
            out.close();
        } catch (Exception e)
        {
            throw new RuntimeException(e);
        }
        return fileToWrite;

    }

    private List<IdmUser> initUser() {
        List<IdmUser> lists = new ArrayList<IdmUser>();
        for (int i = 0; i < 10; i++) {
            IdmUser user = new IdmUser();
            user.setGuestName("guestname:" + i);
            user.setLoginNameForDisplay("LoginName"+i);
            user.setTotalRequestsStr("requestStr"+i);
            lists.add(user);
        }
        return lists;
    }

    private void setRegionStyle(HSSFSheet sheet, CellRangeAddress region, HSSFCellStyle cs)
    {
        int toprowNum = region.getFirstRow();
        for (int i = toprowNum; i <= region.getLastRow(); i++)
        {
            HSSFRow row = HSSFCellUtil.getRow(i, sheet);
            for (int j = region.getFirstColumn(); j <= region.getLastColumn(); j++)
            {
                HSSFCell cell = HSSFCellUtil.getCell(row, (short) j);
                cell.setCellStyle(cs);
            }
        }
    }

    public static void main(String[] args) {
        new ExportXls().exportTopGuests();
    }
}
