package models;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;

/**
 * OutputSheet class is responsible of the design of the sheet and the style
 * contains only static methods each responsible of a specific part of the sheet generation
 */

public class OutputSheet {

    /**
     *  defaultSheetAspects is responsible of setting the default
     *  width and height of each cell on the output sheet.
     *
     * @param outputSheet takes the output sheet as an input
     */
    public static void defaultSheetAspects(Sheet outputSheet){

        outputSheet.setDefaultColumnWidth(13);
        outputSheet.setColumnWidth(1,4000);
        outputSheet.setColumnWidth(2,4000);
        outputSheet.setColumnWidth(4,500);
        outputSheet.setColumnWidth(6,4000);
        outputSheet.setColumnWidth(7,4000);

        outputSheet.setDefaultRowHeightInPoints(10);


    }

    /**
     * getColor is a privet methode which is responsible of converting Strings
     * as "W,G" into the index of the desired colors W-> white, G->Green.
     *
     * @param color string most cases of length==1 which represent the first char of the color in english
     * @return short value which represent the index of the color in the apache.poi library
     */
    static Short getColor(String color){
        HashMap<String,Short> colors = new HashMap<>();
        colors.put("c",IndexedColors.WHITE1.getIndex());
        colors.put("w",IndexedColors.WHITE.getIndex());
        colors.put("l",IndexedColors.BLUE.getIndex());
        colors.put("gy",IndexedColors.GREY_50_PERCENT.getIndex());
        colors.put("y",IndexedColors.YELLOW.getIndex());
        colors.put("br",IndexedColors.BROWN.getIndex());
        colors.put("b",IndexedColors.BLACK.getIndex());
        colors.put("o",IndexedColors.ORANGE.getIndex());
        colors.put("p",IndexedColors.PINK.getIndex());
        colors.put("r",IndexedColors.RED.getIndex());
        colors.put("g",IndexedColors.GREEN.getIndex());
        colors.put("v",IndexedColors.VIOLET.getIndex());
        colors.put("sb",IndexedColors.SKY_BLUE.getIndex());
        colors.put("lg",IndexedColors.LIGHT_GREEN.getIndex());

        if(colors.containsKey(color.toLowerCase())){
            return colors.get(color.toLowerCase());
        }

        return  IndexedColors.WHITE.getIndex();
    }
    
    /**
     * createRows responsible of creating new rows on the output sheet
     *
     * @param outputRows the row objects on the output sheet
     * @param outputSheet the sheet object for the outputSheet
     * @param horizontalPosition the horizontal position which the sheet stopped at
     * @param verticalPosition the vertical Position represent
     *                         how many tickets where already placed on the sheet vertically
     *
     * @return list of Rows at the verticalPosition
     */
    public static List<Row> createRows(List<Row> outputRows, Sheet outputSheet,int horizontalPosition,int verticalPosition){

        if(horizontalPosition==0){
            outputRows = new ArrayList<>();
            outputRows.add(outputSheet.createRow(verticalPosition* 6));
            outputRows.add(outputSheet.createRow(verticalPosition * 6 + 1));
            outputRows.add(outputSheet.createRow(verticalPosition * 6 + 2));
            outputRows.add(outputSheet.createRow(verticalPosition * 6 + 3));
            outputRows.add(outputSheet.createRow(verticalPosition * 6 + 4));


            outputRows.get(1).setHeightInPoints(35);
            outputRows.get(3).setHeightInPoints(45);


        }

        return outputRows;
    }

    /**
     * createCells method is responsible of creating cells in a specific rows
     * and with a specific horizontal position.
     *
     * @param outputRows the rows list which the methode will create cells in
     * @param h the horizontal position which the methode will start creating at h*5 column
     * @return returns a two demotions cells List
     */
    public static List<List<Cell>> createCells(List<Row> outputRows,int h){
        List<List<Cell>> cells = new ArrayList<>();
        for (Row outputRow : outputRows) {
            List<Cell> rowCells = new ArrayList<>();
            for (int j = 0; j < 5; j++) {
                Cell cell = outputRow.createCell(j + h * 5);
                cell.setCellValue("");
                rowCells.add(cell);
            }
            cells.add(rowCells);
        }
        return cells;
    }

    /**
     * updateCells methode is responsible of updating the created cells
     * with the write value from the ticket provided.
     *
     * @param cells the 2D cells list which contains the created cells
     * @param ticket the input data in Ticket object
     */
    public static void updateCells(List<List<Cell>> cells, Ticket ticket){
        //first row
        cells.get(0).get(0).setCellValue(ticket.getCorA());
        cells.get(0).get(1).setCellValue(ticket.getInsertion());
        cells.get(0).get(3).setCellValue("Poste");

        //second row
        cells.get(1).get(0).setCellValue(ticket.getWireCrossSection());
        cells.get(1).get(1).setCellValue(ticket.getSkNumber());
        cells.get(1).get(3).setCellValue(ticket.getPost());

        //third row
        cells.get(2).get(0).setCellValue(ticket.getCorB());
        cells.get(2).get(1).setCellValue("-");

        //forth row
        cells.get(3).get(0).setCellValue(ticket.getSkNumber());

        //fifth row
        cells.get(4).get(0).setCellValue(ticket.getBase());
        cells.get(4).get(1).setCellValue(ticket.getProcess());
        cells.get(4).get(2).setCellValue(ticket.getFollowUp());
        cells.get(4).get(3).setCellValue(ticket.getSequence());

    }

    /** applyStyleOnCells methode responsible of styling the cells
     * with the desired font and colors etc...
     *
     * @param cells the list of cells in a 2D list.
     * @param outputBook the output workBook to add the fonts to it.
     */
    public static void applyStyleOnCells(List<List<Cell>> cells , Workbook outputBook){
        CellStyle cellStyle = null;
        Font font =null;

        //first row first cell
        {
            cellStyle = outputBook.createCellStyle();

            font = outputBook.createFont();

            font.setBold(true);
            font.setFontHeightInPoints((short) 14);

            cellStyle.setFont(font);

            cellStyle.setFillForegroundColor(getColor(cells.get(0).get(0).getStringCellValue()));
            cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

            cellStyle.setBorderTop(BorderStyle.THICK);
            cellStyle.setBorderRight(BorderStyle.THIN);
            cellStyle.setBorderLeft(BorderStyle.THICK);
            cellStyle.setBorderBottom(BorderStyle.THIN);

            cellStyle.setTopBorderColor(IndexedColors.BLACK.getIndex());
            cellStyle.setBottomBorderColor(IndexedColors.BLACK.getIndex());
            cellStyle.setLeftBorderColor(IndexedColors.BLACK.getIndex());
            cellStyle.setRightBorderColor(IndexedColors.BLACK.getIndex());

            cellStyle.setWrapText(true);
            cellStyle.setAlignment(HorizontalAlignment.CENTER);
            cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);


            cells.get(0).get(0).setCellStyle(cellStyle);
        }
        //first row second cell
        {
            cellStyle = outputBook.createCellStyle();

            font = outputBook.createFont();

            font.setBold(false);
            font.setFontHeightInPoints((short) 14);

            cellStyle.setFont(font);


            cellStyle.setBorderTop(BorderStyle.THICK);
            cellStyle.setBorderRight(BorderStyle.NONE);
            cellStyle.setBorderLeft(BorderStyle.NONE);
            cellStyle.setBorderBottom(BorderStyle.NONE);


            cellStyle.setWrapText(true);
            cellStyle.setAlignment(HorizontalAlignment.CENTER);
            cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);


            cells.get(0).get(1).setCellStyle(cellStyle);
        }
        //first row third cell
        {
            cellStyle = outputBook.createCellStyle();


            cellStyle.setBorderTop(BorderStyle.THICK);
            cellStyle.setBorderRight(BorderStyle.NONE);
            cellStyle.setBorderLeft(BorderStyle.NONE);
            cellStyle.setBorderBottom(BorderStyle.NONE);


            cellStyle.setWrapText(true);
            cellStyle.setAlignment(HorizontalAlignment.CENTER);
            cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);


            cells.get(0).get(2).setCellStyle(cellStyle);
        }
        //first row forth cell
        {
            cellStyle = outputBook.createCellStyle();

            font = outputBook.createFont();

            font.setBold(true);
            font.setFontHeightInPoints((short) 14);

            cellStyle.setFont(font);


            cellStyle.setBorderTop(BorderStyle.THICK);
            cellStyle.setBorderRight(BorderStyle.THICK);
            cellStyle.setBorderLeft(BorderStyle.THIN);
            cellStyle.setBorderBottom(BorderStyle.THIN);

            cellStyle.setTopBorderColor(IndexedColors.BLACK.getIndex());
            cellStyle.setBottomBorderColor(IndexedColors.BLACK.getIndex());
            cellStyle.setLeftBorderColor(IndexedColors.BLACK.getIndex());
            cellStyle.setRightBorderColor(IndexedColors.BLACK.getIndex());

            cellStyle.setWrapText(true);
            cellStyle.setAlignment(HorizontalAlignment.CENTER);
            cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);

            cells.get(0).get(3).setCellStyle(cellStyle);
        }

        //second row first cell
        {
            cellStyle = outputBook.createCellStyle();

            font = outputBook.createFont();

            font.setBold(true);
            font.setFontHeightInPoints((short) 14);

            cellStyle.setFont(font);


            cellStyle.setBorderTop(BorderStyle.THIN);
            cellStyle.setBorderRight(BorderStyle.THIN);
            cellStyle.setBorderLeft(BorderStyle.THICK);
            cellStyle.setBorderBottom(BorderStyle.THIN);

            cellStyle.setTopBorderColor(IndexedColors.BLACK.getIndex());
            cellStyle.setBottomBorderColor(IndexedColors.BLACK.getIndex());
            cellStyle.setLeftBorderColor(IndexedColors.BLACK.getIndex());
            cellStyle.setRightBorderColor(IndexedColors.BLACK.getIndex());

            cellStyle.setWrapText(true);
            cellStyle.setAlignment(HorizontalAlignment.CENTER);
            cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);

            cells.get(1).get(0).setCellStyle(cellStyle);
        }
        //second row second cell
        {
            cellStyle = outputBook.createCellStyle();

            font = outputBook.createFont();

            font.setBold(false);
            font.setFontHeightInPoints((short) 48);

            font.setFontName("Code 128");
            cellStyle.setFont(font);



            cellStyle.setBorderTop(BorderStyle.NONE);
            cellStyle.setBorderRight(BorderStyle.NONE);
            cellStyle.setBorderLeft(BorderStyle.NONE);
            cellStyle.setBorderBottom(BorderStyle.NONE);

            cellStyle.setTopBorderColor(IndexedColors.BLACK.getIndex());
            cellStyle.setBottomBorderColor(IndexedColors.BLACK.getIndex());
            cellStyle.setLeftBorderColor(IndexedColors.BLACK.getIndex());
            cellStyle.setRightBorderColor(IndexedColors.BLACK.getIndex());

            cellStyle.setWrapText(true);
            cellStyle.setAlignment(HorizontalAlignment.CENTER);
            cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);


            cells.get(1).get(1).setCellStyle(cellStyle);
        }
        //second row forth cell
        {
            cellStyle = outputBook.createCellStyle();

            font = outputBook.createFont();

            font.setBold(true);
            font.setFontHeightInPoints((short) 14);

            cellStyle.setFont(font);


            cellStyle.setBorderTop(BorderStyle.THIN);
            cellStyle.setBorderRight(BorderStyle.THICK);
            cellStyle.setBorderLeft(BorderStyle.THIN);
            cellStyle.setBorderBottom(BorderStyle.NONE);

            cellStyle.setTopBorderColor(IndexedColors.BLACK.getIndex());
            cellStyle.setBottomBorderColor(IndexedColors.BLACK.getIndex());
            cellStyle.setLeftBorderColor(IndexedColors.BLACK.getIndex());
            cellStyle.setRightBorderColor(IndexedColors.BLACK.getIndex());

            cellStyle.setWrapText(true);
            cellStyle.setAlignment(HorizontalAlignment.CENTER);
            cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);

            cells.get(1).get(3).setCellStyle(cellStyle);
        }

        //third row first cell
        {
            cellStyle = outputBook.createCellStyle();

            font = outputBook.createFont();

            font.setBold(true);
            font.setFontHeightInPoints((short) 14);

            cellStyle.setFont(font);

            cellStyle.setFillForegroundColor(getColor(cells.get(2).get(0).getStringCellValue()));
            cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

            cellStyle.setBorderTop(BorderStyle.THIN);
            cellStyle.setBorderRight(BorderStyle.THIN);
            cellStyle.setBorderLeft(BorderStyle.THICK);
            cellStyle.setBorderBottom(BorderStyle.THIN);

            cellStyle.setTopBorderColor(IndexedColors.BLACK.getIndex());
            cellStyle.setBottomBorderColor(IndexedColors.BLACK.getIndex());
            cellStyle.setLeftBorderColor(IndexedColors.BLACK.getIndex());
            cellStyle.setRightBorderColor(IndexedColors.BLACK.getIndex());

            cellStyle.setWrapText(true);
            cellStyle.setAlignment(HorizontalAlignment.CENTER);
            cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);


            cells.get(2).get(0).setCellStyle(cellStyle);
        }
        //third row second cell
        {
            cellStyle = outputBook.createCellStyle();

            font = outputBook.createFont();

            font.setBold(false);
            font.setFontHeightInPoints((short) 14);

            cellStyle.setFont(font);


            cellStyle.setBorderTop(BorderStyle.NONE);
            cellStyle.setBorderRight(BorderStyle.NONE);
            cellStyle.setBorderLeft(BorderStyle.NONE);
            cellStyle.setBorderBottom(BorderStyle.NONE);


            cellStyle.setWrapText(true);
            cellStyle.setAlignment(HorizontalAlignment.CENTER);
            cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);


            cells.get(2).get(1).setCellStyle(cellStyle);
        }
        //third row forth cell
        {
            cellStyle = outputBook.createCellStyle();

            font = outputBook.createFont();

            font.setBold(true);
            font.setFontHeightInPoints((short) 14);

            cellStyle.setFont(font);


            cellStyle.setBorderTop(BorderStyle.THIN);
            cellStyle.setBorderRight(BorderStyle.THICK);
            cellStyle.setBorderLeft(BorderStyle.THIN);
            cellStyle.setBorderBottom(BorderStyle.THIN);

            cellStyle.setTopBorderColor(IndexedColors.BLACK.getIndex());
            cellStyle.setBottomBorderColor(IndexedColors.BLACK.getIndex());
            cellStyle.setLeftBorderColor(IndexedColors.BLACK.getIndex());
            cellStyle.setRightBorderColor(IndexedColors.BLACK.getIndex());

            cellStyle.setWrapText(true);
            cellStyle.setAlignment(HorizontalAlignment.CENTER);
            cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);

            cells.get(2).get(3).setCellStyle(cellStyle);
        }

        //forth row first cell
        {
            cellStyle = outputBook.createCellStyle();

            font = outputBook.createFont();

            font.setBold(true);
            font.setFontHeightInPoints((short) 36);

            cellStyle.setFont(font);

            cellStyle.setBorderTop(BorderStyle.THIN);
            cellStyle.setBorderRight(BorderStyle.NONE);
            cellStyle.setBorderLeft(BorderStyle.THICK);
            cellStyle.setBorderBottom(BorderStyle.THIN);

            cellStyle.setTopBorderColor(IndexedColors.BLACK.getIndex());
            cellStyle.setBottomBorderColor(IndexedColors.BLACK.getIndex());
            cellStyle.setLeftBorderColor(IndexedColors.BLACK.getIndex());
            cellStyle.setRightBorderColor(IndexedColors.BLACK.getIndex());

            cellStyle.setWrapText(true);
            cellStyle.setAlignment(HorizontalAlignment.CENTER);
            cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);


            cells.get(3).get(0).setCellStyle(cellStyle);
        }
        //forth row second cell
        {
            cellStyle = outputBook.createCellStyle();

            font = outputBook.createFont();

            font.setBold(true);
            font.setFontHeightInPoints((short) 48);

            cellStyle.setFont(font);

            cellStyle.setBorderTop(BorderStyle.THIN);
            cellStyle.setBorderRight(BorderStyle.NONE);
            cellStyle.setBorderLeft(BorderStyle.NONE);
            cellStyle.setBorderBottom(BorderStyle.THIN);

            cellStyle.setTopBorderColor(IndexedColors.BLACK.getIndex());
            cellStyle.setBottomBorderColor(IndexedColors.BLACK.getIndex());
            cellStyle.setLeftBorderColor(IndexedColors.BLACK.getIndex());
            cellStyle.setRightBorderColor(IndexedColors.BLACK.getIndex());

            cellStyle.setWrapText(true);
            cellStyle.setAlignment(HorizontalAlignment.CENTER);
            cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);


            cells.get(3).get(1).setCellStyle(cellStyle);
        }
        //forth row third cell
        {
            cellStyle = outputBook.createCellStyle();

            font = outputBook.createFont();

            font.setBold(true);
            font.setFontHeightInPoints((short) 48);

            cellStyle.setFont(font);

            cellStyle.setBorderTop(BorderStyle.THIN);
            cellStyle.setBorderRight(BorderStyle.NONE);
            cellStyle.setBorderLeft(BorderStyle.NONE);
            cellStyle.setBorderBottom(BorderStyle.THIN);

            cellStyle.setTopBorderColor(IndexedColors.BLACK.getIndex());
            cellStyle.setBottomBorderColor(IndexedColors.BLACK.getIndex());
            cellStyle.setLeftBorderColor(IndexedColors.BLACK.getIndex());
            cellStyle.setRightBorderColor(IndexedColors.BLACK.getIndex());

            cellStyle.setWrapText(true);
            cellStyle.setAlignment(HorizontalAlignment.CENTER);
            cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);


            cells.get(3).get(2).setCellStyle(cellStyle);
        }
        //forth row forth cell
        {
            cellStyle = outputBook.createCellStyle();

            font = outputBook.createFont();

            font.setBold(true);
            font.setFontHeightInPoints((short) 48);

            cellStyle.setFont(font);

            cellStyle.setBorderTop(BorderStyle.THIN);
            cellStyle.setBorderRight(BorderStyle.THICK);
            cellStyle.setBorderLeft(BorderStyle.NONE);
            cellStyle.setBorderBottom(BorderStyle.THIN);

            cellStyle.setTopBorderColor(IndexedColors.BLACK.getIndex());
            cellStyle.setBottomBorderColor(IndexedColors.BLACK.getIndex());
            cellStyle.setLeftBorderColor(IndexedColors.BLACK.getIndex());
            cellStyle.setRightBorderColor(IndexedColors.BLACK.getIndex());

            cellStyle.setWrapText(true);
            cellStyle.setAlignment(HorizontalAlignment.CENTER);
            cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);


            cells.get(3).get(3).setCellStyle(cellStyle);
        }

        //fifth row first cell
        {
            cellStyle = outputBook.createCellStyle();

            font = outputBook.createFont();

            font.setBold(true);
            font.setFontHeightInPoints((short) 14);

            cellStyle.setFont(font);


            cellStyle.setBorderTop(BorderStyle.THIN);
            cellStyle.setBorderRight(BorderStyle.THIN);
            cellStyle.setBorderLeft(BorderStyle.THICK);
            cellStyle.setBorderBottom(BorderStyle.THICK);

            cellStyle.setTopBorderColor(IndexedColors.BLACK.getIndex());
            cellStyle.setBottomBorderColor(IndexedColors.BLACK.getIndex());
            cellStyle.setLeftBorderColor(IndexedColors.BLACK.getIndex());
            cellStyle.setRightBorderColor(IndexedColors.BLACK.getIndex());

            cellStyle.setWrapText(true);
            cellStyle.setAlignment(HorizontalAlignment.CENTER);
            cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);

            cells.get(4).get(0).setCellStyle(cellStyle);
        }
        //fifth row second cell
        {
            cellStyle = outputBook.createCellStyle();

            font = outputBook.createFont();

            font.setBold(true);
            font.setFontHeightInPoints((short) 14);

            cellStyle.setFont(font);


            cellStyle.setBorderTop(BorderStyle.THIN);
            cellStyle.setBorderRight(BorderStyle.THIN);
            cellStyle.setBorderLeft(BorderStyle.THIN);
            cellStyle.setBorderBottom(BorderStyle.THICK);

            cellStyle.setTopBorderColor(IndexedColors.BLACK.getIndex());
            cellStyle.setBottomBorderColor(IndexedColors.BLACK.getIndex());
            cellStyle.setLeftBorderColor(IndexedColors.BLACK.getIndex());
            cellStyle.setRightBorderColor(IndexedColors.BLACK.getIndex());

            cellStyle.setWrapText(true);
            cellStyle.setAlignment(HorizontalAlignment.CENTER);
            cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);

            cells.get(4).get(1).setCellStyle(cellStyle);
        }
        //fifth row third cell
        {
            cellStyle = outputBook.createCellStyle();

            font = outputBook.createFont();

            font.setBold(true);
            font.setFontHeightInPoints((short) 14);

            cellStyle.setFont(font);


            cellStyle.setBorderTop(BorderStyle.THIN);
            cellStyle.setBorderRight(BorderStyle.THIN);
            cellStyle.setBorderLeft(BorderStyle.THIN);
            cellStyle.setBorderBottom(BorderStyle.THICK);

            cellStyle.setTopBorderColor(IndexedColors.BLACK.getIndex());
            cellStyle.setBottomBorderColor(IndexedColors.BLACK.getIndex());
            cellStyle.setLeftBorderColor(IndexedColors.BLACK.getIndex());
            cellStyle.setRightBorderColor(IndexedColors.BLACK.getIndex());

            cellStyle.setWrapText(true);
            cellStyle.setAlignment(HorizontalAlignment.CENTER);
            cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);

            cells.get(4).get(2).setCellStyle(cellStyle);
        }
        //fifth row forth cell
        {
            cellStyle = outputBook.createCellStyle();

            font = outputBook.createFont();

            font.setBold(true);
            font.setFontHeightInPoints((short) 14);

            cellStyle.setFont(font);


            cellStyle.setBorderTop(BorderStyle.THIN);
            cellStyle.setBorderRight(BorderStyle.THICK);
            cellStyle.setBorderLeft(BorderStyle.THIN);
            cellStyle.setBorderBottom(BorderStyle.THICK);

            cellStyle.setTopBorderColor(IndexedColors.BLACK.getIndex());
            cellStyle.setBottomBorderColor(IndexedColors.BLACK.getIndex());
            cellStyle.setLeftBorderColor(IndexedColors.BLACK.getIndex());
            cellStyle.setRightBorderColor(IndexedColors.BLACK.getIndex());

            cellStyle.setWrapText(true);
            cellStyle.setAlignment(HorizontalAlignment.CENTER);
            cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);

            cells.get(4).get(3).setCellStyle(cellStyle);
        }


    }

    /**
     *  mergeCells methode is responsible of merging cells together
     *  to achieve the desired cells Style.
     *
     * @param outputSheet the output sheet
     * @param horizontalPosition the horizontal position which the generation is at the moment
     * @param verticalPosition the vertical position which the generation is at the moment
     */
    public static void mergeCells(Sheet outputSheet,int horizontalPosition,int verticalPosition){

        outputSheet.addMergedRegion(new CellRangeAddress(verticalPosition * 6, verticalPosition * 6, horizontalPosition * 5 + 1,  horizontalPosition * 5+2));
        outputSheet.addMergedRegion(new CellRangeAddress(verticalPosition * 6 + 2, verticalPosition * 6 + 2, 1 + horizontalPosition * 5, 2 + horizontalPosition * 5));
        outputSheet.addMergedRegion(new CellRangeAddress(verticalPosition * 6 + 1, verticalPosition * 6 + 1, 1 + horizontalPosition * 5, 2 + horizontalPosition * 5));
        outputSheet.addMergedRegion(new CellRangeAddress(verticalPosition * 6 + 1, verticalPosition * 6 + 2, 3 + horizontalPosition * 5, 3 + horizontalPosition * 5));
        outputSheet.addMergedRegion(new CellRangeAddress(verticalPosition * 6 + 3, verticalPosition * 6 + 3, horizontalPosition * 5, 3 + horizontalPosition * 5));

    }
}
