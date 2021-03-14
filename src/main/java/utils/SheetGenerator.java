package utils;

import models.OutputSheet;
import models.Ticket;
import org.apache.poi.ss.usermodel.*;
import java.util.ArrayList;
import java.util.List;

public class SheetGenerator {
    public static void sheetGenerator(List<Ticket> ticketList, Sheet outputSheet, Workbook outputBook) {

        int h = 0;
        int v = 0;
        List<Row> outputRows = new ArrayList<>();

        OutputSheet.defaultSheetAspects(outputSheet);

        for (Ticket ticket : ticketList) {

            outputRows = OutputSheet.createRows(outputRows,outputSheet,h,v);

            List<List<Cell>> cells = OutputSheet.createCells(outputRows,h);

            OutputSheet.updateCells(cells,ticket);

            OutputSheet.applyStyleOnCells(cells,outputBook);

            OutputSheet.mergeCells(outputSheet,h,v);

            if (h == 0) {
                h++;
            } else {
                h = 0;
                v++;
            }
        }

    }
}
