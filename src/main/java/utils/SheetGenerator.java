package utils;

import models.OutputSheet;
import models.Ticket;
import org.apache.poi.ss.usermodel.*;
import java.util.ArrayList;
import java.util.List;

/**
 * SheetGenerator class responsible for handling
 * the creation,updating,styling,merging the cells of the output sheet
 * in such a way that it is dynamic and the code that handles
 * all those steps are imported from other classes.
 *
 */
public class SheetGenerator {
    /**
     * sheetGenerator method handles the generation of the output sheet
     * each ticket will be in the format below
     * -----------------------------------------
     * |color a  | insertion            | poste|
     * |section  | barCode              | SPS1 |
     * |color b  |       -              |      |
     * |             sk code                   |
     * |base     | code |  followUp     | dunno|
     *
     * @param ticketList all the input rows placed in a list of Ticket objects.
     * @param outputSheet Sheet object which all the operation will be placed on.
     * @param outputBook WorkBook object which contains the outputSheet.
     */
    public static void sheetGenerator(List<Ticket> ticketList, Sheet outputSheet, Workbook outputBook) {
        // initial horizontal and vertical position
        int h = 0;
        int v = 0;

        //initial output rows
        //they represent the current rows which the creation of cells will be done on
        List<Row> outputRows = new ArrayList<>();

        //initialize the aspects of column width and default height size for each cell
        OutputSheet.defaultSheetAspects(outputSheet);

        //fetch for each ticket in the ticketList
        for (Ticket ticket : ticketList) {

            //create the rows at the specific v,h position
            outputRows = OutputSheet.createRows(outputRows,outputSheet,h,v);

            //create cells on those rows at the h position
            List<List<Cell>> cells = OutputSheet.createCells(outputRows,h);

            //update the values of those cells from default to the matching values from the ticket object
            OutputSheet.updateCells(cells,ticket);

            //apply the desired styling on each cell
            OutputSheet.applyStyleOnCells(cells,outputBook);

            //merge the desired cells
            OutputSheet.mergeCells(outputSheet,h,v);

            //check the horizontal position and adjust the values for the next ticket
            //each vertical position takes two tickets
            if (h == 0) {
                h++;
            } else {
                h = 0;
                v++;
            }
        }

    }
}
