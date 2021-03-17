package utils;

import models.Ticket;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;

/**
 * TicketGenerator class responsible for turning the input data
 * from the xlsx file into a list of ticket objects
 * while the first row represent the header
 * and other rows convert to ticket objects
 */
public class TicketGenerator {

    /**
     * HeaderValuesDetector method fetches throw the header
     * and create a hashMap to understand the schema
     * of the input to detect the values of each column.
     *
     * @param header Row object which represent the first row of the input data
     * @return a Hashmap which indicate the value of each column.
     */
    private static HashMap<String, Integer> headerValuesDetector(Row header){

        Iterator<Cell> cellIterator = header.cellIterator();

        HashMap<String,Integer> headerValues = new HashMap<>();
        int cellNumber=0;

        while(cellIterator.hasNext()){
            Cell currentCell = cellIterator.next();

            String currentCellStringValue = currentCell.getStringCellValue();

            if(currentCellStringValue.toLowerCase().contains("chain")){
                headerValues.put("chain",cellNumber);
            }else
            if(currentCellStringValue.toLowerCase().contains("base")){
                headerValues.put("base",cellNumber);
            }else
            if(currentCellStringValue.toLowerCase().contains("type")){
                headerValues.put("type",cellNumber);
            }else
            if(currentCellStringValue.toLowerCase().contains("process")){
                headerValues.put("process",cellNumber);
            }else
            if(currentCellStringValue.toLowerCase().contains("sk")){
                headerValues.put("sk",cellNumber);
            }else
            if(currentCellStringValue.toLowerCase().contains("follow")){
                headerValues.put("follow",cellNumber);
            }else
            if(currentCellStringValue.toLowerCase().contains("color") && currentCellStringValue.toLowerCase().contains("a")){
                headerValues.put("cor1",cellNumber);
            }else
            if(currentCellStringValue.toLowerCase().contains("color") && currentCellStringValue.toLowerCase().contains("b")){
                headerValues.put("cor2",cellNumber);
            }else
            if(currentCellStringValue.toLowerCase().contains("section")){
                headerValues.put("section",cellNumber);
            }else
            if(currentCellStringValue.toLowerCase().contains("insertion")){
                headerValues.put("insertion",cellNumber);
            }else
            if(currentCellStringValue.toLowerCase().contains("post")){
                headerValues.put("post",cellNumber);
            }else
            if(currentCellStringValue.toLowerCase().contains("sequence")){
                headerValues.put("sequence",cellNumber);
            }else
            if(currentCellStringValue.toLowerCase().contains("size")){
                headerValues.put("size",cellNumber);
            }
            cellNumber++;
        }

        return headerValues;
    }

    /**
     *  convertRowIntoTicket is responsible of matching
     *  the input rows with the headerValues to initialize
     *  a ticket object and place the value of each column
     *  from the input into the correct parameter in the ticket object.
     *
     * @param inputRow Row object which represent a row in the xlsx input file
     * @param headerValues hashMap which represents the values of each column of the input row
     * @return Ticket object with the values from the input row
     */
    private static Ticket convertRowIntoTicket(Row inputRow,HashMap<String,Integer> headerValues){
        Ticket ticket = new Ticket();

        if(headerValues.containsKey("chain")){
            ticket.setChain(inputRow.getCell(headerValues.get("chain")).getStringCellValue());
        }

        if(headerValues.containsKey("base")){
            ticket.setBase(inputRow.getCell(headerValues.get("base")).getStringCellValue());
        }

        if(headerValues.containsKey("cor1")){
            ticket.setCorA(inputRow.getCell(headerValues.get("cor1")).getStringCellValue());
        }

        if(headerValues.containsKey("cor2")){
            ticket.setCorB(inputRow.getCell(headerValues.get("cor2")).getStringCellValue());
        }

        if(headerValues.containsKey("type")){
            ticket.setWireType(inputRow.getCell(headerValues.get("type")).getStringCellValue());
        }

        if(headerValues.containsKey("process")){
            ticket.setProcess(inputRow.getCell(headerValues.get("process")).getStringCellValue());
        }

        if(headerValues.containsKey("sk")){
            ticket.setSkNumber(inputRow.getCell(headerValues.get("sk")).getStringCellValue());
        }

        if(headerValues.containsKey("follow")){
            ticket.setFollowUp(inputRow.getCell(headerValues.get("follow")).getStringCellValue());
        }

        if(headerValues.containsKey("section")){
            ticket.setWireCrossSection(inputRow.getCell(headerValues.get("section")).getStringCellValue());
        }

        if(headerValues.containsKey("insertion")){
            ticket.setInsertion(inputRow.getCell(headerValues.get("insertion")).getStringCellValue());
        }

        if(headerValues.containsKey("post")){
            ticket.setPost(inputRow.getCell(headerValues.get("post")).getStringCellValue());
        }

        if(headerValues.containsKey("sequence")){
            ticket.setSequence(inputRow.getCell(headerValues.get("sequence")).getStringCellValue());
        }

        if(headerValues.containsKey("size")){
            ticket.setSize(inputRow.getCell(headerValues.get("size")).getStringCellValue());
        }

        return ticket;
    }

    /**
     *
     * @param inputSheet Sheet object which is taken from the input workBook
     * @return list of Ticket objects which represent the input value converted into a Ticket object
     */
    public static List<Ticket> generateTicketsFromSheet(Sheet inputSheet){

        Row header = inputSheet.getRow(0);

        HashMap<String, Integer> headerValues = headerValuesDetector(header);

        List<Ticket> tickets = new ArrayList<>(inputSheet.getLastRowNum()-1);

        for(int i =1; i<=inputSheet.getLastRowNum();i++){
            Row currentRow = inputSheet.getRow(i);
            if(currentRow.getLastCellNum()!=0){
                tickets.add(convertRowIntoTicket(currentRow,headerValues));
            }
        }

        return tickets;
    }
}
