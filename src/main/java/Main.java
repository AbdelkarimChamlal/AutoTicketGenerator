import models.Ticket;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import utils.SheetGenerator;
import utils.TicketGenerator;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

/**
 * Main class of the auto ticket project
 * this class will be in direct contact with the user
 *
 */

public class Main {
    public static void main(String[] args)   {
        if(args.length==0){
            System.out.println("Please use -h to get more information on how to use this Jar");
        }else if(args.length==1 && args[0].equals("-h")){


            System.out.println("\n\nto use AutoTicketGenerator correctly please respect the formula of the running command:\n\n\n\t```java -jar jarName.jar -i inputFile.xlsx -o outputFile.xlsx```\n\n");


        }else if(args.length==4 && args[0].equals("-i") && args[2].equals("-o")){
            try{

                System.out.println("thank you for using AutoTicketGenerator ");

                FileInputStream inputStream = new FileInputStream(new File(args[1]));

                Workbook inputBook = new XSSFWorkbook(inputStream);

                inputStream.close();

                Sheet inputSheet = inputBook.getSheetAt(0);

                List<Ticket> ticketList = TicketGenerator.generateTicketsFromSheet(inputSheet);

                Workbook outputBook = new XSSFWorkbook();

                Sheet outputSheet = outputBook.createSheet("tickets");

                SheetGenerator.sheetGenerator(ticketList,outputSheet,outputBook);

                FileOutputStream outputStream = new FileOutputStream(args[3]);

                outputBook.write(outputStream);

                outputStream.close();




            }catch(IOException e){
                System.out.println("input file is not valid or output file is used by another app and can't be overwritten ");
            }

        }else{
            System.out.println("Please use -h to get more information on how to use this Jar");
        }
    }
}
