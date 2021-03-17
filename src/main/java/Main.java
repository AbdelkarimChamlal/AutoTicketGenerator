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
 * Main Class works as the front door of the project
 * the user can run the project using the command line
 * and this class is responsible of handling the format which the project is run in
 * this class contains only the main method
 */

public class Main {

    /**
     * main method which the user runs when starting the project
     * and initializes, handles everything.
     * <p>
     * the project meant to be run as a jar file.
     *
     * @param args string values that indicate the input and the output files of the autoTicketGenerator
     */
    public static void main(String[] args)   {

        //if the jar is runt without any args the program notifies the user to provide args
        //or use the -h arg to get help on how to use the program.
        if(args.length==0){
            System.out.println("\nto run this jar please provide the necessary arguments");
            System.out.println("or use -h to get more information on how to use this Jar\n");
        }else
        //in case the user run the command java -jar auto.jar -h
        //the program responses with a simple guide on how they should use the jar File and run it correctly
        if(args.length==1 && args[0].equals("-h")){

            System.out.println("\nWelcome to the AutoTicketGenerator Help Guide");
            System.out.println("to use this jar File please respect the formula below");
            System.out.println("java -jar jarFile.jar -i inputFile.xlsx -o outputFile.xlsx");
            System.out.println("where:");
            System.out.println("\t\t jarFile.jar should be replaced with the name and extension of this Jar");
            System.out.println("\t\t InputFile.xlsx should be the path to the input data");
            System.out.println("\t\t OutputFIle.xlsx should be the path and name of the desired output location\n");

        }else

        //in case of respecting the correct formula to run this jar
        //the program will try to read the input file
        //and process that input header first
        //then convert the input rows into Ticket objects with the help of the header
        //generate the output sheet and convert the tickets into cells inside the sheet

        if(args.length==4 && args[0].equals("-i") && args[2].equals("-o")){
            try{

                System.out.println("thank you for using AutoTicketGenerator ");
                //read the input and save it as a workbook
                FileInputStream inputStream = new FileInputStream(new File(args[1]));

                Workbook inputBook = new XSSFWorkbook(inputStream);

                inputStream.close();

                //get the sheet at index 0 from the workbook
                Sheet inputSheet = inputBook.getSheetAt(0);

                //generate the tickerList with the help of TicketGenerator class
                List<Ticket> ticketList = TicketGenerator.generateTicketsFromSheet(inputSheet);

                //create the output workbook
                Workbook outputBook = new XSSFWorkbook();

                //create the output sheet and name it tickets
                Sheet outputSheet = outputBook.createSheet("tickets");

                //place the ticketList each individual into a group of cells
                SheetGenerator.sheetGenerator(ticketList,outputSheet,outputBook);

                //initialize a fileStream for the output workbook
                // and write it into the desired path provided by the user

                FileOutputStream outputStream = new FileOutputStream(args[3]);
                outputBook.write(outputStream);
                outputStream.close();
            }catch(IOException e){
                //in case of any error with the IO processes
                System.out.println("input file is not valid or output file is used by another app and can't be overwritten ");
            }

        }else{
            //in case the user provided a false or missing arguments
            System.out.println("Please use -h to get more information on how to use this Jar");
        }
    }
}
