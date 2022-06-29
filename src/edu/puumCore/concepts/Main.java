package edu.puumCore.concepts;

import edu.puumCore.concepts._custom.Assistant;
import edu.puumCore.concepts._interfaces.Excel;

import java.io.File;
import java.io.IOException;

public class Main {

    public static void main(String[] args) {
	// write your code here
        try {
            File file = new File("C:\\Users\\---\\Desktop\\Book1.xlsx");
            Excel excel = new Assistant();
            /*excel.write_to_file(file);
            System.out.println("Completed writing to file");*/
            excel.create_table(file);
            System.out.println("Completed creating table to file");
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
