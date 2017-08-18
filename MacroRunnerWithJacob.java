package com.macro;

import java.io.BufferedReader;
import java.io.File;
import java.io.IOException;
import java.io.InputStreamReader;
import java.io.PrintWriter;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.CellReference;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.ComThread;
import com.jacob.com.Dispatch;
import com.jacob.com.Variant;


public class MacroRunnerWithJacob {

    public static void main(String[] args) throws IOException, InvalidFormatException {
        // TODO Auto-generated method stub
    	/*Macro runner starts*/
        
        Process theProcess = null;
        BufferedReader inStream = null;
   
        System.out.println("CallHelloPgm.main() invoked");
        String elpHome = System.getenv("ELP_Home");
        String script = "Set objExcel = CreateObject(\"Excel.Application\") \n" +
        		//"Set objWorkbook = objExcel.Workbooks.Open(\"E:\\Bala\\Rahul\\xls\\PlotJavaResultV6.xlsm\",0,True) \n" +
        		"Set objWorkbook = objExcel.Workbooks.Open(\""+elpHome+"\\PlotJavaResultV6.xlsm\",0,True) \n" +
        		"objExcel.Application.Visible = False \n"+
        		"objExcel.Application.Run \"uploadFile\", WScript.Arguments(0) \n" +
        		"objExcel.Application.Run \"PrintPlotResult\" \n" +
        		"objExcel.ActiveWorkbook.Close \n" +
        		"objExcel.Application.Quit \n" +
        		"WScript.Quit \n";
        
        
        try{
            PrintWriter writer = new PrintWriter(elpHome+"\\temp.vbs", "UTF-8");
            writer.println(script);
            writer.close();
        } catch (IOException e) {
           // do something
        }
   
        // call the Hello class
        try
        {
        	
            //theProcess = Runtime.getRuntime().exec("wscript E:\\Bala\\Rahul\\jacob-1.14.3-x86\\vbb.vbs "+elpHome+"\\10256025_12569898_0731163917.xls");
        	theProcess = Runtime.getRuntime().exec("wscript "+elpHome+"\\temp.vbs "+elpHome+"\\10256025_12569898_0731163917.xls");
        }
        catch(Exception e)
        {
           System.err.println("Error on exec() method");
           e.printStackTrace();  
        }
          
        // read from the called program's standard output stream
        try
        {
           inStream = new BufferedReader(
                                  new InputStreamReader( theProcess.getInputStream() ));  
           System.out.println("here"+inStream.readLine());
        }
        catch(IOException e)
        {
           System.err.println("Error on inStream.readLine()");
           e.printStackTrace();  
        }
        
    	File temp = new File(elpHome+"\\temp.vbs");
    	temp.delete();

        /*ends */

        /*File file = new File("E:\\Bala\\Rahul\\xls\\PlotJavaResultV6.xlsm");
        
		//XSSFSheet sheet;
        FileInputStream filexl = new FileInputStream(file);
   	 
        //Create Workbook instance holding reference to .xls file
        org.apache.poi.ss.usermodel.Workbook workbook = WorkbookFactory.create(filexl);
        //XSSFWorkbook workbook = new XSSFWorkbook(filexl);

        //Get desired sheet from the workbook
        //sheet = workbook.getSheet("PlotsSheetSystem");
        org.apache.poi.ss.usermodel.Sheet sheet = workbook.getSheet("PlotsSheetSystem");
		 
        //Model Data same for all Revision
        System.out.println(sheet.getRow(0).getCell(1).toString());
         //sheet.getRow(0).getCell(1).setCellValue("E:\\Bala\\Rahul\\xls");
         CellReference cr = new CellReference("A1");
         Row row = sheet.getRow(cr.getRow());
         Cell cell = row.getCell(cr.getCol());
         cell.setCellValue("E:\\Bala\\Rahul\\xls\\");
         System.out.println("cell+"+cell.getStringCellValue());
         System.out.println( sheet.getRow(1).getCell(1)+"after");
         filexl.close();
         FileOutputStream f2 = new FileOutputStream(file);
         workbook.write(f2);
         f2.close();
        
        //String macroName = "Macro";
        //callExcelMacro(file, macroName);
        System.out.println("done");
        String macroName2 = "uploadFile";
        callExcelMacro(file, macroName2);
        System.out.println("done2");
        String macroName3 = "PrintPlotResult";
        callExcelMacro(file, macroName3);
        System.out.println("Completed");*/
        
    }

    private static void callExcelMacro(File file, String macroName) {
        ComThread.InitSTA(true);
        final ActiveXComponent excel = new ActiveXComponent("Excel.Application");
        try{
            excel.setProperty("EnableEvents", new Variant(false));

            Dispatch workbooks = excel.getProperty("Workbooks")
                    .toDispatch();

            Dispatch workBook = Dispatch.call(workbooks, "Open",
                    file.getAbsolutePath()).toDispatch();

            // Calls the macro
            //Variant V1 = new Variant("\'"+file.getName()+"\'"+ macroName);
            Variant V1 =  new Variant(macroName);
            //Variant V1 = new Variant( file.getName() + macroName);
            Variant result = Dispatch.call(excel, "Run", V1);

            // Saves and closes
            Dispatch.call(workBook, "Save");

            com.jacob.com.Variant f = new com.jacob.com.Variant(true);
            Dispatch.call(workBook, "Close", f);

        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            excel.invoke("Quit", new Variant[0]);
            ComThread.Release();
        }
    }
}
