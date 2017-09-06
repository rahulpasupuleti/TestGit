package elp.gui;

import java.io.BufferedReader;
import java.io.File;
import java.io.IOException;
import java.io.InputStreamReader;
import java.io.PrintWriter;
import java.util.ArrayList;
import java.util.List;


public class MacroRunnerWithJacob {

    public static void generatePDF(String excelInput)  {
        // TODO Auto-generated method stub
    	/*Macro runner starts*/
        try{
        Process theProcess = null;
        BufferedReader inStream = null;
   
        //System.out.println("CallHelloPgm.main() invoked");
        //String elpHome = System.getenv("ELP_Home");
        String elpHome = System.getProperty("user.dir");
        
        String script = "Set objExcel = CreateObject(\"Excel.Application\") \n" +
        		"Set objWorkbook = objExcel.Workbooks.Open(\""+elpHome+"\\PlotJavaResultV6.xlsm\",0,False) \n" +
        		"objExcel.Application.WindowState = 1  \n"+
        		"objExcel.Application.ScreenUpdating = False \n"+
        		"objExcel.Application.enableEvents = False \n"+
        		"objExcel.Application.Visible = False \n"+
        		"objExcel.Application.Run \"uploadFile\", WScript.Arguments(0) \n" +
        		//"objExcel.Application.Visible = True \n"+
        		"objExcel.Application.Run \"PrintPlotResult\" \n" +
        		"objExcel.ActiveWorkbook.Close \n" +
        		"objExcel.Application.Quit \n" +
        		"WScript.Quit \n";        
        
        try{
            PrintWriter writer = new PrintWriter(elpHome+"\\temp.vbs", "UTF-8");
            writer.println(script);
            writer.close();
        } catch (IOException e) {
           e.printStackTrace();
        }
   
        // call the Hello class
        try
        {
        	File excelfile= new File(excelInput);
        	String excelFolder = excelfile.getParentFile().getAbsolutePath();
        	//System.out.println(excelFolder);
        	String currentPDF = getLatestFilefromDir(excelFolder);
        //theProcess = Runtime.getRuntime().exec("wscript E:\\Bala\\Rahul\\jacob-1.14.3-x86\\vbb.vbs "+elpHome+"\\10256025_12569898_0731163917.xls");
        	theProcess = Runtime.getRuntime().exec("wscript "+elpHome+"\\temp.vbs "+excelInput);
        	//System.out.println(excelInput);
        	System.out.println(currentPDF);
    	//System.out.println(theProcess.exitValue());
    		int attempt = 0;
    		boolean canKill = false;
    		while(attempt < 5){
    		String newPDF = getLatestFilefromDir(excelFolder);
    		System.out.println(newPDF);
    		//System.out.println(theProcess.exitValue());
    		attempt ++;
    			if(!currentPDF.equals("NoFile")){
    				if(!newPDF.equals("NoFile") && !newPDF.equals(currentPDF)){
    					canKill = true;
    				}
    			}else{
    				if(!newPDF.equals("NoFile")){
    					canKill = true;
    				}    				
    			}
    			if(attempt == 5 || canKill ){
    				try{
    					//theProcess.destroy();
    					isProcessRunning("EXCEL.EXE");
    				}catch(Exception e){
    					System.out.println("process is already killed " + e.getMessage());
    				}
    				break;
    			}
    			Thread.sleep(30000);
    	}}
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
           System.out.println("here "+inStream.readLine());      
        }
        catch(IOException e)
        {
           System.err.println("Error on inStream.readLine()");
           e.printStackTrace();  
        }
        
    	File temp = new File(elpHome+"\\temp.vbs");
    	temp.delete();
        }catch(Exception e){
        	
        }        
    }
    
    
    private static String getLatestFilefromDir(String dirPath){
        File dir = new File(dirPath);
        File[] files = dir.listFiles();
        List<File> plotresultFiles = new ArrayList<File>();
        if (files == null || files.length == 0) {
            return null;
        }
        for(File file : files){
        	if(file.getName().toString().startsWith("PlotResult") && file.getName().toString().endsWith(".pdf")){
        		plotresultFiles.add(file);
        	}
        }
        if(!plotresultFiles.isEmpty()){
	        File lastModifiedFile = plotresultFiles.get(0);
	        for (int i = 1; i < plotresultFiles.size(); i++) {
	           if (lastModifiedFile.lastModified() < plotresultFiles.get(i).lastModified()) {
	               lastModifiedFile =  plotresultFiles.get(i);
	           }
	        }
	        return lastModifiedFile.getName();
        }
        return "NoFile";
    }
    
    private static final String TASKLIST = "tasklist";
	private static final String KILL = "taskkill /F /IM ";

	public static boolean isProcessRunning(String serviceName) throws Exception {
		//Process parent = Runtime.getRuntime().exec("excel");
		serviceName = serviceName.toUpperCase();
	 Process p = Runtime.getRuntime().exec(TASKLIST);
	 BufferedReader reader = new BufferedReader(new InputStreamReader(
	   p.getInputStream()));
	 String line;
	//System.out.println("ID"+p.toString());
	 while ((line = reader.readLine()) != null) {

	  //System.out.println(line+"here");
	  if (line.toUpperCase().contains(serviceName)) {
		  //System.out.println("yes");
		  killProcess(serviceName);
	   return true;
	  }
	 }
	 return false;
	}

	public static void killProcess(String serviceName) throws Exception {
	  Runtime.getRuntime().exec(KILL + serviceName);
	 }
}