package TestRutSinc.TestRutSinc;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.annotations.Test;

public class AppTest{
	
	public List<String> listaTXT = new ArrayList<String>();
	public String miMatriz [][];
	public String Estado[];
	public WebDriver driver;
	
	
	@Test(priority = 1)
	public void SetearDriver() throws IOException {

		System.setProperty("webdriver.chrome.driver", "C:\\webdriver\\chromedriver.exe"); // Para Chrome
		driver = new ChromeDriver(); // para Chrome
		driver.manage().window().maximize();
		String url = "http://localhost:56011/";
		driver.get(url);		
	}
	
	
	@Test(priority = 2)
	public void BuscarRegistro() throws IOException, InterruptedException {
		
	

        //Prepare the parameters of excel file
        String filePath = System.getProperty("user.dir")+"\\src";
        String fileName = "Libro2.xls";
		String sheetName = "Hoja1";

        //Call read file method of the class to read data
        ReadExcelArray(filePath ,fileName, sheetName);
		
               
		for(int i=1; i<miMatriz.length;i++)
		{			
			
				LimpiaInputs();
														
				WebDriverWait wait = new WebDriverWait(driver, 10);				
								
				wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("rutPersona")));
				WebElement TxtRut = driver.findElement(By.id("rutPersona"));
				TxtRut.sendKeys(miMatriz[i][0]); 
								
				wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("rutEmpresa")));
				WebElement TxtRutEmp = driver.findElement(By.id("rutEmpresa"));
				TxtRutEmp.sendKeys(miMatriz[i][1]); 				
				
				wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("CCotizacion")));
				WebElement TxtCentro = driver.findElement(By.id("CCotizacion"));
				TxtCentro.sendKeys(miMatriz[i][2]); 
					
				//wait de 0.5 segundos
				//Thread.sleep(500);
				
				Estado[i]="OK";
				if(miMatriz[i][2].length()>3) {Estado[i]="ERROR";} //Valida Largo Centro Cotizacion
				
				WebElement BtnSincronizar = driver.findElement(By.xpath("//input[@class='btn btn-default']"));
				BtnSincronizar.click();
				
				//wait de 1 segundo
				Thread.sleep(1000);		
				
				driver.navigate().back();
							
		}//cierre for	
		
		//Graba todos los estados
		WriteExcelArray(filePath ,fileName,sheetName, Estado);
	}
	
    public void LimpiaInputs() throws InterruptedException {
    	
    	WebDriverWait wait = new WebDriverWait(driver, 10);
    	
    	wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("ticket")));
		WebElement TxtTicket = driver.findElement(By.id("ticket"));
		TxtTicket.clear();
		
				
		wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("rutPersona")));
		WebElement TxtRut = driver.findElement(By.id("rutPersona"));
		TxtRut.clear();
		
		
		wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("dvPersona")));
		WebElement TxtDV = driver.findElement(By.id("dvPersona"));
		TxtDV.clear();
		
		
		wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("rutEmpresa")));
		WebElement TxtRutEmp = driver.findElement(By.id("rutEmpresa"));
		TxtRutEmp.clear();
		
		
		wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("dvEmpresa")));
		WebElement TxtDVEmp = driver.findElement(By.id("dvEmpresa"));
		TxtDVEmp.clear();
		
		
		wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("CCotizacion")));
		WebElement TxtCentro = driver.findElement(By.id("CCotizacion"));
		TxtCentro.clear();
    	
    }
        
    public void LeeTxtLista() {
    	
    	BufferedReader br = null;

        try {

            String sCurrentLine;           
            
            br = new BufferedReader(new FileReader("C:\\Users\\daniel.manzo\\Documents\\CAJA LOS ANDES\\ruts.txt"));

            while ((sCurrentLine = br.readLine()) != null) {
                //System.out.println(sCurrentLine);
            	listaTXT.add(sCurrentLine);                 
            }

        } catch (IOException e) {
            e.printStackTrace();
        }   finally {
            try {
                if (br != null)br.close();
            } catch (IOException ex) {
                ex.printStackTrace();
            }
        }            	
    }
    
	
    
    
    public void ReadExcelArray(String filePath,String fileName,String sheetName) throws IOException{    	
    	
    	//Create an object of File class to open xlsx file
        File file = new File(filePath+"\\"+fileName);
    	
        //Create an object of FileInputStream class to read excel file
        FileInputStream inputStream = new FileInputStream(file);
    	    
        
        //Find the file extension by splitting file name in substring  and getting only extension name
        String fileExtensionName = fileName.substring(fileName.indexOf("."));
    	
        Sheet hoja = null;
        
        //Check condition if the file is xlsx file(ERROR CON FORMATO XLSX??????)
        if(fileExtensionName.equals(".xlsx")){
        	XSSFWorkbook workbookXSSF  = new XSSFWorkbook(inputStream);//If it is xlsx file then create object of XSSFWorkbook class
        	hoja = workbookXSSF.getSheet(sheetName);//Read sheet inside the workbook by its name
        }//Check condition if the file is xls file
        else if(fileExtensionName.equals(".xls")){
        	HSSFWorkbook workbookHSSF = new HSSFWorkbook(inputStream); //If it is xls file then create object of XSSFWorkbook class
        	hoja = workbookHSSF.getSheet(sheetName);//Read sheet inside the workbook by its name
        }
                  
        //Find number of rows in excel file
        int rowCount = hoja.getLastRowNum()-hoja.getFirstRowNum()+1;
        
        //Asigna Tamaño a mi Matriz
        miMatriz =new String [rowCount][4];
        
        //Asigna Tamaño arreglo Estado
        Estado = new String [rowCount];
        
        //Create a loop over all the rows of excel file to read it
        for (int i = 0; i < rowCount; i++) {

            Row row = hoja.getRow(i);
            
            //Create a loop to print cell values in a row
            for (int j = 0; j < row.getLastCellNum(); j++) {

                //Print Excel data in console
                System.out.print(row.getCell(j).getStringCellValue()+"|| ");
                                
				miMatriz[i][j]= row.getCell(j).getStringCellValue();
                
                
            }//END FOR CELLS   
            System.out.println();
        }//END FOR ROWS
    	 
        System.out.print("FIN DE LECTURA");
    }//FIN METODO LEEEXCELARRAY
    
    
    
    
    
public void WriteExcelArray(String filePath, String fileName, String sheetName, String Estado[]) throws IOException{    	
    	
    	//Create an object of File class to open xlsx file
        File file = new File(filePath+"\\"+fileName);
    	
        //Create an object of FileInputStream class to read excel file
        FileInputStream inputStream = new FileInputStream(file);    	    
        
        //Find the file extension by splitting file name in substring  and getting only extension name
        String fileExtensionName = fileName.substring(fileName.indexOf("."));
    	
        Sheet hoja = null;
        XSSFWorkbook workbookXSSF = null;
        HSSFWorkbook workbookHSSF = null;
        
        //Check condition if the file is xlsx file(ERROR CON FORMATO XLSX??????)
        if(fileExtensionName.equals(".xlsx")){
        	workbookXSSF  = new XSSFWorkbook(inputStream);//If it is xlsx file then create object of XSSFWorkbook class
        	hoja = workbookXSSF.getSheet(sheetName);//Read sheet inside the workbook by its name
        }//Check condition if the file is xls file
        else if(fileExtensionName.equals(".xls")){
        	workbookHSSF = new HSSFWorkbook(inputStream); //If it is xls file then create object of XSSFWorkbook class
        	hoja = workbookHSSF.getSheet(sheetName);//Read sheet inside the workbook by its name
        }
                  
        //Find number of rows in excel file
        int rowCount = hoja.getLastRowNum()-hoja.getFirstRowNum()+1;
                
        //Create a loop over all the rows of excel file to read it
        for (int i = 1; i < rowCount; i++) {

            Row row = hoja.getRow(i);
	
        	//Fill data in row to cell n°3
            Cell cell = row.createCell(3);
            cell.setCellValue(Estado[i]);            	
            	
            //Print Excel data in console
            System.out.print(row.getCell(3).getStringCellValue()+"|| ");
            System.out.println();
        }//END FOR ROWS
    	 
        System.out.print("FIN DE ESCRITURA");
        
        
      //Close input stream
        inputStream.close();

        //Create an object of FileOutputStream class to create write data in excel file
        FileOutputStream outputStream = new FileOutputStream(file);

        //write data in the excel file
        if(fileExtensionName.equals(".xlsx")){workbookXSSF.write(outputStream);}
        if(fileExtensionName.equals(".xls")){workbookHSSF.write(outputStream);}
         
        //close output stream
        outputStream.close();
    }//FIN METODO WRITEEEXCELARRAY
    
}
