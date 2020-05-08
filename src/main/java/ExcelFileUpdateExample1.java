import java.io.*;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import javax.swing.JOptionPane;

/**
 * This program illustrates how to update an existing Microsoft Excel document.
 * Append new rows to an existing sheet.
 * 
 * @author www.codejava.net
 *
 */
public class ExcelFileUpdateExample1 {

	public void CrearExcel(String nombre){
		Workbook libro =  new HSSFWorkbook();
		Sheet hoja = libro.createSheet();
		Row fila = hoja.createRow(0);
		Cell celda = fila.createCell(0);
		celda.setCellValue("No");
		celda = fila.createCell(1);
		celda.setCellValue("Book Title");
		celda = fila.createCell(2);
		celda.setCellValue("Author");
		celda = fila.createCell(3);
		celda.setCellValue("Price");
		String file = nombre+".xls";
		try {
			FileOutputStream out = new FileOutputStream(file);
			libro.write(out);
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}

	}

	public static void UpdateCell(Integer option, String content, Integer id){
		String excelFilePath = "Inventario.xls";
		try {
			FileInputStream inputStream = new FileInputStream(new File(excelFilePath));
			Workbook workbook = WorkbookFactory.create(inputStream);
			Sheet sheet = workbook.getSheetAt(0);
			Integer cell=0;
			if(option == 1)
				/*Esta opcion el usuario indica que quiere modificar el autor */
				cell = 2;
			else
				/*Esta opcion el usuario indica que quiere modificar el precio */
				cell = 3;
				
			Row row = sheet.getRow(id);
			if(row != null){
				Cell cell2Update = sheet.getRow(id).getCell(cell);
				if(cell == 3)
					cell2Update.setCellValue(Integer.parseInt(content));
				else
					cell2Update.setCellValue(content);
				
			}
			else
				JOptionPane.showMessageDialog(null, "El id introducido no existe", "Error", JOptionPane.ERROR_MESSAGE);
			String output="";
			for(Integer i = 1;i<=sheet.getLastRowNum();i++){
				output += sheet.getRow(i).getCell(0) + " " + sheet.getRow(i).getCell(1) + " " + sheet.getRow(i).getCell(2) + " " + sheet.getRow(i).getCell(3) + "\n";
			}

			JOptionPane.showMessageDialog(null, output);

			inputStream.close();
			FileOutputStream outputStream = new FileOutputStream(excelFilePath);
			workbook.write(outputStream);
			workbook.close();
			outputStream.close();
		} catch (IOException | EncryptedDocumentException
				| InvalidFormatException ex) {
			ex.printStackTrace();
		}
	}

	public static void main(String[] args) {
		String excelFilePath = "Inventario.xls";
		File excel = new File(excelFilePath);
		ExcelFileUpdateExample1 variable = new ExcelFileUpdateExample1();
		if (!excel.exists()){
			//excel.createNewFile()
			variable.CrearExcel("Inventario");
			excel = new File("Inventario.xls");
			System.out.println("Se creo el archivo Inventario ya que no existe");
		}

		try{
			String ID ="";
			String autor = "";
			String precio = "";
			String input = JOptionPane.showInputDialog(null, "Seleccione una opcion: \n 1- Actualizar Autor \n 2- Actualizar Precio \n");
			Integer option = Integer.parseInt(input);
			switch(option){
				case 1:
					ID = JOptionPane.showInputDialog(null, "Inserte el ID");
					autor = JOptionPane.showInputDialog(null, "Inserte el autor");
					UpdateCell(1,autor,Integer.parseInt(ID));
					break;
				case 2:
					ID = JOptionPane.showInputDialog(null, "Inserte el ID");
					precio = JOptionPane.showInputDialog(null, "Inserte el precio");
					UpdateCell(1,precio, Integer.parseInt(ID));
					break;
			}
		}
		catch(NumberFormatException e){
			JOptionPane.showMessageDialog(null, "Introdujo un dato invalido", "Error", JOptionPane.ERROR_MESSAGE);
			throw e;
		}
		catch(Exception e){
			throw e;
		}
	}

}
