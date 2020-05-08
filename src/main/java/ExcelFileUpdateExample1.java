import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
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


	public static void main(String[] args) {
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
	
		/* String excelFilePath = "Inventario.xlsx";
		try {
			FileInputStream inputStream = new FileInputStream(new File(excelFilePath));
			Workbook workbook = WorkbookFactory.create(inputStream);
			Object[][] bookData = {
					{"El que se duerme pierde", "Tom Peter", 16},
					{"Sin lugar a duda", "Ana Gutierrez", 26},
					{"El arte de dormir", "Nico", 32},
					{"Buscando a Nemo", "Humble Po", 41},
			};

			int rowCount = sheet.getLastRowNum();

			for (Object[] aBook : bookData) {
				Row row = sheet.createRow(++rowCount);

				int columnCount = 0;
				
				Cell cell = row.createCell(columnCount);
				cell.setCellValue(rowCount);
				
				for (Object field : aBook) {
					cell = row.createCell(++columnCount);
					if (field instanceof String) {
						cell.setCellValue((String) field);
					} else if (field instanceof Integer) {
						cell.setCellValue((Integer) field);
					}
				}

			}
			inputStream.close();
			FileOutputStream outputStream = new FileOutputStream(excelFilePath);
			workbook.write(outputStream);
			workbook.close();
			outputStream.close();


			
		} catch (IOException | EncryptedDocumentException
				| InvalidFormatException ex) {
			ex.printStackTrace();
		} */
	}

	public static void UpdateCell(Integer option, String content, Integer id){
		String excelFilePath = "Inventario.xlsx";
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

}
