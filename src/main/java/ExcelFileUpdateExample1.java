import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
//Adrián: Agrego Iterator para in por cada columna y cada celda mostrando
import java.util.Iterator;
//\

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
			String input = JOptionPane.showInputDialog(null, "Seleccione una opcion: \n 1- Actualizar Autor \n 2- Actualizar Precio \n 3- Llenar celda \n 4- Mostrar contenido del documento");
			Integer option = Integer.parseInt(input);
			switch(option){
				case 3:
					fill();
					break;
				case 4:
					showContent();
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

	public static void fill(){
		String excelFilePath = "Inventario.xlsx";
		try {
			FileInputStream inputStream = new FileInputStream(new File(excelFilePath));
				Workbook workbook = WorkbookFactory.create(inputStream);

				Sheet sheet = workbook.getSheetAt(0);

				Object[][] bookData = {
						{"El que se duerme pierde", "Tom Peter", 16},
						{"Sin lugar a duda", "Ana Gutierrez", 26},
						{"El arte de dormir", "Nico", 32},
						{"Buscando a Nemo", "Humble Po", 41},
				};

				int rowCount = sheet.getLastRowNum();

				for (Object[] aBook : bookData) {
					
					checkSheets(workbook);

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

				showContent();

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

	public static void showContent() {
		String excelFilePath = "Inventario.xlsx";
		try {
			FileInputStream inputStream = new FileInputStream(new File(excelFilePath));
			Workbook workbook = WorkbookFactory.create(inputStream);

			//Código Adrián para mostrar contenido
			for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
				Iterator<Row> iterator = workbook.getSheetAt(i).iterator();

				System.out.println("Hoja: " + workbook.getSheetAt(i).getSheetName());
				while (iterator.hasNext()) {
					Row nextRow = iterator.next();
					Iterator<Cell> cellIterator = nextRow.cellIterator();
					
					while (cellIterator.hasNext()) {
						Cell cell = cellIterator.next();
						
						switch (cell.getCellType()) {
							case Cell.CELL_TYPE_STRING:
								System.out.print(cell.getStringCellValue());
								break;
							case Cell.CELL_TYPE_BOOLEAN:
								System.out.print(cell.getBooleanCellValue());
								break;
							case Cell.CELL_TYPE_NUMERIC:
								System.out.print(cell.getNumericCellValue());
								break;
						}
						if (cellIterator.hasNext())
							System.out.print(" - ");
					}
					System.out.println();
				}
			}
			inputStream.close();

			FileOutputStream outputStream = new FileOutputStream(excelFilePath);
			workbook.write(outputStream);
			workbook.close();
			outputStream.close();
		//\ Finaliza código
		}catch (IOException | EncryptedDocumentException
		| InvalidFormatException ex) {
		ex.printStackTrace();
	}
	}

	public static void checkSheets(Workbook workbook) {
		//Código Adrián para crear hojas nuevas al llegar a 30 celdas
		Sheet sheet = workbook.getSheetAt(0);
		int rowCount = sheet.getLastRowNum();

		if ( (rowCount) >= 30) {

			sheet = workbook.getSheetAt(workbook.getNumberOfSheets()-1);
			rowCount = sheet.getLastRowNum();
			if (rowCount >= 30){
				sheet = workbook.createSheet("Java Books " + (workbook.getNumberOfSheets()+1));

				Object[][] startString = {
					{"No", "Book Title", "Author", "Price"},
				};
				for (Object[] cells : startString) {
					Row r = sheet.getRow(0);
					r = sheet.createRow(0);

					int columnCount = 0;

					Cell c = r.createCell(columnCount);
					for (Object field : cells) {
						c = r.createCell(columnCount++);
						c.setCellValue((String) field);
					}
				}
			}

			rowCount = sheet.getLastRowNum();
		}
		//\ Finaliza código
	}
}
