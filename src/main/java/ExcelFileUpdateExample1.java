import java.io.*;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

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

		try {
			FileInputStream inputStream = new FileInputStream(excel);
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
		}
	}

}
